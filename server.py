from dotenv import load_dotenv
from pathlib import Path
load_dotenv(Path(__file__).parent / ".env")

import os, uuid, secrets, logging, csv, io, base64
from datetime import datetime, timezone, timedelta
from typing import Optional
import bcrypt, jwt, requests
import qrcode
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from fastapi import FastAPI, APIRouter, HTTPException, Depends, Request, UploadFile, File, Form, Header, Query
from fastapi.responses import StreamingResponse, Response
from fastapi.middleware.cors import CORSMiddleware
from motor.motor_asyncio import AsyncIOMotorClient
from pydantic import BaseModel, Field, EmailStr

ROOT_DIR = Path(__file__).parent
client = AsyncIOMotorClient(os.environ["MONGO_URL"])
db = client[os.environ["DB_NAME"]]
app = FastAPI(title="Traxo Inventory API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://traxo-frontend-lfq8.vercel.app",
        "http://localhost:3000"
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

api = APIRouter(prefix="/api")
JWT_SECRET = os.environ.get("JWT_SECRET", "traxo-local-development-secret")
logger = logging.getLogger("traxo")

# ---------- Emergent Object Storage ----------
STORAGE_BASE = (os.environ.get("INTEGRATION_PROXY_URL") or "").strip() or "https://integrations.emergentagent.com"
STORAGE_URL = STORAGE_BASE.rstrip("/") + "/objstore/api/v1/storage"
EMERGENT_KEY = os.environ.get("EMERGENT_LLM_KEY")
APP_NAME = "traxo-inventory"
_storage_key: Optional[str] = None

def init_storage(force: bool = False):
    global _storage_key
    if _storage_key and not force: return _storage_key
    if not EMERGENT_KEY:
        logger.warning("EMERGENT_LLM_KEY missing; object storage disabled")
        return None
    try:
        resp = requests.post(f"{STORAGE_URL}/init", json={"emergent_key": EMERGENT_KEY}, timeout=30)
        resp.raise_for_status()
        _storage_key = resp.json()["storage_key"]
        return _storage_key
    except Exception as exc:
        logger.error(f"Storage init failed: {exc}")
        return None

def put_object(path: str, data: bytes, content_type: str) -> dict:
    key = init_storage()
    if not key: raise HTTPException(503, "File storage unavailable")
    resp = requests.put(f"{STORAGE_URL}/objects/{path}", headers={"X-Storage-Key": key, "Content-Type": content_type}, data=data, timeout=120)
    if resp.status_code == 404:
        key = init_storage(force=True)
        resp = requests.put(f"{STORAGE_URL}/objects/{path}", headers={"X-Storage-Key": key, "Content-Type": content_type}, data=data, timeout=120)
    resp.raise_for_status()
    return resp.json()

def get_object(path: str):
    key = init_storage()
    if not key: raise HTTPException(503, "File storage unavailable")
    resp = requests.get(f"{STORAGE_URL}/objects/{path}", headers={"X-Storage-Key": key}, timeout=60)
    if resp.status_code == 404:
        key = init_storage(force=True)
        resp = requests.get(f"{STORAGE_URL}/objects/{path}", headers={"X-Storage-Key": key}, timeout=60)
    resp.raise_for_status()
    return resp.content, resp.headers.get("Content-Type", "application/octet-stream")

# ---------- Helpers ----------
def now(): return datetime.now(timezone.utc).isoformat()
def safe(doc):
    if not doc: return None
    doc.pop("_id", None)
    return doc
def hash_password(password): return bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
def verify_password(password, hashed): return bcrypt.checkpw(password.encode(), hashed.encode())
def token(user):
    return jwt.encode({"sub": user["id"], "email": user["email"], "role": user["role"], "exp": datetime.now(timezone.utc)+timedelta(hours=8)}, JWT_SECRET, algorithm="HS256")

async def current_user(request: Request):
    raw = request.headers.get("Authorization", "")
    value = request.cookies.get("access_token") or (raw[7:] if raw.startswith("Bearer ") else None)
    if not value: raise HTTPException(401, "Not authenticated")
    try: payload = jwt.decode(value, JWT_SECRET, algorithms=["HS256"])
    except jwt.PyJWTError: raise HTTPException(401, "Invalid or expired session")
    user = await db.users.find_one({"id": payload["sub"]}, {"_id": 0, "password_hash": 0})
    if not user: raise HTTPException(401, "User not found")
    return user

async def user_from_token(token_value: Optional[str]):
    if not token_value: raise HTTPException(401, "Not authenticated")
    try: payload = jwt.decode(token_value, JWT_SECRET, algorithms=["HS256"])
    except jwt.PyJWTError: raise HTTPException(401, "Invalid or expired session")
    user = await db.users.find_one({"id": payload["sub"]}, {"_id": 0, "password_hash": 0})
    if not user: raise HTTPException(401, "User not found")
    return user

# ---------- Models ----------
class Login(BaseModel): email: EmailStr; password: str
class UserCreate(BaseModel): email: EmailStr; password: str; name: str; role: str = "engineer"
class Component(BaseModel):
    name: str; category: str; part_number: str = ""; unit: str = "Nos"; minimum_stock: int = 10; location: str = "Main Store"; unit_price: float = 0; manufacturer: str = ""; hsn: str = ""; gst: float = 0
class Movement(BaseModel):
    component_id: str; quantity: int = Field(gt=0); party: str = ""; reference: str = ""; project: str = ""; remarks: str = ""; warehouse_id: str = "wh-main"; destination_warehouse_id: str = ""
class Workflow(BaseModel):
    title: str; component_id: str = ""; quantity: int = 0; party: str = ""; warehouse_id: str = "wh-main"; destination_warehouse_id: str = ""; notes: str = ""; linked_po_id: str = ""
class Approval(BaseModel): status: str
class WarehouseIn(BaseModel): name: str; code: str; address: str = ""
class VendorIn(BaseModel): name: str; contact: str = ""; phone: str = ""; email: str = ""; gst: str = ""; address: str = ""; notes: str = ""
class EmployeeIn(BaseModel): name: str; department: str = ""; designation: str = ""; phone: str = ""; email: str = ""; employee_code: str = ""

# ---------- Routes ----------
@api.get("/")
async def root(): return {"message": "Traxo Inventory API", "status": "online"}

@api.post("/auth/login")
async def login(data: Login):
    user = await db.users.find_one({"email": data.email.lower()})
    if not user or not verify_password(data.password, user["password_hash"]): raise HTTPException(401, "Incorrect email or password")
    public = {k:v for k,v in user.items() if k not in ["_id", "password_hash"]}
    return {"user": public, "token": token(user)}

@api.post("/auth/logout")
async def logout(): return {"message": "Signed out"}

@api.get("/auth/me")
async def me(user=Depends(current_user)): return user

@api.post("/auth/users")
async def create_user(data: UserCreate, user=Depends(current_user)):
    if user["role"] != "admin": raise HTTPException(403, "Admin access required")
    doc = {"id": str(uuid.uuid4()), "email": data.email.lower(), "name": data.name, "role": data.role, "password_hash": hash_password(data.password), "created_at": now()}
    await db.users.insert_one(doc); return safe({k:v for k,v in doc.items() if k != "password_hash"})

@api.get("/users")
async def users(user=Depends(current_user)):
    return await db.users.find({}, {"_id": 0, "password_hash": 0}).to_list(200)

@api.get("/components")
async def components(user=Depends(current_user)):
    return await db.components.find({}, {"_id": 0}).sort("name", 1).to_list(500)

@api.post("/components")
async def create_component(data: Component, user=Depends(current_user)):
    doc = {"id": str(uuid.uuid4()), "code": f"CMP-{secrets.randbelow(9000)+1000}", **data.model_dump(), "stock": 0, "warehouse_balances": {}, "created_at": now(), "attachments": []}
    await db.components.insert_one(doc)
    await db.audit.insert_one({"id": str(uuid.uuid4()), "action": "Component created", "user": user["name"], "created_at": now(), "details": data.name})
    return safe(doc)

@api.get("/components/{component_id}/balances")
async def component_balances(component_id: str, user=Depends(current_user)):
    component = await db.components.find_one({"id": component_id}, {"_id": 0})
    if not component: raise HTTPException(404, "Component not found")
    warehouses_docs = await db.warehouses.find({}, {"_id": 0}).sort("name", 1).to_list(100)
    balances = component.get("warehouse_balances") or {}
    rows = [{"warehouse_id": w["id"], "warehouse_name": w["name"], "warehouse_code": w.get("code",""), "balance": balances.get(w["id"], 0)} for w in warehouses_docs]
    return {"component_id": component_id, "component_name": component["name"], "total_stock": component.get("stock", 0), "balances": rows}

@api.post("/components/import")
async def import_components(file: UploadFile = File(...), user=Depends(current_user)):
    data = (await file.read()).decode("utf-8", errors="ignore")
    reader = csv.DictReader(io.StringIO(data))
    created, errors = [], []
    for idx, row in enumerate(reader, start=2):
        name = (row.get("name") or row.get("Name") or "").strip()
        if not name:
            errors.append({"row": idx, "reason": "missing name"}); continue
        try:
            opening = int(float(row.get("opening_stock") or row.get("Opening Stock") or 0))
        except ValueError:
            opening = 0
        try:
            price = float(row.get("unit_price") or row.get("Unit Price") or 0)
        except ValueError:
            price = 0
        try:
            minimum = int(float(row.get("minimum_stock") or row.get("Minimum Stock") or 10))
        except ValueError:
            minimum = 10
        default_wh = (row.get("warehouse_id") or row.get("Warehouse") or "wh-main").strip() or "wh-main"
        doc = {"id": str(uuid.uuid4()), "code": (row.get("code") or row.get("Code") or f"CMP-{secrets.randbelow(9000)+1000}").strip(), "name": name, "category": (row.get("category") or row.get("Category") or "General").strip(), "part_number": (row.get("part_number") or row.get("Part Number") or "").strip(), "unit": (row.get("unit") or row.get("Unit") or "Nos").strip(), "minimum_stock": minimum, "location": (row.get("location") or row.get("Location") or "Main Store").strip(), "unit_price": price, "manufacturer": (row.get("manufacturer") or row.get("Manufacturer") or "").strip(), "hsn": (row.get("hsn") or row.get("HSN") or "").strip(), "gst": float(row.get("gst") or row.get("GST") or 0), "stock": opening, "warehouse_balances": {default_wh: opening} if opening else {}, "created_at": now(), "attachments": []}
        await db.components.insert_one(doc)
        if opening > 0:
            await db.movements.insert_one({"id": str(uuid.uuid4()), "type": "purchase", "component_id": doc["id"], "component_name": doc["name"], "quantity": opening, "party": "Opening stock (import)", "reference": f"IMPORT-{doc['code']}", "project": "", "remarks": "Bulk CSV import", "balance": opening, "warehouse_id": default_wh, "destination_warehouse_id": "", "created_by": user["name"], "created_at": now()})
        created.append(safe(doc))
    await db.audit.insert_one({"id": str(uuid.uuid4()), "action": "Components imported", "user": user["name"], "created_at": now(), "details": f"{len(created)} rows imported, {len(errors)} skipped"})
    return {"imported": len(created), "skipped": len(errors), "errors": errors, "components": created}

@api.get("/vendors")
async def vendors(user=Depends(current_user)): return await db.vendors.find({}, {"_id": 0}).sort("name", 1).to_list(500)

@api.post("/vendors")
async def create_vendor(data: VendorIn, user=Depends(current_user)):
    doc = {"id": str(uuid.uuid4()), **data.model_dump(), "created_at": now()}
    await db.vendors.insert_one(doc)
    await db.audit.insert_one({"id": str(uuid.uuid4()), "action": "Vendor created", "user": user["name"], "created_at": now(), "details": data.name})
    return safe(doc)

@api.put("/vendors/{vendor_id}")
async def update_vendor(vendor_id: str, data: VendorIn, user=Depends(current_user)):
    result = await db.vendors.update_one({"id": vendor_id}, {"$set": {**data.model_dump(), "updated_at": now()}})
    if not result.matched_count: raise HTTPException(404, "Vendor not found")
    await db.audit.insert_one({"id": str(uuid.uuid4()), "action": "Vendor updated", "user": user["name"], "created_at": now(), "details": data.name})
    return await db.vendors.find_one({"id": vendor_id}, {"_id": 0})

@api.delete("/vendors/{vendor_id}")
async def delete_vendor(vendor_id: str, user=Depends(current_user)):
    if user["role"] != "admin": raise HTTPException(403, "Admin access required")
    record = await db.vendors.find_one({"id": vendor_id}, {"_id": 0})
    if not record: raise HTTPException(404, "Vendor not found")
    await db.vendors.delete_one({"id": vendor_id})
    await db.audit.insert_one({"id": str(uuid.uuid4()), "action": "Vendor deleted", "user": user["name"], "created_at": now(), "details": record.get("name", vendor_id)})
    return {"deleted": True}

@api.get("/vendors/{vendor_id}/history")
async def vendor_history(vendor_id: str, user=Depends(current_user)):
    vendor = await db.vendors.find_one({"id": vendor_id}, {"_id": 0})
    if not vendor: raise HTTPException(404, "Vendor not found")
    name = vendor.get("name", "")
    purchases = await db.movements.find({"type": "purchase", "party": name}, {"_id": 0}).sort("created_at", -1).to_list(500)
    workflows_list = await db.workflows.find({"$or": [{"party": name}, {"title": {"$regex": name, "$options": "i"}}]}, {"_id": 0}).sort("created_at", -1).to_list(300)
    component_ids = list({p.get("component_id") for p in purchases if p.get("component_id")})
    price_map = {}
    if component_ids:
        async for c in db.components.find({"id": {"$in": component_ids}}, {"_id": 0, "id": 1, "unit_price": 1}):
            price_map[c["id"]] = c.get("unit_price", 0)
    total_value = sum(p.get("quantity", 0) * price_map.get(p.get("component_id"), 0) for p in purchases)
    return {"vendor": vendor, "purchases": purchases, "workflows": workflows_list, "totals": {"orders": len(purchases), "units": sum(m.get("quantity", 0) for m in purchases), "value": round(total_value, 2)}}

@api.post("/vendors/import")
async def import_vendors(file: UploadFile = File(...), user=Depends(current_user)):
    data = (await file.read()).decode("utf-8", errors="ignore")
    reader = csv.DictReader(io.StringIO(data))
    created = []
    for row in reader:
        name = (row.get("name") or row.get("Name") or "").strip()
        if not name: continue
        doc = {"id": str(uuid.uuid4()), "name": name, "contact": row.get("contact") or row.get("Contact") or "", "phone": row.get("phone") or row.get("Phone") or "", "email": row.get("email") or row.get("Email") or "", "gst": row.get("gst") or row.get("GST") or "", "address": row.get("address") or row.get("Address") or "", "notes": row.get("notes") or row.get("Notes") or "", "created_at": now()}
        await db.vendors.insert_one(doc); created.append(safe(doc))
    await db.audit.insert_one({"id": str(uuid.uuid4()), "action": "Vendors imported", "user": user["name"], "created_at": now(), "details": f"{len(created)} rows"})
    return {"imported": len(created), "vendors": created}

@api.get("/employees")
async def employees(user=Depends(current_user)): return await db.employees.find({}, {"_id": 0}).sort("name", 1).to_list(500)

@api.post("/employees")
async def create_employee(data: EmployeeIn, user=Depends(current_user)):
    doc = {"id": str(uuid.uuid4()), **data.model_dump(), "created_at": now()}
    await db.employees.insert_one(doc)
    await db.audit.insert_one({"id": str(uuid.uuid4()), "action": "Employee created", "user": user["name"], "created_at": now(), "details": data.name})
    return safe(doc)

@api.put("/employees/{employee_id}")
async def update_employee(employee_id: str, data: EmployeeIn, user=Depends(current_user)):
    result = await db.employees.update_one({"id": employee_id}, {"$set": {**data.model_dump(), "updated_at": now()}})
    if not result.matched_count: raise HTTPException(404, "Employee not found")
    await db.audit.insert_one({"id": str(uuid.uuid4()), "action": "Employee updated", "user": user["name"], "created_at": now(), "details": data.name})
    return await db.employees.find_one({"id": employee_id}, {"_id": 0})

@api.delete("/employees/{employee_id}")
async def delete_employee(employee_id: str, user=Depends(current_user)):
    if user["role"] != "admin": raise HTTPException(403, "Admin access required")
    record = await db.employees.find_one({"id": employee_id}, {"_id": 0})
    if not record: raise HTTPException(404, "Employee not found")
    await db.employees.delete_one({"id": employee_id})
    await db.audit.insert_one({"id": str(uuid.uuid4()), "action": "Employee deleted", "user": user["name"], "created_at": now(), "details": record.get("name", employee_id)})
    return {"deleted": True}

@api.get("/employees/{employee_id}/history")
async def employee_history(employee_id: str, user=Depends(current_user)):
    employee = await db.employees.find_one({"id": employee_id}, {"_id": 0})
    if not employee: raise HTTPException(404, "Employee not found")
    name = employee.get("name", "")
    issues = await db.movements.find({"type": "issue", "party": name}, {"_id": 0}).sort("created_at", -1).to_list(500)
    returns = await db.movements.find({"type": "return", "party": name}, {"_id": 0}).sort("created_at", -1).to_list(500)
    return {"employee": employee, "issues": issues, "returns": returns, "totals": {"issued": sum(m.get("quantity", 0) for m in issues), "returned": sum(m.get("quantity", 0) for m in returns)}}

@api.post("/employees/import")
async def import_employees(file: UploadFile = File(...), user=Depends(current_user)):
    data = (await file.read()).decode("utf-8", errors="ignore")
    reader = csv.DictReader(io.StringIO(data))
    created = []
    for row in reader:
        name = (row.get("name") or row.get("Name") or "").strip()
        if not name: continue
        doc = {"id": str(uuid.uuid4()), "name": name, "department": row.get("department") or row.get("Department") or "", "designation": row.get("designation") or row.get("Designation") or "", "phone": row.get("phone") or row.get("Phone") or "", "email": row.get("email") or row.get("Email") or "", "employee_code": row.get("employee_code") or row.get("Code") or "", "created_at": now()}
        await db.employees.insert_one(doc); created.append(safe(doc))
    await db.audit.insert_one({"id": str(uuid.uuid4()), "action": "Employees imported", "user": user["name"], "created_at": now(), "details": f"{len(created)} rows"})
    return {"imported": len(created), "employees": created}

@api.get("/warehouses")
async def warehouses(user=Depends(current_user)): return await db.warehouses.find({}, {"_id": 0}).to_list(100)

@api.post("/warehouses")
async def create_warehouse(data: WarehouseIn, user=Depends(current_user)):
    doc = {"id": str(uuid.uuid4()), **data.model_dump(), "active": True, "created_at": now()}
    await db.warehouses.insert_one(doc); return safe(doc)

async def movement(kind, data, user):
    component = await db.components.find_one({"id": data.component_id}, {"_id": 0})
    if not component: raise HTTPException(404, "Component not found")
    balances = dict(component.get("warehouse_balances") or {})
    wh_id = data.warehouse_id or "wh-main"
    current_wh = balances.get(wh_id, 0)
    if kind == "issue":
        if component.get("stock", 0) < data.quantity: raise HTTPException(400, f"Insufficient stock. Available: {component.get('stock', 0)}")
        if current_wh < data.quantity:
            wh = await db.warehouses.find_one({"id": wh_id}, {"_id": 0, "name": 1, "code": 1})
            wh_label = f"{wh['name']} ({wh.get('code','')})" if wh else wh_id
            raise HTTPException(400, f"Insufficient stock at {wh_label}. Available: {current_wh}")
    delta = data.quantity if kind in ["purchase", "return"] else -data.quantity
    balances[wh_id] = current_wh + delta
    total_after = sum(balances.values())
    await db.components.update_one({"id": data.component_id}, {"$set": {"stock": total_after, "warehouse_balances": balances}})
    doc = {"id": str(uuid.uuid4()), "type": kind, "component_id": data.component_id, "component_name": component["name"], "quantity": data.quantity, "party": data.party, "reference": data.reference, "project": data.project, "remarks": data.remarks, "warehouse_id": wh_id, "destination_warehouse_id": "", "balance": total_after, "created_by": user["name"], "created_at": now()}
    await db.movements.insert_one(doc)
    await db.audit.insert_one({"id": str(uuid.uuid4()), "action": kind.title(), "user": user["name"], "created_at": now(), "details": f"{component['name']} · {data.quantity} units · {wh_id}"})
    return safe(doc)

async def _apply_transfer(component_id: str, quantity: int, from_wh: str, to_wh: str, reference: str, user):
    if not from_wh or not to_wh or from_wh == to_wh: raise HTTPException(400, "Choose distinct source and destination warehouses")
    component = await db.components.find_one({"id": component_id}, {"_id": 0})
    if not component: raise HTTPException(404, "Component not found")
    balances = dict(component.get("warehouse_balances") or {})
    src = balances.get(from_wh, 0)
    if src < quantity:
        wh = await db.warehouses.find_one({"id": from_wh}, {"_id": 0, "name": 1, "code": 1})
        wh_label = f"{wh['name']} ({wh.get('code','')})" if wh else from_wh
        raise HTTPException(400, f"Insufficient stock at {wh_label}. Available: {src}")
    balances[from_wh] = src - quantity
    balances[to_wh] = balances.get(to_wh, 0) + quantity
    total_after = sum(balances.values())
    await db.components.update_one({"id": component_id}, {"$set": {"stock": total_after, "warehouse_balances": balances}})
    doc = {"id": str(uuid.uuid4()), "type": "transfer", "component_id": component_id, "component_name": component["name"], "quantity": quantity, "party": "", "reference": reference, "project": "", "remarks": f"{from_wh} → {to_wh}", "warehouse_id": from_wh, "destination_warehouse_id": to_wh, "balance": total_after, "created_by": user["name"], "created_at": now()}
    await db.movements.insert_one(doc)
    await db.audit.insert_one({"id": str(uuid.uuid4()), "action": "Transfer", "user": user["name"], "created_at": now(), "details": f"{component['name']} · {quantity} units · {from_wh} → {to_wh}"})
    return safe(doc)

@api.post("/movements/{kind}")
async def create_movement(kind: str, data: Movement, user=Depends(current_user)):
    if kind not in ["purchase", "issue", "return", "transfer"]: raise HTTPException(400, "Unsupported movement")
    if kind == "transfer":
        return await _apply_transfer(data.component_id, data.quantity, data.warehouse_id or "wh-main", data.destination_warehouse_id, data.reference, user)
    return await movement(kind, data, user)

@api.post("/workflows/{kind}")
async def create_workflow(kind: str, data: Workflow, user=Depends(current_user)):
    if kind not in ["po", "grn", "mrs", "return", "opening", "transfer"]: raise HTTPException(400, "Unsupported workflow")
    if kind in ["return", "opening"] and data.component_id and data.quantity:
        await movement("return" if kind == "return" else "purchase", Movement(component_id=data.component_id, quantity=data.quantity, party=data.party, reference=data.title, remarks=data.notes), user)
    doc = {"id": str(uuid.uuid4()), "kind": kind, **data.model_dump(), "status": "pending", "created_by": user["name"], "created_at": now()}
    await db.workflows.insert_one(doc)
    await db.audit.insert_one({"id": str(uuid.uuid4()), "action": f"{kind.upper()} created", "user": user["name"], "created_at": now(), "details": data.title})
    return safe(doc)

@api.get("/workflows")
async def workflows(user=Depends(current_user)):
    return await db.workflows.find({}, {"_id": 0}).sort("created_at", -1).to_list(300)

@api.patch("/workflows/{workflow_id}/approval")
async def approve_workflow(workflow_id: str, data: Approval, user=Depends(current_user)):
    if data.status not in ["approved", "rejected", "pending"]: raise HTTPException(400, "Invalid approval state")
    record = await db.workflows.find_one({"id": workflow_id}, {"_id": 0})
    if not record: raise HTTPException(404, "Workflow not found")
    executed = None
    # Auto-execute stock effect on approval
    if data.status == "approved" and record.get("status") != "approved":
        kind = record.get("kind")
        wh_id = record.get("warehouse_id") or "wh-main"
        dest = record.get("destination_warehouse_id") or ""
        if kind == "grn":
            # Prefer linked PO's data if present, else use record's own component/quantity
            source = record
            if record.get("linked_po_id"):
                po = await db.workflows.find_one({"id": record["linked_po_id"]}, {"_id": 0})
                if po:
                    source = {**record, "component_id": record.get("component_id") or po.get("component_id",""), "quantity": record.get("quantity") or po.get("quantity",0), "party": record.get("party") or po.get("party",""), "warehouse_id": wh_id}
            if source.get("component_id") and source.get("quantity"):
                executed = await movement("purchase", Movement(component_id=source["component_id"], quantity=int(source["quantity"]), party=source.get("party",""), reference=record.get("title",""), remarks=f"GRN of {record.get('linked_po_id') or 'direct receipt'}", warehouse_id=wh_id), user)
                # Close the loop: mark linked PO as received
                if record.get("linked_po_id"):
                    await db.workflows.update_one({"id": record["linked_po_id"], "status": {"$ne": "received"}}, {"$set": {"status": "received", "received_via_grn_id": workflow_id, "received_at": now()}})
                    await db.audit.insert_one({"id": str(uuid.uuid4()), "action": "PO received", "user": user["name"], "created_at": now(), "details": f"via GRN {record.get('title','')}"})
        elif kind == "transfer" and record.get("component_id") and record.get("quantity") and dest:
            executed = await _apply_transfer(record["component_id"], int(record["quantity"]), wh_id, dest, record.get("title",""), user)
        elif kind == "mrs" and record.get("component_id") and record.get("quantity"):
            executed = await movement("issue", Movement(component_id=record["component_id"], quantity=int(record["quantity"]), party=record.get("party",""), reference=record.get("title",""), warehouse_id=wh_id), user)
    update = {"status": data.status, "approved_by": user["name"], "approved_at": now()}
    if executed: update["executed_movement_id"] = executed.get("id")
    await db.workflows.update_one({"id": workflow_id}, {"$set": update})
    await db.audit.insert_one({"id": str(uuid.uuid4()), "action": f"{record['kind'].upper()} {data.status}", "user": user["name"], "created_at": now(), "details": record["title"]})
    return {**record, **update}

@api.post("/transfers")
async def create_transfer(data: Workflow, user=Depends(current_user)):
    if not data.destination_warehouse_id or data.destination_warehouse_id == data.warehouse_id: raise HTTPException(400, "Choose a different destination warehouse")
    return await create_workflow("transfer", data, user)

# ---------- Reports: CSV ----------
@api.get("/reports/export")
async def export_report(kind: str = "stock"):
    output = io.StringIO(); writer = csv.writer(output)
    if kind == "stock":
        writer.writerow(["Component", "Code", "Category", "Available Stock", "Minimum Stock", "Unit Price", "Stock Value"])
        for c in await db.components.find({}, {"_id": 0}).to_list(500): writer.writerow([c["name"], c["code"], c["category"], c.get("stock", 0), c.get("minimum_stock", 0), c.get("unit_price", 0), c.get("stock", 0)*c.get("unit_price", 0)])
    else:
        writer.writerow(["Date", "Type", "Component", "Quantity", "Party", "Reference", "Balance"])
        for m in await db.movements.find({}, {"_id": 0}).sort("created_at", -1).to_list(1000): writer.writerow([m.get("created_at", "")[:10], m.get("type"), m.get("component_name"), m.get("quantity"), m.get("party"), m.get("reference"), m.get("balance")])
    return StreamingResponse(iter([output.getvalue()]), media_type="text/csv", headers={"Content-Disposition": f"attachment; filename=traxo-{kind}-report.csv"})

# ---------- Reports: PDF ----------
def _pdf_styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TraxoTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=22, textColor=colors.HexColor("#0f172a"), spaceAfter=6))
    styles.add(ParagraphStyle(name="TraxoSub", parent=styles["Normal"], fontName="Helvetica", fontSize=10, textColor=colors.HexColor("#475569"), spaceAfter=14))
    styles.add(ParagraphStyle(name="TraxoSection", parent=styles["Heading3"], fontName="Helvetica-Bold", fontSize=12, textColor=colors.HexColor("#1e40af"), spaceBefore=8, spaceAfter=6))
    return styles

def _pdf_table(headers, rows, col_widths=None):
    data = [headers] + rows
    tbl = Table(data, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0f172a")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("BOTTOMPADDING", (0,0), (-1,0), 8),
        ("TOPPADDING", (0,0), (-1,0), 8),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f1f5f9")]),
        ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#cbd5e1")),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ]))
    return tbl

async def _build_pdf(kind: str) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=1.2*cm, rightMargin=1.2*cm, topMargin=1.2*cm, bottomMargin=1.2*cm)
    styles = _pdf_styles(); story = []
    story.append(Paragraph("Traxo India Automation", styles["TraxoTitle"]))
    title_map = {"stock":"Stock Position Report", "purchases":"Purchase Register", "issues":"Issue Register", "monthly":"Monthly Stock Movement", "audit":"Audit Trail"}
    story.append(Paragraph(f"{title_map.get(kind, 'Report')} · Generated {datetime.now(timezone.utc).strftime('%d %b %Y %H:%M UTC')}", styles["TraxoSub"]))

    if kind == "stock":
        comps = await db.components.find({}, {"_id": 0}).sort("name", 1).to_list(500)
        headers = ["Code", "Component", "Category", "Location", "Available", "Min", "Unit ₹", "Value ₹"]
        rows = [[c["code"], c["name"][:34], c["category"], c.get("location",""), str(c.get("stock",0)), str(c.get("minimum_stock",0)), f"{c.get('unit_price',0):,.2f}", f"{c.get('stock',0)*c.get('unit_price',0):,.2f}"] for c in comps]
        story.append(_pdf_table(headers, rows, col_widths=[2.4*cm, 5.6*cm, 3.2*cm, 3*cm, 2*cm, 1.5*cm, 2.4*cm, 2.8*cm]))
        total = sum(c.get("stock",0)*c.get("unit_price",0) for c in comps)
        story.append(Spacer(1, 12))
        story.append(Paragraph(f"<b>Total inventory value:</b> ₹ {total:,.2f}", styles["TraxoSection"]))
    elif kind in ("purchases", "issues"):
        target = "purchase" if kind == "purchases" else "issue"
        moves = await db.movements.find({"type": target}, {"_id": 0}).sort("created_at", -1).to_list(1000)
        headers = ["Date", "Reference", "Component", "Party", "Qty", "Balance"]
        rows = [[m.get("created_at","")[:10], m.get("reference","—"), m.get("component_name",""), m.get("party","—"), str(m.get("quantity",0)), str(m.get("balance",0))] for m in moves]
        story.append(_pdf_table(headers, rows, col_widths=[2.4*cm, 3.6*cm, 6.4*cm, 4.8*cm, 2*cm, 2.4*cm]))
    elif kind == "monthly":
        moves = await db.movements.find({}, {"_id": 0}).sort("created_at", 1).to_list(2000)
        buckets = {}
        for m in moves:
            key = (m.get("created_at","") or "")[:7] or "current"
            buckets.setdefault(key, {"purchase":0, "issue":0, "return":0})
            buckets[key][m.get("type","purchase")] = buckets[key].get(m.get("type","purchase"),0) + m.get("quantity",0)
        headers = ["Month", "Purchases", "Issues", "Returns", "Net"]
        rows = [[k, str(v["purchase"]), str(v["issue"]), str(v.get("return",0)), str(v["purchase"]+v.get("return",0)-v["issue"])] for k,v in sorted(buckets.items())]
        story.append(_pdf_table(headers, rows or [["—","0","0","0","0"]], col_widths=[3*cm, 3.5*cm, 3.5*cm, 3.5*cm, 3.5*cm]))
    elif kind == "audit":
        events = await db.audit.find({}, {"_id": 0}).sort("created_at", -1).to_list(500)
        headers = ["Timestamp", "Action", "User", "Details"]
        rows = [[(e.get("created_at","")[:19]).replace("T"," "), e.get("action",""), e.get("user",""), (e.get("details","") or "")[:80]] for e in events]
        story.append(_pdf_table(headers, rows, col_widths=[4*cm, 4*cm, 4*cm, 10*cm]))
    else:
        story.append(Paragraph("Unknown report type.", styles["Normal"]))
    story.append(Spacer(1, 18))
    story.append(Paragraph("<font size=8 color='#94a3b8'>Traxo India Automation · Inventory Management System · Confidential</font>", styles["Normal"]))
    doc.build(story)
    return buf.getvalue()

@api.get("/reports/pdf")
async def export_pdf(kind: str = "stock", auth: Optional[str] = Query(None), authorization: Optional[str] = Header(None)):
    token_value = auth or (authorization[7:] if authorization and authorization.startswith("Bearer ") else None)
    await user_from_token(token_value)
    if kind not in ["stock","purchases","issues","monthly","audit"]: raise HTTPException(400, "Unsupported report")
    data = await _build_pdf(kind)
    return Response(content=data, media_type="application/pdf", headers={"Content-Disposition": f'inline; filename="traxo-{kind}-report.pdf"'})

# ---------- Stock Ledger ----------
def _delta(m):
    return m.get("quantity", 0) if m.get("type") in ["purchase", "return"] else -m.get("quantity", 0)

async def _compute_ledger(component_id: str, date_from: Optional[str], date_to: Optional[str], warehouse_id: Optional[str] = None):
    component = await db.components.find_one({"id": component_id}, {"_id": 0})
    if not component: raise HTTPException(404, "Component not found")
    query = {"component_id": component_id}
    if warehouse_id: query["$or"] = [{"warehouse_id": warehouse_id}, {"destination_warehouse_id": warehouse_id}]
    all_moves = await db.movements.find(query, {"_id": 0}).sort("created_at", 1).to_list(2000)
    if warehouse_id:
        balances = component.get("warehouse_balances") or {}
        current_stock = balances.get(warehouse_id, 0)
    else:
        current_stock = component.get("stock", 0)
    def _wh_delta(m):
        # In warehouse-scoped view, transfers move between two warehouses; count as inward for destination, outward for source.
        if not warehouse_id:
            return m.get("quantity", 0) if m.get("type") in ["purchase", "return"] else -m.get("quantity", 0)
        if m.get("type") == "transfer":
            if m.get("destination_warehouse_id") == warehouse_id: return m.get("quantity", 0)
            if m.get("warehouse_id") == warehouse_id: return -m.get("quantity", 0)
            return 0
        return m.get("quantity", 0) if m.get("type") in ["purchase", "return"] else -m.get("quantity", 0)
    from_moves = [m for m in all_moves if (not date_from or (m.get("created_at","") >= date_from))]
    opening = current_stock - sum(_wh_delta(m) for m in from_moves)
    end_cutoff = (date_to + "T23:59:59") if date_to else None
    in_range = [m for m in from_moves if (not end_cutoff or (m.get("created_at","") <= end_cutoff))]
    rows = []
    balance = opening
    for m in in_range:
        d = _wh_delta(m)
        balance += d
        is_inward = d > 0
        rows.append({"date": (m.get("created_at","") or "")[:10], "type": m.get("type"), "reference": m.get("reference",""), "party": m.get("party","") or (m.get("remarks","") if m.get("type")=="transfer" else ""), "inward": abs(d) if is_inward else 0, "outward": abs(d) if d < 0 else 0, "balance": balance, "warehouse_id": m.get("warehouse_id",""), "destination_warehouse_id": m.get("destination_warehouse_id",""), "created_at": m.get("created_at","")})
    return {"component": component, "warehouse_id": warehouse_id, "opening_stock": opening, "closing_stock": balance, "rows": rows, "totals": {"inward": sum(r["inward"] for r in rows), "outward": sum(r["outward"] for r in rows), "movements": len(rows)}}

@api.get("/reports/ledger")
async def stock_ledger(component_id: str, date_from: Optional[str] = None, date_to: Optional[str] = None, warehouse_id: Optional[str] = None, user=Depends(current_user)):
    return await _compute_ledger(component_id, date_from, date_to, warehouse_id)

# ---------- Excel (XLSX) exports ----------
def _xlsx_style(ws, header_row=1):
    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for cell in ws[header_row]:
        cell.fill = header_fill; cell.font = header_font; cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[header_row].height = 22
    for col in ws.columns:
        length = max((len(str(c.value)) for c in col if c.value is not None), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max(length + 3, 12), 42)

async def _build_xlsx(kind: str, component_id: Optional[str] = None, date_from: Optional[str] = None, date_to: Optional[str] = None, warehouse_id: Optional[str] = None) -> bytes:
    wb = Workbook(); ws = wb.active
    if kind == "stock":
        ws.title = "Stock"
        ws.append(["Code", "Component", "Category", "Location", "Available", "Minimum", "Unit ₹", "Value ₹"])
        for c in await db.components.find({}, {"_id": 0}).sort("name", 1).to_list(1000):
            ws.append([c["code"], c["name"], c["category"], c.get("location",""), c.get("stock",0), c.get("minimum_stock",0), c.get("unit_price",0), c.get("stock",0)*c.get("unit_price",0)])
    elif kind in ("purchases", "issues"):
        target = "purchase" if kind == "purchases" else "issue"
        ws.title = kind.title()
        ws.append(["Date", "Reference", "Component", "Party", "Quantity", "Balance", "Created by"])
        for m in await db.movements.find({"type": target}, {"_id": 0}).sort("created_at", -1).to_list(2000):
            ws.append([(m.get("created_at","") or "")[:10], m.get("reference",""), m.get("component_name",""), m.get("party",""), m.get("quantity",0), m.get("balance",0), m.get("created_by","")])
    elif kind == "monthly":
        ws.title = "Monthly"
        ws.append(["Month", "Purchases", "Issues", "Returns", "Net"])
        buckets = {}
        for m in await db.movements.find({}, {"_id": 0}).to_list(3000):
            key = (m.get("created_at","") or "")[:7] or "current"
            buckets.setdefault(key, {"purchase":0, "issue":0, "return":0})
            buckets[key][m.get("type","purchase")] = buckets[key].get(m.get("type","purchase"),0) + m.get("quantity",0)
        for k, v in sorted(buckets.items()):
            ws.append([k, v.get("purchase",0), v.get("issue",0), v.get("return",0), v.get("purchase",0)+v.get("return",0)-v.get("issue",0)])
    elif kind == "vendors":
        ws.title = "Vendors"
        ws.append(["Name", "Contact", "Phone", "Email", "GST", "Address", "Notes"])
        for v in await db.vendors.find({}, {"_id": 0}).sort("name", 1).to_list(1000):
            ws.append([v.get("name",""), v.get("contact",""), v.get("phone",""), v.get("email",""), v.get("gst",""), v.get("address",""), v.get("notes","")])
    elif kind == "employees":
        ws.title = "Employees"
        ws.append(["Name", "Employee Code", "Department", "Designation", "Phone", "Email"])
        for e in await db.employees.find({}, {"_id": 0}).sort("name", 1).to_list(1000):
            ws.append([e.get("name",""), e.get("employee_code",""), e.get("department",""), e.get("designation",""), e.get("phone",""), e.get("email","")])
    elif kind == "ledger":
        if not component_id: raise HTTPException(400, "component_id required for ledger export")
        ledger = await _compute_ledger(component_id, date_from, date_to, warehouse_id)
        ws.title = "Ledger"
        ws.append(["Component", ledger["component"]["name"], "Code", ledger["component"]["code"], "Opening", ledger["opening_stock"], "Closing", ledger["closing_stock"]])
        ws.append([])
        ws.append(["Date", "Type", "Reference", "Party", "Inward", "Outward", "Balance"])
        for r in ledger["rows"]:
            ws.append([r["date"], r["type"], r["reference"], r["party"], r["inward"], r["outward"], r["balance"]])
        _xlsx_style(ws, header_row=3)
        buf = io.BytesIO(); wb.save(buf); return buf.getvalue()
    elif kind == "audit":
        ws.title = "Audit"
        ws.append(["Timestamp", "Action", "User", "Details"])
        for e in await db.audit.find({}, {"_id": 0}).sort("created_at", -1).to_list(2000):
            ws.append([(e.get("created_at","") or "")[:19].replace("T"," "), e.get("action",""), e.get("user",""), e.get("details","")])
    else:
        raise HTTPException(400, "Unsupported export")
    _xlsx_style(ws)
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

@api.get("/reports/xlsx")
async def export_xlsx(kind: str = "stock", component_id: Optional[str] = None, date_from: Optional[str] = None, date_to: Optional[str] = None, warehouse_id: Optional[str] = None, auth: Optional[str] = Query(None), authorization: Optional[str] = Header(None)):
    token_value = auth or (authorization[7:] if authorization and authorization.startswith("Bearer ") else None)
    await user_from_token(token_value)
    data = await _build_xlsx(kind, component_id=component_id, date_from=date_from, date_to=date_to, warehouse_id=warehouse_id)
    return Response(content=data, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f'attachment; filename="traxo-{kind}-report.xlsx"'})

# ---------- QR codes ----------
def _make_qr_png(payload: str) -> bytes:
    qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=10, border=2)
    qr.add_data(payload); qr.make(fit=True)
    img = qr.make_image(fill_color="#0f172a", back_color="white")
    buf = io.BytesIO(); img.save(buf, format="PNG"); return buf.getvalue()

@api.get("/components/{component_id}/qr")
async def component_qr(component_id: str, auth: Optional[str] = Query(None), authorization: Optional[str] = Header(None)):
    token_value = auth or (authorization[7:] if authorization and authorization.startswith("Bearer ") else None)
    await user_from_token(token_value)
    component = await db.components.find_one({"id": component_id}, {"_id": 0})
    if not component: raise HTTPException(404, "Component not found")
    payload = f"TRAXO|{component['code']}|{component_id}|{component['name']}"
    return Response(content=_make_qr_png(payload), media_type="image/png")

@api.get("/components/scan/{code}")
async def component_by_code(code: str, user=Depends(current_user)):
    """Resolve a scanned QR payload back to a component."""
    if "|" in code:
        parts = code.split("|")
        if len(parts) >= 3 and parts[0] == "TRAXO":
            component = await db.components.find_one({"id": parts[2]}, {"_id": 0})
            if component: return safe(component)
    component = await db.components.find_one({"$or": [{"code": code}, {"id": code}, {"part_number": code}]}, {"_id": 0})
    if not component: raise HTTPException(404, "Component not found for scan")
    return component

# ---------- File uploads ----------
@api.post("/files/upload")
async def upload_file(file: UploadFile = File(...), scope: str = Form("general"), scope_id: str = Form(""), user=Depends(current_user)):
    ext = (file.filename or "").rsplit(".", 1)[-1].lower() if "." in (file.filename or "") else "bin"
    if ext not in ["pdf","png","jpg","jpeg","webp","gif","csv","txt","xlsx","docx"]: raise HTTPException(400, "Unsupported file type")
    file_id = str(uuid.uuid4())
    storage_path = f"{APP_NAME}/{scope}/{user['id']}/{file_id}.{ext}"
    data = await file.read()
    if len(data) > 12 * 1024 * 1024: raise HTTPException(413, "File exceeds 12 MB limit")
    result = put_object(storage_path, data, file.content_type or "application/octet-stream")
    record = {"id": file_id, "storage_path": result["path"], "original_filename": file.filename, "content_type": file.content_type, "size": result.get("size", len(data)), "scope": scope, "scope_id": scope_id, "uploaded_by": user["name"], "created_at": now(), "is_deleted": False}
    await db.files.insert_one(record)
    if scope == "component" and scope_id:
        await db.components.update_one({"id": scope_id}, {"$push": {"attachments": file_id}})
    return safe(record)

@api.get("/files/{file_id}")
async def download_file(file_id: str, auth: Optional[str] = Query(None), authorization: Optional[str] = Header(None)):
    token_value = auth or (authorization[7:] if authorization and authorization.startswith("Bearer ") else None)
    await user_from_token(token_value)
    record = await db.files.find_one({"id": file_id, "is_deleted": False}, {"_id": 0})
    if not record: raise HTTPException(404, "File not found")
    data, content_type = get_object(record["storage_path"])
    return Response(content=data, media_type=record.get("content_type") or content_type, headers={"Content-Disposition": f'inline; filename="{record.get("original_filename","file")}"'})

@api.get("/files")
async def list_files(scope: Optional[str] = None, scope_id: Optional[str] = None, user=Depends(current_user)):
    q = {"is_deleted": False}
    if scope: q["scope"] = scope
    if scope_id: q["scope_id"] = scope_id
    return await db.files.find(q, {"_id": 0}).sort("created_at", -1).to_list(200)

# ---------- Read models ----------
@api.get("/movements")
async def movements(user=Depends(current_user)):
    return await db.movements.find({}, {"_id": 0}).sort("created_at", -1).to_list(500)

@api.get("/dashboard")
async def dashboard(user=Depends(current_user)):
    comps = await db.components.find({}, {"_id": 0}).to_list(500); moves = await db.movements.find({}, {"_id": 0}).sort("created_at", -1).to_list(200)
    purchases = sum(x["quantity"] for x in moves if x["type"] == "purchase"); issues = sum(x["quantity"] for x in moves if x["type"] == "issue")
    return {"metrics": {"components": len(comps), "stock_value": round(sum(x.get("stock",0)*x.get("unit_price",0) for x in comps), 2), "purchases": purchases, "issues": issues, "low_stock": sum(1 for x in comps if x.get("stock",0) <= x.get("minimum_stock",0)), "vendors": await db.vendors.count_documents({})}, "components": comps[:8], "movements": moves[:8]}

@api.get("/audit")
async def audit(user=Depends(current_user)): return await db.audit.find({}, {"_id": 0}).sort("created_at", -1).to_list(200)

# ---------- Seed ----------
async def seed():
    if await db.users.count_documents({}) == 0:
        await db.users.insert_one({"id":"usr-admin","email":os.environ.get("ADMIN_EMAIL","admin@traxo.in"),"name":"Aarav Mehta","role":"admin","password_hash":hash_password(os.environ.get("ADMIN_PASSWORD","Traxo@123")),"created_at":now()})
    if await db.components.count_documents({}) == 0:
        items = [{"id":"cmp-esp32","code":"CMP-1001","name":"ESP32 Development Board","category":"Microcontroller","part_number":"ESP32-WROOM-32","unit":"Nos","minimum_stock":25,"location":"Rack A-04","unit_price":420,"manufacturer":"Espressif","hsn":"85423900","gst":18,"stock":142,"warehouse_balances":{"wh-main":142},"attachments":[]},{"id":"cmp-relay","code":"CMP-1002","name":"24V Relay Module","category":"Modules","part_number":"RM-24V-8CH","unit":"Nos","minimum_stock":20,"location":"Rack B-02","unit_price":680,"manufacturer":"Omron","hsn":"85364900","gst":18,"stock":18,"warehouse_balances":{"wh-main":18},"attachments":[]},{"id":"cmp-sensor","code":"CMP-1003","name":"Proximity Sensor","category":"Sensors","part_number":"PRX-M18-NPN","unit":"Nos","minimum_stock":12,"location":"Rack C-01","unit_price":1250,"manufacturer":"Autonics","hsn":"85365090","gst":18,"stock":64,"warehouse_balances":{"wh-main":44,"wh-line":20},"attachments":[]},{"id":"cmp-terminal","code":"CMP-1004","name":"PCB Terminal Block","category":"PCB Parts","part_number":"TB-2P-5.08","unit":"Nos","minimum_stock":100,"location":"Bin D-12","unit_price":18,"manufacturer":"Phoenix","hsn":"85369090","gst":18,"stock":320,"warehouse_balances":{"wh-main":200,"wh-line":120},"attachments":[]}]
        await db.components.insert_many(items)
        await db.vendors.insert_many([{"id":"ven-1","name":"TechCore Components","contact":"Nisha Kapoor","phone":"+91 98765 43210","email":"nisha@techcore.in","gst":"29ABCDE1234F1Z5","address":"Peenya Industrial Area, Bengaluru","notes":"Preferred for ESP32 boards"}, {"id":"ven-2","name":"Industrial Parts Co.","contact":"Vikram Shah","phone":"+91 98200 11884","email":"vikram@industrialparts.co","gst":"27PQRST5678K1Z2","address":"Andheri East, Mumbai","notes":"Bulk terminal blocks"}])
        await db.employees.insert_many([{"id":"emp-1","name":"Rohan Kulkarni","department":"Engineering","designation":"Automation Engineer","phone":"+91 90111 22233","email":"rohan@traxo.in","employee_code":"TRX-E-001"}, {"id":"emp-2","name":"Meera Joshi","department":"Production","designation":"Project Lead","phone":"+91 90222 33344","email":"meera@traxo.in","employee_code":"TRX-E-002"}])
        await db.movements.insert_many([{ "id":"mov-1","type":"purchase","component_id":"cmp-esp32","component_name":"ESP32 Development Board","quantity":50,"party":"TechCore Components","reference":"PO-2024-018","balance":142,"created_by":"Aarav Mehta","created_at":now()},{"id":"mov-2","type":"issue","component_id":"cmp-relay","component_name":"24V Relay Module","quantity":12,"party":"Rohan Kulkarni","project":"Conveyor Retrofit","balance":18,"created_by":"Sonia Rao","created_at":now()}])
    if await db.warehouses.count_documents({}) == 0:
        await db.warehouses.insert_many([{"id":"wh-main","name":"Main Store","code":"WH-01","address":"Traxo India Automation · Bengaluru","active":True},{"id":"wh-line","name":"Assembly Line Store","code":"WH-02","address":"Production floor · Bengaluru","active":True}])

@app.on_event("startup")
async def startup():
    await seed()
    # Backfill warehouse_balances for any legacy component that doesn't have them
    async for c in db.components.find({"warehouse_balances": {"$exists": False}}, {"_id": 0, "id": 1, "stock": 1}):
        await db.components.update_one({"id": c["id"]}, {"$set": {"warehouse_balances": {"wh-main": c.get("stock", 0)}}})
    init_storage()

app.include_router(api)
allowed_origins = [origin.strip() for origin in os.environ.get("CORS_ORIGINS", "").split(",") if origin.strip() and origin.strip() != "*"]
logging.basicConfig(level=logging.INFO)
