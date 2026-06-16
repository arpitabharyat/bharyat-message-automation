from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Optional, List
from dotenv import load_dotenv
from supabase import create_client
import aiosmtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from twilio.rest import Client as TwilioClient
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill
import openpyxl
import shutil
import io
import os

load_dotenv()

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Supabase
supabase = create_client(os.getenv("SUPABASE_URL"), os.getenv("SUPABASE_KEY"))

# Twilio
twilio = TwilioClient(os.getenv("TWILIO_ACCOUNT_SID"), os.getenv("TWILIO_AUTH_TOKEN"))

# ---------- Models ----------

class Customer(BaseModel):
    name: str
    email: Optional[str] = None
    whatsapp: Optional[str] = None

class Recipient(BaseModel):
    name: str
    email: Optional[str] = None
    whatsapp: Optional[str] = None

class SendMessageRequest(BaseModel):
    message: str
    subject: Optional[str] = "Message from Bharyat"
    recipients: List[Recipient]

# ---------- Customers ----------

@app.get("/api/customers")
def get_customers():
    res = supabase.table("customers").select("*").order("created_at", desc=True).execute()
    return res.data

@app.post("/api/customers")
def add_customer(c: Customer):
    if not c.email and not c.whatsapp:
        raise HTTPException(400, "Email or WhatsApp required")
    res = supabase.table("customers").insert({
        "name": c.name,
        "email": c.email,
        "whatsapp": c.whatsapp
    }).execute()
    return res.data[0]

@app.delete("/api/customers/{customer_id}")
def delete_customer(customer_id: int):
    supabase.table("customers").delete().eq("id", customer_id).execute()
    return {"success": True}

# ---------- Send Message ----------

async def send_email(to_email: str, to_name: str, subject: str, body: str):
    msg = MIMEMultipart()
    msg["From"] = os.getenv("ZOHO_EMAIL")
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))
    await aiosmtplib.send(
        msg,
        hostname="smtp.zoho.in",
        port=587,
        start_tls=True,
        username=os.getenv("ZOHO_EMAIL"),
        password=os.getenv("ZOHO_PASSWORD"),
    )

def send_whatsapp(to_number: str, body: str):
    if not to_number.startswith('+'):
        to_number = '+' + to_number
    twilio.messages.create(
        from_=os.getenv("TWILIO_WHATSAPP_FROM"),
        to=f"whatsapp:{to_number}",
        body=body
    )

@app.post("/api/send-message")
async def send_message(req: SendMessageRequest):
    results = {"email": [], "whatsapp": [], "errors": []}

    for r in req.recipients:
        if r.email:
            try:
                await send_email(r.email, r.name, req.subject, req.message)
                results["email"].append(r.name)
                supabase.table("message_logs").insert({
                    "customer_name": r.name,
                    "channel": "email",
                    "subject": req.subject,
                    "message": req.message,
                    "status": "sent"
                }).execute()
            except Exception as e:
                results["errors"].append({"name": r.name, "channel": "email", "error": str(e)})
                supabase.table("message_logs").insert({
                    "customer_name": r.name,
                    "channel": "email",
                    "subject": req.subject,
                    "message": req.message,
                    "status": "failed",
                    "error_text": str(e)
                }).execute()

        if r.whatsapp:
            try:
                send_whatsapp(r.whatsapp, req.message)
                results["whatsapp"].append(r.name)
                supabase.table("message_logs").insert({
                    "customer_name": r.name,
                    "channel": "whatsapp",
                    "message": req.message,
                    "status": "sent"
                }).execute()
            except Exception as e:
                results["errors"].append({"name": r.name, "channel": "whatsapp", "error": str(e)})
                supabase.table("message_logs").insert({
                    "customer_name": r.name,
                    "channel": "whatsapp",
                    "message": req.message,
                    "status": "failed",
                    "error_text": str(e)
                }).execute()

    return {"success": True, "results": results}

# ---------- Logs ----------

@app.get("/api/logs")
def get_logs():
    res = supabase.table("message_logs").select("*").order("sent_at", desc=True).limit(100).execute()
    return res.data

# ---------- Upload Customers Excel ----------

@app.post("/api/customers/upload-excel")
async def upload_excel(file: UploadFile = File(...)):
    contents = await file.read()
    wb = openpyxl.load_workbook(io.BytesIO(contents))
    ws = wb.active

    headers = [str(cell.value).strip().lower() if cell.value else '' for cell in ws[1]]

    name_col = next((i for i, h in enumerate(headers) if 'name' in h), None)
    email_col = next((i for i, h in enumerate(headers) if 'email' in h or 'mail' in h), None)
    wa_col = next((i for i, h in enumerate(headers) if 'whatsapp' in h or 'mobile' in h or 'phone' in h or 'wa' in h), None)

    if name_col is None:
        raise HTTPException(400, "Name column not found in Excel")

    added = []
    skipped = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        name = str(row[name_col]).strip() if row[name_col] else None
        email = str(row[email_col]).strip() if email_col is not None and row[email_col] else None
        wa = str(row[wa_col]).strip() if wa_col is not None and row[wa_col] else None

        if not name or name == 'None':
            skipped += 1
            continue
        if not email and not wa:
            skipped += 1
            continue

        res = supabase.table("customers").insert({
            "name": name,
            "email": email,
            "whatsapp": wa
        }).execute()
        added.append(res.data[0])

    return {"success": True, "added": len(added), "skipped": skipped, "customers": added}

# ---------- RFQ Send (with file upload) ----------

@app.post("/api/rfq/send")
async def send_rfq(
    file: UploadFile = File(...),
    supplier_ids: str = Form(...),
    subject: str = Form(...),
    message: str = Form(...)
):
    temp_path = f"/tmp/{file.filename}"
    with open(temp_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    ids = [int(i) for i in supplier_ids.split(",")]
    res = supabase.table("customers").select("*").in_("id", ids).execute()
    suppliers = res.data
    results = {"sent": [], "errors": []}

    for supplier in suppliers:
        if supplier.get("email"):
            try:
                msg = MIMEMultipart()
                msg["From"] = os.getenv("ZOHO_EMAIL")
                msg["To"] = supplier["email"]
                msg["Subject"] = subject
                msg["Reply-To"] = os.getenv("ZOHO_EMAIL")
                full_message = message + f"\n\nPlease reply to this email with the filled Excel sheet.\nEmail: {os.getenv('ZOHO_EMAIL')}"
                msg.attach(MIMEText(full_message, "plain"))

                with open(temp_path, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename={file.filename}")
                    msg.attach(part)

                await aiosmtplib.send(
                    msg,
                    hostname="smtp.zoho.in",
                    port=587,
                    start_tls=True,
                    username=os.getenv("ZOHO_EMAIL"),
                    password=os.getenv("ZOHO_PASSWORD"),
                )
                results["sent"].append(supplier["name"])
                supabase.table("message_logs").insert({
                    "customer_name": supplier["name"],
                    "channel": "email",
                    "subject": subject,
                    "message": message,
                    "status": "sent"
                }).execute()
            except Exception as e:
                results["errors"].append({"name": supplier["name"], "error": str(e)})

    os.remove(temp_path)
    return {"success": True, "results": results}

# ---------- RFQ Generate + Send ----------

@app.post("/api/rfq/send-with-parts")
async def send_rfq_with_parts(
    supplier_ids: str = Form(...),
    subject: str = Form(...),
    message: str = Form(...),
    part_numbers: str = Form(...),
):
    # Generate Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "RFQ"

    # Add logo if exists
    logo_path = os.path.join(os.path.dirname(__file__), 'logo.png')
    if os.path.exists(logo_path):
        try:
            img = XLImage(logo_path)
            img.width = 200
            img.height = 55
            ws.add_image(img, 'A1')
            ws.append([])
            ws.append([])
            ws.append([])
        except Exception as e:
            print(f"Logo error (skipping): {e}")

    # Headers
    headers = ["Part Number", "Data Code", "Unit Price", "Delivery Time", "Notes"]
    ws.append(headers)

    # Style headers
    header_row = ws.max_row
    for col in range(1, 6):
        cell = ws.cell(row=header_row, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="534AB7")

    # Part numbers
    for pn in part_numbers.split(","):
        if pn.strip():
            ws.append([pn.strip(), "", "", "", ""])

    # Column widths
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 20

    # Save to buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    excel_data = buffer.getvalue()

    # Get suppliers
    ids = [int(i) for i in supplier_ids.split(",")]
    res = supabase.table("customers").select("*").in_("id", ids).execute()
    suppliers = res.data
    results = {"sent": [], "errors": []}

    for supplier in suppliers:
        if supplier.get("email"):
            try:
                msg = MIMEMultipart()
                msg["From"] = os.getenv("ZOHO_EMAIL")
                msg["To"] = supplier["email"]
                msg["Subject"] = subject
                msg["Reply-To"] = os.getenv("ZOHO_EMAIL")
                # msg["To"] = supplier["email"]
                # msg["Cc"] = os.getenv("SIR_EMAIL")  # Sir ko CC
                # msg["Reply-To"] = os.getenv("SIR_EMAIL")  # Reply Sir ke paas jaaye
                full_message = message + f"\n\nPlease reply to this email with the filled Excel sheet.\nEmail: {os.getenv('ZOHO_EMAIL')}"
                msg.attach(MIMEText(full_message, "plain"))

                part = MIMEBase("application", "octet-stream")
                part.set_payload(excel_data)
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", "attachment; filename=RFQ.xlsx")
                msg.attach(part)

                # await aiosmtplib.send(
                #     msg,
                #     hostname="smtp.zoho.in",
                #     port=587,
                #     start_tls=True,
                #     username=os.getenv("ZOHO_EMAIL"),
                #     password=os.getenv("ZOHO_PASSWORD"),
                # )
                await aiosmtplib.send(
                    msg,
                    hostname="smtp.zoho.in",
                    port=587,        # 465 se 587 karo
                    start_tls=True,
                    # use_tls=True,
                    username=os.getenv("ZOHO_EMAIL"),
                    password=os.getenv("ZOHO_PASSWORD"),
                    recipients=[supplier["email"]]
                )
                results["sent"].append(supplier["name"])
                supabase.table("message_logs").insert({
                    "customer_name": supplier["name"],
                    "channel": "email",
                    "subject": subject,
                    "message": message,
                    "status": "sent"
                }).execute()

            except Exception as e:
                print(f"ERROR sending to {supplier['name']}: {e}")
                results["errors"].append({"name": supplier["name"], "error": str(e)})

    return {"success": True, "results": results}