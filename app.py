import os
import sqlite3
import csv
import html
try:
    import openpyxl
except ImportError:
    openpyxl = None

from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file, abort

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

app = Flask(__name__)
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0
app.jinja_env.auto_reload = True

@app.after_request
def add_cache_control_headers(response):
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response

DB_NAME = os.path.join(os.path.dirname(os.path.abspath(__file__)), "billing.db")
PDF_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "generated_pdfs")
os.makedirs(PDF_DIR, exist_ok=True)

# =====================================================================
# COMPANY DETAILS (fixed) - edit these once, they appear on every bill
# =====================================================================
COMPANY = {
    "name": "SRI RADHE ENTERPRISES",
    "tagline": "General Merchants & Commission Forwarding Agents",
    "gstin": "07AEYPN5512J1ZY",
    "address_line1": "6148, Block No.1, Gali No.5, Dev Nagar, New Delhi -110005",
    "address_line2": "BO: 2317, Gali Hinga Beg, Tilak Bazar, Khari Baoli, Old Delhi -110006",
    "mobile": "9868983010",
    "bank_name": "J & K Bank Ltd.",
    "bank_branch": "Karol Bagh, New Delhi",
    "bank_account": "0123020100000232",
    "bank_ifsc": "JAKA0KAROLE",
    "jurisdiction": "All disputes are subjected to Delhi Jurisdiction only",
}


# =====================================================================
# DATABASE
# =====================================================================
def get_db():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def _add_column_if_missing(cur, table, column, coltype):
    cur.execute(f"PRAGMA table_info({table})")
    existing = [row[1] for row in cur.fetchall()]
    if column not in existing:
        cur.execute(f"ALTER TABLE {table} ADD COLUMN {column} {coltype}")


def init_db():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS invoices(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_no TEXT,
        bill_date TEXT,
        delivery_type TEXT,
        customer_name TEXT,
        customer_address TEXT,
        customer_gstin TEXT,
        state_code TEXT,
        transporter TEXT,
        marka TEXT,
        labour_amount REAL,
        remarks TEXT,
        subtotal REAL,
        gst_total REAL,
        grand_total REAL,
        net_weight REAL,
        total_qty REAL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS invoice_items(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_id INTEGER,
        product_name TEXT,
        hsn_code TEXT,
        qty REAL,
        qty_unit TEXT,
        weight REAL,
        rate REAL,
        gst REAL,
        amount REAL,
        tax_amount REAL
    )
    """)

    # Migrate older databases that might be missing some of the newer
    # columns, so nothing breaks on upgrade.
    for col, coltype in [
        ("delivery_type", "TEXT"), ("customer_address", "TEXT"),
        ("customer_gstin", "TEXT"), ("state_code", "TEXT"),
        ("marka", "TEXT"), ("net_weight", "REAL"), ("total_qty", "REAL"),
    ]:
        _add_column_if_missing(cur, "invoices", col, coltype)

    for col, coltype in [("qty_unit", "TEXT"), ("tax_amount", "REAL")]:
        _add_column_if_missing(cur, "invoice_items", col, coltype)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS customers(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE COLLATE NOCASE,
        address TEXT,
        gstin TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS transporters(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE COLLATE NOCASE
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS markas(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE COLLATE NOCASE
    )
    """)

    # Seed default transporters if empty
    cur.execute("SELECT COUNT(*) FROM transporters")
    if cur.fetchone()[0] == 0:
        default_transporters = [
            "Sky Transport",
            "Delhi Gujarat Fleet",
            "ARC (Associated Road Carriers)",
            "Jaipur Golden",
            "V-Trans",
            "TCI Freight",
            "Direct / By Hand",
            "Local Tempo / Self Pickup",
        ]
        cur.executemany("INSERT OR IGNORE INTO transporters(name) VALUES(?)", [(t,) for t in default_transporters])

    # Seed default markas if empty
    cur.execute("SELECT COUNT(*) FROM markas")
    if cur.fetchone()[0] == 0:
        default_markas = [
            "BS/SNR",
            "SR/DEL",
            "KDF/SXR",
            "M&S/ATN",
            "RSE",
            "EXP/DEL",
        ]
        cur.executemany("INSERT OR IGNORE INTO markas(name) VALUES(?)", [(m,) for m in default_markas])

    # Import any existing ones from invoices table if available
    try:
        cur.execute("INSERT OR IGNORE INTO transporters(name) SELECT DISTINCT transporter FROM invoices WHERE transporter IS NOT NULL AND TRIM(transporter) != ''")
        cur.execute("INSERT OR IGNORE INTO markas(name) SELECT DISTINCT marka FROM invoices WHERE marka IS NOT NULL AND TRIM(marka) != ''")
    except Exception:
        pass

    # Products table
    cur.execute("""
    CREATE TABLE IF NOT EXISTS products(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE COLLATE NOCASE,
        hsn TEXT,
        gst REAL DEFAULT 5.0
    )
    """)

    cur.execute("SELECT COUNT(*) FROM products")
    if cur.fetchone()[0] == 0:
        default_products = [
            ("Badamgiri (Almond Kernels) - American / R", "08021200", 5.0),
            ("Badam Mamra", "08021200", 5.0),
            ("Badam Sanora", "08021200", 5.0),
            ("Kaju (Cashew) W240", "08013100", 5.0),
            ("Kaju W320", "08013100", 5.0),
            ("Kaju Tukda / Splits", "08013200", 5.0),
            ("Akhrot Giri (Walnut Kernels)", "08023200", 5.0),
            ("Akhrot Sabut (In-shell)", "08023100", 5.0),
            ("Pista Irani", "08025100", 5.0),
            ("Pista Maghaz", "08025200", 5.0),
            ("Kishmish Indian / Afghan", "08062010", 5.0),
            ("Munakka", "08062010", 5.0),
            ("Anjeer (Figs)", "08042090", 5.0),
            ("Makhana (Fox Nuts)", "19041090", 5.0),
            ("Elaichi (Green Cardamom)", "09083100", 5.0),
        ]
        cur.executemany("INSERT OR IGNORE INTO products(name, hsn, gst) VALUES(?,?,?)", default_products)

    try:
        cur.execute("""
        INSERT OR IGNORE INTO products(name, hsn, gst)
        SELECT DISTINCT product_name, hsn_code, COALESCE(gst, 5.0)
        FROM invoice_items
        WHERE product_name IS NOT NULL AND TRIM(product_name) != ''
        """)
    except Exception:
        pass

    conn.commit()
    conn.close()


def parse_customers_file(file_path):
    customers = []
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".xlsx" and openpyxl:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb.active
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            return []
        headers = [str(h).strip().lower() if h is not None else "" for h in rows[0]]
        name_idx, addr_idx, gst_idx = -1, -1, -1
        for i, h in enumerate(headers):
            if any(k in h for k in ["name", "party", "customer", "firm"]):
                if name_idx == -1: name_idx = i
            elif any(k in h for k in ["addr", "location", "city", "place"]):
                if addr_idx == -1: addr_idx = i
            elif any(k in h for k in ["gst", "tin"]):
                if gst_idx == -1: gst_idx = i
        if name_idx == -1 and len(headers) >= 1: name_idx = 0
        if addr_idx == -1 and len(headers) >= 2: addr_idx = 1
        if gst_idx == -1 and len(headers) >= 3: gst_idx = 2

        for row in rows[1:]:
            if not row or not any(row):
                continue
            name = str(row[name_idx]).strip() if name_idx != -1 and name_idx < len(row) and row[name_idx] is not None else ""
            addr = str(row[addr_idx]).strip() if addr_idx != -1 and addr_idx < len(row) and row[addr_idx] is not None else ""
            gst = str(row[gst_idx]).strip() if gst_idx != -1 and gst_idx < len(row) and row[gst_idx] is not None else ""
            if name:
                customers.append({"name": name, "address": addr, "gstin": gst})

    elif ext == ".csv":
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            reader = csv.reader(f)
            rows = [r for r in reader if r]
            if not rows:
                return []
            headers = [h.strip().lower() for h in rows[0]]
            name_idx, addr_idx, gst_idx = -1, -1, -1
            for i, h in enumerate(headers):
                if any(k in h for k in ["name", "party", "customer", "firm"]):
                    if name_idx == -1: name_idx = i
                elif any(k in h for k in ["addr", "location", "city", "place"]):
                    if addr_idx == -1: addr_idx = i
                elif any(k in h for k in ["gst", "tin"]):
                    if gst_idx == -1: gst_idx = i
            if name_idx == -1 and len(headers) >= 1: name_idx = 0
            if addr_idx == -1 and len(headers) >= 2: addr_idx = 1
            if gst_idx == -1 and len(headers) >= 3: gst_idx = 2

            for row in rows[1:]:
                name = row[name_idx].strip() if name_idx != -1 and name_idx < len(row) else ""
                addr = row[addr_idx].strip() if addr_idx != -1 and addr_idx < len(row) else ""
                gst = row[gst_idx].strip() if gst_idx != -1 and gst_idx < len(row) else ""
                if name:
                    customers.append({"name": name, "address": addr, "gstin": gst})

    return customers


def sync_customers_from_file(file_path):
    customers = parse_customers_file(file_path)
    if not customers:
        return 0

    conn = get_db()
    cur = conn.cursor()
    count = 0
    for c in customers:
        cur.execute("""
        INSERT INTO customers (name, address, gstin)
        VALUES (?, ?, ?)
        ON CONFLICT(name) DO UPDATE SET
            address = excluded.address,
            gstin = excluded.gstin
        """, (c["name"], c["address"], c["gstin"]))
        count += 1
    conn.commit()
    conn.close()
    return count


def auto_import_customers_if_needed():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM customers")
    count = cur.fetchone()[0]
    conn.close()

    base_dir = os.path.dirname(os.path.abspath(__file__))
    if count == 0:
        for candidate in ["customers.xlsx", "customers.csv", "customers_template.xlsx"]:
            path = os.path.join(base_dir, candidate)
            if os.path.exists(path):
                sync_customers_from_file(path)
                break


init_db()
auto_import_customers_if_needed()


PORT = int(os.environ.get("PORT", 5000))


def get_local_ip():
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        local_ip = s.getsockname()[0]
        s.close()
    except Exception:
        local_ip = "127.0.0.1"
    return local_ip


@app.route("/")
def home():
    local_ip = get_local_ip()
    phone_url = f"http://{local_ip}:{PORT}"
    return render_template("index.html", phone_url=phone_url, local_ip=local_ip, port=PORT)


@app.route("/api/presets")
def get_presets():
    return jsonify({
        "products": [
            {"name": "Badamgiri (Almond Kernels) - American / R", "hsn": "08021200", "gst": 5},
            {"name": "Badam Mamra", "hsn": "08021200", "gst": 5},
            {"name": "Badam Sanora", "hsn": "08021200", "gst": 5},
            {"name": "Kaju (Cashew) W240", "hsn": "08013100", "gst": 5},
            {"name": "Kaju W320", "hsn": "08013100", "gst": 5},
            {"name": "Kaju Tukda / Splits", "hsn": "08013200", "gst": 5},
            {"name": "Akhrot Giri (Walnut Kernels)", "hsn": "08023200", "gst": 5},
            {"name": "Akhrot Sabut (In-shell)", "hsn": "08023100", "gst": 5},
            {"name": "Pista Irani", "hsn": "08025100", "gst": 5},
            {"name": "Pista Maghaz", "hsn": "08025200", "gst": 5},
            {"name": "Kishmish Indian / Afghan", "hsn": "08062010", "gst": 5},
            {"name": "Munakka", "hsn": "08062010", "gst": 5},
            {"name": "Anjeer (Figs)", "hsn": "08042090", "gst": 5},
            {"name": "Makhana (Fox Nuts)", "hsn": "19041090", "gst": 5},
            {"name": "Elaichi (Green Cardamom)", "hsn": "09083100", "gst": 5},
        ],
        "transporters": [
            "Sky Transport",
            "Delhi Gujarat Fleet",
            "ARC (Associated Road Carriers)",
            "Jaipur Golden",
            "V-Trans",
            "TCI Freight",
            "Direct / By Hand",
            "Local Tempo / Self Pickup"
        ],
        "markas": [
            "BS/SNR",
            "SR/DEL",
            "KDF/SXR",
            "M&S/ATN",
            "RSE",
            "EXP/DEL"
        ]
    })


@app.route("/api/customers")
def get_customers():
    conn = get_db()
    rows = conn.execute("SELECT id, name, address, gstin FROM customers ORDER BY name COLLATE NOCASE ASC").fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/customers/save", methods=["POST"])
def save_customer():
    data = request.json or {}
    name = (data.get("name") or "").strip()
    address = (data.get("address") or "").strip()
    gstin = (data.get("gstin") or "").strip().upper()

    if not name:
        return jsonify({"success": False, "error": "Customer Name is required"}), 400

    conn = get_db()
    cur = conn.cursor()
    # Check if customer already exists (case-insensitive)
    existing = cur.execute("SELECT id FROM customers WHERE LOWER(name)=?", (name.lower(),)).fetchone()
    if existing:
        cur.execute("UPDATE customers SET address=?, gstin=? WHERE id=?", (address, gstin, existing["id"]))
        cust_id = existing["id"]
        msg = f"Customer '{name}' updated successfully!"
    else:
        cur.execute("INSERT INTO customers (name, address, gstin) VALUES (?, ?, ?)", (name, address, gstin))
        cust_id = cur.lastrowid
        msg = f"Customer '{name}' added to list permanently!"

    conn.commit()
    conn.close()

    return jsonify({
        "success": True,
        "message": msg,
        "customer": {
            "id": cust_id,
            "name": name,
            "address": address,
            "gstin": gstin
        }
    })


@app.route("/api/transporters")
def get_transporters():
    conn = get_db()
    rows = conn.execute("SELECT id, name FROM transporters ORDER BY name COLLATE NOCASE ASC").fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/transporters/save", methods=["POST"])
def save_transporter():
    data = request.json or {}
    name = (data.get("name") or "").strip()

    if not name:
        return jsonify({"success": False, "error": "Transporter Name is required"}), 400

    conn = get_db()
    cur = conn.cursor()
    existing = cur.execute("SELECT id, name FROM transporters WHERE LOWER(name)=?", (name.lower(),)).fetchone()
    if existing:
        t_id = existing["id"]
        t_name = existing["name"]
        msg = f"Transporter '{t_name}' already exists in list!"
    else:
        cur.execute("INSERT INTO transporters (name) VALUES (?)", (name,))
        t_id = cur.lastrowid
        t_name = name
        msg = f"Transporter '{name}' added to list permanently!"

    conn.commit()
    conn.close()

    return jsonify({
        "success": True,
        "message": msg,
        "transporter": {"id": t_id, "name": t_name}
    })


@app.route("/api/markas")
def get_markas():
    conn = get_db()
    rows = conn.execute("SELECT id, name FROM markas ORDER BY name COLLATE NOCASE ASC").fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/markas/save", methods=["POST"])
def save_marka():
    data = request.json or {}
    name = (data.get("name") or "").strip()

    if not name:
        return jsonify({"success": False, "error": "Marka Name is required"}), 400

    conn = get_db()
    cur = conn.cursor()
    existing = cur.execute("SELECT id, name FROM markas WHERE LOWER(name)=?", (name.lower(),)).fetchone()
    if existing:
        m_id = existing["id"]
        m_name = existing["name"]
        msg = f"Marka '{m_name}' already exists in list!"
    else:
        cur.execute("INSERT INTO markas (name) VALUES (?)", (name,))
        m_id = cur.lastrowid
        m_name = name
        msg = f"Marka '{name}' added to list permanently!"

    conn.commit()
    conn.close()

    return jsonify({
        "success": True,
        "message": msg,
        "marka": {"id": m_id, "name": m_name}
    })


# =====================================================================
# HSN CODE CATALOG & ONLINE RESOLVER
# =====================================================================
HSN_CATALOG = [
    # Dry fruits & Nuts
    (["almond", "badam", "badamgiri", "mamra", "sanora", "california badam", "giri"], "08021200", 5.0, "Almonds / Badam"),
    (["cashew", "kaju", "w240", "w320", "w180", "w400", "tukda", "split", "jhula"], "08013100", 5.0, "Cashew Nuts / Kaju"),
    (["walnut", "akhrot", "akhrot giri", "kashmir akhrot"], "08023200", 5.0, "Walnuts / Akhrot"),
    (["pistachio", "pista", "maghaz", "irani pista"], "08025100", 5.0, "Pistachios / Pista"),
    (["raisin", "raisins", "kishmish", "kismis", "munakka", "sultana"], "08062010", 5.0, "Raisins / Kishmish"),
    (["fig", "figs", "anjeer"], "08042090", 5.0, "Figs / Anjeer"),
    (["fox nut", "foxnut", "makhana", "lotus seed", "phool makhana"], "19041090", 5.0, "Makhana / Fox Nuts"),
    (["cardamom", "elaichi", "elachi", "doda", "chhoti elaichi", "badi elaichi"], "09083100", 5.0, "Cardamom / Elaichi"),
    (["date", "dates", "khajoor", "khajur", "chhohara", "chuhara"], "08041020", 5.0, "Dates / Khajoor"),
    (["apricot", "khubani", "khumani", "dried apricot"], "08131000", 5.0, "Apricots / Khubani"),
    (["chironji", "charoli"], "08029000", 5.0, "Chironji / Charoli"),
    (["melon seed", "magaz", "watermelon seed", "muskmelon seed", "charmagaz"], "12077090", 5.0, "Melon Seeds / Magaz"),
    (["saffron", "kesar", "zafran"], "09102010", 5.0, "Saffron / Kesar"),
    (["prune", "prunes", "aloo bukhara"], "08132000", 5.0, "Prunes / Aloo Bukhara"),
    (["hazelnut", "hazelnuts"], "08022200", 5.0, "Hazelnuts"),
    (["pine nut", "chilgoza", "neoza"], "08029000", 5.0, "Pine Nuts / Chilgoza"),
    (["cranberry", "cranberries"], "08134090", 5.0, "Dried Cranberries"),
    (["blueberry", "blueberries"], "08134090", 5.0, "Dried Blueberries"),
    (["chia seed", "chia seeds"], "12079990", 5.0, "Chia Seeds"),
    (["flax seed", "flaxseed", "alsi"], "12040090", 5.0, "Flax Seeds / Alsi"),
    (["sunflower seed", "sunflower seeds"], "12060090", 5.0, "Sunflower Seeds"),
    (["pumpkin seed", "pumpkin seeds"], "12099990", 5.0, "Pumpkin Seeds"),
    
    # Spices & Condiments
    (["cumin", "jeera", "zeera", "jira"], "09093129", 5.0, "Cumin / Jeera"),
    (["turmeric", "haldi"], "09103020", 5.0, "Turmeric / Haldi"),
    (["chilli", "chili", "mirch", "lal mirch", "red chilli"], "09042110", 5.0, "Chilli / Mirch"),
    (["coriander", "dhaniya", "dhania"], "09092110", 5.0, "Coriander / Dhaniya"),
    (["black pepper", "kali mirch", "pepper", "peppercorn"], "09041110", 5.0, "Black Pepper / Kali Mirch"),
    (["white pepper", "safed mirch"], "09041120", 5.0, "White Pepper"),
    (["clove", "cloves", "laung", "lavang"], "09071010", 5.0, "Cloves / Laung"),
    (["cinnamon", "dalchini", "cassia"], "09061910", 5.0, "Cinnamon / Dalchini"),
    (["fennel", "saunf", "sauf"], "09096111", 5.0, "Fennel / Saunf"),
    (["fenugreek", "methi", "methi dana"], "09109912", 5.0, "Fenugreek / Methi"),
    (["carom", "ajwain", "ajowan"], "09109914", 5.0, "Ajwain / Carom"),
    (["mustard seed", "rai", "sarson"], "12075000", 5.0, "Mustard Seeds / Rai"),
    (["asafoetida", "hing", "heeng"], "13019013", 5.0, "Hing / Asafoetida"),
    (["nutmeg", "jaiphal"], "09081110", 5.0, "Nutmeg / Jaiphal"),
    (["mace", "javitri"], "09082110", 5.0, "Mace / Javitri"),
    (["bay leaf", "tejpatta", "tej patta"], "09109929", 5.0, "Bay Leaf / Tejpatta"),
    (["star anise", "chakraphool", "badiyan"], "09096129", 5.0, "Star Anise / Chakraphool"),
    (["dry ginger", "saunth", "sonth"], "09101120", 5.0, "Dry Ginger / Saunth"),
    (["kalonji", "nigella", "black cumin"], "09093119", 5.0, "Kalonji / Nigella"),
    (["kasuri methi", "kasoori methi"], "07129090", 5.0, "Kasuri Methi"),
    (["poppy seed", "khas khas", "khus khus"], "12079100", 5.0, "Poppy Seeds / Khus Khus"),
    (["amchur", "mango powder", "amchoor"], "08134090", 5.0, "Amchur / Mango Powder"),
    (["spice", "spices", "garam masala", "masala", "curry powder"], "09109100", 5.0, "Mixed Spices / Masala"),

    # Grains & Pulses
    (["rice", "basmati", "chawal", "1121", "pusa"], "10063020", 5.0, "Rice / Basmati"),
    (["wheat", "gehu"], "10019910", 5.0, "Wheat / Gehu"),
    (["atta", "flour", "wheat flour", "chakki atta"], "11010000", 5.0, "Wheat Flour / Atta"),
    (["maida", "all purpose flour"], "11010000", 5.0, "Maida"),
    (["suji", "sooji", "rava", "semolina"], "11031120", 5.0, "Suji / Rava"),
    (["besan", "gram flour"], "11061000", 5.0, "Besan / Gram Flour"),
    (["poha", "flattened rice", "chira"], "19041020", 5.0, "Poha / Flattened Rice"),
    (["chana", "chickpeas", "kabuli chana", "kala chana"], "07132000", 5.0, "Chana / Chickpeas"),
    (["moong", "mung", "green gram"], "07133100", 5.0, "Moong Dal"),
    (["urad", "black gram", "mash"], "07133110", 5.0, "Urad Dal"),
    (["toor", "tur", "arhar", "pigeon pea"], "07136000", 5.0, "Toor / Arhar Dal"),
    (["masoor", "red lentil"], "07134000", 5.0, "Masoor Dal"),
    (["rajma", "kidney bean", "kidney beans"], "07133300", 5.0, "Rajma / Kidney Beans"),
    (["lobia", "black eyed pea", "chawli"], "07133500", 5.0, "Lobia / Cowpea"),
    (["matar", "peas", "dry peas"], "07131000", 5.0, "Dry Peas / Matar"),
    (["soybean", "soya", "soya chunks", "bari"], "12019000", 5.0, "Soybean / Soya Chunks"),

    # Grocery & Staples
    (["sugar", "chini", "shakar", "m30", "s30"], "17019990", 5.0, "Sugar / Chini"),
    (["jaggery", "gud", "gur"], "17011490", 5.0, "Jaggery / Gud"),
    (["tea", "chai", "ctc tea", "green tea"], "09024020", 5.0, "Tea / Chai"),
    (["coffee", "coffee powder", "nescafe", "bru"], "09012110", 5.0, "Coffee"),
    (["salt", "namak", "rock salt", "sendha namak", "kala namak"], "25010010", 0.0, "Salt / Namak"),
    (["ghee", "desi ghee"], "04059020", 12.0, "Ghee / Desi Ghee"),
    (["butter", "makkhan", "amul butter"], "04051000", 12.0, "Butter"),
    (["paneer", "cheese"], "04061000", 12.0, "Paneer / Cheese"),
    (["milk", "doodh"], "04012000", 0.0, "Milk"),
    (["mustard oil", "sarson tel", "kachi ghani"], "15149110", 5.0, "Mustard Oil"),
    (["soyabean oil", "refined oil"], "15079010", 5.0, "Refined Soyabean Oil"),
    (["sunflower oil"], "15121910", 5.0, "Sunflower Oil"),
    (["groundnut oil", "peanut oil"], "15089010", 5.0, "Groundnut Oil"),
    (["coconut oil", "nariyal tel"], "15131900", 5.0, "Coconut Oil"),
    (["sesame oil", "til oil", "til tel"], "15155091", 5.0, "Sesame Oil"),
    (["groundnut", "peanut", "peanuts", "mungfali", "moongphali"], "12024200", 5.0, "Groundnut / Peanuts"),
    (["sesame", "til", "safed til", "kala til"], "12074090", 5.0, "Sesame / Til"),

    # FMCG & Bakery
    (["papad", "papadum"], "19059040", 0.0, "Papad"),
    (["biscuit", "biscuits", "cookie", "cookies"], "19053100", 18.0, "Biscuits / Cookies"),
    (["namkeen", "bhujia", "farsan", "sev", "mixture"], "21069099", 12.0, "Namkeen / Bhujia"),
    (["noodle", "noodles", "maggi", "pasta", "macaroni", "vermicelli", "sewai"], "19021900", 12.0, "Noodles / Pasta / Sewai"),
    (["pickle", "achar", "achaar", "murabba"], "20019000", 12.0, "Pickle / Achar"),
    (["sauce", "ketchup", "tomato ketchup"], "21032000", 12.0, "Sauce / Ketchup"),
    (["chocolate", "chocolates", "cadbury"], "18063100", 18.0, "Chocolates"),
    (["honey", "madhu", "shahad"], "04090000", 5.0, "Honey / Shahad"),

    # Packaging & Materials
    (["bardana", "gunny bag", "gunny bags", "jute bag", "jute bags", "burlap"], "63051000", 5.0, "Gunny Bags / Bardana"),
    (["plastic bag", "poly bag", "polythene", "polybag"], "39232100", 18.0, "Plastic Bags / Poly Bags"),
    (["corrugated box", "carton", "khokha", "box", "carton box"], "48191010", 18.0, "Corrugated Box / Carton"),
    (["tape", "adhesive tape", "cello tape", "packing tape"], "39191000", 18.0, "Adhesive Tape"),

    # Textiles & Clothing
    (["cotton fabric", "cloth", "kapda", "cotton cloth"], "52081100", 5.0, "Cotton Fabric"),
    (["shirt", "t-shirt", "pant", "trousers", "garments", "readymade", "clothes"], "62052000", 5.0, "Garments / Clothes"),
    (["saree", "sari"], "54075240", 5.0, "Saree"),
    (["bedsheet", "bed sheet", "pillow cover", "linen"], "63041910", 12.0, "Bedsheet / Linen"),
    (["shoes", "footwear", "chappal", "sandals", "slippers"], "64041990", 12.0, "Footwear / Shoes"),

    # Hardware, Electrical & Misc
    (["soap", "bathing soap", "toilet soap", "detergent"], "34011110", 18.0, "Soap / Detergent"),
    (["shampoo", "hair oil"], "33051090", 18.0, "Shampoo / Hair Oil"),
    (["iron", "steel", "sariya", "tmt rod", "tmt bar", "tmt sariya"], "72142090", 18.0, "Steel / TMT Sariya"),
    (["cement"], "25232940", 28.0, "Cement"),
    (["pipes", "pvc pipe", "cpvc pipe"], "39172300", 18.0, "Pipes / PVC"),
    (["wire", "cable", "electrical wire", "copper wire"], "85444999", 18.0, "Wire / Cable"),
    (["led bulb", "bulb", "light", "tube light"], "85395000", 12.0, "LED Bulb / Light"),
    (["solar panel", "solar module", "solar cell", "solar"], "85414300", 12.0, "Solar Panels / Cells"),
    (["battery", "lithium battery", "lithium ion", "inverter battery", "lead acid battery"], "85072000", 18.0, "Batteries / Accumulators"),
    (["inverter", "ups"], "85044090", 18.0, "Inverter / UPS"),
    (["switch", "socket", "electrical switch"], "85365090", 18.0, "Electrical Switches / Sockets"),
    (["fan", "ceiling fan", "exhaust fan"], "84145110", 18.0, "Electric Fan"),
    (["bicycle", "cycle", "cycle parts"], "87120010", 12.0, "Bicycle / Cycle"),
    (["leather belt", "belt", "leather goods"], "42033000", 28.0, "Leather Belt / Goods"),
    (["screw", "screws", "nut", "bolt", "nuts", "bolts", "washer"], "73181500", 18.0, "Screws, Nuts & Bolts"),
    (["nail", "nails", "keel"], "73170019", 18.0, "Iron Nails / Keel"),
    (["plywood", "ply", "board"], "44123100", 18.0, "Plywood / Boards"),
    (["paint", "asian paint", "distemper", "primer", "enamel"], "32089090", 18.0, "Paints & Enamels"),
    (["tile", "tiles", "ceramic tiles", "vitrified"], "69072100", 18.0, "Ceramic / Floor Tiles"),
    (["sanitary", "wash basin", "toilet seat", "sink"], "69101000", 18.0, "Sanitary Ware"),

    # Paper & Stationery
    (["paper", "a4 paper", "photocopy paper", "copier paper"], "48025610", 12.0, "A4 Paper"),
    (["notebook", "register", "copy"], "48202000", 12.0, "Notebook / Register"),
    (["pen", "ball pen", "gel pen"], "96081019", 18.0, "Ballpoint Pen"),

    # Agricultural, Feeds & Beverages
    (["khal", "churi", "oil cake", "cattle feed", "animal feed"], "23069090", 5.0, "Oil Cake / Cattle Feed / Khal"),
    (["fertilizer", "urea", "dap", "khad"], "31021000", 5.0, "Fertilizer / Urea / DAP"),
    (["pesticide", "insecticide", "herbicide"], "38089199", 18.0, "Pesticides / Insecticides"),
    (["seeds", "crop seeds", "hybrid seeds"], "12099990", 0.0, "Crop Seeds / Sowing Seeds"),
    (["utensils", "bartan", "stainless steel bartan", "cookware"], "73239390", 12.0, "Stainless Steel Utensils"),
    (["cold drink", "pepsi", "coca cola", "soda", "soft drink"], "22021010", 28.0, "Aerated Soft Drinks"),
    (["juice", "fruit juice", "real juice"], "20098990", 12.0, "Fruit Juices"),
    (["mineral water", "packaged drinking water", "bisleri"], "22011010", 18.0, "Packaged Drinking Water"),
    (["candy", "toffee", "confectionery", "chewing gum"], "17049020", 18.0, "Sugar Confectionery / Candies"),
]


def resolve_hsn_code(query):
    """
    Multi-tier HSN resolver:
    1. Check SQLite products database (exact / substring match)
    2. Check comprehensive HSN_CATALOG (token & keyword match)
    3. Query online web search (Brave / DuckDuckGo / web search)
    """
    q = (query or "").strip()
    if not q:
        return None

    # 1. Database check
    try:
        conn = get_db()
        cur = conn.cursor()
        # exact match
        row = cur.execute("SELECT hsn, gst, name FROM products WHERE LOWER(name) = ? AND hsn != ''", (q.lower(),)).fetchone()
        if row and row["hsn"]:
            conn.close()
            return {"hsn": row["hsn"], "gst": float(row["gst"] or 5.0), "source": "Saved Products", "name": row["name"]}

        # contains match
        row = cur.execute("SELECT hsn, gst, name FROM products WHERE (LOWER(name) LIKE ? OR ? LIKE '%' || LOWER(name) || '%') AND hsn != '' LIMIT 1", (f"%{q.lower()}%", q.lower())).fetchone()
        conn.close()
        if row and row["hsn"]:
            return {"hsn": row["hsn"], "gst": float(row["gst"] or 5.0), "source": "Saved Products", "name": row["name"]}
    except Exception:
        pass

    # 2. Comprehensive catalog check
    import re
    q_lower = q.lower()
    words = [w for w in re.sub(r'[^a-zA-Z0-9]', ' ', q_lower).split() if len(w) > 1]
    best_match = None
    max_score = 0
    for keywords, hsn, gst, title in HSN_CATALOG:
        score = 0
        for kw in keywords:
            if kw == q_lower:
                score += 30
            elif " " in kw and kw in q_lower:
                score += 15 + len(kw)
            elif kw in words:
                score += 12 + len(kw)
            elif any(w.startswith(kw) for w in words if len(kw) >= 4):
                score += 5
        if score > max_score:
            max_score = score
            best_match = (hsn, gst, title)

    if best_match and max_score >= 4:
        return {"hsn": best_match[0], "gst": best_match[1], "source": "GST Database", "name": best_match[2]}

    # 3. Live online web fetch
    import urllib.request
    import urllib.parse
    search_engines = [
        ("https://search.brave.com/search?q=", " gst hsn code india"),
        ("https://lite.duckduckgo.com/lite/?q=", " gst hsn code"),
    ]
    for base_url, suffix in search_engines:
        try:
            full_url = base_url + urllib.parse.quote(q + suffix)
            req = urllib.request.Request(full_url, headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
            })
            with urllib.request.urlopen(req, timeout=5) as resp:
                html = resp.read().decode('utf-8', errors='ignore')
                codes = re.findall(r'(?:hsn|chapter)\s*(?:code)?\s*[:\-]?\s*(\d{4,8})', html, re.I)
                codes += re.findall(r'hsn-code-(\d{4,8})', html, re.I)
                codes += re.findall(r'\b(\d{8})\b', html)
                if codes:
                    valid = [c for c in codes if len(c) in (4, 6, 8) and not c.startswith("202") and not c.startswith("19")]
                    if not valid:
                        valid = [c for c in codes if len(c) in (4, 6, 8)]
                    if valid:
                        valid.sort(key=lambda x: len(x), reverse=True)
                        hsn_found = valid[0]
                        return {"hsn": hsn_found, "gst": 5.0, "source": "Online Web Search", "name": q}
        except Exception:
            continue

    return None


@app.route("/api/products")
def get_products():
    conn = get_db()
    rows = conn.execute("SELECT id, name, hsn, gst FROM products ORDER BY name COLLATE NOCASE ASC").fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/products/save", methods=["POST"])
def save_product():
    data = request.json or {}
    name = (data.get("name") or "").strip()
    hsn = (data.get("hsn") or "").strip()
    gst = float(data.get("gst") or 5.0)

    if not name:
        return jsonify({"success": False, "error": "Product Name is required"}), 400

    conn = get_db()
    cur = conn.cursor()
    existing = cur.execute("SELECT id, name FROM products WHERE LOWER(name)=?", (name.lower(),)).fetchone()
    if existing:
        cur.execute("UPDATE products SET hsn=?, gst=? WHERE id=?", (hsn, gst, existing["id"]))
        p_id = existing["id"]
        msg = f"Product '{name}' updated successfully!"
    else:
        cur.execute("INSERT INTO products (name, hsn, gst) VALUES (?, ?, ?)", (name, hsn, gst))
        p_id = cur.lastrowid
        msg = f"Product '{name}' added to dropdown permanently!"

    conn.commit()
    conn.close()

    return jsonify({
        "success": True,
        "message": msg,
        "product": {"id": p_id, "name": name, "hsn": hsn, "gst": gst}
    })


@app.route("/api/lookup_hsn", methods=["GET"])
def lookup_hsn_api():
    name = (request.args.get("name") or "").strip()
    if not name:
        return jsonify({"success": False, "error": "Product name is required"}), 400

    result = resolve_hsn_code(name)
    if result:
        return jsonify({
            "success": True,
            "hsn": result["hsn"],
            "gst": result.get("gst", 5.0),
            "source": result.get("source", "catalog"),
            "matched_name": result.get("name", name)
        })
    else:
        return jsonify({
            "success": False,
            "message": "No HSN found automatically. You can enter it manually."
        })


@app.route("/api/customers/upload", methods=["POST"])
def upload_customers():
    if "file" not in request.files:
        return jsonify({"success": False, "error": "No file provided"}), 400
    file = request.files["file"]
    if not file or not file.filename:
        return jsonify({"success": False, "error": "No file selected"}), 400

    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in [".csv", ".xlsx"]:
        return jsonify({"success": False, "error": "Only .csv or .xlsx files are supported"}), 400

    base_dir = os.path.dirname(os.path.abspath(__file__))
    saved_path = os.path.join(base_dir, f"customers_imported{ext}")
    file.save(saved_path)
    count = sync_customers_from_file(saved_path)
    return jsonify({
        "success": True,
        "count": count,
        "message": f"Successfully loaded {count} customers!",
    })


@app.route("/save_invoice", methods=["POST"])
def save_invoice():
    data = request.json or {}

    items = data.get("items", [])

    subtotal = 0.0
    gst_total = 0.0
    total_qty = 0.0
    net_weight = 0.0

    cleaned_items = []
    for item in items:
        qty = float(item.get("qty") or 0)
        weight = float(item.get("weight") or 0 )
        rate = float(item.get("rate") or 0)
        gst = float(item.get("gst") or 0)

        amount = weight * rate
        tax_amount = amount * gst / 100

        subtotal += amount
        gst_total += tax_amount
        total_qty += qty
        net_weight += weight

        cleaned_items.append({
            "product_name": item.get("product_name", ""),
            "hsn_code": item.get("hsn_code", ""),
            "qty": qty,
            "qty_unit": item.get("qty_unit", ""),
            "weight": weight,
            "rate": rate,
            "gst": gst,
            "amount": amount,
            "tax_amount": tax_amount,
        })

    labour = float(data.get("labour_amount") or 0)
    grand_total = subtotal + gst_total + labour

    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
    INSERT INTO invoices(
        invoice_no, bill_date, delivery_type, customer_name, customer_address,
        customer_gstin, state_code, transporter, marka, labour_amount, remarks,
        subtotal, gst_total, grand_total, net_weight, total_qty
    )
    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (
        data.get("invoice_no", ""),
        data.get("bill_date", ""),
        data.get("delivery_type", ""),
        data.get("customer_name", ""),
        data.get("customer_address", ""),
        data.get("customer_gstin", ""),
        data.get("state_code", ""),
        data.get("transporter", ""),
        data.get("marka", ""),
        labour,
        data.get("remarks", ""),
        subtotal,
        gst_total,
        grand_total,
        net_weight,
        total_qty,
    ))

    invoice_id = cur.lastrowid

    # Auto-save or update customer in customers list permanently
    cust_n = (data.get("customer_name") or "").strip()
    cust_a = (data.get("customer_address") or "").strip()
    cust_g = (data.get("customer_gstin") or "").strip().upper()
    if cust_n:
        try:
            cur.execute("""
            INSERT INTO customers(name, address, gstin)
            VALUES(?,?,?)
            ON CONFLICT(name) DO UPDATE SET
                address=excluded.address,
                gstin=excluded.gstin
            """, (cust_n, cust_a, cust_g))
        except Exception:
            pass

    # Auto-save transporter in transporters list permanently
    transp_n = (data.get("transporter") or "").strip()
    if transp_n:
        try:
            cur.execute("""
            INSERT INTO transporters(name)
            VALUES(?)
            ON CONFLICT(name) DO NOTHING
            """, (transp_n,))
        except Exception:
            pass

    # Auto-save marka in markas list permanently
    marka_n = (data.get("marka") or "").strip()
    if marka_n:
        try:
            cur.execute("""
            INSERT INTO markas(name)
            VALUES(?)
            ON CONFLICT(name) DO NOTHING
            """, (marka_n,))
        except Exception:
            pass

    # Auto-save all products in products table permanently
    for item in cleaned_items:
        p_name = (item.get("product_name") or "").strip()
        p_hsn = (item.get("hsn_code") or "").strip()
        p_gst = float(item.get("gst") or 5.0)
        if p_name:
            try:
                cur.execute("""
                INSERT INTO products(name, hsn, gst)
                VALUES(?,?,?)
                ON CONFLICT(name) DO UPDATE SET
                    hsn = CASE WHEN excluded.hsn != '' THEN excluded.hsn ELSE products.hsn END,
                    gst = excluded.gst
                """, (p_name, p_hsn, p_gst))
            except Exception:
                pass

    for item in cleaned_items:
        cur.execute("""
        INSERT INTO invoice_items(
            invoice_id, product_name, hsn_code, qty, qty_unit, weight, rate,
            gst, amount, tax_amount
        )
        VALUES(?,?,?,?,?,?,?,?,?,?)
        """, (
            invoice_id,
            item["product_name"],
            item["hsn_code"],
            item["qty"],
            item["qty_unit"],
            item["weight"],
            item["rate"],
            item["gst"],
            item["amount"],
            item["tax_amount"],
        ))

    conn.commit()
    conn.close()

    return jsonify({
        "success": True,
        "invoice_id": invoice_id,
        "subtotal": round(subtotal, 2),
        "gst_total": round(gst_total, 2),
        "total": round(grand_total, 2),
        "pdf_url": f"/generate_pdf/{invoice_id}",
        "message": "Invoice Saved Successfully",
    })


@app.route("/invoices")
def invoices():
    conn = get_db()
    rows = conn.execute("""
        SELECT * FROM invoices ORDER BY id DESC
    """).fetchall()
    conn.close()
    return jsonify([dict(row) for row in rows])


@app.route("/invoice/<int:invoice_id>")
def invoice_detail(invoice_id):
    conn = get_db()
    invoice = conn.execute(
        "SELECT * FROM invoices WHERE id=?", (invoice_id,)
    ).fetchone()
    if invoice is None:
        conn.close()
        abort(404)
    items = conn.execute(
        "SELECT * FROM invoice_items WHERE invoice_id=?", (invoice_id,)
    ).fetchall()
    conn.close()
    result = dict(invoice)
    result["items"] = [dict(i) for i in items]
    return jsonify(result)


@app.route("/delete_invoice/<int:invoice_id>", methods=["POST", "DELETE"])
def delete_invoice(invoice_id):
    conn = get_db()
    conn.execute("DELETE FROM invoice_items WHERE invoice_id=?", (invoice_id,))
    conn.execute("DELETE FROM invoices WHERE id=?", (invoice_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


# =====================================================================
# PDF GENERATION - matches the Sri Radhe Enterprises invoice format
# =====================================================================
HAS_TTF = False
for p in [r'C:\Windows\Fonts\arial.ttf', r'C:\Windows\Fonts\calibri.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf']:
    if os.path.exists(p):
        try:
            pdfmetrics.registerFont(TTFont('CustomSans', p))
            bold_p = p.replace('arial.ttf', 'arialbd.ttf').replace('calibri.ttf', 'calibrib.ttf')
            if os.path.exists(bold_p):
                pdfmetrics.registerFont(TTFont('CustomSans-Bold', bold_p))
            else:
                pdfmetrics.registerFont(TTFont('CustomSans-Bold', p))
            HAS_TTF = True
            break
        except Exception:
            pass

FONT_REG = 'CustomSans' if HAS_TTF else 'Helvetica'
FONT_BOLD = 'CustomSans-Bold' if HAS_TTF else 'Helvetica-Bold'
SYM = '₹' if HAS_TTF else 'Rs. '


def xml_esc(val):
    if val is None:
        return ""
    return html.escape(str(val))


def format_date(raw_date):
    if not raw_date:
        return datetime.now().strftime("%d-%b-%Y")
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            dt = datetime.strptime(raw_date.strip(), fmt)
            return dt.strftime("%d-%b-%Y")
        except Exception:
            pass
    return str(raw_date)


def draw_border(canvas, doc):
    canvas.saveState()
    canvas.setLineWidth(0.75)
    canvas.setStrokeColor(colors.black)
    m = 5 * mm
    canvas.rect(m, m, 210 * mm - 2 * m, 297 * mm - 2 * m)
    canvas.restoreState()


@app.route("/generate_pdf/<int:invoice_id>")
def generate_pdf(invoice_id):
    conn = get_db()
    invoice = conn.execute(
        "SELECT * FROM invoices WHERE id=?", (invoice_id,)
    ).fetchone()

    if invoice is None:
        conn.close()
        return "Invoice not found", 404

    items = conn.execute(
        "SELECT * FROM invoice_items WHERE invoice_id=?", (invoice_id,)
    ).fetchall()
    conn.close()

    pdf_path = os.path.join(PDF_DIR, f"invoice_{invoice_id}.pdf")

    doc = SimpleDocTemplate(
        pdf_path, pagesize=A4,
        leftMargin=8 * mm, rightMargin=8 * mm,
        topMargin=8 * mm, bottomMargin=8 * mm,
    )

    W = doc.width
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('CompTitle', parent=styles['Normal'], alignment=TA_CENTER, fontName='Times-Bold', fontSize=19, leading=23)
    sub_style = ParagraphStyle('CompSub', parent=styles['Normal'], alignment=TA_CENTER, fontName=FONT_REG, fontSize=7.5, leading=10)
    bold_sub = ParagraphStyle('CompBoldSub', parent=styles['Normal'], alignment=TA_CENTER, fontName=FONT_BOLD, fontSize=8.5, leading=11)
    inv_hdr = ParagraphStyle('InvHdr', parent=styles['Normal'], alignment=TA_CENTER, fontName='Times-Bold', fontSize=15, leading=18)

    cell_s = ParagraphStyle('CellS', parent=styles['Normal'], fontName=FONT_REG, fontSize=7.5, leading=9.5)
    cell_s_center = ParagraphStyle('CellSCenter', parent=styles['Normal'], alignment=TA_CENTER, fontName=FONT_REG, fontSize=7.5, leading=9.5)
    cell_b = ParagraphStyle('CellB', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=8.5, leading=10.5)
    cell_b_center = ParagraphStyle('CellBCenter', parent=styles['Normal'], alignment=TA_CENTER, fontName=FONT_BOLD, fontSize=8.5, leading=10.5)
    grand_total_style = ParagraphStyle('GrandTotal', parent=styles['Normal'], alignment=TA_CENTER, fontName=FONT_BOLD, fontSize=13, leading=15)

    story = []

    # 1. HEADER BOX WITH RADHA KRISHNA LOGOS
    base_dir = os.path.dirname(os.path.abspath(__file__))
    logo_file = os.path.join(base_dir, "logo.png")
    if os.path.exists(logo_file):
        logo_img1 = RLImage(logo_file, width=42, height=58)
        logo_img2 = RLImage(logo_file, width=42, height=58)
    else:
        logo_img1 = ""
        logo_img2 = ""

    header_text = [
        Paragraph(COMPANY['name'], title_style),
        Spacer(1, 2),
        Paragraph(COMPANY['tagline'], sub_style),
        Paragraph(f"GSTIN: {COMPANY['gstin']}", bold_sub),
        Paragraph(COMPANY['address_line1'], sub_style),
        Paragraph(COMPANY['address_line2'], sub_style),
        Paragraph(f"Mobile: {COMPANY['mobile']}", bold_sub),
    ]

    hdr_table_data = [
        [logo_img1, header_text, logo_img2],
        [Paragraph('INVOICE', inv_hdr), '', '']
    ]
    hdr_table = Table(hdr_table_data, colWidths=[52, W - 104, 52])
    hdr_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.75, colors.black),
        ('SPAN', (0, 1), (2, 1)),
        ('LINEABOVE', (0, 1), (2, 1), 0.75, colors.black),
        ('ALIGN', (0, 0), (0, 0), 'CENTER'),
        ('ALIGN', (2, 0), (2, 0), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, 0), 4),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
        ('TOPPADDING', (0, 1), (-1, 1), 3),
        ('BOTTOMPADDING', (0, 1), (-1, 1), 3),
    ]))
    story.append(hdr_table)
    story.append(Spacer(1, 4))

    # 2. DATE ROW & DELIVERY BOX
    deliv = (invoice["delivery_type"] or "").strip()
    deliv_label = f"({deliv.upper()})" if deliv else "(HOME DELIVERY)"

    date_table_data = [
        [
            Paragraph('Date:', cell_b),
            Paragraph(format_date(invoice['bill_date']), cell_b_center),
            '',
            Paragraph(deliv_label, cell_b_center)
        ]
    ]
    date_table = Table(date_table_data, colWidths=[35, 140, W - 345, 170])
    date_table.setStyle(TableStyle([
        ('BOX', (0, 0), (1, 0), 0.75, colors.black),
        ('INNERGRID', (0, 0), (1, 0), 0.75, colors.black),
        ('BOX', (3, 0), (3, 0), 0.75, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(date_table)
    story.append(Spacer(1, 4))

    # 3. CUSTOMER INFO GRID - 4 columns
    c0, c1, c2, c3 = W * 0.38, W * 0.31, W * 0.17, W * 0.14

    state_code = (invoice["state_code"] or "").strip()
    cust_gstin = (invoice["customer_gstin"] or "").strip()
    if not state_code and cust_gstin and len(cust_gstin) >= 2 and cust_gstin[:2].isdigit():
        state_code = str(int(cust_gstin[:2]))

    addr_full = (invoice["customer_address"] or "").strip()
    addr_lines = [l.strip() for l in addr_full.split("\n") if l.strip()]
    if len(addr_lines) >= 2:
        addr1 = addr_lines[0]
        addr2 = ", ".join(addr_lines[1:])
    elif len(addr_lines) == 1:
        parts = [p.strip() for p in addr_lines[0].split(",") if p.strip()]
        if len(parts) >= 2:
            addr1 = ", ".join(parts[:-1])
            addr2 = parts[-1]
        else:
            addr1 = addr_lines[0]
            addr2 = ""
    else:
        addr1 = ""
        addr2 = ""

    grand_total_num = float(invoice["grand_total"] or 0)
    bill_amt_str = f"{SYM}{round(grand_total_num):,}"

    net_wt_val = invoice["net_weight"] or 0
    tot_qty_val = invoice["total_qty"] or 0
    net_wt_str = f"{net_wt_val:g}" if isinstance(net_wt_val, (int, float)) else str(net_wt_val)
    tot_qty_str = f"{tot_qty_val:g}" if isinstance(tot_qty_val, (int, float)) else str(tot_qty_val)

    info_data = [
        [Paragraph('Invoice for', cell_s_center), Paragraph('Invoice No.', cell_s_center), Paragraph('Transporter Details', cell_s_center), ''],
        [Paragraph(xml_esc(invoice['customer_name'] or ''), cell_b_center), Paragraph(xml_esc(str(invoice['invoice_no'] or invoice_id)), cell_b_center), Paragraph(xml_esc(invoice['transporter'] or ''), cell_b_center), ''],
        [Paragraph(xml_esc(addr1), cell_s_center), Paragraph('State Code', cell_s_center), Paragraph('Marka', cell_s_center), ''],
        [Paragraph(xml_esc(addr2), cell_b_center), Paragraph(xml_esc(state_code), cell_b_center), Paragraph(xml_esc(invoice['marka'] or ''), cell_b_center), ''],
        [Paragraph('GSTIN', cell_s_center), Paragraph('Bill Amount', cell_s_center), Paragraph('Net Weight in K.G', cell_s_center), Paragraph('Total Qty', cell_s_center)],
        [Paragraph(xml_esc(cust_gstin), cell_b_center), Paragraph(bill_amt_str, cell_b_center), Paragraph(net_wt_str, cell_b_center), Paragraph(tot_qty_str, cell_b_center)]
    ]

    info_tbl = Table(info_data, colWidths=[c0, c1, c2, c3])
    info_tbl.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.75, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.75, colors.black),
        ('SPAN', (2, 0), (3, 0)),
        ('SPAN', (2, 1), (3, 1)),
        ('SPAN', (2, 2), (3, 2)),
        ('SPAN', (2, 3), (3, 3)),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 2),
        ('RIGHTPADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(info_tbl)
    story.append(Spacer(1, 4))

    # 4. ITEMS TABLE (15 rows padded)
    items_header = [
        Paragraph('<b>Sr.<br/>No.</b>', cell_s_center),
        Paragraph('<b>Description of<br/>Goods</b>', cell_s_center),
        Paragraph('<b>HSN/<br/>ACS Code</b>', cell_s_center),
        Paragraph('<b>Qty</b>', cell_s_center),
        Paragraph('<b>Net Weight/<br/>Qty. in K.G</b>', cell_s_center),
        Paragraph('<b>Rate/<br/>K.G</b>', cell_s_center),
        Paragraph('<b>Amount<br/>in Rs.</b>', cell_s_center),
        Paragraph('<b>Tax%</b>', cell_s_center),
        Paragraph('<b>Tax<br/>Amount</b>', cell_s_center),
    ]

    col_w = [
        26,
        W * 0.26,
        55,
        46,
        52,
        48,
        66,
        30,
        W - (26 + W * 0.26 + 55 + 46 + 52 + 48 + 66 + 30)
    ]

    items_data = [items_header]
    calc_total_qty = 0.0
    calc_total_wt = 0.0
    calc_total_amt = 0.0
    calc_total_tax = 0.0

    num_rows = max(15, len(items))
    for i in range(1, num_rows + 1):
        if i <= len(items):
            it = items[i - 1]
            q_val = it['qty'] or 0
            w_val = it['weight'] or 0
            r_val = it['rate'] or 0
            a_val = it['amount'] or 0
            g_val = it['gst'] or 0
            t_val = it['tax_amount'] or 0

            calc_total_qty += q_val
            calc_total_wt += w_val
            calc_total_amt += a_val
            calc_total_tax += t_val

            unit_str = f" {it['qty_unit']}" if it['qty_unit'] else ""
            qty_text = f"{q_val:g}{unit_str}"

            items_data.append([
                Paragraph(str(i), cell_s_center),
                Paragraph(xml_esc(it['product_name'] or ''), cell_s),
                Paragraph(xml_esc(it['hsn_code'] or ''), cell_s_center),
                Paragraph(xml_esc(qty_text), cell_s_center),
                Paragraph(f"{w_val:g}", cell_s_center),
                Paragraph(f"{SYM}{r_val:,.2f}", cell_s_center),
                Paragraph(f"{SYM}{a_val:,.2f}", cell_s_center),
                Paragraph(f"{g_val:g}", cell_s_center),
                Paragraph(f"{SYM}{t_val:,.2f}", cell_s_center),
            ])
        else:
            items_data.append([
                Paragraph(str(i), cell_s_center),
                '', '', '', '', '', '', '', ''
            ])

    display_qty = tot_qty_val if tot_qty_val else calc_total_qty
    display_wt = net_wt_val if net_wt_val else calc_total_wt
    items_data.append([
        Paragraph('<b>Total</b>', cell_b_center),
        '', '',
        Paragraph(f'<b>{display_qty:g}</b>', cell_b_center),
        Paragraph(f'<b>{display_wt:g}</b>', cell_b_center),
        '',
        Paragraph(f'<b>{SYM}{calc_total_amt:,.2f}</b>', cell_b_center),
        '',
        Paragraph(f'<b>{SYM}{calc_total_tax:,.2f}</b>', cell_b_center),
    ])

    items_table = Table(items_data, colWidths=col_w)
    items_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.75, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('LEFTPADDING', (0,0), (-1,-1), 2),
        ('RIGHTPADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(items_table)
    story.append(Spacer(1, 4))

    # 5. BOTTOM DETAILS: BANK, REMARKS, TOTALS
    bank_text = [
        Paragraph('<b>Bank Details:</b>', cell_s),
        Paragraph(f"<b>{COMPANY['bank_name']}</b>", cell_b),
        Paragraph(COMPANY['bank_branch'], cell_s),
        Paragraph(f"A/c No.- {COMPANY['bank_account']}", cell_s),
        Paragraph(f"IFSC Code: <b>{COMPANY['bank_ifsc']}</b>", cell_s),
    ]

    custom_remark = (invoice["remarks"] or "").strip()
    if not custom_remark:
        custom_remark = f"MARK {invoice['marka'] or ''}".strip()
    remarks_text = [
        Paragraph('<b>Remarks:-</b>', cell_s),
        Paragraph(f"<b>{xml_esc(custom_remark)}</b>", cell_b),
    ]

    # Subtotal before labour bardana
    subtotal_with_tax = float(invoice["subtotal"] or calc_total_amt) + float(invoice["gst_total"] or calc_total_tax)
    labour_amount = float(invoice["labour_amount"] or 0)
    labour_str = f"{SYM}{labour_amount:,.2f}" if labour_amount > 0 else ""

    totals_subtable = Table([
        [Paragraph('Subtotal', cell_s_center), Paragraph(f'{SYM}{subtotal_with_tax:,.2f}', cell_b_center)],
        [Paragraph('Labour Bardana', cell_s_center), Paragraph(labour_str, cell_b_center)],
        [Paragraph('<b>Total Amount in Rs:</b><br/><font size=6.5>(Rounded off)</font>', cell_s_center),
         Paragraph(f'{SYM}{round(grand_total_num):,}', grand_total_style)]
    ], colWidths=[W * 0.22, W * 0.16])
    totals_subtable.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.75, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ]))

    bottom_data = [
        [bank_text, remarks_text, totals_subtable]
    ]
    bottom_tbl = Table(bottom_data, colWidths=[W * 0.34, W * 0.28, W * 0.38])
    bottom_tbl.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.75, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.75, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(bottom_tbl)
    story.append(Spacer(1, 4))

    # 6. SIGNATURE & JURISDICTION
    sig_file = os.path.join(base_dir, "signature.png")
    if os.path.exists(sig_file):
        sig_img = RLImage(sig_file, width=50, height=25)
    else:
        sig_img = Paragraph("", cell_s)

    sig_data = [
        ['', Paragraph(f'<b>For {COMPANY["name"]}</b>', cell_b_center)],
        ['', sig_img],
        ['', Paragraph('<font size=7>Authorised Signatory</font>', cell_s_center)],
        [Paragraph(f'<font size=6.5><i>{COMPANY["jurisdiction"]}</i></font>', cell_s), '']
    ]
    sig_tbl = Table(sig_data, colWidths=[W - 140, 140])
    sig_tbl.setStyle(TableStyle([
        ('ALIGN', (1,0), (1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 1),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1),
        ('LINEABOVE', (0, 3), (0, 3), 0.5, colors.black),
    ]))
    story.append(sig_tbl)

    doc.build(story, onFirstPage=draw_border)

    return send_file(
        os.path.abspath(pdf_path),
        download_name=f"Invoice_{invoice['invoice_no'] or invoice_id}.pdf",
        as_attachment=False,
        mimetype="application/pdf",
    ) 


def _print_network_info():
    local_ip = get_local_ip()
    print("=" * 60)
    print(" Sri Radhe Enterprises - Billing System is running!")
    print("=" * 60)
    print(f" On THIS computer, open:            http://127.0.0.1:{PORT}")
    print(f" On your PHONE (same WiFi), open:   http://{local_ip}:{PORT}")
    print("=" * 60)


if __name__ == "__main__":
    _print_network_info()
    # host="0.0.0.0" makes the app reachable from other devices (like a
    # phone) on the same WiFi network, not just this computer.
    app.run(host="0.0.0.0", port=PORT, debug=False)
