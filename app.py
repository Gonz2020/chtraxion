# -*- coding: utf-8 -*-
"""
Plataforma Integral de Capital Humano Traxion - Fase 2.7 a 3.0
De registro a laborar: expediente digital, capacitación/DC3, firma electrónica,
integración Access Kiosk, dashboard corporativo, IA avanzada y multi UDN.
"""
from __future__ import annotations
import os, json, sqlite3
from functools import wraps
from pathlib import Path
from datetime import datetime, date
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file, session
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
UPLOAD_DIR = APP_DIR / "uploads"
REPORT_DIR = APP_DIR / "reports"
LOG_DIR = APP_DIR / "logs"
for d in [DATA_DIR, UPLOAD_DIR, REPORT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

DB_PATH = DATA_DIR / "capital_humano_traxion.db"
EXCEL_BASE = DATA_DIR / "DESARROLLO AT - SISTEMA.xlsx"

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "traxion_ch_fase3_admin_seguridad")

ADMIN_USER = os.environ.get("ADMIN_USER", "Administrador")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "Tequila2026")
PUBLIC_ENDPOINTS = {"login", "health", "static"}

ETAPAS = [
    "Registro", "Entrevista", "Oferta Laboral", "Servicio Medico",
    "Evaluacion Tecnica", "Documentacion", "Induccion", "Entrega Operaciones", "Activo"
]
ESTATUS_MACRO = {
    "Registro": "Reclutamiento",
    "Entrevista": "Reclutamiento",
    "Oferta Laboral": "Reclutamiento",
    "Servicio Medico": "Reclutamiento",
    "Evaluacion Tecnica": "Reclutamiento",
    "Documentacion": "Contratacion",
    "Induccion": "Contratacion",
    "Entrega Operaciones": "Contratacion",
    "Activo": "Activo",
}
CHECK_OFERTA = ["Sueldo", "Prestaciones", "Bonos", "Horario", "Lugar de trabajo", "Fecha primer pago"]
CHECK_DOCS = ["INE", "CURP", "RFC / Constancia fiscal", "NSS", "Cuenta bancaria", "Comprobante domicilio", "Acta nacimiento", "Solicitud/CV"]
RESULTADOS = ["Pendiente", "Apto", "Apto Condicionado", "No Apto"]

ENTREVISTA_BASE = {
    "MONTACARGUISTA": [
        "¿Cuánto tiempo de experiencia tiene operando montacargas?",
        "¿Qué tipo de montacargas ha operado?",
        "¿Cuenta con licencia o constancia de manejo de montacargas?",
        "¿Ha trabajado con WMS, SAP o handheld?",
        "¿Ha tenido incidentes o accidentes operando equipo?",
        "¿Está disponible para rolar turnos y trabajar tiempo extra?"
    ],
    "PATINERO": [
        "¿Cuánto tiempo de experiencia tiene usando patín hidráulico o eléctrico?",
        "¿Ha realizado surtido, recibo o embarques?",
        "¿Ha trabajado con handheld, WMS o SAP?",
        "¿Puede realizar actividades de carga, descarga y acomodo?",
        "¿Está disponible para rolar turnos?"
    ],
    "GENERAL": [
        "Cuéntame brevemente su experiencia laboral reciente.",
        "¿Por qué le interesa la vacante?",
        "¿Tiene disponibilidad de horario?",
        "¿Cuenta con documentación completa?",
        "¿Tiene experiencia en almacén o logística?"
    ]
}

INDUCCION_CHECKLIST = ["Inducción general", "Reglamento interno", "Seguridad y EPP", "Recorrido operativo", "Entrega a jefe directo", "Alta biométrica", "Gafete", "Uniforme/EPP"]

SLA_DIAS = {
    "Registro": 1,
    "Entrevista": 1,
    "Oferta Laboral": 1,
    "Servicio Medico": 1,
    "Evaluacion Tecnica": 1,
    "Documentacion": 2,
    "Induccion": 1,
    "Entrega Operaciones": 1,
    "Activo": 0,
}


def now_str(): return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
def today_str(): return date.today().strftime("%Y-%m-%d")

def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def execute(sql, params=()):
    with get_conn() as conn:
        cur = conn.cursor(); cur.execute(sql, params); conn.commit(); return cur.lastrowid

def query(sql, params=(), one=False):
    with get_conn() as conn:
        cur = conn.cursor(); cur.execute(sql, params); rows = cur.fetchall()
        return (rows[0] if rows else None) if one else rows

def clean(v):
    if pd.isna(v): return ""
    s = str(v).strip()
    if s.lower() == "nan": return ""
    if s.endswith(".0") and s[:-2].isdigit(): s = s[:-2]
    return s

def column_exists(conn, table, column):
    return any(r["name"] == column for r in conn.execute(f"PRAGMA table_info({table})").fetchall())

def ensure_column(conn, table, column, definition):
    if not column_exists(conn, table, column):
        try: conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")
        except Exception: pass


def init_db():
    print("[1/4] Inicializando base de datos...")
    with get_conn() as conn:
        c = conn.cursor()
        c.execute("""
        CREATE TABLE IF NOT EXISTS candidatos(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            folio TEXT UNIQUE,
            nombre TEXT NOT NULL,
            telefono TEXT,
            email TEXT,
            region TEXT,
            localidad TEXT,
            udn TEXT,
            puesto TEXT,
            reclutador TEXT,
            etapa TEXT DEFAULT 'Registro',
            estatus_macro TEXT DEFAULT 'Reclutamiento',
            fecha_entrevista TEXT,
            fecha_registro TEXT,
            fecha_actualizacion TEXT,
            motivo_rechazo TEXT,
            comentarios TEXT
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS hoja_ruta(
            candidato_id INTEGER PRIMARY KEY,
            ruta_transporte TEXT, parada TEXT,
            experiencia_json TEXT, handheld TEXT, wms TEXT, sap TEXT, excel TEXT,
            patin TEXT, montacargas TEXT,
            reingreso_traxion TEXT, reingreso_cliente TEXT,
            emergencia_nombre TEXT, emergencia_parentesco TEXT, emergencia_telefono TEXT,
            updated_at TEXT
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS oferta_laboral(
            candidato_id INTEGER PRIMARY KEY,
            sueldo TEXT, prestaciones TEXT, bonos TEXT, horario TEXT, lugar_trabajo TEXT,
            fecha_primer_pago TEXT, checklist_json TEXT, updated_at TEXT
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS evaluaciones(
            candidato_id INTEGER PRIMARY KEY,
            medico_resultado TEXT, medico_motivo TEXT,
            tecnico_resultado TEXT, tecnico_tipo TEXT, tecnico_motivo TEXT,
            rfc_actualizado TEXT DEFAULT 'No', costo_rfc REAL DEFAULT 0,
            updated_at TEXT
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS documentacion(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            candidato_id INTEGER,
            documento TEXT, estado TEXT DEFAULT 'Pendiente', archivo TEXT, updated_at TEXT
        )""")
        c.execute("CREATE TABLE IF NOT EXISTS catalogo_puestos(id INTEGER PRIMARY KEY AUTOINCREMENT, puesto TEXT UNIQUE, perfil TEXT DEFAULT 'OPERATIVO', activo INTEGER DEFAULT 1)")
        c.execute("CREATE TABLE IF NOT EXISTS catalogo_udn(id INTEGER PRIMARY KEY AUTOINCREMENT, almacen TEXT UNIQUE, region TEXT, localidad TEXT, activo INTEGER DEFAULT 1)")
        c.execute("CREATE TABLE IF NOT EXISTS catalogo_reclutadores(id INTEGER PRIMARY KEY AUTOINCREMENT, reclutador TEXT UNIQUE, region TEXT, activo INTEGER DEFAULT 1)")
        c.execute("CREATE TABLE IF NOT EXISTS headcount(id INTEGER PRIMARY KEY AUTOINCREMENT, udn TEXT, puesto TEXT, requerido INTEGER DEFAULT 0, activo INTEGER DEFAULT 0)")
        c.execute("CREATE TABLE IF NOT EXISTS actividad(id INTEGER PRIMARY KEY AUTOINCREMENT, candidato_id INTEGER, accion TEXT, detalle TEXT, usuario TEXT, fecha TEXT)")
        c.execute("""
        CREATE TABLE IF NOT EXISTS entrevistas(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            candidato_id INTEGER,
            tipo TEXT,
            preguntas_json TEXT,
            respuestas_json TEXT,
            resultado TEXT DEFAULT 'Pendiente',
            evaluador TEXT,
            comentarios TEXT,
            updated_at TEXT
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS induccion(
            candidato_id INTEGER PRIMARY KEY,
            fecha_induccion TEXT,
            checklist_json TEXT,
            responsable TEXT,
            observaciones TEXT,
            resultado TEXT DEFAULT 'Pendiente',
            updated_at TEXT
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS expediente_digital(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            candidato_id INTEGER,
            categoria TEXT,
            documento TEXT,
            archivo TEXT,
            estatus TEXT DEFAULT 'Cargado',
            observaciones TEXT,
            fecha TEXT
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS capacitacion_dc3(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            candidato_id INTEGER,
            curso TEXT,
            fecha_curso TEXT,
            instructor TEXT,
            resultado TEXT DEFAULT 'Pendiente',
            dc3_generado TEXT DEFAULT 'No',
            archivo_dc3 TEXT,
            observaciones TEXT,
            fecha TEXT
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS firmas_electronicas(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            candidato_id INTEGER,
            documento TEXT,
            firmante TEXT,
            archivo_firma TEXT,
            archivo_pdf TEXT,
            fecha TEXT
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS access_kiosk_integracion(
            candidato_id INTEGER PRIMARY KEY,
            empleado_id TEXT,
            qr_generado TEXT DEFAULT 'No',
            estatus TEXT DEFAULT 'Pendiente',
            fecha TEXT
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS usuarios_roles(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT,
            usuario TEXT UNIQUE,
            rol TEXT,
            udn TEXT,
            activo INTEGER DEFAULT 1
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS usuarios_sistema(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            usuario TEXT UNIQUE NOT NULL,
            email TEXT,
            password_hash TEXT NOT NULL,
            rol TEXT NOT NULL DEFAULT 'Consulta',
            udn_permitidas TEXT DEFAULT 'TODAS',
            activo INTEGER DEFAULT 1,
            ultimo_acceso TEXT,
            fecha_creacion TEXT
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS roles_sistema(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rol TEXT UNIQUE NOT NULL,
            descripcion TEXT,
            permisos_json TEXT,
            activo INTEGER DEFAULT 1
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS auditoria_sistema(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT,
            accion TEXT,
            modulo TEXT,
            detalle TEXT,
            ip TEXT,
            fecha TEXT
        )""")
        c.execute("""
        CREATE TABLE IF NOT EXISTS configuracion_sistema(
            clave TEXT PRIMARY KEY,
            valor TEXT,
            descripcion TEXT,
            updated_at TEXT
        )""")
        conn.commit()
    seed_security_once()
    seed_from_excel_once()




def seed_security_once():
    roles = {
        "Administrador": ["*"] ,
        "RH Corporativo": ["dashboard", "candidatos", "reportes", "catalogos", "admin"],
        "Reclutador": ["dashboard", "candidatos", "entrevistas", "hoja_ruta"],
        "Servicio Médico": ["dashboard", "servicio_medico", "evaluacion"],
        "Operaciones": ["dashboard", "evaluacion_tecnica", "induccion", "entrega_operaciones"],
        "Capacitación": ["dashboard", "induccion", "capacitacion", "dc3"],
        "Consulta": ["dashboard", "reportes"]
    }
    for rol, permisos in roles.items():
        try:
            execute("INSERT INTO roles_sistema(rol,descripcion,permisos_json,activo) VALUES(?,?,?,1)", (rol, f"Rol {rol}", json.dumps(permisos, ensure_ascii=False)))
        except Exception:
            pass
    # Usuario administrador inicial / recuperación.
    # Se fuerza la actualización del password y permisos para evitar que una base previa
    # deje bloqueado el acceso en Render o en pruebas locales.
    admin = query("SELECT id FROM usuarios_sistema WHERE LOWER(usuario)=LOWER(?)", (ADMIN_USER,), one=True)
    admin_hash = generate_password_hash(ADMIN_PASSWORD)

    if not admin:
        execute("""INSERT INTO usuarios_sistema(nombre,usuario,email,password_hash,rol,udn_permitidas,activo,fecha_creacion)
                   VALUES(?,?,?,?,?,?,1,?)""",
                ("Administrador del Sistema", ADMIN_USER, "", admin_hash, "Administrador", "TODAS", now_str()))
    else:
        execute("""UPDATE usuarios_sistema
                   SET usuario=?, nombre=?, password_hash=?, rol=?, udn_permitidas=?, activo=1
                   WHERE id=?""",
                (ADMIN_USER, "Administrador del Sistema", admin_hash, "Administrador", "TODAS", admin["id"]))


def audit(accion, modulo="Sistema", detalle=""):
    try:
        execute("INSERT INTO auditoria_sistema(usuario,accion,modulo,detalle,ip,fecha) VALUES(?,?,?,?,?,?)",
                (session.get("usuario", "Sistema"), accion, modulo, detalle, request.headers.get("X-Forwarded-For", request.remote_addr or ""), now_str()))
    except Exception:
        pass


def is_logged_in():
    return bool(session.get("user_id"))


def current_user():
    if not session.get("user_id"):
        return None
    return {
        "id": session.get("user_id"),
        "nombre": session.get("nombre", ""),
        "usuario": session.get("usuario", ""),
        "rol": session.get("rol", "Consulta"),
        "udn_permitidas": session.get("udn_permitidas", "TODAS"),
    }


def has_role(*roles):
    return session.get("rol") in roles or session.get("rol") == "Administrador"


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not is_logged_in():
            return redirect(url_for("login", next=request.path))
        return fn(*args, **kwargs)
    return wrapper


def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not is_logged_in():
            return redirect(url_for("login", next=request.path))
        if not has_role("Administrador", "RH Corporativo"):
            flash("No tienes permiso para acceder a Administración.", "error")
            audit("Acceso denegado", "Administración", request.path)
            return redirect(url_for("dashboard"))
        return fn(*args, **kwargs)
    return wrapper


@app.context_processor
def inject_user():
    return {"current_user": current_user()}


@app.before_request
def require_login_global():
    endpoint = request.endpoint or ""
    if endpoint in PUBLIC_ENDPOINTS or endpoint.startswith("static"):
        return None
    if not is_logged_in():
        return redirect(url_for("login", next=request.path))
    return None


@app.route("/health")
def health():
    return jsonify({"status": "ok", "app": "Traxion Capital Humano", "version": "Fase 3.1 Seguridad", "time": now_str()})


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        # Acepta name="usuario" y name="username" para evitar errores entre versiones de login.html.
        usuario = (request.form.get("usuario") or request.form.get("username") or "").strip()
        password = request.form.get("password", "")
        row = query("SELECT * FROM usuarios_sistema WHERE LOWER(usuario)=LOWER(?) AND activo=1", (usuario,), one=True)
        if row and check_password_hash(row["password_hash"], password):
            session.clear()
            session["user_id"] = row["id"]
            session["nombre"] = row["nombre"]
            session["usuario"] = row["usuario"]
            session["rol"] = row["rol"]
            session["udn_permitidas"] = row["udn_permitidas"] or "TODAS"
            execute("UPDATE usuarios_sistema SET ultimo_acceso=? WHERE id=?", (now_str(), row["id"]))
            audit("Inicio de sesión", "Seguridad", f"Usuario {usuario}")
            return redirect(request.args.get("next") or url_for("inicio_plataforma"))
        flash("Usuario o contraseña incorrectos.", "error")
        audit("Login fallido", "Seguridad", usuario)
    return render_template("login.html", title="Acceso Plataforma CH")


@app.route("/logout")
def logout():
    audit("Cierre de sesión", "Seguridad", session.get("usuario", ""))
    session.clear()
    return redirect(url_for("login"))

def log(candidato_id, accion, detalle, usuario="Sistema"):
    execute("INSERT INTO actividad(candidato_id,accion,detalle,usuario,fecha) VALUES(?,?,?,?,?)", (candidato_id, accion, detalle, usuario, now_str()))

def next_folio():
    """Genera un folio CH-#### sin repetir.

    Antes se usaba COUNT(*) y eso podia chocar cuando la base ya tenia
    folios antiguos, eliminados o importaciones repetidas. Ahora busca el
    mayor consecutivo existente y valida que el folio no exista antes de
    devolverlo.
    """
    rows = query("SELECT folio FROM candidatos WHERE folio LIKE 'CH-%'")
    max_num = 1000
    for r in rows:
        folio = (r["folio"] or "").strip()
        try:
            num = int(folio.replace("CH-", "").strip())
            if num > max_num:
                max_num = num
        except Exception:
            continue

    next_num = max_num + 1
    while True:
        folio = f"CH-{next_num}"
        exists = query("SELECT id FROM candidatos WHERE folio=?", (folio,), one=True)
        if not exists:
            return folio
        next_num += 1

def ensure_docs(candidato_id):
    row = query("SELECT COUNT(*) c FROM documentacion WHERE candidato_id=?", (candidato_id,), one=True)
    if row and row["c"] > 0: return
    for doc in CHECK_DOCS:
        execute("INSERT INTO documentacion(candidato_id,documento,estado,updated_at) VALUES(?,?,?,?)", (candidato_id, doc, "Pendiente", now_str()))


def seed_from_excel_once():
    if not EXCEL_BASE.exists():
        print("[2/4] No existe Excel base; se crean catálogos mínimos.")
        for p in ["AYUDANTE GENERAL", "MONTACARGUISTA", "PATINERO", "SURTIDOR", "AUDITOR", "SUPERVISOR"]:
            try: execute("INSERT INTO catalogo_puestos(puesto) VALUES(?)", (p,))
            except Exception: pass
        return
    already = query("SELECT COUNT(*) c FROM catalogo_puestos", one=True)["c"]
    if already and already > 0:
        print("[2/4] Catálogos existentes; no se sobrescriben.")
        return
    print("[2/4] Cargando catálogos iniciales desde Excel...")
    try:
        xl = pd.ExcelFile(EXCEL_BASE)
        puestos, udns, reclutadores = set(), set(), set()
        if "RyS" in xl.sheet_names:
            df = pd.read_excel(EXCEL_BASE, sheet_name="RyS").fillna("")
            for v in df.get("PUESTO", []):
                if clean(v): puestos.add(clean(v).upper())
            for v in df.get("ALMACEN", []):
                if clean(v): udns.add(clean(v).upper())
            for v in df.get("RECLUTADOR", []):
                if clean(v): reclutadores.add(clean(v).upper())
        if "Catálogo de puestos" in xl.sheet_names:
            raw = pd.read_excel(EXCEL_BASE, sheet_name="Catálogo de puestos", header=None).fillna("")
            for row in raw.values:
                for cell in row:
                    txt = clean(cell).upper()
                    if txt and len(txt) > 3 and len(txt) < 70 and txt not in ["PUESTO", "CÓDIGO DE PUESTO"]:
                        if any(k in txt for k in ["AYUDANTE", "MONTAC", "PATIN", "SURTID", "AUDITOR", "SUPERV", "OPERADOR", "AUXILIAR", "ANALISTA", "COORD", "ENFERM", "CAPTUR"]):
                            puestos.add(txt)
        for p in sorted(puestos):
            try: execute("INSERT INTO catalogo_puestos(puesto) VALUES(?)", (p,))
            except Exception: pass
        for u in sorted(udns):
            try: execute("INSERT INTO catalogo_udn(almacen, region, localidad) VALUES(?,?,?)", (u, "POR DEFINIR", "POR DEFINIR"))
            except Exception: pass
        for r in sorted(reclutadores):
            try: execute("INSERT INTO catalogo_reclutadores(reclutador) VALUES(?)", (r,))
            except Exception: pass
        print(f"[2/4] Catálogos: puestos={len(puestos)}, udn={len(udns)}, reclutadores={len(reclutadores)}")
    except Exception as e:
        print("Error al cargar catálogos:", e)


def import_rys_excel(path):
    xl = pd.ExcelFile(path)
    df = pd.read_excel(path, sheet_name="RyS").fillna("") if "RyS" in xl.sheet_names else pd.read_excel(path).fillna("")
    created, updated = 0, 0
    for _, row in df.iterrows():
        nombre = clean(row.get("Nombre de Empleado\n(1er apelido + 2do + nombres en Mayusculas)", row.get("Nombre Completo", row.get("Nombre", "")))).upper()
        telefono = clean(row.get("TELEFONO", row.get("Telefono", "")))
        puesto = clean(row.get("PUESTO", row.get("Puesto", ""))).upper()
        almacen = clean(row.get("ALMACEN", row.get("Almacen", row.get("UDN", "")))).upper()
        reclutador = clean(row.get("RECLUTADOR", row.get("Reclutador", ""))).upper()
        fecha_ent = clean(row.get("FECHA DE ENTREVISTA", ""))
        motivo = clean(row.get("MOTIVO DE RECHAZO", ""))
        estatus = clean(row.get("ESTATUS", "")).upper()
        if not nombre or not puesto: continue
        etapa = "Registro"
        if "ENTREGADO" in estatus or "ACTIVO" in estatus: etapa = "Activo"
        elif "SELECCION" in estatus or "ENTREVISTA" in estatus: etapa = "Entrevista"
        elif "DECLINADO" in estatus or "RECHAZ" in estatus: etapa = "Evaluacion Tecnica"
        macro = ESTATUS_MACRO.get(etapa, "Reclutamiento")
        exists = query("SELECT id FROM candidatos WHERE nombre=? AND puesto=?", (nombre, puesto), one=True)
        if exists:
            execute("""UPDATE candidatos SET telefono=?, udn=?, reclutador=?, fecha_entrevista=?, etapa=?, estatus_macro=?, motivo_rechazo=?, fecha_actualizacion=? WHERE id=?""",
                    (telefono, almacen, reclutador, fecha_ent, etapa, macro, motivo, now_str(), exists["id"]))
            updated += 1
            cid = exists["id"]
        else:
            # Insercion robusta: evita que una importacion se detenga por folio duplicado.
            # Si existe choque de folio, se genera otro consecutivo y se reintenta.
            cid = None
            last_error = None
            for _ in range(10):
                try:
                    cid = execute("""INSERT INTO candidatos(folio,nombre,telefono,puesto,udn,reclutador,fecha_entrevista,etapa,estatus_macro,motivo_rechazo,fecha_registro,fecha_actualizacion,comentarios)
                                 VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                              (next_folio(), nombre, telefono, puesto, almacen, reclutador, fecha_ent, etapa, macro, motivo, now_str(), now_str(), "Importado desde Excel RyS"))
                    created += 1
                    break
                except sqlite3.IntegrityError as e:
                    last_error = e
                    if "candidatos.folio" in str(e) or "UNIQUE constraint failed" in str(e):
                        continue
                    raise
            if cid is None:
                raise last_error or Exception("No fue posible generar un folio unico para el candidato")
        ensure_docs(cid)
        for table, col, val in [("catalogo_puestos", "puesto", puesto), ("catalogo_udn", "almacen", almacen), ("catalogo_reclutadores", "reclutador", reclutador)]:
            if val:
                try:
                    if table == "catalogo_udn": execute("INSERT INTO catalogo_udn(almacen,region,localidad) VALUES(?,?,?)", (val,"POR DEFINIR","POR DEFINIR"))
                    elif table == "catalogo_reclutadores": execute("INSERT INTO catalogo_reclutadores(reclutador) VALUES(?)", (val,))
                    else: execute("INSERT INTO catalogo_puestos(puesto) VALUES(?)", (val,))
                except Exception: pass
    return created, updated


def get_kpis():
    total = query("SELECT COUNT(*) c FROM candidatos", one=True)["c"] or 0
    proceso = query("SELECT COUNT(*) c FROM candidatos WHERE estatus_macro IN ('Reclutamiento','Contratacion')", one=True)["c"] or 0
    activos = query("SELECT COUNT(*) c FROM candidatos WHERE estatus_macro='Activo' OR etapa='Activo'", one=True)["c"] or 0
    docs = query("SELECT COUNT(*) c FROM candidatos WHERE etapa='Documentacion'", one=True)["c"] or 0
    hoy = query("SELECT COUNT(*) c FROM candidatos WHERE substr(fecha_registro,1,10)=?", (today_str(),), one=True)["c"] or 0
    rfc = query("SELECT COUNT(*) c, COALESCE(SUM(costo_rfc),0) costo FROM evaluaciones WHERE rfc_actualizado='Si'", one=True)
    return {"total": total, "proceso": proceso, "activos": activos, "docs": docs, "hoy": hoy, "rfc": rfc["c"] or 0, "costo_rfc": rfc["costo"] or 0}


def cobertura_por_puesto(limit=12):
    # Si existe headcount cargado, la cobertura se calcula contra requerido real.
    hc_count = query("SELECT COUNT(*) c FROM headcount", one=True)["c"] or 0
    if hc_count:
        rows = query("""
            SELECT h.puesto,
                   SUM(h.requerido) requerido,
                   SUM(h.activo) activo_hc,
                   (SELECT COUNT(*) FROM candidatos c WHERE c.puesto=h.puesto AND (c.etapa='Activo' OR c.estatus_macro='Activo')) activos
            FROM headcount h
            GROUP BY h.puesto
            ORDER BY SUM(h.requerido) DESC
            LIMIT ?
        """, (limit,))
        out=[]
        for r in rows:
            req = r["requerido"] or 0
            act = r["activos"] or r["activo_hc"] or 0
            vac = max(req-act, 0)
            pct = int((act/req)*100) if req else 0
            out.append({"puesto": r["puesto"], "total": req, "activos": act, "vacantes": vac, "pct": min(pct,100)})
        return out

    # Fallback: candidatos activos contra total importado por puesto.
    rows = query("""
        SELECT puesto,
               COUNT(*) total,
               SUM(CASE WHEN etapa='Activo' OR estatus_macro='Activo' THEN 1 ELSE 0 END) activos
        FROM candidatos
        WHERE puesto IS NOT NULL AND puesto<>''
        GROUP BY puesto ORDER BY total DESC LIMIT ?
    """, (limit,))
    out = []
    for r in rows:
        total = r["total"] or 0; act = r["activos"] or 0
        pct = int((act / total) * 100) if total else 0
        out.append({"puesto": r["puesto"], "total": total, "activos": act, "vacantes": total-act, "pct": pct})
    return out


def etapa_counts():
    rows = query("SELECT etapa, COUNT(*) total FROM candidatos GROUP BY etapa")
    d = {e: 0 for e in ETAPAS}
    for r in rows: d[r["etapa"]] = r["total"]
    return d


def parse_fecha(valor):
    if not valor:
        return None
    txt = str(valor).strip()
    for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%Y %H:%M:%S"]:
        try:
            return datetime.strptime(txt[:19] if fmt.endswith('%S') else txt[:10], fmt)
        except Exception:
            pass
    try:
        return pd.to_datetime(txt).to_pydatetime()
    except Exception:
        return None


def dias_en_etapa(c):
    base = parse_fecha(c["fecha_actualizacion"] or c["fecha_registro"])
    if not base:
        return 0
    return max(0, (datetime.now() - base).days)


def sla_estado(c):
    etapa = c["etapa"] or "Registro"
    limite = SLA_DIAS.get(etapa, 1)
    dias = dias_en_etapa(c)
    if etapa == "Activo":
        return {"color": "verde", "texto": "Finalizado", "dias": dias, "limite": limite}
    if dias > limite:
        return {"color": "rojo", "texto": f"Vencido {dias-limite} día(s)", "dias": dias, "limite": limite}
    if dias == limite:
        return {"color": "amarillo", "texto": "Por vencer", "dias": dias, "limite": limite}
    return {"color": "verde", "texto": "En tiempo", "dias": dias, "limite": limite}


def enriquecer_candidatos(rows):
    out = []
    for r in rows:
        d = dict(r)
        s = sla_estado(r)
        d.update({"sla_color": s["color"], "sla_texto": s["texto"], "dias_etapa": s["dias"], "sla_limite": s["limite"]})
        out.append(d)
    return out


def seguimiento_resumen():
    rows = query("SELECT * FROM candidatos WHERE etapa<>'Activo' ORDER BY fecha_actualizacion ASC")
    enriched = enriquecer_candidatos(rows)
    return {
        "vencidos": sum(1 for x in enriched if x["sla_color"] == "rojo"),
        "por_vencer": sum(1 for x in enriched if x["sla_color"] == "amarillo"),
        "en_tiempo": sum(1 for x in enriched if x["sla_color"] == "verde"),
        "total": len(enriched),
    }


def agenda_entrevistas(limit=20):
    rows = query("""
        SELECT * FROM candidatos
        WHERE fecha_entrevista IS NOT NULL AND fecha_entrevista<>'' AND etapa IN ('Registro','Entrevista','Oferta Laboral')
        ORDER BY fecha_entrevista ASC LIMIT ?
    """, (limit,))
    return rows



def preguntas_por_puesto(puesto, tipo="Entrevista 1"):
    p = (puesto or "").upper()
    if "MONTAC" in p:
        base = ENTREVISTA_BASE["MONTACARGUISTA"]
    elif "PATIN" in p or "PATÍN" in p:
        base = ENTREVISTA_BASE["PATINERO"]
    else:
        base = ENTREVISTA_BASE["GENERAL"]
    if tipo == "Entrevista 2":
        return base + ["Validación final de disponibilidad", "Confirmación de condiciones de oferta", "Comentarios del jefe operativo"]
    return base


def import_headcount_excel(path):
    df = pd.read_excel(path).fillna("")
    # Columnas flexibles: UDN/ALMACEN/CUENTA, PUESTO, HC/REQUERIDO/NOM/VAC
    created = 0
    execute("DELETE FROM headcount")
    for _, row in df.iterrows():
        udn = clean(row.get("UDN", row.get("ALMACEN", row.get("CUENTA", row.get("Cuenta", ""))))).upper()
        puesto = clean(row.get("PUESTO", row.get("Puesto", ""))).upper()
        if not puesto:
            continue
        requerido = clean(row.get("HC", row.get("REQUERIDO", row.get("Headcount", row.get("Requerido", 0)))))
        activo = clean(row.get("NOM", row.get("ACTIVO", row.get("Activos", 0))))
        try: requerido = int(float(requerido or 0))
        except Exception: requerido = 0
        try: activo = int(float(activo or 0))
        except Exception: activo = 0
        execute("INSERT INTO headcount(udn,puesto,requerido,activo) VALUES(?,?,?,?)", (udn, puesto, requerido, activo))
        created += 1
        if puesto:
            try: execute("INSERT INTO catalogo_puestos(puesto) VALUES(?)", (puesto,))
            except Exception: pass
        if udn:
            try: execute("INSERT INTO catalogo_udn(almacen,region,localidad) VALUES(?,?,?)", (udn,"POR DEFINIR","POR DEFINIR"))
            except Exception: pass
    return created


def candidatos_por_modulo(etapa=None, macro=None, q=''):
    sql = "SELECT * FROM candidatos WHERE 1=1"
    params = []
    if etapa:
        if isinstance(etapa, (list, tuple)):
            sql += " AND etapa IN (%s)" % ','.join(['?']*len(etapa))
            params.extend(etapa)
        else:
            sql += " AND etapa=?"
            params.append(etapa)
    if macro:
        sql += " AND estatus_macro=?"
        params.append(macro)
    if q:
        like=f"%{q.upper()}%"
        sql += " AND (UPPER(nombre) LIKE ? OR UPPER(puesto) LIKE ? OR UPPER(udn) LIKE ? OR UPPER(reclutador) LIKE ?)"
        params += [like, like, like, like]
    sql += " ORDER BY id DESC"
    return query(sql, tuple(params))

@app.route("/")
def root():
    return redirect(url_for("inicio_plataforma"))

@app.route("/inicio")
def inicio_plataforma():
    return render_template("inicio.html", active="inicio", page_title="Inicio")

@app.route("/dashboard")
def dashboard():
    kpis = get_kpis(); cobertura = cobertura_por_puesto(14); pipeline = etapa_counts()
    recientes = query("SELECT * FROM candidatos ORDER BY id DESC LIMIT 10")
    udns = query("SELECT udn, COUNT(*) total FROM candidatos WHERE udn<>'' GROUP BY udn ORDER BY total DESC LIMIT 12")
    seguimiento = seguimiento_resumen()
    agenda = agenda_entrevistas(8)
    return render_template("dashboard.html", kpis=kpis, cobertura=cobertura, pipeline=pipeline, recientes=recientes, etapas=ETAPAS, udns=udns, seguimiento=seguimiento, agenda=agenda)

@app.route("/candidatos")
def candidatos():
    q = request.args.get("q", "").strip(); etapa = request.args.get("etapa", "").strip(); puesto = request.args.get("puesto", "").strip(); udn = request.args.get("udn", "").strip()
    sql = "SELECT * FROM candidatos WHERE 1=1"; params=[]
    if q:
        sql += " AND (nombre LIKE ? OR puesto LIKE ? OR udn LIKE ? OR reclutador LIKE ?)"; like=f"%{q}%"; params += [like,like,like,like]
    if etapa: sql += " AND etapa=?"; params.append(etapa)
    if puesto: sql += " AND puesto=?"; params.append(puesto)
    if udn: sql += " AND udn=?"; params.append(udn)
    sql += " ORDER BY id DESC LIMIT 700"
    rows = query(sql, params)
    return render_template("candidatos.html", rows=rows, q=q, etapas=ETAPAS, etapa_sel=etapa)

@app.route("/candidatos/nuevo", methods=["GET", "POST"])
def candidato_nuevo():
    puestos = query("SELECT puesto FROM catalogo_puestos WHERE activo=1 ORDER BY puesto")
    udns = query("SELECT almacen FROM catalogo_udn WHERE activo=1 ORDER BY almacen")
    reclutadores = query("SELECT reclutador FROM catalogo_reclutadores WHERE activo=1 ORDER BY reclutador")
    if request.method == "POST":
        nombre = request.form.get("nombre", "").strip().upper(); puesto = request.form.get("puesto", "").strip().upper()
        if not nombre or not puesto:
            flash("Nombre y puesto son obligatorios.", "error"); return redirect(url_for("candidato_nuevo"))
        cid = None
        last_error = None
        for _ in range(10):
            try:
                cid = execute("""INSERT INTO candidatos(folio,nombre,telefono,email,region,localidad,udn,puesto,reclutador,fecha_entrevista,etapa,estatus_macro,fecha_registro,fecha_actualizacion,comentarios)
                             VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                          (next_folio(), nombre, request.form.get("telefono",""), request.form.get("email",""), request.form.get("region","").upper(), request.form.get("localidad","").upper(), request.form.get("udn","").upper(), puesto, request.form.get("reclutador","").upper(), request.form.get("fecha_entrevista",""), "Registro", "Reclutamiento", now_str(), now_str(), "Registro manual"))
                break
            except sqlite3.IntegrityError as e:
                last_error = e
                if "candidatos.folio" in str(e) or "UNIQUE constraint failed" in str(e):
                    continue
                raise
        if cid is None:
            flash(f"No fue posible generar un folio unico: {last_error}", "error")
            return redirect(url_for("candidato_nuevo"))
        ensure_docs(cid); log(cid, "Registro", "Alta manual", "RH")
        flash("Candidato registrado correctamente.", "success")
        return redirect(url_for("candidato_detalle", candidato_id=cid))
    return render_template("candidato_form.html", puestos=puestos, udns=udns, reclutadores=reclutadores)

@app.route("/candidatos/importar", methods=["POST"])
def importar_candidatos():
    f = request.files.get("archivo")
    if not f or not f.filename:
        flash("Selecciona un archivo Excel.", "error"); return redirect(url_for("candidatos"))
    dest = UPLOAD_DIR / secure_filename(f.filename); f.save(dest)
    try:
        created, updated = import_rys_excel(dest)
        flash(f"Importación finalizada. Nuevos: {created} | Actualizados: {updated}", "success")
    except Exception as e:
        flash(f"Error al importar: {e}", "error")
    return redirect(url_for("dashboard"))

@app.route("/candidatos/<int:candidato_id>", methods=["GET", "POST"])
def candidato_detalle(candidato_id):
    cand = query("SELECT * FROM candidatos WHERE id=?", (candidato_id,), one=True)
    if not cand: flash("Candidato no encontrado.", "error"); return redirect(url_for("candidatos"))
    if request.method == "POST":
        etapa = request.form.get("etapa", cand["etapa"]); macro = ESTATUS_MACRO.get(etapa, cand["estatus_macro"])
        execute("UPDATE candidatos SET etapa=?, estatus_macro=?, fecha_actualizacion=? WHERE id=?", (etapa, macro, now_str(), candidato_id))
        log(candidato_id, "Cambio etapa", f"Nueva etapa: {etapa}", "RH"); flash("Etapa actualizada.", "success")
        return redirect(url_for("candidato_detalle", candidato_id=candidato_id))
    docs = query("SELECT * FROM documentacion WHERE candidato_id=? ORDER BY id", (candidato_id,))
    ruta = query("SELECT * FROM hoja_ruta WHERE candidato_id=?", (candidato_id,), one=True)
    evals = query("SELECT * FROM evaluaciones WHERE candidato_id=?", (candidato_id,), one=True)
    oferta = query("SELECT * FROM oferta_laboral WHERE candidato_id=?", (candidato_id,), one=True)
    actividad = query("SELECT * FROM actividad WHERE candidato_id=? ORDER BY id DESC LIMIT 10", (candidato_id,))
    return render_template("candidato_detalle.html", c=cand, etapas=ETAPAS, docs=docs, ruta=ruta, evals=evals, oferta=oferta, actividad=actividad)

@app.route("/hoja-ruta/<int:candidato_id>", methods=["GET", "POST"])
def hoja_ruta(candidato_id):
    cand = query("SELECT * FROM candidatos WHERE id=?", (candidato_id,), one=True)
    if not cand: return redirect(url_for("candidatos"))
    if request.method == "POST":
        vals = {k: request.form.get(k, "") for k in ["ruta_transporte", "parada", "handheld", "wms", "sap", "excel", "patin", "montacargas", "reingreso_traxion", "reingreso_cliente", "emergencia_nombre", "emergencia_parentesco", "emergencia_telefono"]}
        vals["experiencia_json"] = request.form.get("experiencia_json", "")
        execute("""INSERT INTO hoja_ruta(candidato_id,ruta_transporte,parada,experiencia_json,handheld,wms,sap,excel,patin,montacargas,reingreso_traxion,reingreso_cliente,emergencia_nombre,emergencia_parentesco,emergencia_telefono,updated_at)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                   ON CONFLICT(candidato_id) DO UPDATE SET ruta_transporte=excluded.ruta_transporte, parada=excluded.parada, experiencia_json=excluded.experiencia_json,
                   handheld=excluded.handheld, wms=excluded.wms, sap=excluded.sap, excel=excluded.excel, patin=excluded.patin, montacargas=excluded.montacargas,
                   reingreso_traxion=excluded.reingreso_traxion, reingreso_cliente=excluded.reingreso_cliente, emergencia_nombre=excluded.emergencia_nombre,
                   emergencia_parentesco=excluded.emergencia_parentesco, emergencia_telefono=excluded.emergencia_telefono, updated_at=excluded.updated_at""",
                (candidato_id, vals["ruta_transporte"], vals["parada"], vals["experiencia_json"], vals["handheld"], vals["wms"], vals["sap"], vals["excel"], vals["patin"], vals["montacargas"], vals["reingreso_traxion"], vals["reingreso_cliente"], vals["emergencia_nombre"], vals["emergencia_parentesco"], vals["emergencia_telefono"], now_str()))
        execute("UPDATE candidatos SET etapa='Entrevista', estatus_macro='Reclutamiento', fecha_actualizacion=? WHERE id=?", (now_str(), candidato_id))
        log(candidato_id, "Hoja Ruta", "Hoja de ruta guardada", "RH"); flash("Hoja de ruta guardada.", "success")
        return redirect(url_for("candidato_detalle", candidato_id=candidato_id))
    data = query("SELECT * FROM hoja_ruta WHERE candidato_id=?", (candidato_id,), one=True)
    return render_template("hoja_ruta.html", c=cand, data=data)

@app.route("/oferta/<int:candidato_id>", methods=["GET", "POST"])
def oferta_laboral(candidato_id):
    cand = query("SELECT * FROM candidatos WHERE id=?", (candidato_id,), one=True)
    if not cand: return redirect(url_for("candidatos"))
    if request.method == "POST":
        checked = {x: ("Si" if request.form.get(x) else "No") for x in CHECK_OFERTA}
        execute("""INSERT INTO oferta_laboral(candidato_id,sueldo,prestaciones,bonos,horario,lugar_trabajo,fecha_primer_pago,checklist_json,updated_at)
                   VALUES(?,?,?,?,?,?,?,?,?) ON CONFLICT(candidato_id) DO UPDATE SET sueldo=excluded.sueldo,prestaciones=excluded.prestaciones,bonos=excluded.bonos,horario=excluded.horario,lugar_trabajo=excluded.lugar_trabajo,fecha_primer_pago=excluded.fecha_primer_pago,checklist_json=excluded.checklist_json,updated_at=excluded.updated_at""",
                (candidato_id, request.form.get("sueldo",""), request.form.get("prestaciones",""), request.form.get("bonos",""), request.form.get("horario",""), request.form.get("lugar_trabajo",""), request.form.get("fecha_primer_pago",""), json.dumps(checked, ensure_ascii=False), now_str()))
        execute("UPDATE candidatos SET etapa='Oferta Laboral', estatus_macro='Reclutamiento', fecha_actualizacion=? WHERE id=?", (now_str(), candidato_id))
        log(candidato_id, "Oferta", "Oferta laboral capturada", "RH"); flash("Oferta laboral guardada.", "success")
        return redirect(url_for("candidato_detalle", candidato_id=candidato_id))
    data = query("SELECT * FROM oferta_laboral WHERE candidato_id=?", (candidato_id,), one=True)
    checks = json.loads(data["checklist_json"]) if data and data["checklist_json"] else {}
    return render_template("oferta.html", c=cand, data=data, checks=checks, checklist=CHECK_OFERTA)

@app.route("/evaluacion/<int:candidato_id>", methods=["GET", "POST"])
def evaluacion(candidato_id):
    cand = query("SELECT * FROM candidatos WHERE id=?", (candidato_id,), one=True)
    if not cand: return redirect(url_for("candidatos"))
    if request.method == "POST":
        med = request.form.get("medico_resultado", "Pendiente")
        tec = request.form.get("tecnico_resultado", "Pendiente")
        if med == "No Apto" and not request.form.get("medico_motivo", "").strip():
            flash("Cuando el resultado médico es No Apto, el motivo es obligatorio.", "error"); return redirect(url_for("evaluacion", candidato_id=candidato_id))
        if tec == "No Apto" and not request.form.get("tecnico_motivo", "").strip():
            flash("Cuando la evaluación técnica es No Apto, el motivo es obligatorio.", "error"); return redirect(url_for("evaluacion", candidato_id=candidato_id))
        execute("""INSERT INTO evaluaciones(candidato_id,medico_resultado,medico_motivo,tecnico_resultado,tecnico_tipo,tecnico_motivo,rfc_actualizado,costo_rfc,updated_at)
                   VALUES(?,?,?,?,?,?,?,?,?) ON CONFLICT(candidato_id) DO UPDATE SET medico_resultado=excluded.medico_resultado,medico_motivo=excluded.medico_motivo,tecnico_resultado=excluded.tecnico_resultado,tecnico_tipo=excluded.tecnico_tipo,tecnico_motivo=excluded.tecnico_motivo,rfc_actualizado=excluded.rfc_actualizado,costo_rfc=excluded.costo_rfc,updated_at=excluded.updated_at""",
                (candidato_id, med, request.form.get("medico_motivo",""), tec, request.form.get("tecnico_tipo",""), request.form.get("tecnico_motivo",""), request.form.get("rfc_actualizado","No"), float(request.form.get("costo_rfc") or 0), now_str()))
        new_stage = "Servicio Medico" if med != "Pendiente" and tec == "Pendiente" else "Evaluacion Tecnica"
        execute("UPDATE candidatos SET etapa=?, estatus_macro='Reclutamiento', fecha_actualizacion=? WHERE id=?", (new_stage, now_str(), candidato_id))
        log(candidato_id, "Evaluacion", "Evaluación médico/técnica actualizada", "RH"); flash("Evaluación guardada.", "success")
        return redirect(url_for("candidato_detalle", candidato_id=candidato_id))
    data = query("SELECT * FROM evaluaciones WHERE candidato_id=?", (candidato_id,), one=True)
    return render_template("evaluacion.html", c=cand, data=data, resultados=RESULTADOS)

@app.route("/documentacion/<int:candidato_id>", methods=["GET", "POST"])
def documentacion(candidato_id):
    cand = query("SELECT * FROM candidatos WHERE id=?", (candidato_id,), one=True)
    if not cand: return redirect(url_for("candidatos"))
    ensure_docs(candidato_id)
    if request.method == "POST":
        for d in query("SELECT * FROM documentacion WHERE candidato_id=?", (candidato_id,)):
            estado = request.form.get(f"estado_{d['id']}", d["estado"])
            execute("UPDATE documentacion SET estado=?, updated_at=? WHERE id=?", (estado, now_str(), d["id"]))
        execute("UPDATE candidatos SET etapa='Documentacion', estatus_macro='Contratacion', fecha_actualizacion=? WHERE id=?", (now_str(), candidato_id))
        log(candidato_id, "Documentacion", "Checklist documental actualizado", "RH"); flash("Documentación actualizada.", "success")
        return redirect(url_for("candidato_detalle", candidato_id=candidato_id))
    docs = query("SELECT * FROM documentacion WHERE candidato_id=? ORDER BY id", (candidato_id,))
    return render_template("documentacion.html", c=cand, docs=docs)

@app.route("/hoja-ruta/<int:candidato_id>/pdf")
def hoja_ruta_pdf(candidato_id):
    cand = query("SELECT * FROM candidatos WHERE id=?", (candidato_id,), one=True)
    ruta = query("SELECT * FROM hoja_ruta WHERE candidato_id=?", (candidato_id,), one=True)
    oferta = query("SELECT * FROM oferta_laboral WHERE candidato_id=?", (candidato_id,), one=True)
    evals = query("SELECT * FROM evaluaciones WHERE candidato_id=?", (candidato_id,), one=True)
    if not cand: return redirect(url_for("candidatos"))
    out = REPORT_DIR / f"Hoja_Ruta_{cand['folio']}.pdf"
    c = canvas.Canvas(str(out), pagesize=letter)
    w,h=letter
    c.setFont("Helvetica-Bold", 18); c.drawString(0.75*inch, h-0.75*inch, "HOJA DE RUTA - CAPITAL HUMANO")
    c.setFillColorRGB(0.82,0.87,0); c.rect(0.75*inch,h-1.1*inch,7*inch,0.15*inch,fill=1,stroke=0); c.setFillColorRGB(0,0,0)
    y=h-1.45*inch; c.setFont("Helvetica-Bold", 11); c.drawString(0.75*inch,y,f"Folio: {cand['folio']}"); c.drawString(3.2*inch,y,f"Nombre: {cand['nombre'][:45]}")
    y-=0.3*inch; c.setFont("Helvetica",10); fields=[("Puesto",cand['puesto']),("UDN",cand['udn']),("Teléfono",cand['telefono']),("Reclutador",cand['reclutador']),("Etapa",cand['etapa'])]
    for lab,val in fields:
        c.drawString(0.75*inch,y,f"{lab}: {val or ''}"); y-=0.22*inch
    y-=0.12*inch; c.setFont("Helvetica-Bold",12); c.drawString(0.75*inch,y,"Hoja de Ruta"); y-=0.28*inch; c.setFont("Helvetica",10)
    if ruta:
        for lab,key in [("Ruta transporte","ruta_transporte"),("Parada","parada"),("Handheld","handheld"),("WMS","wms"),("SAP","sap"),("Excel","excel"),("Patín","patin"),("Montacargas","montacargas")]:
            c.drawString(0.75*inch,y,f"{lab}: {ruta[key] or ''}"); y-=0.2*inch
    else:
        c.drawString(0.75*inch,y,"Sin hoja de ruta capturada."); y-=0.25*inch
    y-=0.1*inch; c.setFont("Helvetica-Bold",12); c.drawString(0.75*inch,y,"Oferta / Evaluaciones"); y-=0.25*inch; c.setFont("Helvetica",10)
    c.drawString(0.75*inch,y,f"Oferta: {'Capturada' if oferta else 'Pendiente'}"); y-=0.2*inch
    if evals:
        c.drawString(0.75*inch,y,f"Médico: {evals['medico_resultado'] or ''}    Técnico: {evals['tecnico_resultado'] or ''}    RFC actualizado: {evals['rfc_actualizado'] or 'No'}")
    else: c.drawString(0.75*inch,y,"Evaluaciones pendientes.")
    c.showPage(); c.save()
    return send_file(out, as_attachment=True)


@app.route("/entrevista/<int:candidato_id>", methods=["GET", "POST"])
def entrevista(candidato_id):
    cand = query("SELECT * FROM candidatos WHERE id=?", (candidato_id,), one=True)
    if not cand: return redirect(url_for("candidatos"))
    tipo = request.args.get("tipo", request.form.get("tipo", "Entrevista 1"))
    preguntas = preguntas_por_puesto(cand["puesto"], tipo)
    data = query("SELECT * FROM entrevistas WHERE candidato_id=? AND tipo=? ORDER BY id DESC LIMIT 1", (candidato_id, tipo), one=True)
    respuestas = json.loads(data["respuestas_json"]) if data and data["respuestas_json"] else {}
    if request.method == "POST":
        respuestas = {f"p{i}": request.form.get(f"p{i}", "") for i in range(len(preguntas))}
        resultado = request.form.get("resultado", "Pendiente")
        evaluador = request.form.get("evaluador", "")
        comentarios = request.form.get("comentarios", "")
        if data:
            execute("""UPDATE entrevistas SET preguntas_json=?, respuestas_json=?, resultado=?, evaluador=?, comentarios=?, updated_at=? WHERE id=?""",
                    (json.dumps(preguntas, ensure_ascii=False), json.dumps(respuestas, ensure_ascii=False), resultado, evaluador, comentarios, now_str(), data["id"]))
        else:
            execute("""INSERT INTO entrevistas(candidato_id,tipo,preguntas_json,respuestas_json,resultado,evaluador,comentarios,updated_at) VALUES(?,?,?,?,?,?,?,?)""",
                    (candidato_id, tipo, json.dumps(preguntas, ensure_ascii=False), json.dumps(respuestas, ensure_ascii=False), resultado, evaluador, comentarios, now_str()))
        execute("UPDATE candidatos SET etapa='Entrevista', estatus_macro='Reclutamiento', fecha_actualizacion=? WHERE id=?", (now_str(), candidato_id))
        log(candidato_id, "Entrevista", f"{tipo} guardada con resultado {resultado}", evaluador or "RH")
        flash("Entrevista guardada correctamente.", "success")
        return redirect(url_for("candidato_detalle", candidato_id=candidato_id))
    return render_template("entrevista.html", c=cand, tipo=tipo, preguntas=preguntas, respuestas=respuestas, data=data, resultados=RESULTADOS)

@app.route("/induccion/<int:candidato_id>", methods=["GET", "POST"])
def induccion(candidato_id):
    cand = query("SELECT * FROM candidatos WHERE id=?", (candidato_id,), one=True)
    if not cand: return redirect(url_for("candidatos"))
    data = query("SELECT * FROM induccion WHERE candidato_id=?", (candidato_id,), one=True)
    checks = json.loads(data["checklist_json"]) if data and data["checklist_json"] else {}
    if request.method == "POST":
        checks = {x: ("Si" if request.form.get(x) else "No") for x in INDUCCION_CHECKLIST}
        resultado = request.form.get("resultado", "Pendiente")
        execute("""INSERT INTO induccion(candidato_id,fecha_induccion,checklist_json,responsable,observaciones,resultado,updated_at)
                   VALUES(?,?,?,?,?,?,?) ON CONFLICT(candidato_id) DO UPDATE SET fecha_induccion=excluded.fecha_induccion,checklist_json=excluded.checklist_json,responsable=excluded.responsable,observaciones=excluded.observaciones,resultado=excluded.resultado,updated_at=excluded.updated_at""",
                (candidato_id, request.form.get("fecha_induccion",""), json.dumps(checks, ensure_ascii=False), request.form.get("responsable",""), request.form.get("observaciones",""), resultado, now_str()))
        etapa = "Activo" if resultado == "Completada" else "Induccion"
        macro = ESTATUS_MACRO.get(etapa, "Contratacion")
        execute("UPDATE candidatos SET etapa=?, estatus_macro=?, fecha_actualizacion=? WHERE id=?", (etapa, macro, now_str(), candidato_id))
        log(candidato_id, "Induccion", f"Inducción: {resultado}", request.form.get("responsable", "RH"))
        flash("Inducción actualizada correctamente.", "success")
        return redirect(url_for("candidato_detalle", candidato_id=candidato_id))
    return render_template("induccion.html", c=cand, data=data, checklist=INDUCCION_CHECKLIST, checks=checks)


@app.route("/entrevistas")
def entrevistas_panel():
    q = request.args.get("q", "")
    rows = candidatos_por_modulo(etapa=["Registro", "Entrevista", "Oferta Laboral"], q=q)
    return render_template("modulo_lista.html", active="entrevista", page_title="Entrevistas", titulo="Entrevistas", descripcion="Candidatos en etapa de registro, entrevista u oferta. Desde aquí puedes capturar Entrevista 1 o Entrevista 2.", rows=rows, q=q, action1="Entrevista 1", action2="Entrevista 2", endpoint1="entrevista", endpoint2="entrevista")

@app.route("/hoja-ruta")
def hoja_ruta_panel():
    q = request.args.get("q", "")
    rows = candidatos_por_modulo(etapa=["Registro", "Entrevista"], q=q)
    return render_template("modulo_lista.html", active="hoja_ruta", page_title="Hoja de Ruta", titulo="Hoja de Ruta Digital", descripcion="Candidatos pendientes o en entrevista para completar los datos de ruta, transporte, herramientas y experiencia.", rows=rows, q=q, action1="Capturar ruta", action2=None, endpoint1="hoja_ruta", endpoint2=None)

@app.route("/oferta-laboral")
def oferta_panel():
    q = request.args.get("q", "")
    rows = candidatos_por_modulo(etapa=["Entrevista", "Oferta Laboral"], q=q)
    return render_template("modulo_lista.html", active="oferta", page_title="Oferta Laboral", titulo="Oferta Laboral", descripcion="Checklist de explicación de sueldo, prestaciones, bonos, horario, lugar de trabajo y fecha de primer pago.", rows=rows, q=q, action1="Capturar oferta", action2=None, endpoint1="oferta_laboral", endpoint2=None)

@app.route("/servicio-medico")
def servicio_medico_panel():
    q = request.args.get("q", "")
    rows = candidatos_por_modulo(etapa=["Oferta Laboral", "Servicio Medico", "Evaluacion Tecnica"], q=q)
    return render_template("modulo_lista.html", active="evaluacion", page_title="Servicio Médico", titulo="Servicio Médico", descripcion="Captura de resultado médico: Apto, Apto condicionado o No apto con motivo obligatorio.", rows=rows, q=q, action1="Capturar médico", action2=None, endpoint1="evaluacion", endpoint2=None)

@app.route("/evaluacion-tecnica")
def evaluacion_tecnica_panel():
    q = request.args.get("q", "")
    rows = candidatos_por_modulo(etapa=["Servicio Medico", "Evaluacion Tecnica"], q=q)
    return render_template("modulo_lista.html", active="evaluacion", page_title="Evaluación Técnica", titulo="Evaluación Técnica", descripcion="Captura prueba técnica por perfil: patín, montacargas, aritmética, picking, SAP o WMS.", rows=rows, q=q, action1="Capturar evaluación", action2=None, endpoint1="evaluacion", endpoint2=None)

@app.route("/documentacion-panel")
def documentacion_panel():
    q = request.args.get("q", "")
    rows = candidatos_por_modulo(etapa=["Evaluacion Tecnica", "Documentacion", "Induccion"], q=q)
    return render_template("modulo_lista.html", active="documentacion", page_title="Documentación", titulo="Documentación", descripcion="Checklist documental del candidato: INE, CURP, RFC, NSS, cuenta bancaria, comprobante, acta y solicitud/CV.", rows=rows, q=q, action1="Checklist docs", action2=None, endpoint1="documentacion", endpoint2=None)

@app.route("/induccion-panel")
def induccion_panel():
    q = request.args.get("q", "")
    rows = candidatos_por_modulo(etapa=["Documentacion", "Induccion", "Entrega Operaciones"], q=q)
    return render_template("modulo_lista.html", active="induccion", page_title="Inducción", titulo="Inducción", descripcion="Control de inducción, seguridad, EPP, recorrido operativo, alta biométrica, gafete y entrega a jefe directo.", rows=rows, q=q, action1="Capturar inducción", action2=None, endpoint1="induccion", endpoint2=None)

@app.route("/reportes")
def reportes_panel():
    rows = query("SELECT etapa, estatus_macro, COUNT(*) total FROM candidatos GROUP BY etapa, estatus_macro ORDER BY etapa")
    rfc = query("SELECT COUNT(*) total, COALESCE(SUM(costo_rfc),0) costo FROM evaluaciones WHERE rfc_actualizado='Si'", one=True)
    return render_template("reportes.html", rows=rows, kpis=get_kpis(), rfc=rfc)


@app.route("/seguimiento")
def seguimiento():
    q = request.args.get("q", "").strip()
    etapa = request.args.get("etapa", "").strip()
    rows = candidatos_por_modulo(etapa=etapa if etapa else None, q=q)
    rows = [r for r in rows if r["etapa"] != "Activo"]
    enriched = enriquecer_candidatos(rows)
    color = request.args.get("sla", "").strip()
    if color:
        enriched = [r for r in enriched if r["sla_color"] == color]
    resumen = seguimiento_resumen()
    return render_template("seguimiento.html", active="seguimiento", rows=enriched, q=q, etapas=ETAPAS, etapa_sel=etapa, sla_sel=color, resumen=resumen)

@app.route("/agenda")
def agenda():
    rows = agenda_entrevistas(200)
    return render_template("agenda.html", active="agenda", rows=rows)

@app.route("/candidatos/<int:candidato_id>/nota", methods=["POST"])
def agregar_nota(candidato_id):
    nota = request.form.get("nota", "").strip()
    usuario = request.form.get("usuario", "RH").strip() or "RH"
    if nota:
        log(candidato_id, "Nota", nota, usuario)
        flash("Nota agregada a la bitácora.", "success")
    else:
        flash("Captura una nota antes de guardar.", "error")
    return redirect(url_for("candidato_detalle", candidato_id=candidato_id))

@app.route("/plantillas/candidatos")
def plantilla_candidatos():
    out = REPORT_DIR / "Plantilla_Candidatos_RyS.xlsx"
    cols = ["Nombre Completo", "TELEFONO", "Email", "REGION", "LOCALIDAD", "ALMACEN", "PUESTO", "RECLUTADOR", "FECHA DE ENTREVISTA", "ESTATUS", "MOTIVO DE RECHAZO"]
    pd.DataFrame(columns=cols).to_excel(out, index=False)
    return send_file(out, as_attachment=True)

@app.route("/plantillas/headcount")
def plantilla_headcount():
    out = REPORT_DIR / "Plantilla_Headcount.xlsx"
    pd.DataFrame(columns=["UDN", "PUESTO", "HC", "NOM"]).to_excel(out, index=False)
    return send_file(out, as_attachment=True)

@app.route("/headcount/importar", methods=["POST"])
def importar_headcount():
    f = request.files.get("archivo")
    if not f or not f.filename:
        flash("Selecciona un Excel de headcount.", "error"); return redirect(url_for("headcount"))
    dest = UPLOAD_DIR / secure_filename(f.filename); f.save(dest)
    try:
        total = import_headcount_excel(dest)
        flash(f"Headcount importado correctamente: {total} registros.", "success")
    except Exception as e:
        flash(f"Error al importar headcount: {e}", "error")
    return redirect(url_for("headcount"))

@app.route("/catalogos")
def catalogos():
    return render_template("catalogos.html", puestos=query("SELECT * FROM catalogo_puestos ORDER BY puesto"), udns=query("SELECT * FROM catalogo_udn ORDER BY almacen"), reclutadores=query("SELECT * FROM catalogo_reclutadores ORDER BY reclutador"))

@app.route("/headcount")
def headcount():
    rows = query("SELECT * FROM headcount ORDER BY udn, puesto LIMIT 1000")
    return render_template("headcount.html", cobertura=cobertura_por_puesto(30), kpis=get_kpis(), rows=rows)
@app.route("/ia")
def ia():
    return render_template("ia.html", candidatos=query("SELECT id,folio,nombre,puesto,udn,etapa FROM candidatos ORDER BY id DESC LIMIT 200"), puestos=query("SELECT puesto FROM catalogo_puestos ORDER BY puesto"), preguntas=None)
@app.route("/exportar")
def exportar():
    rows = query("SELECT * FROM candidatos ORDER BY id DESC")
    out = REPORT_DIR / f"Capital_Humano_Candidatos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    pd.DataFrame([dict(r) for r in rows]).to_excel(out, index=False)
    return send_file(out, as_attachment=True)
@app.route("/api/kpis")
def api_kpis(): return jsonify(get_kpis())


# ==============================
# FASE 2.7 - 3.0 - Administración, IA y reportes ejecutivos
# ==============================
def candidato_full(candidato_id):
    cand = query("SELECT * FROM candidatos WHERE id=?", (candidato_id,), one=True)
    if not cand:
        return None
    ruta = query("SELECT * FROM hoja_ruta WHERE candidato_id=?", (candidato_id,), one=True)
    oferta = query("SELECT * FROM oferta_laboral WHERE candidato_id=?", (candidato_id,), one=True)
    evals = query("SELECT * FROM evaluaciones WHERE candidato_id=?", (candidato_id,), one=True)
    entrevistas = query("SELECT * FROM entrevistas WHERE candidato_id=? ORDER BY id", (candidato_id,))
    docs = query("SELECT documento,estado FROM documentacion WHERE candidato_id=? ORDER BY id", (candidato_id,))
    return {"cand": cand, "ruta": ruta, "oferta": oferta, "evals": evals, "entrevistas": entrevistas, "docs": docs}


def generar_analisis_ia(candidato_id):
    data = candidato_full(candidato_id)
    if not data:
        return None
    c = data["cand"]
    ruta = data["ruta"]
    evals = data["evals"]
    docs = data["docs"]
    entrevistas = data["entrevistas"]
    score = 50
    fortalezas, riesgos, recomendaciones = [], [], []
    puesto = (c["puesto"] or "").upper()
    if ruta:
        for campo, etiqueta in [("handheld", "Handheld"), ("wms", "WMS"), ("sap", "SAP"), ("excel", "Excel"), ("patin", "Patín"), ("montacargas", "Montacargas")]:
            val = (ruta[campo] or "").upper() if campo in ruta.keys() else ""
            if val in ["SI", "SÍ", "YES", "1"]:
                score += 6
                fortalezas.append(f"Experiencia declarada en {etiqueta}.")
        if ruta["experiencia_json"]:
            score += 8
            fortalezas.append("Cuenta con antecedentes laborales capturados en hoja de ruta.")
        if (ruta["ruta_transporte"] or ""):
            score += 3
        else:
            riesgos.append("Falta validar ruta de transporte y distancia al sitio.")
    else:
        riesgos.append("No se ha capturado la Hoja de Ruta.")
    if evals:
        med = evals["medico_resultado"] or "Pendiente"
        tec = evals["tecnico_resultado"] or "Pendiente"
        if med == "Apto": score += 10; fortalezas.append("Resultado médico apto.")
        if med == "Apto Condicionado": riesgos.append("Resultado médico apto condicionado; revisar observaciones.")
        if med == "No Apto": score -= 30; riesgos.append("Resultado médico no apto.")
        if tec == "Apto": score += 10; fortalezas.append("Evaluación técnica apta.")
        if tec == "Apto Condicionado": riesgos.append("Evaluación técnica apta condicionada.")
        if tec == "No Apto": score -= 25; riesgos.append("Evaluación técnica no apta.")
    else:
        riesgos.append("No se ha capturado evaluación médica/técnica.")
    if entrevistas:
        score += min(10, len(entrevistas)*5)
        fortalezas.append(f"Tiene {len(entrevistas)} entrevista(s) registrada(s).")
    else:
        recomendaciones.append("Programar o capturar entrevista inicial.")
    if docs:
        ok = sum(1 for d in docs if d["estado"] == "Validado")
        total = len(docs)
        pct_docs = int(ok/total*100) if total else 0
        if pct_docs >= 80:
            score += 10
            fortalezas.append("Documentación con avance alto.")
        elif pct_docs < 40:
            riesgos.append("Documentación con avance bajo.")
    if "MONTAC" in puesto and not (ruta and (ruta["montacargas"] or "").upper() in ["SI", "SÍ"]):
        riesgos.append("El puesto requiere validar experiencia en montacargas.")
        recomendaciones.append("Aplicar prueba práctica de montacargas antes de avanzar.")
    if "PATIN" in puesto and not (ruta and (ruta["patin"] or "").upper() in ["SI", "SÍ"]):
        riesgos.append("El puesto requiere validar experiencia en patín.")
        recomendaciones.append("Aplicar prueba práctica de patín y seguridad operativa.")
    if not recomendaciones:
        recomendaciones.append("Continuar con la siguiente etapa del flujo operativo.")
    score = max(0, min(100, score))
    if score >= 80:
        nivel = "Alta compatibilidad"
    elif score >= 60:
        nivel = "Compatibilidad media"
    else:
        nivel = "Revisión requerida"
    resumen = f"{c['nombre']} presenta {nivel.lower()} para el puesto {c['puesto']} en {c['udn'] or 'UDN por definir'}."
    return {"score": score, "nivel": nivel, "resumen": resumen, "fortalezas": fortalezas[:6], "riesgos": riesgos[:6], "recomendaciones": recomendaciones[:6]}


@app.route("/ia/candidato/<int:candidato_id>")
def ia_candidato(candidato_id):
    cand = query("SELECT * FROM candidatos WHERE id=?", (candidato_id,), one=True)
    analisis = generar_analisis_ia(candidato_id)
    if not cand or not analisis:
        flash("Candidato no encontrado.", "error")
        return redirect(url_for("ia"))
    return render_template("ia_resultado.html", c=cand, analisis=analisis)


@app.route("/ia/preguntas", methods=["POST"])
def ia_preguntas():
    puesto = request.form.get("puesto", "GENERAL")
    tipo = request.form.get("tipo", "Entrevista 1")
    preguntas = preguntas_por_puesto(puesto, tipo)
    return render_template("ia.html", preguntas=preguntas, puesto_sel=puesto, tipo_sel=tipo, candidatos=query("SELECT id,folio,nombre,puesto,udn,etapa FROM candidatos ORDER BY id DESC LIMIT 200"), puestos=query("SELECT puesto FROM catalogo_puestos ORDER BY puesto"))


@app.route("/catalogos/agregar", methods=["POST"])
def catalogos_agregar():
    tipo = request.form.get("tipo")
    valor = (request.form.get("valor") or "").strip().upper()
    region = (request.form.get("region") or "").strip().upper()
    localidad = (request.form.get("localidad") or "").strip().upper()
    if not valor:
        flash("Captura un valor para el catálogo.", "error")
        return redirect(url_for("catalogos"))
    try:
        if tipo == "puesto":
            execute("INSERT INTO catalogo_puestos(puesto) VALUES(?)", (valor,))
        elif tipo == "udn":
            execute("INSERT INTO catalogo_udn(almacen,region,localidad) VALUES(?,?,?)", (valor, region or "POR DEFINIR", localidad or "POR DEFINIR"))
        elif tipo == "reclutador":
            execute("INSERT INTO catalogo_reclutadores(reclutador,region) VALUES(?,?)", (valor, region))
        flash("Catálogo actualizado correctamente.", "success")
    except Exception as e:
        flash(f"No se pudo agregar el registro. Puede que ya exista. Detalle: {e}", "error")
    return redirect(url_for("catalogos"))


@app.route("/headcount/agregar", methods=["POST"])
def headcount_agregar():
    udn = (request.form.get("udn") or "").strip().upper()
    puesto = (request.form.get("puesto") or "").strip().upper()
    try:
        requerido = int(float(request.form.get("requerido") or 0))
        activo = int(float(request.form.get("activo") or 0))
    except Exception:
        requerido, activo = 0, 0
    if not puesto:
        flash("Captura el puesto para agregar headcount.", "error")
        return redirect(url_for("headcount"))
    execute("INSERT INTO headcount(udn,puesto,requerido,activo) VALUES(?,?,?,?)", (udn, puesto, requerido, activo))
    flash("Headcount agregado manualmente.", "success")
    return redirect(url_for("headcount"))


def export_df(rows, filename):
    out = REPORT_DIR / filename
    pd.DataFrame([dict(r) for r in rows]).to_excel(out, index=False)
    return send_file(out, as_attachment=True)


@app.route("/reportes/pipeline")
def reporte_pipeline():
    rows = query("SELECT folio,nombre,telefono,email,region,localidad,udn,puesto,reclutador,etapa,estatus_macro,fecha_entrevista,fecha_registro,fecha_actualizacion,motivo_rechazo FROM candidatos ORDER BY id DESC")
    return export_df(rows, f"Reporte_Pipeline_CH_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")


@app.route("/reportes/documentacion")
def reporte_documentacion():
    rows = query("""
        SELECT c.folio,c.nombre,c.udn,c.puesto,c.etapa,d.documento,d.estado,d.updated_at
        FROM candidatos c LEFT JOIN documentacion d ON d.candidato_id=c.id
        ORDER BY c.nombre,d.documento
    """)
    return export_df(rows, f"Reporte_Documentacion_CH_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")


@app.route("/reportes/rfc")
def reporte_rfc():
    rows = query("""
        SELECT c.folio,c.nombre,c.udn,c.puesto,e.rfc_actualizado,e.costo_rfc,e.medico_resultado,e.tecnico_resultado,e.updated_at
        FROM candidatos c JOIN evaluaciones e ON e.candidato_id=c.id
        WHERE e.rfc_actualizado='Si'
        ORDER BY e.updated_at DESC
    """)
    return export_df(rows, f"Reporte_RFC_Actualizados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")


@app.route("/reportes/headcount")
def reporte_headcount():
    rows = query("SELECT * FROM headcount ORDER BY udn,puesto")
    return export_df(rows, f"Reporte_Headcount_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")



# ==============================
# FASE 2.7 a 3.0 - Expediente, DC3, firma, Access Kiosk y corporativo
# ==============================
def ensure_employee_id(candidato_id):
    row = query("SELECT * FROM access_kiosk_integracion WHERE candidato_id=?", (candidato_id,), one=True)
    if row and row["empleado_id"]:
        return row["empleado_id"]
    empleado_id = f"10{100000 + int(candidato_id):06d}"[-8:]
    execute("""INSERT INTO access_kiosk_integracion(candidato_id,empleado_id,qr_generado,estatus,fecha)
               VALUES(?,?,?,?,?) ON CONFLICT(candidato_id) DO UPDATE SET empleado_id=excluded.empleado_id, fecha=excluded.fecha""",
            (candidato_id, empleado_id, "No", "Pendiente", now_str()))
    return empleado_id

@app.route("/expediente-digital")
def expediente_digital_panel():
    rows = query("SELECT * FROM candidatos WHERE etapa IN ('Activo','Entrega Operaciones','Induccion','Documentacion') ORDER BY id DESC LIMIT 500")
    return render_template("expediente_digital.html", rows=rows, active="expediente", page_title="Expediente Digital")

@app.route("/expediente-digital/<int:candidato_id>", methods=["GET","POST"])
def expediente_digital(candidato_id):
    cand = query("SELECT * FROM candidatos WHERE id=?", (candidato_id,), one=True)
    if not cand: return redirect(url_for("expediente_digital_panel"))
    folder = UPLOAD_DIR / f"expediente_{candidato_id}"
    folder.mkdir(parents=True, exist_ok=True)
    if request.method == "POST":
        categoria = request.form.get("categoria", "General")
        documento = request.form.get("documento", "Documento")
        obs = request.form.get("observaciones", "")
        f = request.files.get("archivo")
        archivo_path = ""
        if f and f.filename:
            safe = secure_filename(f.filename)
            dest = folder / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{safe}"
            f.save(dest)
            archivo_path = str(dest)
        execute("INSERT INTO expediente_digital(candidato_id,categoria,documento,archivo,observaciones,fecha) VALUES(?,?,?,?,?,?)", (candidato_id,categoria,documento,archivo_path,obs,now_str()))
        log(candidato_id, "Expediente Digital", f"Documento agregado: {documento}", "RH")
        flash("Documento agregado al expediente.", "success")
        return redirect(url_for("expediente_digital", candidato_id=candidato_id))
    docs = query("SELECT * FROM expediente_digital WHERE candidato_id=? ORDER BY id DESC", (candidato_id,))
    return render_template("expediente_digital.html", c=cand, docs=docs, active="expediente", page_title="Expediente Digital")

@app.route("/capacitacion-dc3")
def capacitacion_dc3_panel():
    rows = query("SELECT * FROM candidatos WHERE etapa IN ('Induccion','Entrega Operaciones','Activo') ORDER BY id DESC LIMIT 500")
    cursos = query("SELECT cd.*, c.nombre, c.puesto, c.udn FROM capacitacion_dc3 cd JOIN candidatos c ON c.id=cd.candidato_id ORDER BY cd.id DESC LIMIT 50")
    return render_template("capacitacion_dc3.html", rows=rows, cursos=cursos, active="capacitacion", page_title="Capacitación y DC3")

@app.route("/capacitacion-dc3/<int:candidato_id>", methods=["GET","POST"])
def capacitacion_dc3(candidato_id):
    cand = query("SELECT * FROM candidatos WHERE id=?", (candidato_id,), one=True)
    if not cand: return redirect(url_for("capacitacion_dc3_panel"))
    if request.method == "POST":
        execute("INSERT INTO capacitacion_dc3(candidato_id,curso,fecha_curso,instructor,resultado,dc3_generado,observaciones,fecha) VALUES(?,?,?,?,?,?,?,?)",
                (candidato_id, request.form.get("curso",""), request.form.get("fecha_curso",""), request.form.get("instructor",""), request.form.get("resultado","Pendiente"), request.form.get("dc3_generado","No"), request.form.get("observaciones",""), now_str()))
        log(candidato_id, "Capacitación", f"Curso capturado: {request.form.get('curso','')}", "RH/SHE")
        flash("Curso/DC3 capturado correctamente.", "success")
        return redirect(url_for("capacitacion_dc3", candidato_id=candidato_id))
    cursos = query("SELECT * FROM capacitacion_dc3 WHERE candidato_id=? ORDER BY id DESC", (candidato_id,))
    return render_template("capacitacion_dc3.html", c=cand, cursos=cursos, active="capacitacion", page_title="Capacitación y DC3")

@app.route("/firma-electronica")
def firma_electronica_panel():
    rows = query("SELECT * FROM candidatos WHERE etapa IN ('Documentacion','Induccion','Entrega Operaciones','Activo') ORDER BY id DESC LIMIT 500")
    return render_template("firma_electronica.html", rows=rows, active="firma", page_title="Firma Electrónica")

@app.route("/firma-electronica/<int:candidato_id>", methods=["GET","POST"])
def firma_electronica(candidato_id):
    cand = query("SELECT * FROM candidatos WHERE id=?", (candidato_id,), one=True)
    if not cand: return redirect(url_for("firma_electronica_panel"))
    folder = UPLOAD_DIR / f"firmas_{candidato_id}"
    folder.mkdir(parents=True, exist_ok=True)
    if request.method == "POST":
        documento = request.form.get("documento", "Documento firmado")
        firmante = request.form.get("firmante", cand["nombre"])
        f = request.files.get("archivo_pdf")
        archivo_pdf = ""
        if f and f.filename:
            dest = folder / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{secure_filename(f.filename)}"
            f.save(dest); archivo_pdf = str(dest)
        execute("INSERT INTO firmas_electronicas(candidato_id,documento,firmante,archivo_pdf,fecha) VALUES(?,?,?,?,?)", (candidato_id,documento,firmante,archivo_pdf,now_str()))
        log(candidato_id, "Firma", f"Documento firmado/cargado: {documento}", "RH")
        flash("Documento firmado/cargado correctamente.", "success")
        return redirect(url_for("firma_electronica", candidato_id=candidato_id))
    firmas = query("SELECT * FROM firmas_electronicas WHERE candidato_id=? ORDER BY id DESC", (candidato_id,))
    return render_template("firma_electronica.html", c=cand, firmas=firmas, active="firma", page_title="Firma Electrónica")

@app.route("/access-kiosk")
def access_kiosk_panel():
    rows = query("""
        SELECT c.*, a.empleado_id, a.qr_generado, a.estatus AS kiosk_estatus
        FROM candidatos c LEFT JOIN access_kiosk_integracion a ON a.candidato_id=c.id
        WHERE c.etapa='Activo' OR c.estatus_macro='Activo'
        ORDER BY c.id DESC
    """)
    return render_template("access_kiosk.html", rows=rows, active="access", page_title="Integración Access Kiosk")

@app.route("/access-kiosk/generar/<int:candidato_id>")
def access_kiosk_generar(candidato_id):
    emp = ensure_employee_id(candidato_id)
    execute("UPDATE access_kiosk_integracion SET qr_generado='Si', estatus='Listo para sincronizar', fecha=? WHERE candidato_id=?", (now_str(), candidato_id))
    log(candidato_id, "Access Kiosk", f"Empleado/QR generado: {emp}", "Sistema")
    flash(f"Empleado {emp} listo para Access Kiosk.", "success")
    return redirect(url_for("access_kiosk_panel"))

@app.route("/dashboard-corporativo")
def dashboard_corporativo():
    k = get_kpis()
    cobertura = cobertura_por_puesto(12)
    rfc = query("SELECT COUNT(*) c, COALESCE(SUM(costo_rfc),0) total FROM evaluaciones WHERE rfc_actualizado='Si'", one=True)
    activos = query("SELECT COUNT(*) c FROM candidatos WHERE etapa='Activo' OR estatus_macro='Activo'", one=True)["c"]
    reclutamiento = query("SELECT COUNT(*) c FROM candidatos WHERE estatus_macro='Reclutamiento'", one=True)["c"]
    contratacion = query("SELECT COUNT(*) c FROM candidatos WHERE estatus_macro='Contratacion'", one=True)["c"]
    return render_template("dashboard_corporativo.html", kpis=k, cobertura=cobertura, rfc=rfc, activos=activos, reclutamiento=reclutamiento, contratacion=contratacion, active="corp", page_title="Dashboard Ejecutivo Corporativo")

@app.route("/ia-avanzada")
def ia_avanzada():
    rows = query("SELECT id, folio, nombre, puesto, udn, etapa FROM candidatos ORDER BY id DESC LIMIT 200")
    return render_template("ia_avanzada.html", rows=rows, active="ia", page_title="IA Avanzada Capital Humano")



# ==============================
# FASE 3.1 - Administración, roles, UDN y auditoría
# ==============================
@app.route("/admin")
@admin_required
def admin_panel():
    usuarios = query("SELECT id,nombre,usuario,email,rol,udn_permitidas,activo,ultimo_acceso,fecha_creacion FROM usuarios_sistema ORDER BY id")
    roles = query("SELECT * FROM roles_sistema ORDER BY rol")
    udns = query("SELECT almacen FROM catalogo_udn WHERE activo=1 ORDER BY almacen")
    auditoria = query("SELECT * FROM auditoria_sistema ORDER BY id DESC LIMIT 100")
    return render_template("admin.html", active="admin", page_title="Administración", usuarios=usuarios, roles=roles, udns=udns, auditoria=auditoria)


@app.route("/admin/usuarios/nuevo", methods=["POST"])
@admin_required
def admin_usuario_nuevo():
    nombre = request.form.get("nombre", "").strip()
    usuario = request.form.get("usuario", "").strip()
    email = request.form.get("email", "").strip()
    password = request.form.get("password", "").strip()
    rol = request.form.get("rol", "Consulta").strip()
    udn_list = request.form.getlist("udn")
    udn_permitidas = "TODAS" if "TODAS" in udn_list or not udn_list else "|".join(udn_list)
    if not nombre or not usuario or not password:
        flash("Captura nombre, usuario y contraseña.", "error")
        return redirect(url_for("admin_panel"))
    try:
        execute("""INSERT INTO usuarios_sistema(nombre,usuario,email,password_hash,rol,udn_permitidas,activo,fecha_creacion)
                   VALUES(?,?,?,?,?,?,1,?)""", (nombre, usuario, email, generate_password_hash(password), rol, udn_permitidas, now_str()))
        audit("Alta usuario", "Administración", f"{usuario} - {rol}")
        flash("Usuario creado correctamente.", "success")
    except Exception as e:
        flash(f"No se pudo crear usuario. Puede que el usuario ya exista. Detalle: {e}", "error")
    return redirect(url_for("admin_panel"))


@app.route("/admin/usuarios/<int:user_id>/actualizar", methods=["POST"])
@admin_required
def admin_usuario_actualizar(user_id):
    rol = request.form.get("rol", "Consulta").strip()
    activo = 1 if request.form.get("activo") == "1" else 0
    udn_list = request.form.getlist("udn")
    udn_permitidas = "TODAS" if "TODAS" in udn_list or not udn_list else "|".join(udn_list)
    execute("UPDATE usuarios_sistema SET rol=?, udn_permitidas=?, activo=? WHERE id=?", (rol, udn_permitidas, activo, user_id))
    audit("Actualiza usuario", "Administración", f"ID {user_id} rol={rol} activo={activo}")
    flash("Usuario actualizado.", "success")
    return redirect(url_for("admin_panel"))


@app.route("/admin/usuarios/<int:user_id>/reset", methods=["POST"])
@admin_required
def admin_usuario_reset(user_id):
    new_pass = request.form.get("password", "").strip()
    if not new_pass:
        flash("Captura una nueva contraseña.", "error")
        return redirect(url_for("admin_panel"))
    execute("UPDATE usuarios_sistema SET password_hash=? WHERE id=?", (generate_password_hash(new_pass), user_id))
    audit("Reset password", "Administración", f"ID {user_id}")
    flash("Contraseña actualizada.", "success")
    return redirect(url_for("admin_panel"))


@app.route("/admin/roles/agregar", methods=["POST"])
@admin_required
def admin_roles_agregar():
    rol = request.form.get("rol", "").strip()
    descripcion = request.form.get("descripcion", "").strip()
    permisos = request.form.get("permisos", "").strip()
    if not rol:
        flash("Captura el nombre del rol.", "error")
        return redirect(url_for("admin_panel"))
    try:
        permisos_json = json.dumps([p.strip() for p in permisos.split(",") if p.strip()], ensure_ascii=False)
        execute("INSERT INTO roles_sistema(rol,descripcion,permisos_json,activo) VALUES(?,?,?,1)", (rol, descripcion, permisos_json))
        audit("Alta rol", "Administración", rol)
        flash("Rol creado correctamente.", "success")
    except Exception as e:
        flash(f"No se pudo crear el rol: {e}", "error")
    return redirect(url_for("admin_panel"))


@app.route("/admin/auditoria/exportar")
@admin_required
def admin_auditoria_exportar():
    rows = query("SELECT * FROM auditoria_sistema ORDER BY id DESC")
    return export_df(rows, f"Auditoria_Plataforma_CH_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

# Inicialización compatible con Render/Gunicorn y ejecución local.
# Gunicorn importa app.py, por eso la base se prepara al cargar el módulo.
try:
    init_db()
    print("[OK] Plataforma CH inicializada correctamente.")
except Exception as e:
    print(f"[WARN] No se pudo inicializar la base al arrancar: {e}")

if __name__ == "__main__":
    print("="*70); print("PLATAFORMA INTEGRAL DE CAPITAL HUMANO TRAXION - FASE 3.1 SEGURIDAD"); print("="*70)
    print("[3/4] Preparando servidor Flask..."); print("[4/4] Abre en tu navegador: http://127.0.0.1:5000")
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False, use_reloader=False)
