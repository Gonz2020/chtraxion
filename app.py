# -*- coding: utf-8 -*-
"""
TRAXION Access Kiosk Definitivo
Web Flask + SQLite + Scanner QR USB + Cámara OpenCV/pyzbar opcional.

Flujo recomendado:
- Producción: scanner QR USB HID.
- Pruebas / respaldo: botón Cámara OpenCV.
"""

import os
import re
import sqlite3
import uuid
from datetime import datetime, date, time as dtime
from pathlib import Path

import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file
import qrcode

try:
    import cv2
    from pyzbar.pyzbar import decode
except Exception:
    cv2 = None
    decode = None

try:
    import pyttsx3
except Exception:
    pyttsx3 = None


APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "Data"
REPORT_DIR = APP_DIR / "Reportes"
QR_DIR = APP_DIR / "QR_Visitantes"
PHOTO_DIR = APP_DIR / "static" / "photos"
DB_PATH = DATA_DIR / "traxion_access.sqlite"
EXCEL_EMPLEADOS = DATA_DIR / "Base_Credenciales_Traxion_actualizado.xlsx"

for d in [DATA_DIR, REPORT_DIR, QR_DIR, PHOTO_DIR]:
    d.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)


# ==============================
# DB
# ==============================
def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def column_exists(conn, table, column):
    rows = conn.execute(f"PRAGMA table_info({table})").fetchall()
    return any(r["name"] == column for r in rows)


def ensure_column(conn, table, column, sql_type):
    if not column_exists(conn, table, column):
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {sql_type}")


def init_db():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS empleados (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            empleado TEXT UNIQUE NOT NULL,
            nombre TEXT,
            apellido TEXT,
            nombre_completo TEXT,
            puesto TEXT,
            area TEXT,
            turno TEXT,
            tipo_personal TEXT DEFAULT 'OPERATIVO',
            archivo_foto TEXT,
            activo INTEGER DEFAULT 1,
            fecha_importacion TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS registros (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            empleado TEXT,
            nombre_completo TEXT,
            fecha TEXT,
            hora TEXT,
            tipo_evento TEXT,
            turno_detectado TEXT,
            estatus TEXT,
            mensaje TEXT,
            origen TEXT DEFAULT 'SCANNER',
            raw_qr TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS visitantes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            qr_id TEXT UNIQUE,
            nombre TEXT,
            empresa TEXT,
            destino_empresa TEXT DEFAULT 'TRAXION',
            tipo_visita TEXT,
            persona_visita TEXT,
            motivo TEXT,
            activo INTEGER DEFAULT 1,
            fecha_registro TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS visitas_registros (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            qr_id TEXT,
            nombre TEXT,
            empresa TEXT,
            destino_empresa TEXT DEFAULT 'TRAXION',
            tipo_visita TEXT,
            persona_visita TEXT,
            fecha TEXT,
            hora TEXT,
            tipo_evento TEXT,
            mensaje TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Migracion segura para bases existentes
    ensure_column(conn, "visitantes", "destino_empresa", "TEXT DEFAULT 'TRAXION'")
    ensure_column(conn, "visitas_registros", "destino_empresa", "TEXT DEFAULT 'TRAXION'")
    ensure_column(conn, "visitas_registros", "tipo_visita", "TEXT")
    ensure_column(conn, "visitas_registros", "persona_visita", "TEXT")

    conn.commit()
    conn.close()



def normalizar_empleado(valor):
    """Normaliza el numero de empleado para que Excel/QR/SQLite coincidan.
    - Acepta 10116507, 10116507.0, ID:10116507
    - Prioriza empleados Traxion que inician con 10 u 11 y tienen 8 digitos
    """
    if valor is None:
        return ""
    txt = str(valor).strip()
    if not txt or txt.lower() == "nan":
        return ""

    # Si Excel lo lee como numero flotante: 10116507.0
    m_float = re.fullmatch(r"(\d+)\.0+", txt)
    if m_float:
        txt = m_float.group(1)

    # Quitar espacios, guiones y separadores comunes
    txt_compacto = re.sub(r"[^0-9]", "", txt)

    # Buscar empleado valido: inicia 10 u 11 y tiene 8 digitos
    m_emp = re.search(r"(?:10|11)\d{6}", txt_compacto)
    if m_emp:
        return m_emp.group(0)

    # Respaldo: si solo trae digitos, devolverlos sin .0
    if txt_compacto:
        return txt_compacto

    return txt

def parse_bool_activo(valor):
    if pd.isna(valor):
        return 1
    txt = str(valor).strip().upper()
    if txt in ["BAJA", "INACTIVO", "NO", "0", "FALSE", "FALSO", "CANCELADO"]:
        return 0
    return 1


def detectar_tipo_personal(row):
    texto = " ".join(str(v) for v in row.to_dict().values()).upper()
    if "ADMIN" in texto or "ADMINISTRATIVO" in texto or "ESPECIAL" in texto:
        return "ADMINISTRATIVO"
    return "OPERATIVO"


def importar_empleados_excel(path=EXCEL_EMPLEADOS):
    if not Path(path).exists():
        return 0

    df = pd.read_excel(path)
    colmap = {str(c).lower().strip(): c for c in df.columns}

    def col(*names):
        for n in names:
            if n.lower() in colmap:
                return colmap[n.lower()]
        return None

    c_emp = col("empleado", "num_empleado", "id", "id empleado", "numero empleado")
    c_nom = col("nombre", "name")
    c_ape = col("apellido", "apellidos", "apellido paterno")
    c_puesto = col("puesto", "posicion", "cargo")
    c_area = col("area", "departamento")
    c_turno = col("turno", "shift")
    c_foto = col("archivo_foto", "foto", "photofile")
    c_estatus = col("estatus", "status", "activo", "estado")

    if not c_emp:
        raise ValueError("No se encontró columna 'empleado' en el Excel.")

    conn = get_conn()
    cur = conn.cursor()
    total = 0

    for _, row in df.iterrows():
        empleado = normalizar_empleado(row.get(c_emp, ""))
        if not empleado:
            continue

        nombre = str(row.get(c_nom, "")).strip() if c_nom else ""
        apellido = str(row.get(c_ape, "")).strip() if c_ape else ""
        puesto = str(row.get(c_puesto, "")).strip() if c_puesto else ""
        area = str(row.get(c_area, "")).strip() if c_area else ""
        turno = str(row.get(c_turno, "")).strip() if c_turno else ""
        foto = str(row.get(c_foto, "")).strip() if c_foto else ""
        activo = parse_bool_activo(row.get(c_estatus, "")) if c_estatus else 1
        nombre_completo = f"{nombre} {apellido}".strip() or empleado
        tipo_personal = detectar_tipo_personal(row)

        cur.execute("""
            INSERT INTO empleados(empleado,nombre,apellido,nombre_completo,puesto,area,turno,tipo_personal,archivo_foto,activo)
            VALUES(?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(empleado) DO UPDATE SET
                nombre=excluded.nombre,
                apellido=excluded.apellido,
                nombre_completo=excluded.nombre_completo,
                puesto=excluded.puesto,
                area=excluded.area,
                turno=excluded.turno,
                tipo_personal=excluded.tipo_personal,
                archivo_foto=excluded.archivo_foto,
                activo=excluded.activo,
                fecha_importacion=CURRENT_TIMESTAMP
        """, (empleado, nombre, apellido, nombre_completo, puesto, area, turno, tipo_personal, foto, activo))
        total += 1

    conn.commit()
    conn.close()
    return total


# ==============================
# Reglas
# ==============================
def limpiar_qr(valor):
    """
    Limpia la lectura del scanner/camara y devuelve:
    - VISITA:XXXXXXXXXXXX cuando sea QR de visitante/proveedor.
    - Numero de empleado Traxion de 8 digitos que inicia con 10 o 11.

    Importante:
    Muchos gafetes traen fechas/vigencias como 2026. Antes el sistema tomaba
    el primer numero encontrado y por eso buscaba "2026" en empleados.
    Esta version primero busca patrones de empleado 10xxxxxx / 11xxxxxx.
    """
    if valor is None:
        return ""

    txt_original = str(valor).strip()
    if not txt_original:
        return ""

    txt = txt_original.replace("\\r", " ").replace("\\n", " ")
    txt = txt.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    txt = re.sub(r"\s+", " ", txt).strip()
    upper = txt.upper()

    # QR de visita/proveedor generado por esta aplicacion
    m_visita = re.search(r"VISITA\s*[:=\-]\s*([A-Z0-9]{8,40})", upper)
    if m_visita:
        return "VISITA:" + m_visita.group(1).strip()
    if upper.startswith("VISITA:"):
        return upper

    # Patrones directos: ID:10116507, Empleado:10116507, etc.
    patrones = [
        r'"EMPLEADO"\s*:\s*"?((?:10|11)\d{6})"?',
        r"'EMPLEADO'\s*:\s*'?((?:10|11)\d{6})'?",
        r"\bEMPLEADO\b\s*[:=|\- ]+\s*((?:10|11)\d{6})",
        r"\bNUM[_ ]?EMPLEADO\b\s*[:=|\- ]+\s*((?:10|11)\d{6})",
        r"\bNO[_ ]?EMPLEADO\b\s*[:=|\- ]+\s*((?:10|11)\d{6})",
        r"\bID\b\s*[:=|\- ]+\s*((?:10|11)\d{6})",
        r"\bEMP\b\s*[:=|\- ]+\s*((?:10|11)\d{6})",
        r"\bCLAVE\b\s*[:=|\- ]+\s*((?:10|11)\d{6})",
    ]
    for patron in patrones:
        m = re.search(patron, upper, flags=re.IGNORECASE)
        if m:
            return m.group(1).strip()

    # Busqueda global: toma solo numeros de empleado que empiezan con 10 o 11.
    # Sirve para lecturas largas del QR con CURP, NSS, vigencia, nombre, etc.
    m = re.search(r"(?<!\d)((?:10|11)\d{6})(?!\d)", upper)
    if m:
        return m.group(1).strip()

    # Si el scanner concatena todo sin espacios, buscar dentro de todos los digitos.
    compact_digits = re.sub(r"\D", "", upper)
    m = re.search(r"(?:10|11)\d{6}", compact_digits)
    if m:
        return m.group(0).strip()

    # Respaldo para pruebas: aceptar 8 digitos exactos, pero evitar 2026 y fechas cortas.
    m = re.fullmatch(r"\D*(\d{8})\D*", upper)
    if m:
        return m.group(1).strip()

    return ""
def entre(h, ini, fin):
    return ini <= h <= fin


def tiene_entrada_manana(conn, empleado, fecha):
    row = conn.execute("""
        SELECT id FROM registros
        WHERE empleado=? AND fecha=? AND tipo_evento='ENTRADA'
          AND hora BETWEEN '05:00:00' AND '10:59:59'
        LIMIT 1
    """, (empleado, fecha)).fetchone()
    return row is not None


def ultimo_evento_hoy(conn, empleado, fecha):
    return conn.execute("""
        SELECT * FROM registros
        WHERE empleado=? AND fecha=? AND estatus IN ('OK','RETARDO','FUERA HORARIO')
        ORDER BY id DESC LIMIT 1
    """, (empleado, fecha)).fetchone()


def decidir_evento(conn, emp, fecha, now):
    empleado = emp["empleado"]
    tipo_personal = (emp["tipo_personal"] or "OPERATIVO").upper()
    h = now.time()
    ultimo = ultimo_evento_hoy(conn, empleado, fecha)

    if tipo_personal in ["ADMINISTRATIVO", "ESPECIAL", "ADMIN"]:
        if ultimo and ultimo["tipo_evento"] == "ENTRADA":
            return "SALIDA", "ADMINISTRATIVO", "OK", "Hasta luego"
        return "ENTRADA", "ADMINISTRATIVO", "OK", "Bienvenido"

    if entre(h, dtime(5, 40), dtime(6, 15)):
        if ultimo and ultimo["tipo_evento"] == "ENTRADA":
            return "DUPLICADO", "TURNO 1", "ALERTA", "Ya tienes entrada registrada"
        return "ENTRADA", "TURNO 1", "OK", "Bienvenido"

    if entre(h, dtime(6, 16), dtime(11, 59)):
        if ultimo and ultimo["tipo_evento"] == "ENTRADA":
            return "DUPLICADO", "TURNO 1", "ALERTA", "Ya tienes entrada registrada"
        return "ENTRADA", "TURNO 1", "RETARDO", "Entrada con retardo"

    if entre(h, dtime(13, 30), dtime(14, 30)):
        if tiene_entrada_manana(conn, empleado, fecha):
            if ultimo and ultimo["tipo_evento"] == "SALIDA":
                return "DUPLICADO", "TURNO 1", "ALERTA", "Salida ya registrada"
            return "SALIDA", "TURNO 1", "OK", "Hasta luego"
        estatus = "OK" if h <= dtime(14, 15) else "RETARDO"
        return "ENTRADA", "TURNO 2", estatus, "Bienvenido" if estatus == "OK" else "Entrada con retardo"

    if entre(h, dtime(14, 31), dtime(16, 30)):
        if ultimo and ultimo["tipo_evento"] == "ENTRADA":
            return "SALIDA", ultimo["turno_detectado"] or "TURNO 1", "OK", "Hasta luego"
        return "ENTRADA", "TURNO 2", "RETARDO", "Entrada con retardo"

    if entre(h, dtime(21, 0), dtime(23, 59)):
        if ultimo and ultimo["tipo_evento"] == "ENTRADA":
            return "SALIDA", ultimo["turno_detectado"] or "TURNO 2", "OK", "Hasta luego"
        return "ALERTA", "TURNO 2", "ALERTA", "No existe entrada activa"

    if ultimo and ultimo["tipo_evento"] == "ENTRADA":
        return "SALIDA", ultimo["turno_detectado"] or "NO DEFINIDO", "FUERA HORARIO", "Hasta luego"
    return "ENTRADA", "NO DEFINIDO", "FUERA HORARIO", "Registro fuera de horario"


def speak_async(text):
    if not pyttsx3 or not text:
        return
    import threading
    def run():
        try:
            engine = pyttsx3.init()
            engine.setProperty("rate", 165)
            engine.say(text)
            engine.runAndWait()
        except Exception:
            pass
    threading.Thread(target=run, daemon=True).start()


def registrar_scan(raw_qr, origen="SCANNER"):

    codigo = limpiar_qr(raw_qr)
    if codigo and not codigo.upper().startswith("VISITA:"):
        codigo = normalizar_empleado(codigo)

    # ==========================================
    # Ignorar lecturas que no contienen empleado
    # ==========================================
    if not codigo:
        return {
            "ignore": True
        }

    fecha = datetime.now().strftime("%Y-%m-%d")
    hora = datetime.now().strftime("%H:%M:%S")
    now = datetime.now()

    conn = get_conn()


    if codigo.upper().startswith("VISITA:"):
        return registrar_visitante_scan(conn, codigo)

    emp = conn.execute("SELECT * FROM empleados WHERE empleado=?", (codigo,)).fetchone()

    if not emp:
        msg = "Llama a Capital Humano para su asistencia"
        conn.execute("""
            INSERT INTO registros(empleado,nombre_completo,fecha,hora,tipo_evento,turno_detectado,estatus,mensaje,origen,raw_qr)
            VALUES(?,?,?,?,?,?,?,?,?,?)
        """, (codigo, "", fecha, hora, "NO ENCONTRADO", "", "ALERTA", msg, origen, raw_qr))
        conn.commit()
        conn.close()
        speak_async(msg)
        return {"ok": False, "status": "NO ENCONTRADO", "message": msg, "event_type": "NO ENCONTRADO", "time": hora}

    if int(emp["activo"] or 0) == 0:
        msg = "Acceso denegado"
        conn.execute("""
            INSERT INTO registros(empleado,nombre_completo,fecha,hora,tipo_evento,turno_detectado,estatus,mensaje,origen,raw_qr)
            VALUES(?,?,?,?,?,?,?,?,?,?)
        """, (codigo, emp["nombre_completo"], fecha, hora, "DENEGADO", "", "ALERTA", msg, origen, raw_qr))
        conn.commit()
        data = dict(emp)
        conn.close()
        speak_async(msg)
        return {"ok": False, "status": "ACCESO DENEGADO", "message": msg, "employee": data, "event_type": "DENEGADO", "time": hora}

    reciente = ultimo_registro_reciente(conn, codigo, minutos=5)
    if reciente:
        msg = "Ya fue registrado hace menos de 5 minutos"
        data = dict(emp)
        conn.close()
        speak_async(msg)
        return {
            "ok": False,
            "duplicate": True,
            "employee": data,
            "event_type": "DUPLICADO",
            "shift": reciente["turno_detectado"],
            "status": "DUPLICADO",
            "message": msg,
            "time": hora
        }

    tipo, turno, estatus, mensaje = decidir_evento(conn, emp, fecha, now)

    conn.execute("""
        INSERT INTO registros(empleado,nombre_completo,fecha,hora,tipo_evento,turno_detectado,estatus,mensaje,origen,raw_qr)
        VALUES(?,?,?,?,?,?,?,?,?,?)
    """, (codigo, emp["nombre_completo"], fecha, hora, tipo, turno, estatus, mensaje, origen, raw_qr))
    conn.commit()
    data = dict(emp)
    conn.close()

    text_voice = f"{mensaje}, {emp['nombre_completo']}" if tipo in ["ENTRADA", "SALIDA"] else mensaje
    speak_async(text_voice)

    return {
        "ok": estatus != "ALERTA",
        "employee": data,
        "event_type": tipo,
        "shift": turno,
        "status": estatus,
        "message": text_voice,
        "time": hora
    }


def registrar_visitante_scan(conn, qr_id):
    fecha = datetime.now().strftime("%Y-%m-%d")
    hora = datetime.now().strftime("%H:%M:%S")

    vis = conn.execute("SELECT * FROM visitantes WHERE qr_id=? AND activo=1", (qr_id,)).fetchone()
    if not vis:
        conn.close()
        return {
            "ok": False,
            "status": "VISITA NO VALIDA",
            "message": "QR de visitante no registrado o inactivo",
            "event_type": "DENEGADO",
            "time": hora
        }

    ultimo = conn.execute("""
        SELECT * FROM visitas_registros
        WHERE qr_id=? AND fecha=?
        ORDER BY id DESC LIMIT 1
    """, (qr_id, fecha)).fetchone()

    tipo = "ENTRADA" if not ultimo or ultimo["tipo_evento"] == "SALIDA" else "SALIDA"
    msg = "Bienvenido" if tipo == "ENTRADA" else "Hasta luego"

    destino_empresa = vis["destino_empresa"] if "destino_empresa" in vis.keys() and vis["destino_empresa"] else "TRAXION"
    tipo_visita = vis["tipo_visita"] if "tipo_visita" in vis.keys() else "VISITA"
    persona_visita = vis["persona_visita"] if "persona_visita" in vis.keys() else ""

    conn.execute("""
        INSERT INTO visitas_registros(
            qr_id,nombre,empresa,destino_empresa,tipo_visita,persona_visita,
            fecha,hora,tipo_evento,mensaje
        )
        VALUES(?,?,?,?,?,?,?,?,?,?)
    """, (
        qr_id,
        vis["nombre"],
        vis["empresa"],
        destino_empresa,
        tipo_visita,
        persona_visita,
        fecha,
        hora,
        tipo,
        msg
    ))

    conn.commit()
    data = dict(vis)
    conn.close()

    speak_async(f"{msg}, {vis['nombre']}")

    return {
        "ok": True,
        "visitor": data,
        "event_type": tipo,
        "shift": "VISITANTE",
        "status": "OK",
        "message": f"{msg}, {vis['nombre']}",
        "time": hora
    }



def ultimo_registro_reciente(conn, empleado, minutos=5):
    """
    Valida si el empleado ya fue registrado en los ultimos N minutos.
    Si existe, NO se debe guardar otro registro.
    """
    row = conn.execute("""
        SELECT *
        FROM registros
        WHERE empleado=?
          AND datetime(created_at) >= datetime('now', ?)
          AND estatus IN ('OK','RETARDO','FUERA HORARIO')
        ORDER BY id DESC
        LIMIT 1
    """, (empleado, f"-{minutos} minutes")).fetchone()
    return row

# ==============================
# Camera OpenCV
# ==============================
def leer_qr_camara_opencv(timeout=20):
    if cv2 is None or decode is None:
        return {"ok": False, "message": "Faltan librerías opencv-python o pyzbar."}

    cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
    if not cap or not cap.isOpened():
        return {"ok": False, "message": "No se pudo abrir la cámara. Revisa si está ocupada."}

    start = datetime.now()
    codigo = None

    try:
        while (datetime.now() - start).total_seconds() < timeout:
            ret, frame = cap.read()
            if not ret:
                continue

            codigos = decode(frame)
            if codigos:
                codigo = codigos[0].data.decode("utf-8", errors="ignore")
                break

            # Ventana local para pruebas
            cv2.putText(frame, "Prueba camara QR - ESC para cancelar", (20, 35),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 255, 255), 2)
            cv2.imshow("TRAXION - Camara QR Prueba", frame)
            if cv2.waitKey(1) & 0xFF == 27:
                break
    finally:
        cap.release()
        cv2.destroyAllWindows()

    if not codigo:
        return {"ok": False, "message": "No se leyó ningún QR."}

    result = registrar_scan(codigo, origen="CAMARA_OPENCV")
    result["camera_qr"] = codigo
    return result


# ==============================
# Routes
# ==============================
@app.route("/")
def index():
    return render_template("index.html", stats=get_stats())



@app.route("/api/test_parser", methods=["POST"])
def api_test_parser():
    data = request.get_json(force=True)
    raw = data.get("raw", "")
    empleado = limpiar_qr(raw)
    if empleado and not empleado.upper().startswith("VISITA:"):
        empleado = normalizar_empleado(empleado)
    conn = get_conn()
    emp = conn.execute("SELECT empleado,nombre_completo,puesto,area,turno,activo FROM empleados WHERE empleado=?", (empleado,)).fetchone()
    conn.close()
    return jsonify({
        "raw": raw,
        "empleado_detectado": empleado,
        "encontrado": emp is not None,
        "empleado": dict(emp) if emp else None
    })

@app.route("/api/scan", methods=["POST"])
def api_scan():
    data = request.get_json(force=True)
    scan = data.get("scan", "")
    return jsonify(registrar_scan(scan, origen="SCANNER_WEB"))


@app.route("/api/camera_scan", methods=["POST"])
def api_camera_scan():
    return jsonify(leer_qr_camara_opencv())


@app.route("/api/stats")
def api_stats():
    return jsonify(get_stats())


@app.route("/api/recent")
def api_recent():
    conn = get_conn()
    fecha = date.today().strftime("%Y-%m-%d")

    rows = conn.execute("""
        SELECT
            hora,
            empleado,
            nombre_completo,
            tipo_evento,
            turno_detectado,
            estatus,
            mensaje,
            origen,
            created_at
        FROM registros
        WHERE fecha=?

        UNION ALL

        SELECT
            hora,
            qr_id AS empleado,
            nombre || ' | ' || COALESCE(tipo_visita,'VISITA') || ' | Empresa visita: ' || COALESCE(destino_empresa,'TRAXION') AS nombre_completo,
            tipo_evento,
            'VISITANTE' AS turno_detectado,
            'OK' AS estatus,
            mensaje,
            'QR_VISITANTE' AS origen,
            created_at
        FROM visitas_registros
        WHERE fecha=?

        ORDER BY created_at DESC
        LIMIT 50
    """, (fecha, fecha)).fetchall()

    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/visitas")
def api_visitas():
    fecha = request.args.get("fecha") or date.today().strftime("%Y-%m-%d")
    conn = get_conn()
    rows = conn.execute("""
        SELECT
            fecha,
            hora,
            qr_id,
            nombre,
            empresa,
            destino_empresa,
            tipo_visita,
            persona_visita,
            tipo_evento,
            mensaje,
            created_at
        FROM visitas_registros
        WHERE fecha=?
        ORDER BY created_at DESC
    """, (fecha,)).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/visitas")
def visitas():
    fecha = request.args.get("fecha") or date.today().strftime("%Y-%m-%d")
    conn = get_conn()
    rows = conn.execute("""
        SELECT
            fecha,
            hora,
            qr_id,
            nombre,
            empresa,
            destino_empresa,
            tipo_visita,
            persona_visita,
            tipo_evento,
            mensaje
        FROM visitas_registros
        WHERE fecha=?
        ORDER BY hora DESC
    """, (fecha,)).fetchall()
    conn.close()
    return render_template("visitas.html", fecha=fecha, visitas=[dict(r) for r in rows])


@app.route("/exportar_visitas")
def exportar_visitas():
    fecha = request.args.get("fecha") or date.today().strftime("%Y-%m-%d")
    conn = get_conn()
    rows = conn.execute("""
        SELECT fecha,hora,qr_id,nombre,empresa,destino_empresa,tipo_visita,persona_visita,tipo_evento,mensaje,created_at
        FROM visitas_registros
        WHERE fecha=?
        ORDER BY hora ASC
    """, (fecha,)).fetchall()
    conn.close()

    df = pd.DataFrame([dict(r) for r in rows])
    out = REPORT_DIR / f"Visitas_Proveedores_{fecha.replace('-','')}.xlsx"
    df.to_excel(out, index=False)
    return send_file(out, as_attachment=True)



# =====================================================
# SOLO PRUEBAS / DESARROLLO
# DESHABILITAR EN PRODUCCION
# Esta ruta elimina los registros de asistencia y visitas.
# En produccion comentar desde @app.route hasta el return.
# =====================================================
@app.route("/api/admin/limpiar_registros", methods=["POST"])
def api_limpiar_registros():
    conn = get_conn()

    conn.execute("DELETE FROM registros")
    conn.execute("DELETE FROM visitas_registros")

    try:
        conn.execute("DELETE FROM sqlite_sequence WHERE name='registros'")
        conn.execute("DELETE FROM sqlite_sequence WHERE name='visitas_registros'")
    except Exception:
        pass

    conn.commit()
    conn.close()

    return jsonify({
        "ok": True,
        "message": "Registros eliminados correctamente"
    })


@app.route("/parser-test")
def parser_test():
    return render_template("parser_test.html")

@app.route("/visitantes", methods=["GET", "POST"])
def visitantes():
    if request.method == "POST":
        nombre = request.form.get("nombre", "").strip()
        empresa = request.form.get("empresa", "").strip()
        destino_empresa = request.form.get("destino_empresa", "TRAXION").strip().upper()
        tipo_visita = request.form.get("tipo_visita", "VISITA").strip().upper()
        persona_visita = request.form.get("persona_visita", "").strip()
        motivo = request.form.get("motivo", "").strip()

        if destino_empresa not in ["TRAXION", "HENKEL"]:
            destino_empresa = "TRAXION"

        if tipo_visita not in ["VISITA", "PROVEEDOR", "CONTRATISTA", "CLIENTE", "AUDITORIA", "OTRO"]:
            tipo_visita = "VISITA"

        if nombre:
            qr_id = "VISITA:" + uuid.uuid4().hex[:12].upper()
            conn = get_conn()
            conn.execute("""
                INSERT INTO visitantes(qr_id,nombre,empresa,destino_empresa,tipo_visita,persona_visita,motivo)
                VALUES(?,?,?,?,?,?,?)
            """, (qr_id, nombre, empresa, destino_empresa, tipo_visita, persona_visita, motivo))
            conn.commit()
            conn.close()

            qr_img = qrcode.make(qr_id)
            qr_path = QR_DIR / f"{qr_id.replace(':','_')}.png"
            qr_img.save(qr_path)

            return render_template(
                "visitantes.html",
                generado=True,
                qr_id=qr_id,
                qr_file=qr_path.name,
                nombre=nombre,
                empresa=empresa,
                destino_empresa=destino_empresa,
                tipo_visita=tipo_visita,
                persona_visita=persona_visita,
                motivo=motivo
            )

    return render_template("visitantes.html", generado=False)


@app.route("/qr/<filename>")
def qr_file(filename):
    return send_file(QR_DIR / filename, mimetype="image/png")



@app.route("/reportes")
def reportes():
    fecha = request.args.get("fecha") or date.today().strftime("%Y-%m-%d")
    conn = get_conn()
    rows = conn.execute("""
        SELECT fecha,hora,empleado,nombre_completo,tipo_evento,turno_detectado,estatus,mensaje,origen
        FROM registros
        WHERE fecha=?
        ORDER BY hora ASC
    """, (fecha,)).fetchall()
    conn.close()
    return render_template("reportes.html", fecha=fecha, registros=[dict(r) for r in rows])


@app.route("/api/reportes")
def api_reportes():
    fecha = request.args.get("fecha") or date.today().strftime("%Y-%m-%d")
    conn = get_conn()
    rows = conn.execute("""
        SELECT fecha,hora,empleado,nombre_completo,tipo_evento,turno_detectado,estatus,mensaje,origen
        FROM registros
        WHERE fecha=?
        ORDER BY hora ASC
    """, (fecha,)).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route("/empleados")
def empleados():
    conn = get_conn()
    rows = conn.execute("SELECT empleado,nombre_completo,puesto,area,turno,tipo_personal,activo FROM empleados ORDER BY nombre_completo LIMIT 1000").fetchall()
    conn.close()
    return render_template("empleados.html", empleados=[dict(r) for r in rows])


@app.route("/importar")
def importar():
    try:
        total = importar_empleados_excel()
        return jsonify({"ok": True, "message": f"Empleados importados/sincronizados: {total}"})
    except Exception as e:
        return jsonify({"ok": False, "message": str(e)})


@app.route("/exportar")
def exportar():
    fecha = request.args.get("fecha") or date.today().strftime("%Y-%m-%d")
    conn = get_conn()
    rows = conn.execute("""
        SELECT fecha,hora,empleado,nombre_completo,tipo_evento,turno_detectado,estatus,mensaje,origen,created_at
        FROM registros WHERE fecha=? ORDER BY hora ASC
    """, (fecha,)).fetchall()
    conn.close()

    df = pd.DataFrame([dict(r) for r in rows])
    out = REPORT_DIR / f"Asistencia_Traxion_{fecha.replace('-','')}.xlsx"
    df.to_excel(out, index=False)
    return send_file(out, as_attachment=True)


def get_stats():
    conn = get_conn()
    fecha = date.today().strftime("%Y-%m-%d")
    def q(sql):
        return conn.execute(sql, (fecha,)).fetchone()[0]
    stats = {
        "entries": q("SELECT COUNT(*) FROM registros WHERE fecha=? AND tipo_evento='ENTRADA'"),
        "exits": q("SELECT COUNT(*) FROM registros WHERE fecha=? AND tipo_evento='SALIDA'"),
        "late": q("SELECT COUNT(*) FROM registros WHERE fecha=? AND estatus IN ('RETARDO','FUERA HORARIO')"),
        "alerts": q("SELECT COUNT(*) FROM registros WHERE fecha=? AND estatus='ALERTA'"),
    }
    conn.close()
    return stats


init_db()
try:
    importar_empleados_excel()
except Exception:
    pass

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
