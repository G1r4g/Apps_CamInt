#!/usr/bin/env python3
"""
Inventario de equipos de grabación (NVR/DVR/XVR) Dahua y Hikvision -> CSV.
Interfaz gráfica (tkinter) con el estilo del portal GIRAG. Máximo 20 equipos. Solo consultas de lectura.
Cada equipo tiene su propia contraseña (se piden una a una; solo viven en memoria).

Requiere:  pip install requests
Uso:       python inventario_grabadores.py

Genera dos CSV (separador ';', listos para Excel/GLPI):
  - inventario_equipos_AAAAMMDD_HHMM.csv   (un renglón por equipo)
  - inventario_camaras_AAAAMMDD_HHMM.csv   (un renglón por canal/cámara: nombre, IP, estado)
La lista de equipos se guarda en equipos.json (SIN usuario ni contraseña).
"""
import csv, json, os, queue, random, re, threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
import xml.etree.ElementTree as ET
import requests, urllib3
from requests.auth import HTTPDigestAuth

urllib3.disable_warnings()

MAX_EQUIPOS = 20
ARCHIVO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "equipos.json")

# (base, nombre, ip, puerto, marca) - editable desde la interfaz
POR_DEFECTO = [
    ("BOG TC1", "BOGNVR01", "100.100.102.6", 80, "Dahua"),
    ("BOG TC1", "BOGNVR02", "100.100.102.7", 80, "Dahua"),
    ("BOG TC2", "BOGDVR03", "100.100.200.203", 80, "Hikvision"),
    ("BOG TC2", "BOGDVR04", "100.100.200.206", 80, "Hikvision"),
    ("BOG TC2", "BOGNVR05", "100.100.200.207", 80, "Dahua"),
    ("BOG TC2", "BOGNVR06", "100.100.200.208", 80, "Dahua"),
    ("FONTIBON", "BOGNVR07", "100.100.161.4", 80, "Dahua"),
    ("MDE", "MDEDVR08", "100.100.141.116", 80, "Dahua"),
    ("MDE", "MDENVR09", "100.100.141.122", 80, "Hikvision"),
    ("MDE", "MDENVR10", "100.100.141.121", 80, "Hikvision"),
    ("MDE", "MDENVR11", "100.100.141.13", 80, "Dahua"),
    ("MDE", "MDEDVR12", "100.100.141.202", 80, "Dahua"),
    ("MDE", "MDEDVR13", "100.100.141.203", 80, "Hikvision"),
]

COLS_EQUIPO = ["Base", "Nombre", "IP", "Marca", "Tipo", "Modelo", "Serial", "MAC", "Firmware",
               "Canales_total", "Canales_en_uso", "Discos", "Capacidad_TB", "Usado_TB",
               "Estado_discos", "Estado_consulta", "Fecha_consulta"]
COLS_CAMARA = ["Equipo", "Canal", "Nombre_camara", "IP_camara", "Estado"]


# --------------------------------------------------------------------------- HTTP
class AuthError(Exception):
    pass


def http_get(ses, host, port, path, timeout):
    """GET con digest. 401 -> AuthError (se corta de inmediato para no bloquear la cuenta).
    Otro código distinto de 200 -> None (endpoint no soportado). Sin conexión -> ConnectionError."""
    ultimo = None
    for esquema, p in (("http", port), ("https", 443)):
        try:
            r = ses.get(f"{esquema}://{host}:{p}{path}", timeout=timeout, verify=False)
        except requests.RequestException as e:
            ultimo = e
            continue
        if r.status_code == 401:
            raise AuthError("Credenciales rechazadas (401)")
        return r.text if r.status_code == 200 else None
    raise ConnectionError(f"Sin respuesta HTTP/HTTPS ({ultimo.__class__.__name__})")


def opcional(ses, host, port, path, timeout, dbg, etiqueta):
    try:
        t = http_get(ses, host, port, path, timeout)
    except ConnectionError:
        t = None
    dbg.append(f"--- {etiqueta} ---\n{t}\n")
    return t or ""


def kv(texto):
    d = {}
    for linea in (texto or "").splitlines():
        if "=" in linea:
            k, v = linea.split("=", 1)
            d[k.strip()] = v.strip()
    return d


def pick(d, *nombres):
    bajo = {k.lower(): v for k, v in d.items()}
    for n in nombres:
        if n.lower() in bajo:
            return bajo[n.lower()]
    return ""


def xml(texto):
    if not texto:
        return None
    texto = re.sub(r'\sxmlns(:\w+)?="[^"]*"', "", texto)
    try:
        return ET.fromstring(texto)
    except ET.ParseError:
        return None


def xt(el, ruta):
    if el is None:
        return ""
    n = el.find(ruta)
    return (n.text or "").strip() if n is not None and n.text else ""


def nueva_fila(d):
    return {c: "" for c in COLS_EQUIPO} | {
        "Base": d["base"], "Nombre": d["nombre"], "IP": d["ip"], "Marca": d["marca"],
        "Fecha_consulta": datetime.now().strftime("%Y-%m-%d %H:%M")}


def tipo_por_modelo(modelo, marca):
    m = (modelo or "").upper()
    if "XVR" in m or "DVR" in m:
        return "DVR"
    if marca == "Hikvision" and re.search(r"HWI|HQHI|HUHI|HGHI|HTHI", m):
        return "DVR"
    if "NVR" in m or "NI" in m:
        return "NVR"
    return ""


# --------------------------------------------------------------------------- Dahua
def consultar_dahua(d, user, pwd, timeout, dbg):
    row, cams = nueva_fila(d), []
    ses = requests.Session()
    ses.auth = HTTPDigestAuth(user, pwd)
    h, p = d["ip"], d["puerto"]

    def g(path, etiqueta):
        return opcional(ses, h, p, path, timeout, dbg, etiqueta)

    # Primera consulta SIN capturar errores: si falla por 401/conexión se corta aquí.
    t = http_get(ses, h, p, "/cgi-bin/magicBox.cgi?action=getSystemInfo", timeout)
    if t is None:
        raise ConnectionError("El equipo no respondió como Dahua (¿marca correcta?)")
    dbg.append(f"--- sys ---\n{t}\n")
    sysd = kv(t)
    row["Serial"] = pick(sysd, "serialNumber")
    tipo = kv(g("/cgi-bin/magicBox.cgi?action=getDeviceType", "type"))
    row["Modelo"] = pick(tipo, "type") or pick(sysd, "deviceType")
    row["Firmware"] = pick(kv(g("/cgi-bin/magicBox.cgi?action=getSoftwareVersion", "ver")), "version")
    row["Tipo"] = tipo_por_modelo(row["Modelo"], "Dahua")

    # MAC
    net = kv(g("/cgi-bin/configManager.cgi?action=getConfig&name=Network", "net"))
    row["MAC"] = pick(net, "table.Network.eth0.PhysicalAddress") or next(
        (v for k, v in net.items() if k.lower().endswith("physicaladdress")), "")

    # Canales totales
    total = 0
    c1 = kv(g("/cgi-bin/magicBox.cgi?action=getProductDefinition&name=MaxRemoteInputChannels", "chan_ip"))
    c2 = kv(g("/cgi-bin/devVideoInput.cgi?action=getCollect", "chan_analog"))
    for v in (pick(c1, "table.MaxRemoteInputChannels", "MaxRemoteInputChannels"), pick(c2, "result")):
        try:
            total = max(total, int(v))
        except ValueError:
            pass

    # Nombres de canal
    titulos = {}
    for k, v in kv(g("/cgi-bin/configManager.cgi?action=getConfig&name=ChannelTitle", "titles")).items():
        m = re.match(r"table\.ChannelTitle\[(\d+)\]\.Name$", k, re.I)
        if m:
            titulos[int(m.group(1))] = v

    # Estado de conexión por canal
    estados = {}
    for k, v in kv(g("/cgi-bin/api/LogicDeviceManager/getCameraState?channel=All", "state")).items():
        m = re.match(r"states\[(\d+)\]\.(\w+)$", k, re.I)
        if m:
            estados.setdefault(int(m.group(1)), {})[m.group(2).lower()] = v

    # Cámaras remotas (IP y nombre)
    remotos = {}
    for k, v in kv(g("/cgi-bin/configManager.cgi?action=getConfig&name=RemoteDevice", "remote")).items():
        m = re.match(r"table\.RemoteDevice[.\[]([^.\]]+)\]?\.(.+)$", k, re.I)
        if m:
            remotos.setdefault(m.group(1), {})[m.group(2).lower()] = v

    def datos_remoto(r):
        nombre = r.get("name") or r.get("videoinputs[0].name") or r.get("devicename") or ""
        return r.get("address", ""), nombre

    usados_remotos = set()
    n_canales = max(total, len(estados), (max(titulos) + 1) if titulos else 0)
    for i in range(n_canales):
        st = estados.get(i, {})
        conectado = st.get("connectionstate", "").lower() == "connected"
        ip_cam = ""
        uuid = st.get("uuid", "")
        if uuid and uuid in remotos:
            ip_cam = datos_remoto(remotos[uuid])[0]
            usados_remotos.add(uuid)
        if not ip_cam:
            # 1) el id del remoto suele terminar en el índice del canal (..._INFO_0 = canal 1)
            for rid, r in remotos.items():
                m = re.search(r"(\d+)$", rid)
                if rid not in usados_remotos and m and int(m.group(1)) == i and datos_remoto(r)[0]:
                    ip_cam = datos_remoto(r)[0]
                    usados_remotos.add(rid)
                    break
        if not ip_cam and titulos.get(i):
            # 2) por nombre de cámara
            for rid, r in remotos.items():
                ip_r, nom_r = datos_remoto(r)
                if rid not in usados_remotos and ip_r and nom_r.lower() == titulos[i].lower():
                    ip_cam = ip_r
                    usados_remotos.add(rid)
                    break
        if st:
            estado = "Conectada" if conectado else (st.get("connectionstate") or "Sin conexión")
        else:
            estado = ""
        cams.append({"Equipo": d["nombre"], "Canal": i + 1, "Nombre_camara": titulos.get(i, ""),
                     "IP_camara": ip_cam, "Estado": estado})
    # Cámaras remotas que no se pudieron asociar a un canal: se listan igual
    for rid, r in remotos.items():
        if rid in usados_remotos:
            continue
        ip_cam, nombre = datos_remoto(r)
        if ip_cam:
            cams.append({"Equipo": d["nombre"], "Canal": "", "Nombre_camara": nombre,
                         "IP_camara": ip_cam,
                         "Estado": "Habilitada" if r.get("enable", "").lower() == "true" else "Deshabilitada"})

    row["Canales_total"] = total or n_canales or ""
    if estados:
        row["Canales_en_uso"] = sum(1 for s in estados.values()
                                    if s.get("connectionstate", "").lower() == "connected")
    elif remotos:
        row["Canales_en_uso"] = sum(1 for r in remotos.values() if r.get("enable", "").lower() == "true")

    # Discos
    disk = kv(g("/cgi-bin/storageDevice.cgi?action=getDeviceAllInfo", "disk"))
    caps, usados, estd = {}, {}, {}
    for k, v in disk.items():
        m = re.match(r"list\.info\[(\d+)\]\.detail\[(\d+)\]\.(\w+)$", k, re.I)
        if m:
            idx, f = (m.group(1), m.group(2)), m.group(3).lower()
            try:
                if f == "totalbytes":
                    caps[idx] = float(v)
                elif f == "usedbytes":
                    usados[idx] = float(v)
            except ValueError:
                pass
            if f == "state":
                estd[idx] = v
    if caps:
        row["Discos"] = len(caps)
        row["Capacidad_TB"] = round(sum(caps.values()) / 1e12, 2)
        if usados:
            row["Usado_TB"] = round(sum(usados.values()) / 1e12, 2)
    if estd:
        row["Estado_discos"] = "; ".join(f"D{i + 1}:{estd[k]}" for i, k in enumerate(sorted(estd)))
    row["Estado_consulta"] = "OK"
    return row, cams


# --------------------------------------------------------------------------- Hikvision
def consultar_hikvision(d, user, pwd, timeout, dbg):
    row, cams = nueva_fila(d), []
    ses = requests.Session()
    ses.auth = HTTPDigestAuth(user, pwd)
    h, p = d["ip"], d["puerto"]

    def g(path, etiqueta):
        return opcional(ses, h, p, path, timeout, dbg, etiqueta)

    t = http_get(ses, h, p, "/ISAPI/System/deviceInfo", timeout)
    if t is None:
        raise ConnectionError("El equipo no respondió como Hikvision (¿marca correcta?)")
    dbg.append(f"--- deviceInfo ---\n{t}\n")
    info = xml(t)
    row["Modelo"] = xt(info, "model")
    row["Serial"] = xt(info, "serialNumber")
    row["MAC"] = xt(info, "macAddress")
    row["Firmware"] = xt(info, "firmwareVersion")
    row["Tipo"] = tipo_por_modelo(row["Modelo"], "Hikvision")

    # Cámaras IP (NVR / DVR híbrido)
    proxy = {}
    raiz = xml(g("/ISAPI/ContentMgmt/InputProxy/channels", "proxy"))
    if raiz is not None:
        for c in raiz.iter("InputProxyChannel"):
            proxy[xt(c, "id")] = {"nombre": xt(c, "name"),
                                  "ip": xt(c, "sourceInputPortDescriptor/ipAddress")}
    online = {}
    raiz = xml(g("/ISAPI/ContentMgmt/InputProxy/channels/status", "proxy_status"))
    if raiz is not None:
        for c in raiz.iter("InputProxyChannelStatus"):
            online[xt(c, "id")] = xt(c, "online").lower() == "true"

    # Canales de video (analógicos o ranuras totales)
    entradas = {}
    raiz = xml(g("/ISAPI/System/Video/inputs/channels", "inputs"))
    if raiz is not None:
        for c in raiz.iter("VideoInputChannel"):
            entradas[xt(c, "id")] = {"nombre": xt(c, "name"),
                                     "hab": xt(c, "videoInputEnabled").lower() != "false"}

    def clave(x):
        return int(x) if x.isdigit() else 0

    if proxy:
        for cid in sorted(proxy, key=clave):
            cams.append({"Equipo": d["nombre"], "Canal": cid,
                         "Nombre_camara": proxy[cid]["nombre"] or entradas.get(cid, {}).get("nombre", ""),
                         "IP_camara": proxy[cid]["ip"],
                         "Estado": "Conectada" if online.get(cid) else "Sin conexión"})
        row["Canales_en_uso"] = sum(1 for v in online.values() if v) if online else ""
        row["Canales_total"] = max(len(entradas), len(proxy))
    elif entradas:
        for cid in sorted(entradas, key=clave):
            cams.append({"Equipo": d["nombre"], "Canal": cid, "Nombre_camara": entradas[cid]["nombre"],
                         "IP_camara": "", "Estado": "Habilitado" if entradas[cid]["hab"] else "Deshabilitado"})
        row["Canales_total"] = len(entradas)
        row["Canales_en_uso"] = sum(1 for e in entradas.values() if e["hab"])

    # Discos
    raiz = xml(g("/ISAPI/ContentMgmt/Storage", "storage"))
    if raiz is not None:
        discos = list(raiz.iter("hdd"))
        cap = lib = 0.0
        estd = []
        for i, hd in enumerate(discos):
            try:
                cap += float(xt(hd, "capacity"))
                lib += float(xt(hd, "freeSpace"))
            except ValueError:
                pass
            estd.append(f"D{i + 1}:{xt(hd, 'status')}")
        if discos:
            row["Discos"] = len(discos)
            row["Capacidad_TB"] = round(cap / 1024 / 1024, 2)   # el equipo informa en MB
            row["Usado_TB"] = round((cap - lib) / 1024 / 1024, 2)
            row["Estado_discos"] = "; ".join(estd)
    row["Estado_consulta"] = "OK"
    return row, cams


# --------------------------------------------------------------------------- Estilo (portal GIRAG)
BG_TOP, BG_BOT = "#09102f", "#0c1545"
BG = "#0a1239"
CARD, CARD_BRD = "#0d1642", "#1f2f7a"
FIELD, FIELD_BRD = "#080e32", "#26367d"
FOCUS, BLUE = "#3d6bff", "#2a62ff"
TXT, MUTED, DIM = "#f3f6ff", "#b9c5ee", "#7f90cf"
ORANGE, ORANGE_H = "#ff7a0f", "#ff9333"
OK_C, ERR_C, WARN_C = "#6fe0aa", "#ff8f8f", "#ffb866"
FONT = "Segoe UI"
CARD_W, CARD_H = 392, 584


def esp(t, sep="\u200a"):
    """Texto con espaciado entre letras (como los rótulos del portal)."""
    return sep.join(t)


def mezcla(c1, c2, t):
    a = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02x%02x%02x" % tuple(round(x + (y - x) * t) for x, y in zip(a, b))


def rrect(cv, x1, y1, x2, y2, r, **kw):
    """Rectángulo redondeado dibujado como polígono suavizado."""
    p = [x1 + r, y1, x1 + r, y1, x2 - r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y1 + r,
         x2, y2 - r, x2, y2 - r, x2, y2, x2 - r, y2, x2 - r, y2, x1 + r, y2, x1 + r, y2,
         x1, y2, x1, y2 - r, x1, y2 - r, x1, y1 + r, x1, y1 + r, x1, y1]
    return cv.create_polygon(p, smooth=True, **kw)


def estilos(root):
    s = ttk.Style(root)
    s.theme_use("clam")
    s.configure("Dark.Treeview", background=FIELD, fieldbackground=FIELD, foreground=TXT,
                rowheight=26, borderwidth=0, bordercolor=FIELD, lightcolor=FIELD,
                darkcolor=FIELD, font=(FONT, 9))
    s.configure("Dark.Treeview.Heading", background="#111b55", foreground="#9fb2f0",
                relief="flat", borderwidth=0, padding=(8, 6), font=(FONT, 8, "bold"))
    s.map("Dark.Treeview", background=[("selected", "#22409f")], foreground=[("selected", "#ffffff")])
    s.map("Dark.Treeview.Heading", background=[("active", "#16226a")])
    s.configure("Dark.Vertical.TScrollbar", background="#1c2a6e", troughcolor=FIELD,
                bordercolor=FIELD, lightcolor="#1c2a6e", darkcolor="#1c2a6e",
                arrowcolor="#8fa3e8", relief="flat")
    s.map("Dark.Vertical.TScrollbar", background=[("active", "#2a3c93")])
    s.configure("Dark.TCombobox", fieldbackground=FIELD, background=FIELD, foreground=TXT,
                arrowcolor=MUTED, bordercolor=FIELD, lightcolor=FIELD, darkcolor=FIELD,
                arrowsize=14, padding=2)
    s.map("Dark.TCombobox",
          fieldbackground=[("readonly", FIELD)], foreground=[("readonly", TXT)],
          selectbackground=[("readonly", FIELD)], selectforeground=[("readonly", TXT)],
          background=[("active", FIELD), ("readonly", FIELD)])
    root.option_add("*TCombobox*Listbox.background", FIELD)
    root.option_add("*TCombobox*Listbox.foreground", TXT)
    root.option_add("*TCombobox*Listbox.selectBackground", "#22409f")
    root.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")
    root.option_add("*TCombobox*Listbox.font", (FONT, 10))


class RoundField(tk.Canvas):
    """Campo de texto (o lista desplegable) con bordes redondeados; opcional ojo para ver la clave."""

    def __init__(self, parent, ancho, values=None, password=False, bg=CARD):
        super().__init__(parent, width=ancho, height=44, bg=bg, highlightthickness=0)
        self.password, self.visible = password, False
        self.var = tk.StringVar()
        self.box = rrect(self, 1, 1, ancho - 1, 43, 9, fill=FIELD, outline=FIELD_BRD, width=1)
        extra = 46 if password else 0
        if values is not None:
            self.inp = ttk.Combobox(self, textvariable=self.var, values=values, state="readonly",
                                    style="Dark.TCombobox", font=(FONT, 10))
        else:
            self.inp = tk.Entry(self, textvariable=self.var, bd=0, relief="flat", bg=FIELD, fg=TXT,
                                insertbackground=TXT, highlightthickness=0, font=(FONT, 10),
                                show="•" if password else "")
        self.create_window(14, 22, window=self.inp, anchor="w", width=ancho - 28 - extra, height=30)
        if password:
            x0 = ancho - 46
            self.create_line(x0, 6, x0, 38, fill=FIELD_BRD)
            cx = x0 + 23
            self.create_oval(cx - 10, 16, cx + 10, 28, outline=MUTED, width=2, tags="eye")
            self.create_oval(cx - 3, 19, cx + 3, 25, fill=MUTED, outline=MUTED, tags="eye")
            self.slash = self.create_line(cx - 9, 31, cx + 9, 13, fill=MUTED, width=2, state="hidden")
            self.tag_bind("eye", "<Button-1>", self._ver)
            self.tag_bind("eye", "<Enter>", lambda e: self.config(cursor="hand2"))
            self.tag_bind("eye", "<Leave>", lambda e: self.config(cursor=""))
        self.inp.bind("<FocusIn>", lambda e: self.itemconfigure(self.box, outline=FOCUS))
        self.inp.bind("<FocusOut>", lambda e: self.itemconfigure(self.box, outline=FIELD_BRD))

    def _ver(self, _e=None):
        self.visible = not self.visible
        self.inp.configure(show="" if self.visible else "•")
        self.itemconfigure(self.slash, state="normal" if self.visible else "hidden")

    def get(self):
        return self.var.get()

    def set(self, v):
        self.var.set(v)

    def foco(self):
        self.inp.focus_set()

    def on(self, secuencia, cb):
        self.inp.bind(secuencia, cb)


class RoundButton(tk.Canvas):
    """Botón de esquinas redondeadas (relleno naranja o solo borde)."""

    def __init__(self, parent, texto, comando, ancho, alto=44, fill=ORANGE, hover=ORANGE_H,
                 fg="#0a1033", outline=None, bg=CARD, size=10):
        super().__init__(parent, width=ancho, height=alto, bg=bg, highlightthickness=0, cursor="hand2")
        self.cmd, self.fill, self.hover, self.fg, self.activo = comando, fill, hover, fg, True
        self.shape = rrect(self, 1, 1, ancho - 1, alto - 1, 9, fill=fill, outline=outline or fill)
        self.txt = self.create_text(ancho / 2, alto / 2, text=texto, fill=fg, font=(FONT, size, "bold"))
        self.bind("<Enter>", lambda e: self.activo and self.itemconfigure(self.shape, fill=self.hover))
        self.bind("<Leave>", lambda e: self.activo and self.itemconfigure(self.shape, fill=self.fill))
        self.bind("<Button-1>", self._click)

    def _click(self, _e):
        if self.activo and self.cmd:
            self.cmd()

    def texto(self, t):
        self.itemconfigure(self.txt, text=t)

    def habilitar(self, v):
        self.activo = v
        self.itemconfigure(self.shape, fill=self.fill if v else "#1b2563")
        self.itemconfigure(self.txt, fill=self.fg if v else "#6c7bbd")
        self.config(cursor="hand2" if v else "")


def clave(d):
    """Identifica un equipo para guardar su contraseña (solo en memoria)."""
    return (d["nombre"], d["ip"])


# --------------------------------------------------------------------------- Interfaz
class App:
    def __init__(self, root):
        self.root = root
        root.title("Inventario de equipos de grabación · GIRAG")
        root.geometry("1280x800")
        root.minsize(1150, 700)
        root.configure(bg=BG)
        estilos(root)
        self.cola = queue.Queue()
        self.equipos = []            # lista de dicts
        self.creds = {}              # clave(equipo) -> (usuario, contraseña)  [solo memoria]
        self.res_equipos = {}        # idx -> fila
        self.res_camaras = {}        # idx -> lista
        self.trabajando = False
        self.pendientes = 0
        self._msg_pend = None
        self.pg_acceso = tk.Canvas(root, bg=BG, highlightthickness=0)
        self.pg_result = tk.Frame(root, bg=BG)
        self._build_acceso()
        self._build_resultados()
        self.cargar_lista()
        self.pg_acceso.pack(fill="both", expand=True)

    # ---- utilidades de interfaz
    def _tree(self, parent, defs, height=None):
        marco = tk.Frame(parent, bg=FIELD, highlightthickness=1, highlightbackground=CARD_BRD)
        kw = {"height": height} if height else {}
        tree = ttk.Treeview(marco, columns=[d[0] for d in defs], show="headings",
                            selectmode="browse", style="Dark.Treeview", **kw)
        for i, (c, t, w) in enumerate(defs):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor="w", stretch=(i == len(defs) - 1))
        sb = ttk.Scrollbar(marco, command=tree.yview, style="Dark.Vertical.TScrollbar")
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)
        tree.tag_configure("ok", foreground=OK_C)
        tree.tag_configure("err", foreground=ERR_C)
        tree.tag_configure("lista", foreground=TXT)
        tree.tag_configure("pend", foreground=WARN_C)
        return marco, tree

    def _link(self, parent, texto, cmd, bg=BG):
        b = tk.Label(parent, text=texto, bg=bg, fg="#8fa8ff", cursor="hand2", font=(FONT, 9, "bold"))
        b.bind("<Button-1>", lambda e: cmd())
        b.bind("<Enter>", lambda e: b.configure(fg="#ffffff"))
        b.bind("<Leave>", lambda e: b.configure(fg="#8fa8ff"))
        return b

    def msg(self, texto, color=MUTED):
        self.cv.itemconfigure(self.t_msg, text=texto, fill=color)

    # ---- pantalla 1: acceso por equipo (réplica del portal)
    def _build_acceso(self):
        cv = self.cv = self.pg_acceso
        self.items = []

        def fijo(modo, x, y, id_):
            self.items.append((id_, modo, x, y))
            return id_

        def txt(modo, x, y, **kw):
            kw.setdefault("anchor", "nw")
            return fijo(modo, x, y, cv.create_text(0, 0, **kw))

        def win(modo, x, y, w, **kw):
            kw.setdefault("anchor", "nw")
            return fijo(modo, x, y, cv.create_window(0, 0, window=w, **kw))

        # --- columna izquierda
        txt("tl", 88, 70, anchor="w", text=esp("PORTAL DE INVENTARIO · GRABADORES CCTV"),
            fill=DIM, font=(FONT, 8))
        marco, self.tree = self._tree(cv, (("base", "Base", 80), ("nombre", "Nombre", 100),
                                           ("ip", "IP", 115), ("marca", "Marca", 75),
                                           ("cred", "Contraseña", 100), ("estado", "Estado de consulta", 200)))
        self.win_tabla = win("tl", 54, 112, marco)
        self.tree.bind("<<TreeviewSelect>>", self._al_seleccionar)

        self.tools = tk.Frame(cv, bg=BG)
        for t, cmd in (("+ Agregar", self.agregar), ("Editar", self.editar), ("Eliminar", self.eliminar),
                       ("Guardar lista", self.guardar_lista), ("Cargar lista…", self.cargar_archivo)):
            self._link(self.tools, t, cmd).pack(side="left", padx=(0, 16))
        self.v_debug, self.v_timeout = tk.BooleanVar(value=False), tk.StringVar(value="8")
        tk.Checkbutton(self.tools, text="Guardar respuestas crudas", variable=self.v_debug, bg=BG,
                       fg=MUTED, selectcolor=FIELD, activebackground=BG, activeforeground=TXT,
                       highlightthickness=0, bd=0, font=(FONT, 8)).pack(side="right")
        tk.Entry(self.tools, textvariable=self.v_timeout, width=3, bg=FIELD, fg=TXT, bd=0,
                 insertbackground=TXT, justify="center", font=(FONT, 9)).pack(side="right", padx=(4, 18))
        tk.Label(self.tools, text="Timeout (s)", bg=BG, fg=MUTED, font=(FONT, 8)).pack(side="right")
        self.win_tools = win("bl", 54, -292, self.tools, height=28)

        txt("bl", 54, -252, text="GIRAG", fill="#ffffff", font=("Arial Black", 40))
        txt("bl", 54, -188, text="Inventario de equipos\nde grabación CCTV", fill=TXT,
            font=(FONT, 15, "bold"))
        for x, cod, sub in ((54, "BOG", "Bogotá · El Dorado"),
                            (177, "MDE", "Rionegro · José María Córdova"),
                            (354, "CLO", "Cali · Alfonso Bonilla Aragón")):
            txt("bl", x, -108, text=esp(cod), fill="#6f86d6", font=(FONT, 14, "bold"))
            txt("bl", x, -82, text=sub, fill=DIM, font=(FONT, 8))
        txt("bl", 70, -44, anchor="w", text=esp("GROUND HANDLING · CARGA · RAMPA — DESDE 1974"),
            fill=MUTED, font=(FONT, 8))

        # --- tarjeta de acceso (derecha)
        cw = CARD_W - 60
        txt("card", 30, 28, text="Acceso al equipo", fill=TXT, font=(FONT, 17, "bold"))
        self.t_marca = txt("card", CARD_W - 73, 39, anchor="center", text="—", fill=TXT, font=(FONT, 9))
        txt("card", 30, 72, width=cw, fill=MUTED, font=(FONT, 9),
            text="Identifícate con la contraseña propia del equipo de grabación seleccionado.")
        txt("card", 30, 134, text=esp("EQUIPO DE GRABACIÓN"), fill=TXT, font=(FONT, 8, "bold"))
        self.f_equipo = RoundField(cv, cw, values=[])
        win("card", 30, 150, self.f_equipo)
        txt("card", 30, 214, text=esp("USUARIO DEL EQUIPO"), fill=TXT, font=(FONT, 8, "bold"))
        self.f_user = RoundField(cv, cw)
        self.f_user.set("admin")
        win("card", 30, 230, self.f_user)
        txt("card", 30, 294, text=esp("CONTRASEÑA DEL EQUIPO"), fill=TXT, font=(FONT, 8, "bold"))
        self.f_pwd = RoundField(cv, cw, password=True)
        win("card", 30, 310, self.f_pwd)
        self.t_msg = txt("card", 30, 366, width=cw, fill=MUTED, font=(FONT, 9),
                         text="Selecciona un equipo y escribe su contraseña.")
        self.btn_guardar = RoundButton(cv, "Guardar y siguiente", self.guardar_cred, cw)
        win("card", 30, 392, self.btn_guardar)
        self.btn_cons = RoundButton(cv, "", self.consultar, cw, fill=CARD, hover="#16205f",
                                    fg=ORANGE, outline=ORANGE)
        win("card", 30, 450, self.btn_cons)
        txt("card", CARD_W / 2, 524, anchor="n", fill=DIM, font=(FONT, 8),
            text="Credenciales solo en memoria · no se guardan en disco")
        self.t_cnt = txt("card", CARD_W / 2, 540, anchor="n", fill=MUTED, font=(FONT, 8), text="")
        txt("card", CARD_W / 2, 556, anchor="n", fill=DIM, font=(FONT, 8),
            text="GIRAG S.A.S. · Servicios aeroportuarios")

        self.f_equipo.on("<<ComboboxSelected>>", self._al_combo)
        self.f_user.on("<Return>", lambda e: self.f_pwd.foco())
        self.f_pwd.on("<Return>", self.guardar_cred)
        cv.bind("<Configure>", self._al_redimensionar)

    def _al_redimensionar(self, e):
        w, h = e.width, e.height
        if w < 100 or h < 100:
            return
        cx, cy = w - CARD_W - 70, max(30, (h - CARD_H) // 2)
        for id_, modo, x, y in self.items:
            if modo == "tl":
                self.cv.coords(id_, x, y)
            elif modo == "bl":
                self.cv.coords(id_, x, h + y)
            else:
                self.cv.coords(id_, cx + x, cy + y)
        ancho_tabla = max(300, cx - 50 - 54)
        self.cv.itemconfigure(self.win_tabla, width=ancho_tabla, height=max(120, h - 300 - 112))
        self.cv.itemconfigure(self.win_tools, width=ancho_tabla)
        self._dibujar(w, h, cx, cy)

    def _dibujar(self, w, h, cx, cy):
        cv = self.cv
        cv.delete("bg")
        cv.delete("dyn")
        n = 48
        for i in range(n):
            cv.create_rectangle(0, int(h * i / n), w, int(h * (i + 1) / n) + 1,
                                fill=mezcla(BG_TOP, BG_BOT, i / (n - 1)), outline="", tags="bg")
        for r in (260, 200):
            cv.create_arc(110 - r, 250 - r, 110 + r, 250 + r, start=20, extent=150, style="arc",
                          outline="#1a2a6e", width=2, tags="bg")
        rnd = random.Random(7)
        for _ in range(70):
            x, y, r = rnd.randint(0, w), rnd.randint(int(h * .36), int(h * .52)), rnd.choice((1, 1, 2))
            cv.create_oval(x - r, y - r, x + r, y + r, outline="", tags="bg",
                           fill=rnd.choice(("#3b4f9c", "#c8b777", "#5f7be0", "#2a3a82")))
        cv.create_line(54, 70, 78, 70, fill=DIM, tags="dyn")
        cv.create_line(54, h - 122, 544, h - 122, fill="#1c2a6a", tags="dyn")
        cv.create_oval(50, h - 48, 58, h - 40, fill=ORANGE, outline="", tags="dyn")
        rrect(cv, cx, cy, cx + CARD_W, cy + CARD_H, 12, fill=CARD, outline=CARD_BRD, tags="dyn")
        cv.create_line(cx + 14, cy + 1, cx + CARD_W - 14, cy + 1, fill=BLUE, width=3, tags="dyn")
        rrect(cv, cx + CARD_W - 116, cy + 24, cx + CARD_W - 30, cy + 54, 15, fill=FIELD,
              outline=FIELD_BRD, tags="dyn")
        cv.create_line(cx, cy + 512, cx + CARD_W, cy + 512, fill=CARD_BRD, tags="dyn")
        cv.tag_lower("dyn")
        cv.tag_lower("bg")

    # ---- pantalla 2: resultados
    def _build_resultados(self):
        f = self.pg_result
        top = tk.Frame(f, bg=BG)
        top.pack(fill="x", padx=54, pady=(30, 8))
        tk.Label(top, text="GIRAG", bg=BG, fg="#ffffff", font=("Arial Black", 18)).pack(side="left")
        tk.Label(top, text=esp("  RESULTADOS DE LA CONSULTA"), bg=BG, fg=DIM,
                 font=(FONT, 8)).pack(side="left", padx=(14, 0), pady=(10, 0))
        RoundButton(top, "Exportar CSV…", self.exportar, 170, 38, bg=BG).pack(side="right")
        self._link(top, "← Volver a equipos", lambda: self._ir("acceso")).pack(side="right", padx=22)

        tk.Label(f, text=esp("EQUIPOS"), bg=BG, fg=TXT, font=(FONT, 8, "bold")).pack(anchor="w", padx=54)
        marco, self.tr_eq = self._tree(f, (("nombre", "Nombre", 100), ("ip", "IP", 110), ("marca", "Marca", 70),
                                           ("tipo", "Tipo", 45), ("modelo", "Modelo", 130), ("serial", "Serial", 170),
                                           ("mac", "MAC", 110), ("total", "Canales", 55), ("uso", "En uso", 50),
                                           ("cap", "TB", 55), ("estado", "Estado", 190)), height=9)
        marco.pack(fill="x", padx=54, pady=(4, 12))
        self.tr_eq.bind("<<TreeviewSelect>>", self._filtrar_camaras)

        tk.Label(f, text=esp("CÁMARAS") + "   (selecciona un equipo para filtrar; sin selección se muestran todas)",
                 bg=BG, fg=TXT, font=(FONT, 8, "bold")).pack(anchor="w", padx=54)
        marco, self.tr_cam = self._tree(f, (("equipo", "Equipo", 110), ("canal", "Canal", 55),
                                            ("nombre", "Nombre cámara", 240), ("ip", "IP cámara", 140),
                                            ("estado", "Estado", 120)))
        marco.pack(fill="both", expand=True, padx=54, pady=(4, 30))

    def _ir(self, pagina):
        self.pg_acceso.pack_forget()
        self.pg_result.pack_forget()
        (self.pg_acceso if pagina == "acceso" else self.pg_result).pack(fill="both", expand=True)

    # ---- diálogo para agregar / editar un equipo
    def _dialogo(self, d=None):
        top = tk.Toplevel(self.root)
        top.title("Equipo de grabación")
        top.configure(bg=CARD)
        top.resizable(False, False)
        top.transient(self.root)
        rx, ry = self.root.winfo_rootx() + 300, self.root.winfo_rooty() + 150
        top.geometry(f"388x390+{rx}+{ry}")
        d = d or {"base": "", "nombre": "", "ip": "", "puerto": 80, "marca": "Dahua"}
        res = {}

        def etiqueta(t, x, y):
            tk.Label(top, text=esp(t), bg=CARD, fg=TXT, font=(FONT, 8, "bold")).place(x=x, y=y)

        etiqueta("BASE", 24, 22)
        f_base = RoundField(top, 340); f_base.set(d["base"]); f_base.place(x=24, y=40)
        etiqueta("NOMBRE", 24, 96)
        f_nom = RoundField(top, 340); f_nom.set(d["nombre"]); f_nom.place(x=24, y=114)
        etiqueta("IP", 24, 170)
        f_ip = RoundField(top, 340); f_ip.set(d["ip"]); f_ip.place(x=24, y=188)
        etiqueta("PUERTO", 24, 244)
        f_pto = RoundField(top, 140); f_pto.set(str(d["puerto"])); f_pto.place(x=24, y=262)
        etiqueta("MARCA", 184, 244)
        f_mar = RoundField(top, 180, values=("Dahua", "Hikvision")); f_mar.set(d["marca"])
        f_mar.place(x=184, y=262)

        def aceptar():
            nombre, ip = f_nom.get().strip(), f_ip.get().strip()
            if not nombre or not ip:
                messagebox.showwarning("Datos incompletos", "Nombre e IP son obligatorios.", parent=top)
                return
            try:
                puerto = int(f_pto.get())
            except ValueError:
                messagebox.showwarning("Puerto inválido", "El puerto debe ser un número (normalmente 80).", parent=top)
                return
            res.update(base=f_base.get().strip(), nombre=nombre, ip=ip, puerto=puerto,
                       marca=f_mar.get() or "Dahua")
            top.destroy()

        RoundButton(top, "Guardar", aceptar, 164).place(x=24, y=326)
        RoundButton(top, "Cancelar", top.destroy, 164, fill=CARD, hover="#16205f", fg=ORANGE,
                    outline=ORANGE).place(x=200, y=326)
        f_nom.foco()
        top.grab_set()
        self.root.wait_window(top)
        return res or None

    # ---- manejo de la lista
    def _vals(self, d):
        cred = "● Lista" if clave(d) in self.creds else "○ Pendiente"
        return (d["base"], d["nombre"], d["ip"], d["marca"], cred, d.get("estado", ""))

    def _tag(self, d):
        est = d.get("estado", "")
        if est == "OK":
            return "ok"
        if est and est not in ("En cola...", "Consultando..."):
            return "err"
        return "lista" if clave(d) in self.creds else "pend"

    def _actualizar_fila(self, i):
        d = self.equipos[i]
        self.tree.item(str(i), values=self._vals(d), tags=(self._tag(d),))

    def _cont(self):
        n = sum(1 for d in self.equipos if clave(d) in self.creds)
        self.cv.itemconfigure(self.t_cnt, text=f"{n} de {len(self.equipos)} equipos con contraseña")
        self.btn_cons.texto(f"▶  Consultar equipos ({n}/{len(self.equipos)})")

    def _refrescar_lista(self, sel=None):
        self.tree.delete(*self.tree.get_children())
        for i, d in enumerate(self.equipos):
            self.tree.insert("", "end", iid=str(i), values=self._vals(d), tags=(self._tag(d),))
        self.f_equipo.inp.configure(values=[f"{d['nombre']} · {d['ip']}" for d in self.equipos])
        self._cont()
        if sel is not None and sel < len(self.equipos):
            self.tree.selection_set(str(sel))
        else:
            self.f_equipo.set("")
            self.cv.itemconfigure(self.t_marca, text="—")

    def _sel_idx(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def _al_seleccionar(self, _e=None):
        i = self._sel_idx()
        if i is None:
            return
        d = self.equipos[i]
        self.f_equipo.set(f"{d['nombre']} · {d['ip']}")
        self.cv.itemconfigure(self.t_marca, text=d["marca"])
        u, p = self.creds.get(clave(d), (self.f_user.get() or "admin", ""))
        self.f_user.set(u)
        self.f_pwd.set(p)
        if self._msg_pend:
            self.msg(*self._msg_pend)
            self._msg_pend = None
        elif clave(d) in self.creds:
            self.msg(f"{d['nombre']} ya tiene contraseña. Puedes cambiarla y guardar de nuevo.", OK_C)
        else:
            self.msg(f"Escribe la contraseña de {d['nombre']}.", MUTED)

    def _al_combo(self, _e=None):
        i = self.f_equipo.inp.current()
        if i >= 0:
            self.tree.selection_set(str(i))
            self.tree.see(str(i))
            self.f_equipo.inp.selection_clear()

    def guardar_cred(self, _e=None):
        i = self._sel_idx()
        if i is None:
            self.msg("Selecciona primero un equipo.", WARN_C)
            return
        u, p = self.f_user.get().strip(), self.f_pwd.get()
        if not u or not p:
            self.msg("Escribe usuario y contraseña de este equipo.", WARN_C)
            return
        d = self.equipos[i]
        self.creds[clave(d)] = (u, p)
        self._actualizar_fila(i)
        self._cont()
        orden = list(range(i + 1, len(self.equipos))) + list(range(0, i))
        sig = next((j for j in orden if clave(self.equipos[j]) not in self.creds), None)
        if sig is None:
            self.msg("✔ Todos los equipos tienen contraseña. Ya puedes consultar.", OK_C)
        else:
            self._msg_pend = (f"✔ {d['nombre']} guardado. Sigue {self.equipos[sig]['nombre']}.", OK_C)
            self.tree.selection_set(str(sig))
            self.tree.see(str(sig))
            self.f_pwd.foco()

    def agregar(self):
        if len(self.equipos) >= MAX_EQUIPOS:
            messagebox.showwarning("Límite", f"Máximo {MAX_EQUIPOS} equipos.")
            return
        d = self._dialogo()
        if d:
            self.equipos.append(d)
            self._refrescar_lista(len(self.equipos) - 1)

    def editar(self):
        i = self._sel_idx()
        if i is None:
            self.msg("Selecciona un equipo para editarlo.", WARN_C)
            return
        d = self._dialogo(self.equipos[i])
        if d:
            vieja = clave(self.equipos[i])
            if vieja in self.creds and clave(d) != vieja:
                self.creds[clave(d)] = self.creds.pop(vieja)
            self.equipos[i] = d
            self._refrescar_lista(i)

    def eliminar(self):
        i = self._sel_idx()
        if i is None:
            self.msg("Selecciona un equipo para eliminarlo.", WARN_C)
            return
        self.creds.pop(clave(self.equipos[i]), None)
        del self.equipos[i]
        self._refrescar_lista()

    def guardar_lista(self, ruta=None, silencioso=False):
        ruta = ruta or ARCHIVO
        try:
            with open(ruta, "w", encoding="utf-8") as fh:
                json.dump([{k: v for k, v in d.items() if k != "estado"} for d in self.equipos],
                          fh, indent=2, ensure_ascii=False)
            if not silencioso:
                self.msg(f"Lista guardada en {os.path.basename(ruta)} (sin contraseñas).", OK_C)
        except OSError as e:
            messagebox.showerror("Error", str(e))

    def cargar_lista(self, ruta=None):
        ruta = ruta or ARCHIVO
        try:
            with open(ruta, encoding="utf-8") as fh:
                self.equipos = json.load(fh)[:MAX_EQUIPOS]
        except (OSError, ValueError):
            self.equipos = [dict(base=b, nombre=n, ip=i, puerto=p, marca=m) for b, n, i, p, m in POR_DEFECTO]
        self._refrescar_lista()

    def cargar_archivo(self):
        ruta = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if ruta:
            self.cargar_lista(ruta)

    # ---- consulta
    def consultar(self):
        if self.trabajando:
            return
        if not self.equipos:
            messagebox.showwarning("Sin equipos", "Agrega al menos un equipo.")
            return
        faltan = [d["nombre"] for d in self.equipos if clave(d) not in self.creds]
        if len(faltan) == len(self.equipos):
            messagebox.showwarning("Contraseñas", "Guarda la contraseña de al menos un equipo.")
            return
        if faltan and not messagebox.askyesno(
                "Equipos sin contraseña",
                f"{len(faltan)} equipo(s) no tienen contraseña y se omitirán:\n\n"
                + ", ".join(faltan[:15]) + "\n\n¿Consultar solo los demás?"):
            return
        try:
            timeout = max(2, int(self.v_timeout.get()))
        except ValueError:
            timeout = 8
        self.guardar_lista(silencioso=True)
        self.res_equipos.clear()
        self.res_camaras.clear()
        for d in self.equipos:
            d["estado"] = "En cola..."
        self._refrescar_lista(self._sel_idx())
        creds = {i: self.creds.get(clave(d)) for i, d in enumerate(self.equipos)}
        self.trabajando = True
        self.btn_cons.habilitar(False)
        self.pendientes = len(self.equipos)
        self.dbg_lineas = {}
        threading.Thread(target=self._trabajo, args=(list(self.equipos), creds, timeout), daemon=True).start()
        self.root.after(150, self._sondear)

    def _trabajo(self, equipos, creds, timeout):
        def uno(i, d):
            dbg = [f"\n===== {d['nombre']} {d['ip']} =====\n"]
            cred = creds.get(i)
            if not cred:
                row, cams = nueva_fila(d), []
                row["Estado_consulta"] = "Sin contraseña (omitido)"
                self.cola.put(("fin", i, row["Estado_consulta"], (row, cams), "".join(dbg)))
                return
            self.cola.put(("estado", i, "Consultando...", None, None))
            try:
                fn = consultar_dahua if d["marca"] == "Dahua" else consultar_hikvision
                row, cams = fn(d, cred[0], cred[1], timeout, dbg)
            except (AuthError, ConnectionError) as e:
                row, cams = nueva_fila(d), []
                row["Estado_consulta"] = str(e)
            except Exception as e:  # noqa: BLE001
                row, cams = nueva_fila(d), []
                row["Estado_consulta"] = f"Error: {e.__class__.__name__}: {e}"
            self.cola.put(("fin", i, row["Estado_consulta"], (row, cams), "".join(dbg)))

        with ThreadPoolExecutor(max_workers=4) as ex:
            for i, d in enumerate(equipos):
                ex.submit(uno, i, d)

    def _sondear(self):
        try:
            while True:
                tipo, i, estado, datos, dbg = self.cola.get_nowait()
                self.equipos[i]["estado"] = estado
                self._actualizar_fila(i)
                if tipo == "fin":
                    self.res_equipos[i], self.res_camaras[i] = datos
                    self.dbg_lineas[i] = dbg
                    self.pendientes -= 1
                    self.msg(f"Consultando… faltan {self.pendientes} equipo(s).", MUTED)
        except queue.Empty:
            pass
        if self.pendientes > 0:
            self.root.after(150, self._sondear)
            return
        self.trabajando = False
        self.btn_cons.habilitar(True)
        ok = sum(1 for r in self.res_equipos.values() if r["Estado_consulta"] == "OK")
        self.msg(f"Listo: {ok}/{len(self.equipos)} equipos consultados con éxito.", OK_C if ok else WARN_C)
        if self.v_debug.get():
            with open("debug_grabadores.txt", "w", encoding="utf-8") as fh:
                fh.write("".join(self.dbg_lineas[i] for i in sorted(self.dbg_lineas)))
        self._mostrar_resultados()
        self._ir("result")

    # ---- resultados
    def _mostrar_resultados(self):
        self.tr_eq.delete(*self.tr_eq.get_children())
        for i in sorted(self.res_equipos):
            r = self.res_equipos[i]
            self.tr_eq.insert("", "end", iid=str(i), tags=("ok" if r["Estado_consulta"] == "OK" else "err",),
                              values=(r["Nombre"], r["IP"], r["Marca"], r["Tipo"], r["Modelo"], r["Serial"],
                                      r["MAC"], r["Canales_total"], r["Canales_en_uso"], r["Capacidad_TB"],
                                      r["Estado_consulta"]))
        self._filtrar_camaras()

    @staticmethod
    def _tag_cam(estado):
        e = (estado or "").lower()
        if e in ("conectada", "habilitada", "habilitado"):
            return "ok"      # verde: activa
        if e:
            return "err"     # rojo: sin conexión / deshabilitada
        return "lista"

    def _filtrar_camaras(self, _e=None):
        self.tr_cam.delete(*self.tr_cam.get_children())
        sel = self.tr_eq.selection()
        idxs = [int(sel[0])] if sel else sorted(self.res_camaras)
        for i in idxs:
            for c in self.res_camaras.get(i, []):
                self.tr_cam.insert("", "end", tags=(self._tag_cam(c["Estado"]),),
                                   values=(c["Equipo"], c["Canal"], c["Nombre_camara"], c["IP_camara"], c["Estado"]))

    def exportar(self):
        if not self.res_equipos:
            messagebox.showinfo("Sin datos", "Primero consulta los equipos.")
            return
        carpeta = filedialog.askdirectory(title="Carpeta donde guardar los CSV")
        if not carpeta:
            return
        sello = datetime.now().strftime("%Y%m%d_%H%M")
        rutas = []
        for nombre, cols, filas in (
                (f"inventario_equipos_{sello}.csv", COLS_EQUIPO,
                 [self.res_equipos[i] for i in sorted(self.res_equipos)]),
                (f"inventario_camaras_{sello}.csv", COLS_CAMARA,
                 [c for i in sorted(self.res_camaras) for c in self.res_camaras[i]])):
            ruta = os.path.join(carpeta, nombre)
            with open(ruta, "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.DictWriter(fh, fieldnames=cols, delimiter=";")
                w.writeheader()
                w.writerows(filas)
            rutas.append(ruta)
        messagebox.showinfo("Exportado", "Archivos generados:\n" + "\n".join(rutas))


if __name__ == "__main__":
    raiz = tk.Tk()
    App(raiz)
    raiz.mainloop()
