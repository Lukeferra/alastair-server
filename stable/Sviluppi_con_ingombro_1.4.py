# -*- coding: utf-8 -*-
import rhinoscriptsyntax as rs
import re
import scriptcontext as sc
import Rhino

# ----------------- CONFIG -----------------
PARENT_LAYER   = "UT_INGOMBRI_UNROLL"
LAYER_UNROLL   = PARENT_LAYER + "::UNROLL_SMASH"
LAYER_TABELLA  = PARENT_LAYER + "::TABELLA"

DEFAULT_SER_LAYER = "SERIGRAFIA"
DEFAULT_SER_TOL   = 0.2
DEFAULT_GAP       = 10.0

# sviluppi "fuori dalle scatole" (come vuoi tu: sotto asse X)
OUTPUT_ORIGIN = (0.0, -30000.0, 0.0)

# tabella (sempre su World XY, Z=0)
TABLE_BASE_POINT   = (30000.0, 0.0, 0.0)

# CSV headers (come nei tuoi export)
CSV_HEADER_MUST_CONTAIN = "COMMESSA"
CSV_COL_ID   = "ID"
CSV_COL_POS  = "Posizione"
CSV_COL_COMP = "Composizione"

# DimStyle tabella (standard)
TABLE_DIMSTYLE_NAME = "Sviluppi"
TABLE_TEXT_HEIGHT_PAPER = 2.5   # "carta"
TABLE_MODEL_SCALE = 5.0         # scala modello
TABLE_TEXT_HEIGHT_MODEL = TABLE_TEXT_HEIGHT_PAPER * TABLE_MODEL_SCALE  # 12.5

# ----------------- LAYER HELPERS -----------------
def ensure_layer(path):
    if rs.IsLayer(path):
        return path
    parts = path.split("::")
    cur = ""
    for i, p in enumerate(parts):
        cur = p if i == 0 else cur + "::" + p
        if not rs.IsLayer(cur):
            rs.AddLayer(cur)
    return path

# ----------------- DIMSTYLE -----------------
def ensure_dimstyle(name, text_height_paper=2.5, model_scale=5.0):
    """
    Crea/aggiorna DimStyle 'name' senza andare in crash se alcune proprietà non esistono.
    Ritorna (dimstyle_id, text_height_model) dove text_height_model = paper * model_scale.
    """
    text_height_model = float(text_height_paper) * float(model_scale)

    def _try_set(ds, prop_name, value):
        try:
            if hasattr(ds, prop_name):
                setattr(ds, prop_name, value)
                return True
        except:
            pass
        return False

    # trova o crea
    try:
        idx = sc.doc.DimStyles.Find(name, True)
    except:
        idx = -1

    if idx < 0:
        ds = Rhino.DocObjects.DimensionStyle()
        ds.Name = name

        # Impostazioni sicure
        _try_set(ds, "TextHeight", float(text_height_paper))

        # Proviamo tutte le varianti possibili per la scala (dipende dalla build)
        # Se nessuna esiste, pace: non è critica per la tabella.
        _try_set(ds, "DimScale", float(model_scale))
        _try_set(ds, "Scale", float(model_scale))
        _try_set(ds, "DimensionScale", float(model_scale))
        _try_set(ds, "ModelSpaceScale", float(model_scale))
        _try_set(ds, "AnnotationScale", float(model_scale))

        try:
            idx = sc.doc.DimStyles.Add(ds)
        except:
            idx = -1
    else:
        try:
            ds = sc.doc.DimStyles[idx]
            _try_set(ds, "TextHeight", float(text_height_paper))

            _try_set(ds, "DimScale", float(model_scale))
            _try_set(ds, "Scale", float(model_scale))
            _try_set(ds, "DimensionScale", float(model_scale))
            _try_set(ds, "ModelSpaceScale", float(model_scale))
            _try_set(ds, "AnnotationScale", float(model_scale))

            sc.doc.DimStyles.Modify(ds, idx, True)
        except:
            pass

    # best-effort: set corrente
    try:
        sc.doc.DimStyles.Current = idx
    except:
        pass

    dim_id = None
    try:
        dim_id = sc.doc.DimStyles[idx].Id
    except:
        dim_id = None

    return dim_id, text_height_model


# ----------------- AXIS / PLANE -----------------
def unitize(v):
    l = rs.VectorLength(v)
    if not l or l == 0:
        return None
    return rs.VectorScale(v, 1.0 / l)

def make_axis_plane_from_line(line_id):
    a = rs.CurveStartPoint(line_id)
    b = rs.CurveEndPoint(line_id)
    v = rs.VectorCreate(b, a)
    v.Z = 0.0
    xaxis = unitize(v)
    if not xaxis:
        return None
    zaxis = (0, 0, 1)
    yaxis = rs.VectorCrossProduct(zaxis, xaxis)
    yaxis = unitize(yaxis)
    if not yaxis:
        return None
    return rs.PlaneFromFrame((0, 0, 0), xaxis, yaxis)

# ----------------- DOC OBJECT TRACKING -----------------
def all_doc_ids():
    return set(rs.AllObjects(select=False, include_lights=False, include_grips=False) or [])

def get_new_objects(before_ids):
    after_ids = all_doc_ids()
    return list(after_ids - before_ids)

# ----------------- UNROLL / SMASH -----------------
def run_unroll_or_smash_with_additional(method, surface_id, additional_ids, rel_tol=0.01):
    rs.UnselectAllObjects()
    rs.SelectObject(surface_id)
    if additional_ids:
        rs.SelectObjects(additional_ids)

    if method == "UnrollSrf":
        cmd = (
            '_-UnrollSrf '
            '_Explode=No '
            '_Labels=No '
            '_KeepProperties=No '
            '_RelativeTolerance={} '
            '_Enter'
        ).format(rel_tol)
    else:
        cmd = (
            '_-Smash '
            '_Explode=Yes '
            '_Labels=No '
            '_KeepProperties=No '
            '_Enter'
        )
    return rs.Command(cmd, echo=False)

def choose_main_flattened_object(new_ids):
    # preferisci superfici/brep
    srf_like = [i for i in new_ids if rs.IsSurface(i) or rs.IsPolysurface(i)]
    if srf_like:
        best = None
        best_area = -1
        for i in srf_like:
            a = rs.SurfaceArea(i)
            if a and a[0] > best_area:
                best_area = a[0]
                best = i
        return best
    # fallback
    return new_ids[0] if new_ids else None

# ----------------- BBOX DIMENSIONS -----------------
def oriented_bbox_dimensions(obj_id, axis_plane):
    bb = rs.BoundingBox(obj_id, axis_plane)
    if not bb or len(bb) != 8:
        return None
    min_pt = bb[0]
    max_pt = bb[6]
    base = max_pt.X - min_pt.X
    alt  = max_pt.Y - min_pt.Y
    return base, alt

# ----------------- SERIGRAFIA: LAYER BASED + TOL -----------------
def get_layer_objects(layer_name):
    if not rs.IsLayer(layer_name):
        return []
    ids = rs.ObjectsByLayer(layer_name, select=False) or []
    out = []
    for i in ids:
        if rs.IsCurve(i) or rs.IsPoint(i):
            out.append(i)
    return out

def point_to_brep_distance(brep_id, pt):
    if rs.IsSurface(brep_id):
        uv = rs.SurfaceClosestPoint(brep_id, pt)
        if not uv: return None
        p2 = rs.EvaluateSurface(brep_id, uv[0], uv[1])
        return rs.Distance(pt, p2)
    if rs.IsPolysurface(brep_id):
        cp = rs.BrepClosestPoint(brep_id, pt)
        if not cp: return None
        return rs.Distance(pt, cp[0])
    return None

def curve_on_surface_within_tol(srf_id, crv_id, tol):
    try:
        dom = rs.CurveDomain(crv_id)
        t0, t1 = dom[0], dom[1]
        tm = (t0 + t1) / 2.0
        pts = [rs.EvaluateCurve(crv_id, t0),
               rs.EvaluateCurve(crv_id, tm),
               rs.EvaluateCurve(crv_id, t1)]
    except:
        return False

    for p in pts:
        if not p:
            return False
        d = point_to_brep_distance(srf_id, p)
        if d is None or d > tol:
            return False
    return True

def point_on_surface_within_tol(srf_id, pt_id, tol):
    p = rs.PointCoordinates(pt_id)
    if not p:
        return False
    d = point_to_brep_distance(srf_id, p)
    return (d is not None and d <= tol)

def collect_serigraphy_for_surface(srf_id, ser_ids, tol):
    keep = []
    for sid in ser_ids:
        if rs.IsCurve(sid):
            if curve_on_surface_within_tol(srf_id, sid, tol):
                keep.append(sid)
        elif rs.IsPoint(sid):
            if point_on_surface_within_tol(srf_id, sid, tol):
                keep.append(sid)
    return keep

# ----------------- CSV: ORDER + COMPOSIZIONE -----------------
def normalize_key(s):
    if not s:
        return None
    return re.sub(r"[\s\-_]", "", s.strip().upper())

def read_csv_maps(csv_path):
    """
    Ritorna:
      comp_by_pos: Posizione(normalizzata) -> Composizione
      comp_by_id:  ID -> Composizione
    """
    try:
        raw = open(csv_path, "rb").read()
        try:
            txt = raw.decode("utf-8-sig")
        except:
            try:
                txt = raw.decode("utf-8")
            except:
                txt = raw.decode("latin-1")
    except Exception as e:
        print("Errore lettura CSV: {}".format(e))
        return None, None

    lines = [l for l in txt.splitlines() if l.strip()]
    if not lines:
        return None, None

    delim = ";" if any(";" in l for l in lines[:150]) else ","

    header_idx = None
    header_cols = None
    for i, line in enumerate(lines):
        cols = [c.strip() for c in line.split(delim)]
        if (CSV_HEADER_MUST_CONTAIN in cols and CSV_COL_ID in cols and CSV_COL_POS in cols):
            header_idx = i
            header_cols = cols
            break

    if header_idx is None:
        print("Header tabella non trovato nel CSV.")
        return None, None

    id_idx  = header_cols.index(CSV_COL_ID)
    pos_idx = header_cols.index(CSV_COL_POS)
    comp_idx = header_cols.index(CSV_COL_COMP) if (CSV_COL_COMP in header_cols) else None

    comp_by_pos = {}
    comp_by_id  = {}

    for line in lines[header_idx + 1:]:
        cols = [c.strip() for c in line.split(delim)]
        if len(cols) <= max(id_idx, pos_idx):
            continue

        cid = cols[id_idx]
        cpos = normalize_key(cols[pos_idx])

        comp = ""
        if comp_idx is not None and comp_idx < len(cols):
            comp = cols[comp_idx]

        if cid and comp and cid not in comp_by_id:
            comp_by_id[cid] = comp

        if cpos and comp and cpos not in comp_by_pos:
            comp_by_pos[cpos] = comp

    return comp_by_pos, comp_by_id

def extract_id_from_name(name):
    if not name:
        return None
    s = name.strip()
    m = re.search(r"(\d+)\s*$", s)
    if m:
        return m.group(1)
    m = re.search(r"\b(\d+)\b", s)
    return m.group(1) if m else None

def extract_pos_from_name(name):
    """
    Estrae Posizione dal nome e normalizza.
    Esempi: L01STB, L 01 STB, L01-STB, PS 12, STB-03...
    """
    if not name:
        return None
    n = name.strip().upper()

    m = re.search(r"\bL\s*\d{1,3}(?:\s*-\s*\d{1,3})?\s*(?:STB|SB)\b", n)
    if m:
        return normalize_key(m.group(0))

    m = re.search(r"\bPS\s*[-_ ]?\s*\d+\b", n)
    if m:
        return normalize_key(m.group(0))

    m = re.search(r"\b(?:STB|SB)\s*[-_ ]?\s*\d+\b", n)
    if m:
        return normalize_key(m.group(0))

    return None

def find_comp_from_name(name, comp_by_pos, comp_by_id):
    """
    1) match per Posizione estratta
    2) match per ID estratto
    3) fallback: cerca nel nome normalizzato una qualunque Posizione presente nel CSV
    """
    if not name:
        return ""

    nkey = normalize_key(name) or ""
    pos_key = extract_pos_from_name(name)
    id_key  = extract_id_from_name(name)

    if comp_by_pos and pos_key and pos_key in comp_by_pos:
        return comp_by_pos.get(pos_key, "")

    if comp_by_id and id_key and id_key in comp_by_id:
        return comp_by_id.get(id_key, "")

    if comp_by_pos and nkey:
        keys = sorted(comp_by_pos.keys(), key=lambda k: len(k), reverse=True)
        for k in keys:
            if k and (k in nkey):
                return comp_by_pos.get(k, "")

    return ""

# ----------------- PICK ORDER (SURFACES) -----------------
def get_surfaces_pick_order(prompt):
    """
    Selezione in ordine (click o finestra). Rhino decide l'ordine interno della finestra,
    ma NON facciamo nessun sort alfabetico dopo.
    """
    go = Rhino.Input.Custom.GetObject()
    go.SetCommandPrompt(prompt)
    go.GeometryFilter = Rhino.DocObjects.ObjectType.Surface | Rhino.DocObjects.ObjectType.Brep
    go.SubObjectSelect = False
    go.EnablePreSelect(True, True)
    go.GetMultiple(1, 0)
    if go.CommandResult() != Rhino.Commands.Result.Success:
        return None

    ids = []
    for i in range(go.ObjectCount):
        ids.append(go.Object(i).ObjectId)
    return ids

# ----------------- PACKING -----------------
def bbox_2d(ids):
    bb = rs.BoundingBox(ids)
    if not bb or len(bb) != 8:
        return None
    minx = bb[0].X
    miny = bb[0].Y
    maxx = bb[6].X
    maxy = bb[6].Y
    return minx, miny, maxx, maxy

def move_group_to_cursor(ids, cursor_x, baseline_y=0.0):
    b = bbox_2d(ids)
    if not b:
        return None
    minx, miny, maxx, maxy = b
    dx = cursor_x - minx
    dy = baseline_y - miny
    rs.MoveObjects(ids, (dx, dy, 0))
    return (maxx - minx)

# ----------------- TABLE -----------------
def safe_text(x):
    if x is None:
        return u"-"
    try:
        if isinstance(x, unicode):
            s = x
        else:
            s = unicode(x)
    except:
        try:
            s = unicode(str(x), "utf-8", "ignore")
        except:
            s = u""

    s = s.replace(u"\r\n", u" ").replace(u"\n", u" ").replace(u"\r", u" ").replace(u"\t", u" ")
    s = u"".join(ch for ch in s if ord(ch) >= 32)
    s = s.strip()
    return s if s else u"-"

def create_text_table(data, base_point=TABLE_BASE_POINT, text_height_model=12.5):
    """
    Tabella con rs.AddText e CPlane temporaneo World XY (no rotazioni strane).
    """
    view = rs.CurrentView()
    old_cplane = rs.ViewCPlane(view)
    try:
        rs.ViewCPlane(view, rs.WorldXYPlane())

        bp = rs.CreatePoint(float(base_point[0]), float(base_point[1]), 0.0)

        row_h = 150
        col_w = [1600, 900, 900, 2200]  # Nome, Base, Altezza, Composizione

        for r, row in enumerate(data):
            for c, cell in enumerate(row):
                xoff = sum(col_w[:c])
                pt = rs.PointAdd(bp, (xoff, -r * row_h, 0))
                pt = rs.CreatePoint(pt.X, pt.Y, 0.0)

                txt = safe_text(cell)
                tid = rs.AddText(txt, pt, height=float(text_height_model), font="Arial", justification=2)
                if tid:
                    rs.ObjectLayer(tid, LAYER_TABELLA)

    finally:
        rs.ViewCPlane(view, old_cplane)

# ----------------- MAIN -----------------
def main():
    ensure_layer(PARENT_LAYER)
    ensure_layer(LAYER_UNROLL)
    ensure_layer(LAYER_TABELLA)

    # standard tabella
    _dim_id, table_h_model = ensure_dimstyle(
        TABLE_DIMSTYLE_NAME,
        text_height_paper=TABLE_TEXT_HEIGHT_PAPER,
        model_scale=TABLE_MODEL_SCALE
    )

    rs.EnableRedraw(True)

    # UI: CSV solo per composizione (ordine = selezione)
    opts = rs.GetBoolean(
        "Opzioni script",
        (("UsaCSVComposizione", "No", "Si"),
         ("IncludiSerigrafiaDaLayer", "No", "Si"),
         ("DisposizioneInFila", "No", "Si")),
        (True, True, True)
    )
    if opts is None:
        return

    use_csv_comp, include_ser, do_pack = opts

    axis_line = rs.GetObject("Seleziona la LINEA asse nave (direzione BASE)", rs.filter.curve, preselect=False)
    if not axis_line:
        print("Nessuna linea selezionata.")
        return

    axis_plane = make_axis_plane_from_line(axis_line)
    if not axis_plane:
        print("Linea asse non valida.")
        return

    method = rs.GetString("Metodo di stesura?", "UnrollSrf", ["UnrollSrf", "Smash"]) or "UnrollSrf"

    # serigrafia
    ser_ids = []
    ser_tol = DEFAULT_SER_TOL
    if include_ser:
        ser_layer = rs.GetString("Nome layer serigrafia", DEFAULT_SER_LAYER) or DEFAULT_SER_LAYER
        ser_tol = rs.GetReal("Tolleranza serigrafia (mm)", DEFAULT_SER_TOL, minimum=0.0) or DEFAULT_SER_TOL
        ser_ids = get_layer_objects(ser_layer)
        if not ser_ids:
            print("Nota: nessuna curva/punto su layer '{}'. Continuo senza serigrafia.".format(ser_layer))
            include_ser = False

    # CSV composizione
    comp_by_pos = None
    comp_by_id = None
    if use_csv_comp:
        csv_path = rs.OpenFileName("Seleziona CSV lista vetri (per Composizione)", "CSV (*.csv)|*.csv||")
        if not csv_path:
            use_csv_comp = False
        else:
            comp_by_pos, comp_by_id = read_csv_maps(csv_path)
            if not comp_by_pos and not comp_by_id:
                print("CSV non utilizzabile: composizione non compilata.")
                use_csv_comp = False

    gap = DEFAULT_GAP
    if do_pack:
        gap = rs.GetReal("Spaziatura tra sagome (mm)", DEFAULT_GAP, minimum=0.0) or DEFAULT_GAP

    # superfici in ordine di selezione
    srfs = get_surfaces_pick_order("Seleziona le SUPERFICI da sviluppare (ordine = selezione)")
    if not srfs:
        print("Nessuna superficie selezionata.")
        return

    rs.EnableRedraw(False)

    results = []
    all_group_ids = []

    for srf in srfs:
        before = all_doc_ids()

        additional = []
        if include_ser:
            additional = collect_serigraphy_for_surface(srf, ser_ids, ser_tol)

        ok = run_unroll_or_smash_with_additional(method, srf, additional, rel_tol=0.01)
        if not ok:
            print("Comando {} fallito per {}".format(method, srf))
            continue

        new_ids = get_new_objects(before)
        if not new_ids:
            print("Nessun output creato per {}".format(srf))
            continue

        for nid in new_ids:
            try:
                rs.ObjectLayer(nid, LAYER_UNROLL)
            except:
                pass

        main_flat = choose_main_flattened_object(new_ids)
        if not main_flat:
            print("Nessun output principale per {}".format(srf))
            continue

        dims = oriented_bbox_dimensions(main_flat, axis_plane)
        if not dims:
            print("BBox non calcolabile per {}".format(main_flat))
            continue

        base, alt = dims
        name = rs.ObjectName(srf) or "ID:{}".format(srf)

        comp = ""
        if use_csv_comp and (comp_by_pos or comp_by_id):
            comp = find_comp_from_name(name, comp_by_pos, comp_by_id)

        all_group_ids.extend(list(new_ids))

        results.append({
            "name": name,
            "base": int(round(abs(base))),
            "alt":  int(round(abs(alt))),
            "comp": comp,
            "group_ids": list(new_ids)
        })

    if not results:
        rs.EnableRedraw(True)
        print("Nessun risultato valido.")
        return

    # packing (segue l'ordine di selezione, perché NON sortiamo)
    if do_pack:
        cursor_x = 0.0
        baseline_y = 0.0
        for it in results:
            w = move_group_to_cursor(it["group_ids"], cursor_x, baseline_y=baseline_y)
            if w is None:
                continue
            cursor_x += w + gap

    # offset globale sviluppi
    if all_group_ids:
        rs.MoveObjects(all_group_ids, OUTPUT_ORIGIN)

    # tabella
    table = [("Nome", "Base", "Altezza", "Composizione")]
    for it in results:
        table.append((it["name"], str(it["base"]), str(it["alt"]), it.get("comp", "")))

    create_text_table(table, base_point=TABLE_BASE_POINT, text_height_model=table_h_model)

    rs.EnableRedraw(True)
    print("Fatto. Righe tabella: {}".format(len(table) - 1))

main()
