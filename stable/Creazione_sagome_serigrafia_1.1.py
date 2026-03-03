# -*- coding: utf-8 -*-
"""
Creazione_sagome_serigrafia_1.1
Rhino 7 - IronPython 2.7

------------------------------------------------------------------------------
  Copyright (c) 2024-2025  Luca Ferrari
  Tutti i diritti riservati.

  Il presente script, la sua architettura, logica e design sono opera
  originale di Luca Ferrari e sono protetti dalle leggi sul diritto d'autore.
  E' vietata la riproduzione, distribuzione o modifica, anche parziale,
  senza esplicito consenso scritto dell'autore.

  Author  : Luca Ferrari
  Version : 1.1
  Contact : (inserire contatto se desiderato)
------------------------------------------------------------------------------

Funzionalita' principali:
- Selezione bordo vetro e serigrafia
- G0 Relief automatico sugli angoli della serigrafia (cerchi R10)
- Maschiature: orientamento rigido dal file libreria, micro-offset per CurveBoolean
- Divisione automatica tratti > 1800mm (maschiatura aggiuntiva al centro)
- Ganci automatici a distanza configurabile da ogni maschiatura,
  con profilo ridotto per gap < 45mm
- Raccordo R5 tra gancio e bordo vetro
- Regioni (CurveBoolean unica): filtro microsliver e regione grande
- Etichette testo per ogni regione (Arial 4mm, giallo)
- Export DXF per regione con schema configurabile
- Layer finali: colore bianco, solo Maschere::REGIONI e Maschere::TESTI visibili
"""

import os
import math
import Rhino
import rhinoscriptsyntax as rs
import scriptcontext as sc
import System
from Rhino.Geometry import (
    Point3d, Vector3d, Plane, Transform, Curve,
    CurveOffsetCornerStyle, AreaMassProperties, LineCurve
)

# ----------------------------
# CONFIG
# ----------------------------

PROD_MODE = False

DELETE_INPUT_PROFILES = False if not PROD_MODE else True
CREATE_GAP_DEBUG_LINES = False
KEEP_CENTRAL_CURVE = False

DO_REGIONS_BOOLEAN = True
REGION_KEEP_RATIO = 0.07
REGION_MIN_AREA = 500.0   # mm² - scarta microsliver da micro-offset
REGION_MAX_AREA_RATIO = 0.80  # scarta la regione "grande" (interno vetro) se > 80% dell'area totale
REGION_LAYER = "Maschere::REGIONI"

LIB_3DM_PATH = r"C:\AUTOMAZIONE\Maschere serigrafiche\Maschere_serigrafiche.3dm"

LIB_LAYERS = {
    "Maschiatura_1": {"min_gap": 51.0, "total_width": 94.08},
    "Maschiatura_2": {"min_gap": 20.0, "total_width": 63.03},
    "Maschiatura_3": {"min_gap": 90.0, "total_width": 134.08},
    "Maschiatura_4": {"min_gap": 16.0, "total_width": 40.0},
}

MICRO_OFFSET = 0.05

# --- G0 Relief sulla serigrafia ---
G0_RELIEF_RADIUS = 10.0  # raggio cerchio di relief sugli angoli G0 della serigrafia

# --- Gancio ---
GANCIO_LAYER         = "Gancio"
GANCIO_LAYER_SMALL   = "Gancio_<45mm"   # profilo alternativo per gap < 45mm
GANCIO_GAP_THRESHOLD = 45.0             # mm: sotto questa soglia usa GANCIO_LAYER_SMALL
GANCIO_LAYER_OUT     = "Maschere::GANCI"  # layer dove finiscono i ganci raccordati
GANCIO_OFFSET_MM         = 85.0   # distanza gancio standard dalla maschiatura (mm)
GANCIO_OFFSET_SMALL_MM   = 115.0  # distanza gancio <45mm dalla maschiatura (mm)

# --- Testi etichette ---
TEXT_LAYER         = "Maschere::TESTI"
TEXT_HEIGHT        = 4.0          # altezza testo in mm
TEXT_LINE_SPACING  = 6.0          # interlinea in mm (TEXT_HEIGHT * 1.5)
TEXT_FONT          = "Arial"
TEXT_COLOR_YELLOW  = None         # impostato a runtime (System.Drawing.Color.Yellow)

# --- Export DXF ---
DXF_ROOT_PATH      = r"C:\MASCHERE SERIGRAFICHE"  # cartella radice
DXF_VERSION        = "R2010"                       # AutoCAD 2010 (R18)
DXF_SCHEME         = "ESPORTAZIONE DXF PER SKELET" # schema export DXF/DWG

# --- Divisione automatica tratti lunghi ---
MAX_SEGMENT_LENGTH = 1800.0  # mm: se un tratto tra maschiature supera questo valore,
                             # viene aggiunta automaticamente una maschiatura a meta'

EXTEND_EXTRA = 5.0
EXTEND_STEP = 10.0
EXTEND_MAX_ITERS = 25

DEBUG_REGIONS = (not PROD_MODE)  # stampa diagnostica regioni


# ----------------------------
# DOC TOL
# ----------------------------

def doc_tol():
    return sc.doc.ModelAbsoluteTolerance

def doc_angle_tol():
    try:
        return sc.doc.ModelAngleToleranceRadians
    except:
        return Rhino.RhinoMath.ToRadians(1.0)


# ----------------------------
# G0 RELIEF SULLA SERIGRAFIA
# ----------------------------

def g0_points_via_discontinuity(curve_geom):
    """
    Trova i punti di discontinuita' G1 (angoli G0) su una curva.
    Ritorna lista di Point3d.
    """
    pts = []
    if not curve_geom:
        return pts
    dom = curve_geom.Domain
    t0 = dom.T0
    t1 = dom.T1
    t = t0
    safety = 0
    while True:
        safety += 1
        if safety > 10000:
            break
        rc, t_next = curve_geom.GetNextDiscontinuity(
            Rhino.Geometry.Continuity.G1_continuous, t, t1
        )
        if not rc:
            break
        p = curve_geom.PointAt(t_next)
        if not pts:
            pts.append(p)
        else:
            if p.DistanceTo(pts[-1]) > doc_tol() * 2:
                pts.append(p)
        t = t_next + doc_tol() * 10
        if t >= t1:
            break
    return pts


def apply_g0_relief(ser_id):
    """
    Trova gli angoli G0 sulla serigrafia, crea cerchi di raggio G0_RELIEF_RADIUS
    su ciascuno, poi esegue CreateBooleanUnion con la serigrafia stessa.
    Ritorna (new_ser_id, n_relief, g0_pts):
      - new_ser_id: GUID della serigrafia modificata (o originale se nessun G0)
      - n_relief: numero di cerchi applicati
      - g0_pts: lista dei punti G0 trovati
    La curva originale ser_id NON viene cancellata: il chiamante decide.
    """
    ser_geom = rs.coercecurve(ser_id)
    if not ser_geom:
        return ser_id, 0, []

    g0_pts = g0_points_via_discontinuity(ser_geom)

    if not g0_pts:
        return ser_id, 0, []

    # Crea cerchi di relief
    circle_geoms = []
    circle_ids = []
    for p in g0_pts:
        cid = rs.AddCircle(p, G0_RELIEF_RADIUS)
        if cid:
            circle_ids.append(cid)
            cg = rs.coercecurve(cid)
            if cg:
                circle_geoms.append(cg)

    if not circle_geoms:
        return ser_id, 0, g0_pts

    # Union 2D: serigrafia + tutti i cerchi
    try:
        res = Curve.CreateBooleanUnion([ser_geom] + circle_geoms, doc_tol())
    except Exception as e:
        print("G0 relief union fallita:", e)
        res = None

    # Pulizia cerchi temporanei
    for cid in circle_ids:
        try:
            rs.DeleteObject(cid)
        except:
            pass

    if res and len(res) > 0:
        new_id = sc.doc.Objects.AddCurve(res[0])
        if new_id and new_id != System.Guid.Empty:
            set_layer(new_id, "Maschere")
            print("G0 relief: {} angoli processati, serigrafia aggiornata.".format(len(g0_pts)))
            return new_id, len(g0_pts), g0_pts

    # Fallback: ritorna originale
    print("G0 relief: union fallita, serigrafia invariata.")
    return ser_id, 0, g0_pts


# ----------------------------
# LAYERS
# ----------------------------

def ensure_layers():
    for ln in ["Maschere", "Maschere::01", "Maschere::DEBUG", REGION_LAYER, GANCIO_LAYER_OUT, TEXT_LAYER]:
        if not rs.IsLayer(ln):
            rs.AddLayer(ln)
    try:
        rs.LayerVisible("Maschere::DEBUG", (not PROD_MODE))
    except:
        pass

def set_layer(obj_id, layer_name):
    try:
        if obj_id and rs.IsObject(obj_id):
            rs.ObjectLayer(obj_id, layer_name)
    except:
        pass


# ----------------------------
# INPUT
# ----------------------------

def get_closed_profile(prompt):
    ids = rs.GetObjects(prompt, rs.filter.curve, preselect=False, select=True)
    if not ids:
        return None

    if len(ids) == 1:
        crv = rs.coercecurve(ids[0])
        if not crv or not crv.IsClosed:
            print("Curva non chiusa. Seleziona più segmenti unibili oppure un profilo chiuso.")
            return None
        return ids[0]

    joined = rs.JoinCurves(ids, delete_input=False)
    if not joined or len(joined) != 1:
        print("Non riesco a unire le curve in un unico profilo chiuso.")
        return None

    j_id = joined[0]
    j_crv = rs.coercecurve(j_id)
    if not j_crv or not j_crv.IsClosed:
        print("Il profilo unito non è chiuso.")
        return None

    return j_id

def get_point_on_curve_repeated(curve_id):
    pts = []
    while True:
        gp = rs.GetPointOnCurve(curve_id, "Clicca punti divisione sul BORDO (Invio per finire)")
        if not gp:
            break
        pts.append(Point3d(gp.X, gp.Y, gp.Z))
    return pts

def try_get_curve_plane(crv):
    try:
        ok, pl = crv.TryGetPlane()
        if ok:
            return pl
    except:
        pass
    return Plane.WorldXY


# ----------------------------
# GAP
# ----------------------------

def gap_from_border_point_to_serigrafia(p_on_bordo, serigrafia_id):
    s_crv = rs.coercecurve(serigrafia_id)
    if not s_crv:
        return None, None, None

    pl = try_get_curve_plane(s_crv)

    try:
        ok, ts = s_crv.ClosestPoint(p_on_bordo)
    except:
        ok, ts = (False, None)

    if not ok:
        try:
            p_proj = pl.ClosestPoint(p_on_bordo)
            ok, ts = s_crv.ClosestPoint(p_proj)
            if ok:
                p_on_bordo = p_proj
        except:
            ok = False

    if not ok:
        return None, None, None

    ip = s_crv.PointAt(ts)
    v = ip - p_on_bordo
    if v.IsTiny(doc_tol()):
        return ip, None, 0.0

    gap = p_on_bordo.DistanceTo(ip)
    v.Unitize()
    return ip, v, gap


# ----------------------------
# TEMPLATE PICK
# ----------------------------

def pick_maschiatura_key(gap):
    """
    Sceglie il profilo maschiatura piu' adatto al gap misurato.

    min_gap  = distanza minima utilizzabile per quel profilo (sotto questa soglia
               il profilo non puo' fisicamente essere inserito)
    total_width = ampiezza nominale del profilo; se il gap e' maggiore,
                  le curve vengono estese da ensure_intersections_by_extending

    Logica: scegli il profilo con min_gap piu' alto che sia ancora <= gap.
    Cosi' si usa sempre il profilo piu' "generoso" compatibile con lo spazio
    disponibile, lasciando all'estensione il compito di raggiungere bordo e serigrafia.
    Se nessun profilo ha min_gap <= gap, ritorna None (punto saltato).
    """
    if gap is None:
        return None

    # Ordina per min_gap decrescente: proviamo prima il profilo piu' esigente
    opts = sorted(
        [(name, float(meta["min_gap"])) for name, meta in LIB_LAYERS.items()],
        key=lambda x: x[1],
        reverse=True
    )

    for name, mg in opts:
        if gap >= mg:
            return name

    return None  # gap inferiore al min_gap di tutti i profili


# ----------------------------
# LIB LOAD (curve + 2 guide points)
# ----------------------------

def pick_two_farthest_points(pt_list):
    if not pt_list or len(pt_list) < 2:
        return None, None
    bi, bj = 0, 1
    bd2 = pt_list[0].DistanceToSquared(pt_list[1])
    n = len(pt_list)
    for i in range(n):
        for j in range(i + 1, n):
            d2 = pt_list[i].DistanceToSquared(pt_list[j])
            if d2 > bd2:
                bd2, bi, bj = d2, i, j
    return pt_list[bi], pt_list[bj]

def load_template_from_3dm(layer_name, lib_path):
    """
    layer_name: "Maschiatura_1" ecc.
    Cerca:
    - curva sul layer Maschiatura_X
    - 2 guide points sul layer Maschiatura_X (tutti i punti NON chiamati 'pick')
    - pick point sul layer Maschiatura_X::PICK (o su qualunque layer che finisca con ::PICK sotto Maschiatura_X)
      oppure punto con Name == 'pick'
    Ritorna: (template_curve, guide_p1, guide_p2, pick_pt)
    """
    if not os.path.exists(lib_path):
        print("Libreria non trovata:", lib_path)
        return None, None, None, None

    f3dm = Rhino.FileIO.File3dm.Read(lib_path)
    if not f3dm:
        print("Errore lettura libreria:", lib_path)
        return None, None, None, None

    base_layer_index = None
    pick_layer_indexes = set()

    # trova indici layer base e layer pick figli
    for lay in f3dm.Layers:
        if not lay:
            continue
        if lay.FullPath == layer_name or lay.Name == layer_name:
            base_layer_index = lay.Index

        # es: "Maschiatura_1::PICK" (qualsiasi profondità)
        if lay.FullPath.startswith(layer_name + "::") and lay.FullPath.upper().endswith("::PICK"):
            pick_layer_indexes.add(lay.Index)

    if base_layer_index is None:
        print("Layer non trovato in libreria:", layer_name)
        return None, None, None, None

    template_curve = None
    guide_pts = []
    pick_pt = None

    for obj in f3dm.Objects:
        if not obj:
            continue
        att = obj.Attributes
        if not att:
            continue
        geo = obj.Geometry
        if not geo:
            continue

        # --- sul layer base: curva + guide points + eventuale pick per Name ---
        if att.LayerIndex == base_layer_index:
            if template_curve is None and isinstance(geo, Rhino.Geometry.Curve):
                template_curve = geo.DuplicateCurve()
                continue

            if isinstance(geo, Rhino.Geometry.Point):
                nm = ""
                try:
                    nm = (att.Name or "").strip().lower()
                except:
                    nm = ""
                if nm == "pick" and pick_pt is None:
                    pick_pt = Point3d(geo.Location)
                else:
                    guide_pts.append(Point3d(geo.Location))
                continue

        # --- sul layer PICK dedicato: prendo il primo point (o quello chiamato pick) ---
        if att.LayerIndex in pick_layer_indexes and isinstance(geo, Rhino.Geometry.Point):
            if pick_pt is None:
                pick_pt = Point3d(geo.Location)
            else:
                # se ne hai più di uno e uno si chiama pick, preferiscilo
                try:
                    nm = (att.Name or "").strip().lower()
                    if nm == "pick":
                        pick_pt = Point3d(geo.Location)
                except:
                    pass

    if template_curve is None:
        print("Nessuna curva trovata sul layer:", layer_name)
        return None, None, None, None

    p1, p2 = (None, None)
    if len(guide_pts) >= 2:
        p1, p2 = pick_two_farthest_points(guide_pts)

    return template_curve, p1, p2, pick_pt


# ----------------------------
# ORIENT RIGIDO (NO SCALE)
# ----------------------------

def point_at_distance_from_end(crv, dist, from_start=True):
    total_len = crv.GetLength()
    if total_len <= dist + doc_tol():
        return None
    if from_start:
        ok, t = crv.LengthParameter(dist)
    else:
        ok, t = crv.LengthParameter(total_len - dist)
    if not ok:
        return None
    return crv.PointAt(t)

def build_template_plane(template_crv, guide_p1, guide_p2):
    if guide_p1 is not None and guide_p2 is not None:
        x = guide_p2 - guide_p1
        if not x.IsTiny(doc_tol()):
            x.Unitize()
            y = Vector3d.CrossProduct(Vector3d(0, 0, 1), x)
            if y.IsTiny():
                y = Vector3d(-x.Y, x.X, 0.0)
            y.Unitize()
            origin = (guide_p1 + guide_p2) * 0.5
            return Plane(origin, x, y)

    pA = point_at_distance_from_end(template_crv, 10.0, True)
    pB = point_at_distance_from_end(template_crv, 10.0, False)
    if pA is None or pB is None:
        pA = template_crv.PointAtStart
        pB = template_crv.PointAtEnd

    x = pB - pA
    if x.IsTiny(doc_tol()):
        return None
    x.Unitize()
    y = Vector3d.CrossProduct(Vector3d(0, 0, 1), x)
    if y.IsTiny():
        y = Vector3d(-x.Y, x.X, 0.0)
    y.Unitize()
    origin = (pA + pB) * 0.5
    return Plane(origin, x, y)

def orient_template_rigid_with_xform(template_crv, tpl_plane, target_origin, target_xaxis):
    x = Vector3d(target_xaxis)
    if x.IsTiny(doc_tol()):
        return None, None
    x.Unitize()

    y = Vector3d.CrossProduct(Vector3d(0, 0, 1), x)
    if y.IsTiny():
        y = Vector3d(-x.Y, x.X, 0.0)
    y.Unitize()

    target_plane = Plane(Point3d(target_origin), x, y)

    xf = Transform.PlaneToPlane(tpl_plane, target_plane)
    c = template_crv.DuplicateCurve()
    c.Transform(xf)
    return c, xf


# ----------------------------
# INTERSECTIONS + ESTENSIONE ROBUSTA (append/prepend lines)
# ----------------------------

def curve_intersects_target(curve_geom, target_id):
    target = rs.coercecurve(target_id)
    if not target or not curve_geom:
        return False
    events = Rhino.Geometry.Intersect.Intersection.CurveCurve(curve_geom, target, doc_tol(), doc_tol())
    return (events is not None and events.Count > 0)

def _tangent_safe(crv, t, fallback):
    try:
        tan = crv.TangentAt(t)
        if tan.IsTiny():
            return fallback
        tan.Unitize()
        return tan
    except:
        return fallback

def extend_curve_ends_by_lines(curve_geom, step):
    """
    Robust extension:
    - crea una LineCurve in start e end lungo la tangente
    - join [line_start, curve, line_end]
    """
    crv = curve_geom.DuplicateCurve()
    if crv.IsClosed:
        return crv

    sp = crv.PointAtStart
    ep = crv.PointAtEnd

    # fallback axis: linea start->end
    fb = ep - sp
    if fb.IsTiny():
        fb = Vector3d(1, 0, 0)
    fb.Unitize()

    t0 = crv.Domain.T0
    t1 = crv.Domain.T1

    tan_s = _tangent_safe(crv, t0, fb)
    tan_e = _tangent_safe(crv, t1, fb)

    # start: estendo indietro
    ls0 = LineCurve(sp - tan_s * step, sp)
    # end: estendo avanti
    ls1 = LineCurve(ep, ep + tan_e * step)

    joined = Curve.JoinCurves([ls0, crv, ls1], doc_tol())
    if joined and len(joined) > 0:
        return joined[0]
    return crv

def ensure_intersections_by_extending(curve_geom, bordo_id, ser_id):
    crv = curve_geom
    # extra iniziale
    crv = extend_curve_ends_by_lines(crv, EXTEND_EXTRA)

    for _ in range(EXTEND_MAX_ITERS):
        ok_b = curve_intersects_target(crv, bordo_id)
        ok_s = curve_intersects_target(crv, ser_id)
        if ok_b and ok_s:
            return crv
        crv = extend_curve_ends_by_lines(crv, EXTEND_STEP)

    return crv


# ----------------------------
# MICRO OFFSET
# ----------------------------

def micro_offset_both_sides(curve_id, dist, plane, layer_name):
    crv = rs.coercecurve(curve_id)
    if not crv:
        return []
    tol = doc_tol()
    ids = []
    for d in (dist, -dist):
        try:
            offs = crv.Offset(plane, d, tol, CurveOffsetCornerStyle.Round)
        except:
            offs = None
        if not offs:
            continue
        for oc in offs:
            oid = sc.doc.Objects.AddCurve(oc)
            if oid:
                rs.ObjectLayer(oid, layer_name)
                ids.append(oid)
    return ids


# ----------------------------
# REGIONI (centroid/area da curve loop)
# ----------------------------

def _containment(curve_id, pt, plane):
    crv = rs.coercecurve(curve_id)
    if not crv:
        return None
    try:
        return crv.Contains(pt, plane, doc_tol())
    except:
        return None

def _curve_area_centroid(closed_curve):
    try:
        amp = AreaMassProperties.Compute(closed_curve)
        if not amp:
            return None, None
        return amp.Area, amp.Centroid
    except:
        return None, None

def extract_closed_loops_from_brep_naked_edges(brep, tol):
    """
    Estrae i contorni (outer + inner) da un brep usando gli edge NAKED,
    poi li join-a per ottenere loop chiusi.
    Ritorna lista di Curve chiuse.
    """
    if brep is None:
        return []

    edge_curves = []
    try:
        for e in brep.Edges:
            # Naked = bordo esterno "a vista" (inclusi i fori)
            if e.Valence == Rhino.Geometry.EdgeAdjacency.Naked:
                c = e.DuplicateCurve()
                if c:
                    edge_curves.append(c)
    except:
        pass

    if not edge_curves:
        return []

    # Join in loop
    joined = Rhino.Geometry.Curve.JoinCurves(edge_curves, tol)
    if not joined:
        return []

    closed = []
    for c in joined:
        if c and c.IsClosed:
            closed.append(c)

    return closed

def _format_pt_cmd(p):
    # Rhino command line point "x,y,z"
    return "{:.6f},{:.6f},{:.6f}".format(p.X, p.Y, p.Z)

def _area_of_closed_curve(crv_id):
    try:
        a = rs.CurveArea(crv_id)
        if a and len(a) > 0:
            return float(a[0])
    except:
        pass
    return None

def _pick_inward_offset_curve(bordo_id, ser_id, offset_dist, plane):
    """
    Crea offset interno del bordo (verso la serigrafia) di offset_dist mm.
    Strategia: prova +d e -d, sceglie quello con area MINORE del bordo
    (l'offset interno rimpicciolisce la curva).
    Ritorna GUID della curva aggiunta al doc, oppure None.
    La curva va eliminata dal chiamante dopo l'uso.
    """
    bordo = rs.coercecurve(bordo_id)
    if not bordo:
        return None

    tol = doc_tol()

    # Area del bordo originale
    bordo_area = None
    try:
        amp0 = AreaMassProperties.Compute(bordo)
        if amp0:
            bordo_area = amp0.Area
    except:
        pass

    best = None
    best_area_diff = None

    for d in (offset_dist, -offset_dist):
        try:
            offs = bordo.Offset(plane, d, tol, Rhino.Geometry.CurveOffsetCornerStyle.Round)
        except:
            offs = None
        if not offs:
            continue
        c = offs[0]
        if not c.IsClosed:
            continue

        # L'offset interno ha area minore del bordo
        if bordo_area is not None:
            try:
                amp = AreaMassProperties.Compute(c)
                if amp:
                    diff = bordo_area - amp.Area  # positivo = piu' piccolo = interno
                    if diff > 0 and (best_area_diff is None or diff < best_area_diff):
                        best = c
                        best_area_diff = diff
            except:
                pass
        else:
            # fallback: prendi il primo valido
            if best is None:
                best = c

    if best is None:
        return None

    cid = sc.doc.Objects.AddCurve(best)
    return cid if cid else None

def _get_border_params_sorted(bordo_id, division_pts):
    """
    Ritorna lista di parametri t sul bordo in ordine crescente (lungo la curva).
    """
    bordo = rs.coercecurve(bordo_id)
    if not bordo:
        return []

    ts = []
    for p in division_pts:
        try:
            ok, t = bordo.ClosestPoint(p)
            if ok:
                ts.append(t)
        except:
            pass

    if not ts:
        return []

    # ordina e rimuovi duplicati “quasi uguali”
    ts.sort()
    cleaned = [ts[0]]
    for t in ts[1:]:
        if abs(t - cleaned[-1]) > doc_tol():
            cleaned.append(t)

    return cleaned

def _make_pick_points_on_inner_offset(bordo_id, inner_offset_id, division_pts):
    """
    Genera un pick point (interno alla fascia) per ogni settore tra due divisioni successive,
    usando l’offset interno a 5mm.
    """
    bordo = rs.coercecurve(bordo_id)
    inner = rs.coercecurve(inner_offset_id)
    if not bordo or not inner:
        return []

    ts = _get_border_params_sorted(bordo_id, division_pts)
    if len(ts) < 2:
        return []

    dom = bordo.Domain
    picks = []

    def pt_on_inner_near_border_param(t_mid):
        pb = bordo.PointAt(t_mid)
        ok, ti = inner.ClosestPoint(pb)
        if not ok:
            return None
        pi = inner.PointAt(ti)
        return pi

    # segmenti tra t[i] e t[i+1]
    for i in range(len(ts) - 1):
        t0 = ts[i]
        t1 = ts[i + 1]
        tmid = (t0 + t1) * 0.5
        p = pt_on_inner_near_border_param(tmid)
        if p:
            picks.append(p)

    # segmento wrap-around: da ultimo a primo passando per fine dominio
    # (funziona per curve chiuse con dominio continuo)
    t_last = ts[-1]
    t_first = ts[0]
    # midpoint “wrap”: (t_last + (t_first + (dom.T1-dom.T0)))/2 -> poi rimappa nel dominio
    span = (dom.T1 - dom.T0)
    tmid_wrap = (t_last + (t_first + span)) * 0.5
    # rimappa
    while tmid_wrap > dom.T1:
        tmid_wrap -= span

    pwrap = pt_on_inner_near_border_param(tmid_wrap)
    if pwrap:
        picks.append(pwrap)

    return picks

def build_regions_with_pickpoints(bordo_id, ser_id, offset_ids, pick_points):
    """
    Usa Curve.CreateBooleanRegions (RhinoCommon API) invece di RunScript/_-CurveBoolean.
    Raccoglie tutte le curve input come geometria, calcola le regioni booleane,
    filtra per area, aggiunge al doc sul layer REGIONI.
    """
    if not pick_points:
        print("Regioni: nessun pick point disponibile.")
        return []

    tol = doc_tol()
    angle_tol = doc_angle_tol()

    # --- Raccogli le geometrie curve ---
    input_ids = [bordo_id, ser_id] + list(offset_ids)
    input_curves = []
    for cid in input_ids:
        crv = rs.coercecurve(cid)
        if crv:
            input_curves.append(crv)

    if not input_curves:
        print("Regioni: nessuna curva input valida.")
        return []

    # --- Piano di lavoro: usa il piano della prima curva chiusa (bordo) ---
    bordo_crv = rs.coercecurve(bordo_id)
    if bordo_crv:
        plane = try_get_curve_plane(bordo_crv)
    else:
        plane = Plane.WorldXY

    if DEBUG_REGIONS:
        print("Regioni: curve input =", len(input_curves), "| pick points =", len(pick_points))
        print("Regioni: piano usato =", plane)

    # --- Chiama CreateBooleanRegions ---
    try:
        regions = Curve.CreateBooleanRegions(input_curves, plane, pick_points, False, tol)
    except Exception as e:
        print("Regioni: CreateBooleanRegions ha sollevato eccezione:", e)
        regions = None

    if regions is None or regions.RegionCount == 0:
        print("Regioni: CreateBooleanRegions non ha prodotto output.")
        if DEBUG_REGIONS:
            print("  Verifica che le curve si intersechino correttamente nel piano:", plane)
        return []

    if DEBUG_REGIONS:
        print("Regioni trovate (raw):", regions.RegionCount)
        print("Pick points totali:", len(pick_points))
        try:
            for i in range(regions.RegionCount):
                pi = regions.RegionPointIndex(i)
                pt = pick_points[pi] if 0 <= pi < len(pick_points) else None
                print("  Regione {} <- pick_point[{}] = {}".format(i, pi, pt))
        except Exception as e:
            print("  RegionPointIndex debug fallito:", e)

    # --- Aggiungi al doc, filtra per area ---
    # API reale (da dir()): RegionCurves(i) -> array di Curve per la regione i
    # BoundaryCount(i) -> numero di loop (outer+holes) della regione i
    # SegmentCount(i, j) -> numero segmenti del loop j della regione i
    # SegmentDetails(i, j, k) -> (curve_index, reversed) per il segmento k
    # PlanarCurve(curve_index) -> Curve geometrica
    areas = []
    for i in range(regions.RegionCount):
        closed_crv = None

        # Metodo 1: RegionCurves(i) - ritorna direttamente le curve del contorno
        try:
            region_crvs = regions.RegionCurves(i)
            if region_crvs is not None and len(region_crvs) > 0:
                if len(region_crvs) == 1:
                    closed_crv = region_crvs[0]
                else:
                    joined = Curve.JoinCurves(list(region_crvs), tol)
                    closed_crv = joined[0] if joined else None
        except Exception as e1:
            if DEBUG_REGIONS:
                print("  RegionCurves({}) fallito: {}".format(i, e1))

        # Metodo 2: ricostruisci da BoundaryCount + SegmentCount + SegmentDetails + PlanarCurve
        if closed_crv is None:
            try:
                n_boundaries = regions.BoundaryCount(i)
                all_segs = []
                for j in range(n_boundaries):
                    n_segs = regions.SegmentCount(i, j)
                    for k in range(n_segs):
                        det = regions.SegmentDetails(i, j, k)
                        crv_idx = det[0]
                        is_rev = det[1]
                        pc = regions.PlanarCurve(crv_idx)
                        if pc:
                            seg = pc.DuplicateCurve()
                            if is_rev:
                                seg.Reverse()
                            all_segs.append(seg)
                if all_segs:
                    joined = Curve.JoinCurves(all_segs, tol)
                    closed_crv = joined[0] if joined else None
            except Exception as e2:
                if DEBUG_REGIONS:
                    print("  BoundaryCount/SegmentDetails({}) fallito: {}".format(i, e2))

        if closed_crv is None:
            if DEBUG_REGIONS:
                print("  Regione {} ignorata: nessuna curva estratta.".format(i))
            continue

        # Calcola area
        area = None
        try:
            amp = AreaMassProperties.Compute(closed_crv)
            if amp:
                area = amp.Area
        except:
            pass

        if DEBUG_REGIONS:
            print("  Regione {}: area={}".format(i, round(area, 1) if area else None))

        if area is not None and area > tol:
            areas.append((closed_crv, area))

    if not areas:
        print("Regioni: nessuna regione con area valida.")
        return []

    areas.sort(key=lambda x: x[1], reverse=True)

    # Calcola area totale per rilevare la regione "grande" (interno vetro)
    total_area = sum(a for _, a in areas)

    if DEBUG_REGIONS:
        print("Regioni con area: ", [(round(a, 1)) for _, a in areas])
        print("Area totale:", round(total_area, 1))

    # Filtro 1: scarta microsliver assoluti (es. da micro-offset +-0.05mm)
    areas = [(crv, a) for crv, a in areas if a >= REGION_MIN_AREA]

    if not areas:
        print("Regioni: tutte le regioni sono sotto la soglia minima di area ({} mm2).".format(REGION_MIN_AREA))
        return []

    # Filtro 2: scarta la regione "grande" se occupa piu' di REGION_MAX_AREA_RATIO dell'area totale
    # (e' l'interno del vetro, non una fascia di serigrafia)
    areas_filtered = [(crv, a) for crv, a in areas if a / total_area <= REGION_MAX_AREA_RATIO]

    if not areas_filtered:
        # Fallback: se tutto e' stato scartato, usa la soglia relativa classica
        areas_filtered = areas

    if DEBUG_REGIONS:
        print("Regioni dopo filtri: ", [(round(a, 1)) for _, a in areas_filtered])

    # Filtro 3: soglia relativa rispetto alla piu' grande rimasta
    max_a = areas_filtered[0][1]
    threshold = max_a * REGION_KEEP_RATIO

    if DEBUG_REGIONS:
        print("Area massima (filtrata):", round(max_a, 1), "| soglia keep:", round(threshold, 1))

    keep = []
    for crv, a in areas_filtered:
        if a >= threshold:
            oid = sc.doc.Objects.AddCurve(crv)
            if oid:
                try:
                    rs.ObjectLayer(oid, REGION_LAYER)
                except:
                    pass
                keep.append(oid)

    return keep


# ----------------------------
# ONE POINT PIPELINE
# ----------------------------

def build_one_maschiatura(bordo_id, ser_id, click_pt):
    """
    Ritorna: (off_ids, pick_world)
      - off_ids: lista di GUID delle 2 curve offset (±0.05)
      - pick_world: Point3d trasformato del punto "pick" del template (o None)
    """

    bordo_crv = rs.coercecurve(bordo_id)
    if not bordo_crv:
        return [], None

    # 1) proietta/aggancia il click al bordo
    try:
        ok, tb = bordo_crv.ClosestPoint(click_pt)
    except:
        ok, tb = (False, None)

    if not ok:
        return [], None

    p_on_bordo = bordo_crv.PointAt(tb)

    # 2) GAP verso serigrafia (closest point)
    ip, dir_vec, gap = gap_from_border_point_to_serigrafia(p_on_bordo, ser_id)
    if ip is None or dir_vec is None:
        print("Impossibile calcolare GAP. Punto saltato.")
        return [], None

    if CREATE_GAP_DEBUG_LINES and not PROD_MODE:
        seg_id = rs.AddLine(p_on_bordo, ip)
        set_layer(seg_id, "Maschere::DEBUG")

    print("GAP:", round(gap, 2), "mm")

    # 3) scegli maschiatura in base al gap
    key = pick_maschiatura_key(gap)
    if key is None:
        print("GAP troppo piccolo (%.2f mm). Punto saltato." % gap)
        return [], None

    # 4) carica template + guide + pick
    template, gp1, gp2, pick_pt = load_template_from_3dm(key, LIB_3DM_PATH)
    if template is None:
        print("Template non disponibile:", key)
        return [], None

    tpl_plane = build_template_plane(template, gp1, gp2)
    if tpl_plane is None:
        print("Impossibile costruire piano template.")
        return [], None

    # 5) centro sul GAP (midpoint bordo-serigrafia)
    mid = (p_on_bordo + ip) * 0.5

    # 6) orientamento rigido (NO scaling) + prendo trasformazione
    placed, xf = orient_template_rigid_with_xform(template, tpl_plane, mid, dir_vec)
    if placed is None or xf is None:
        print("Errore orientamento maschiatura.")
        return [], None

    # 7) estendi finché interseca sia bordo che serigrafia
    placed2 = ensure_intersections_by_extending(placed, bordo_id, ser_id)

    # 8) aggiungi curva centrale (solo per fare offset), poi micro-offset
    central_id = sc.doc.Objects.AddCurve(placed2)
    if not central_id:
        return [], None
    set_layer(central_id, "Maschere")

    off_ids = micro_offset_both_sides(central_id, MICRO_OFFSET, Plane.WorldXY, "Maschere::01")

    # elimina curva centrale
    if not KEEP_CENTRAL_CURVE:
        try:
            rs.DeleteObject(central_id)
        except:
            pass

    # I pick points per le regioni sono ora generati in main() via offset interno del bordo
    return off_ids, None


def _cut_bordo_at_gancio_small(bordo_crv, gancio_geom, expected_len=180.0, tolerance_pct=0.20):
    """
    Per il gancio <45mm (che sporge verso esterno):
    - Trova le 2 intersezioni tra la curva gancio e il bordo
    - Splitta il bordo nei 2 punti di intersezione
    - Elimina il tratto di lunghezza ~expected_len mm (+-tolerance_pct)

    Accetta una curva bordo (geometria), NON un GUID.
    Ritorna la curva bordo tagliata, oppure None se fallisce.
    Il bordo originale nel documento NON viene modificato.
    """
    tol = doc_tol()
    if not bordo_crv or not gancio_geom:
        return None

    # Intersezioni gancio / bordo
    events = Rhino.Geometry.Intersect.Intersection.CurveCurve(
        gancio_geom, bordo_crv, tol, tol * 100
    )
    if events is None or events.Count < 2:
        if not PROD_MODE:
            print("  _cut_bordo: meno di 2 intersezioni ({})".format(
                events.Count if events else 0))
        return None

    # Raccogli tutti i parametri sul bordo e i punti
    ix_tb = []
    for i in range(events.Count):
        ev = events[i]
        if ev.IsPoint:
            ix_tb.append(ev.ParameterB)

    if len(ix_tb) < 2:
        return None

    ix_tb.sort()

    # Valuta tutti i possibili tratti tra coppie di parametri successivi
    # e scegli quello la cui lunghezza e' piu' vicina a expected_len
    dom = bordo_crv.Domain
    total_len = bordo_crv.GetLength()
    best_pair = None
    best_delta = float('inf')

    for i in range(len(ix_tb)):
        for j in range(i+1, len(ix_tb)):
            t0, t1 = ix_tb[i], ix_tb[j]
            seg = bordo_crv.Trim(t0, t1)
            if seg is None:
                continue
            seg_len = seg.GetLength()
            delta = abs(seg_len - expected_len)
            if delta < best_delta:
                best_delta = delta
                best_pair = (t0, t1, seg_len)

    if best_pair is None:
        return None

    t0, t1, seg_len = best_pair
    tol_len = expected_len * tolerance_pct
    if best_delta > tol_len:
        if not PROD_MODE:
            print("  _cut_bordo: tratto trovato {:.1f}mm, atteso {}mm +/-{}mm -- skippato".format(
                seg_len, expected_len, tol_len))
        return None

    if not PROD_MODE:
        print("  _cut_bordo: eliminato tratto {:.1f}mm tra t={:.4f} e t={:.4f}".format(
            seg_len, t0, t1))

    # Splitta il bordo nei 2 parametri e sostituiscilo con i pezzi rimasti
    pieces = bordo_crv.Split([t0, t1])
    if not pieces or len(pieces) < 2:
        return None

    # Tieni tutti i pezzi tranne quello da ~expected_len
    kept = []
    for pc in pieces:
        pc_len = pc.GetLength()
        if abs(pc_len - expected_len) > tol_len:
            kept.append(pc)

    if not kept:
        if not PROD_MODE:
            print("  _cut_bordo: nessun pezzo mantenuto dopo split")
        return None

    # Join dei pezzi mantenuti -> nuovo bordo (aperto nei due punti di taglio)
    if len(kept) == 1:
        new_bordo_crv = kept[0]
    else:
        joined = Curve.JoinCurves(kept, tol * 10)
        if not joined:
            return None
        new_bordo_crv = joined[0]

    return new_bordo_crv


def _cut_bordo_at_endpoints(bordo_crv, t0, t1, min_seg_len=50.0):
    """
    Splitta bordo_crv ai parametri t0 e t1 ed elimina il tratto
    compreso tra loro (quello piu' corto, >= min_seg_len).
    Ritorna la curva risultante o None se fallisce.
    """
    tol = doc_tol()
    if t0 > t1:
        t0, t1 = t1, t0

    pieces = bordo_crv.Split([t0, t1])
    if not pieces or len(pieces) < 2:
        return None

    # Elimina il pezzo piu' corto tra t0 e t1
    kept = []
    for pc in pieces:
        pc_len = pc.GetLength()
        # Il tratto da eliminare e' quello con lunghezza minima
        # e che inizia/finisce vicino a t0..t1
        mid_t = (pc.Domain.T0 + pc.Domain.T1) * 0.5
        # verifica se il punto medio del pezzo cade tra t0 e t1 nel dominio orig
        # (indicativo del tratto interno)
        is_inner = (t0 < mid_t < t1)
        if is_inner and pc_len < (t1 - t0) * 2:
            continue  # scarta questo tratto
        kept.append(pc)

    if not kept:
        return None

    if len(kept) == 1:
        return kept[0]

    joined = Curve.JoinCurves(kept, tol * 10)
    if not joined:
        return None
    return joined[0]


# ----------------------------
# GANCIO
# ----------------------------

def load_gancio_from_3dm(lib_path, layer_name=None):
    """
    Carica dal layer specificato (default: GANCIO_LAYER) del file libreria:
      - 1 curva aperta  -> profilo gancio
      - 2 punti "Posizione_gancio" -> definiscono asse del piano template
      - 1 punto "dir"   -> direzione verso l'interno (serigrafia)
    Ritorna: (gancio_curve, pos1, pos2, dir_pt)
    """
    if layer_name is None:
        layer_name = GANCIO_LAYER
    if not os.path.exists(lib_path):
        print("Libreria non trovata:", lib_path)
        return None, None, None, None, []

    f3dm = Rhino.FileIO.File3dm.Read(lib_path)
    if not f3dm:
        print("Errore lettura libreria:", lib_path)
        return None, None, None, None, []

    layer_index = None
    for lay in f3dm.Layers:
        if lay and (lay.FullPath == layer_name or lay.Name == layer_name):
            layer_index = lay.Index
            break

    if layer_index is None:
        print("Layer '{}' non trovato in libreria.".format(layer_name))
        return None, None, None, None, []

    gancio_curve = None
    pos_pts      = []
    dir_pt       = None
    bool_pts     = []   # punti "booleana": pick points nel template

    for obj in f3dm.Objects:
        if not obj:
            continue
        att = obj.Attributes
        geo = obj.Geometry
        if not att or not geo:
            continue
        if att.LayerIndex != layer_index:
            continue

        if isinstance(geo, Rhino.Geometry.Curve):
            if gancio_curve is None:
                gancio_curve = geo.DuplicateCurve()

        elif isinstance(geo, Rhino.Geometry.Point):
            nm = ""
            try:
                nm = (att.Name or "").strip().lower()
            except:
                pass
            if nm == "dir":
                dir_pt = Point3d(geo.Location)
            elif nm == "posizione_gancio":
                pos_pts.append(Point3d(geo.Location))
            elif nm == "booleana":
                bool_pts.append(Point3d(geo.Location))

    if gancio_curve is None:
        print("Nessuna curva trovata sul layer Gancio.")
        return None, None, None, None, []

    if len(pos_pts) < 2:
        print("Servono 2 punti \'Posizione_gancio\' sul layer Gancio, trovati:", len(pos_pts))
        return None, None, None, None, []

    # I due punti piu' lontani tra loro come asse
    p1, p2 = pick_two_farthest_points(pos_pts)
    if bool_pts:
        print("  Punti booleana caricati da {}: {}".format(layer_name, len(bool_pts)))
    return gancio_curve, p1, p2, dir_pt, bool_pts


def place_gancio(gancio_crv, pos1, pos2, dir_pt, click_pt, bordo_id, ser_id,
                 bool_pts_template=None, bordo_crv_for_fillet=None):
    """
    Orienta il gancio sul punto clic sul bordo.
    Strategia orientamento:
      - Asse X target = tangente al bordo nel punto clic
      - Asse Y target = normale interna (verso serigrafia), calcolata come:
          vettore da p_on_bordo al closest point sulla serigrafia,
          proiettato sul piano XY e perpendicolarizzato alla tangente
      - Se dir_pt (dopo orient) non punta verso la serigrafia -> specchia Y

    bool_pts_template: lista di Point3d nel sistema template da trasformare
                       con lo stesso xf della curva -> pick points mondo.
    Ritorna: (guid, [pick_pts_mondo]) oppure (guid, []) se no bool_pts
    """
    tol = doc_tol()

    # 1. Piano template da pos1/pos2
    tpl_plane = build_template_plane(gancio_crv, pos1, pos2)
    if tpl_plane is None:
        print("Gancio: impossibile costruire piano template.")
        return None

    # 2. Snap al bordo
    bordo_crv = rs.coercecurve(bordo_id)
    ser_crv   = rs.coercecurve(ser_id)
    if not bordo_crv:
        return None

    ok, tb = bordo_crv.ClosestPoint(click_pt)
    if not ok:
        return None
    p_on_bordo = bordo_crv.PointAt(tb)

    tangent = Vector3d(bordo_crv.TangentAt(tb))
    if tangent.IsTiny():
        print("Gancio: tangente nulla.")
        return None
    tangent.Unitize()

    # 3. Calcola normale interna: direzione da p_on_bordo verso il closest point
    #    sulla serigrafia, poi proiettata perpendicolare alla tangente (sul piano XY)
    normal_in = None
    if ser_crv:
        ok_s, ts = ser_crv.ClosestPoint(p_on_bordo)
        if ok_s:
            v = ser_crv.PointAt(ts) - p_on_bordo
            v.Z = 0.0
            # rimuovi componente parallela alla tangente -> perpendicolare pura
            dot = v.X * tangent.X + v.Y * tangent.Y
            v = v - tangent * dot
            if not v.IsTiny():
                v.Unitize()
                normal_in = v

    if normal_in is None:
        # fallback: perpendicolare sinistra della tangente
        normal_in = Vector3d(-tangent.Y, tangent.X, 0.0)
        normal_in.Unitize()

    # 4. Costruisci piano target con asse X = tangente, asse Y = normale interna
    #    PlaneToPlane mappa tpl_plane.XAxis -> target X, tpl_plane.YAxis -> target Y
    #    Dobbiamo quindi costruire il target plane con X=tangente, Y=normal_in
    target_plane = Plane(Point3d(p_on_bordo), tangent, normal_in)
    xf = Transform.PlaneToPlane(tpl_plane, target_plane)
    placed = gancio_crv.DuplicateCurve()
    placed.Transform(xf)

    # 5. Verifica con dir_pt: se dopo la trasformazione dir non punta verso
    #    la serigrafia, specchia rispetto all'asse X (ribalta Y -> -Y)
    needs_mirror  = False
    mirror_plane  = None
    if dir_pt is not None and ser_crv is not None:
        dir_world = Point3d(dir_pt)
        dir_world.Transform(xf)

        ok_s2, ts2 = ser_crv.ClosestPoint(dir_world)
        ok_b2, tb2 = bordo_crv.ClosestPoint(dir_world)
        if ok_s2 and ok_b2:
            dist_to_ser   = dir_world.DistanceTo(ser_crv.PointAt(ts2))
            dist_to_bordo = dir_world.DistanceTo(bordo_crv.PointAt(tb2))

            needs_mirror = dist_to_bordo < dist_to_ser
            if needs_mirror:
                mirror_plane = Plane(Point3d(p_on_bordo), normal_in)
                placed.Transform(Transform.Mirror(mirror_plane))
                if not PROD_MODE:
                    print("Gancio: specchiato (dir puntava verso esterno).")
            else:
                if not PROD_MODE:
                    print("Gancio: orientamento ok (dir verso interno).")

    # 6. Aggiungi al doc (il raccordo viene fatto DOPO la CurveBoolean,
    #    direttamente sulle regioni create)
    final = placed
    oid = sc.doc.Objects.AddCurve(final)
    if oid and oid != System.Guid.Empty:
        set_layer(oid, "Maschere::DEBUG" if not PROD_MODE else "Maschere")

        # 8. Trasforma i punti booleana con xf + mirror (se applicato alla curva)
        bool_pts_world = []
        if bool_pts_template:
            for bp in bool_pts_template:
                bp_w = Point3d(bp)
                bp_w.Transform(xf)
                if needs_mirror and mirror_plane is not None:
                    bp_w.Transform(Transform.Mirror(mirror_plane))
                bool_pts_world.append(bp_w)
            if not PROD_MODE:
                print("  Punti booleana trasformati (mirror={}): {}".format(
                    needs_mirror, bool_pts_world))

        return oid, bool_pts_world
    return None, []


def _make_fillet_arc(crv1, t1, crv2, t2, radius):
    """
    Costruisce manualmente un arco di raccordo tangente a crv1 in t1
    e a crv2 in t2, con raggio dato.

    Algoritmo:
      - Tangente a crv1 in t1 -> normale perpendicolare
      - Tangente a crv2 in t2 -> normale perpendicolare
      - Centro arco = intersezione delle due rette normali spostate di 'radius'
        nella direzione giusta
      - Arco da punto di tangenza su crv1 a punto di tangenza su crv2

    Ritorna (arc_curve, t1_trim, t2_trim) oppure None se fallisce.
    t1_trim e t2_trim sono i parametri di taglio sulle curve originali.
    """
    tol = doc_tol()

    pt1 = crv1.PointAt(t1)
    pt2 = crv2.PointAt(t2)

    tan1 = crv1.TangentAt(t1)
    tan2 = crv2.TangentAt(t2)
    tan1.Unitize()
    tan2.Unitize()

    # Normali perpendicolari (ruota 90 gradi in XY)
    # Due candidati per lato: +perp e -perp
    def perps(tan):
        return (Vector3d(-tan.Y, tan.X, 0), Vector3d(tan.Y, -tan.X, 0))

    perps1 = perps(tan1)
    perps2 = perps(tan2)

    best_center = None
    best_dist   = float('inf')

    # Prova tutte le combinazioni di lato (+/-)
    for n1 in perps1:
        c1 = pt1 + n1 * radius
        for n2 in perps2:
            c2 = pt2 + n2 * radius
            # Il centro deve essere equidistante (radius) da entrambi i punti
            d1 = c1.DistanceTo(pt1)
            d2 = c2.DistanceTo(pt2)
            # Usa la media come centro candidato se i due centri sono vicini
            if c1.DistanceTo(c2) < radius * 0.5:
                cand = Point3d(
                    (c1.X + c2.X) * 0.5,
                    (c1.Y + c2.Y) * 0.5,
                    0
                )
                err = abs(cand.DistanceTo(pt1) - radius) + abs(cand.DistanceTo(pt2) - radius)
                if err < best_dist:
                    best_dist   = err
                    best_center = cand

    if best_center is None or best_dist > radius * 0.3:
        # Fallback: intersezione linee normali con algebra
        # retta1: pt1 + s*n1, retta2: pt2 + s*n2
        # Risolvi per n1 e n2 che minimizzano distanza
        for n1 in perps1:
            for n2 in perps2:
                # Sistema: pt1 + s*n1 = pt2 + t*n2
                # Risolto per s con least squares (2D)
                dx = pt2.X - pt1.X
                dy = pt2.Y - pt1.Y
                denom = n1.X * n2.Y - n1.Y * n2.X
                if abs(denom) < 1e-10:
                    continue
                s = (dx * n2.Y - dy * n2.X) / denom
                cand = Point3d(pt1.X + s * n1.X, pt1.Y + s * n1.Y, 0)
                err = abs(cand.DistanceTo(pt1) - radius) + abs(cand.DistanceTo(pt2) - radius)
                if err < best_dist:
                    best_dist   = err
                    best_center = cand

    if best_center is None or best_dist > radius * 0.5:
        return None

    # Punti di tangenza effettivi sul cerchio di raccordo
    # = proiezione di pt1/pt2 sul cerchio centrato in best_center
    v1 = pt1 - best_center
    v1.Unitize()
    tang_pt1 = best_center + v1 * radius

    v2 = pt2 - best_center
    v2.Unitize()
    tang_pt2 = best_center + v2 * radius

    # Aggiorna i parametri di trim sulle curve originali
    ok1, t1_trim = crv1.ClosestPoint(tang_pt1)
    ok2, t2_trim = crv2.ClosestPoint(tang_pt2)
    if not ok1 or not ok2:
        return None

    # Costruisci arco da tang_pt1 a tang_pt2 con centro best_center
    try:
        arc = Rhino.Geometry.Arc(tang_pt1, best_center, tang_pt2)
        if not arc.IsValid:
            # Prova con Arc(plane, radius, angle)
            return None
        arc_crv = Rhino.Geometry.ArcCurve(arc)
        return arc_crv, t1_trim, t2_trim
    except:
        return None


def _point_at_distance_along_bordo(bordo_crv, p_start, dist_mm):
    """
    Dato un punto p_start sulla curva bordo_crv, ritorna il punto
    a dist_mm di distanza curvilinea in avanti (verso parametri crescenti).
    Gestisce il wrap-around sul bordo chiuso.
    """
    dom   = bordo_crv.Domain
    total = bordo_crv.GetLength()

    # Parametro del punto di partenza
    ok, t0 = bordo_crv.ClosestPoint(p_start)
    if not ok:
        return None

    # Lunghezza cumulativa dal t0 del dominio fino a t0
    L0 = bordo_crv.GetLength(Rhino.Geometry.Interval(dom.T0, t0))

    # Lunghezza target con wrap-around
    L_target = L0 + dist_mm
    if L_target > total:
        L_target -= total

    ok2, t_target = bordo_crv.LengthParameter(L_target)
    if not ok2:
        return None
    return bordo_crv.PointAt(t_target)


def auto_place_ganci_from_division_pts(
        bordo_id, ser_id, division_pts_on_bordo,
        gancio_crv, pos1, pos2, dir_pt,
        gancio_crv_small=None, pos1_small=None, pos2_small=None, dir_pt_small=None,
        bool_pts_small=None):
    """
    Posiziona automaticamente un gancio a GANCIO_OFFSET_MM di distanza
    (in avanti lungo il bordo) da ogni punto di divisione (maschiatura).

    Sceglie il profilo in base al GAP nel punto di inserimento:
      - gap >= GANCIO_GAP_THRESHOLD  -> profilo standard (gancio_crv), sporge verso interno
      - gap <  GANCIO_GAP_THRESHOLD  -> profilo ridotto  (gancio_crv_small), sporge verso esterno

    Ritorna: lista GUID ganci creati.
    Per i ganci <45mm splitta il bordo nel punto di intersezione
    eliminando il tratto da ~180mm, cosi\' la CurveBoolean li gestisce
    come qualsiasi altro gancio standard.
    """
    if not gancio_crv:
        print("Auto ganci: profilo gancio non disponibile.")
        return [], []

    bordo_crv = rs.coercecurve(bordo_id)
    ser_crv   = rs.coercecurve(ser_id)
    if not bordo_crv:
        return [], []

    gancio_ids_std      = []  # ganci standard -> partecipano alla CurveBoolean
    gancio_ids_small    = []  # ganci <45mm -> solo output, NON alla CurveBoolean
    bordo_crv           = rs.coercecurve(bordo_id)

    for p_div in division_pts_on_bordo:
        _, _, gap = gap_from_border_point_to_serigrafia(p_div, ser_id)

        use_small = (gap is not None and gap < GANCIO_GAP_THRESHOLD
                     and gancio_crv_small is not None)

        offset_mm = GANCIO_OFFSET_SMALL_MM if use_small else GANCIO_OFFSET_MM
        p_gancio  = _point_at_distance_along_bordo(bordo_crv, p_div, offset_mm)
        if p_gancio is None:
            print("  WARN: impossibile calcolare punto gancio per divisore", p_div)
            continue

        if use_small:
            crv_use = gancio_crv_small
            p1_use  = pos1_small
            p2_use  = pos2_small
            dir_use = dir_pt_small
            if not PROD_MODE:
                print("  Gancio <45mm (gap={:.1f}mm, offset={}mm)".format(gap, int(offset_mm)))

            # Piazza il gancio normalmente
            gid, _ = place_gancio(
                crv_use, p1_use, p2_use, dir_use,
                p_gancio, bordo_id, ser_id
            )
            if gid:
                gancio_ids_small.append(gid)
            # Salta il place_gancio generico sotto
            continue
        else:
            crv_use = gancio_crv
            p1_use  = pos1
            p2_use  = pos2
            dir_use = dir_pt
            if not PROD_MODE and gap is not None:
                print("  Gancio standard (gap={:.1f}mm, offset={}mm)".format(gap, int(offset_mm)))

        # Solo per ganci standard (i <45mm usano 'continue' sopra)
        gid, _ = place_gancio(crv_use, p1_use, p2_use, dir_use, p_gancio, bordo_id, ser_id)
        if gid:
            gancio_ids_std.append(gid)

    n_tot = len(gancio_ids_std) + len(gancio_ids_small)
    print("Ganci automatici posizionati: {} / {} (std={}, small={})".format(
        n_tot, len(division_pts_on_bordo),
        len(gancio_ids_std), len(gancio_ids_small)))
    return gancio_ids_std, gancio_ids_small


# ----------------------------
# ETO DIALOG
# ----------------------------

def show_maschera_dialog():
    """
    Mostra dialog ETO per raccogliere i metadati della maschera.
    Ritorna dict con chiavi:
      cod_attr, cod_prod, n_dis, vista_interna (bool)
    oppure None se l'utente annulla.
    """
    import Rhino.UI
    import Eto.Forms as forms
    import Eto.Drawing as drawing

    class MascheraDialog(forms.Dialog):
        def __init__(self):
            self.Title   = "Dati Maschera Serigrafia"
            self.Padding = drawing.Padding(12)
            self.Resizable = False

            lbl_attr  = forms.Label(Text="Codice attrezzatura:")
            lbl_prod  = forms.Label(Text="Codice prodotto:")
            lbl_ndis  = forms.Label(Text="N° disegno:")
            lbl_vista = forms.Label(Text="Vista:")

            self.txt_attr  = forms.TextBox()
            self.txt_prod  = forms.TextBox()
            self.txt_ndis  = forms.TextBox()

            self.rb_interna  = forms.RadioButton(Text="Vista interna")
            self.rb_esterna  = forms.RadioButton(Text="Vista esterna")
            self.rb_interna.Checked = True

            # Radio group
            rb_group = forms.StackLayout()
            rb_group.Orientation = forms.Orientation.Horizontal
            rb_group.Spacing = 12
            rb_group.Items.Add(forms.StackLayoutItem(self.rb_interna))
            rb_group.Items.Add(forms.StackLayoutItem(self.rb_esterna))

            btn_ok     = forms.Button(Text="OK")
            btn_cancel = forms.Button(Text="Annulla")
            btn_ok.Click     += self._on_ok
            btn_cancel.Click += self._on_cancel
            self.DefaultButton = btn_ok
            self.AbortButton   = btn_cancel

            btn_row = forms.StackLayout()
            btn_row.Orientation = forms.Orientation.Horizontal
            btn_row.Spacing = 8
            btn_row.Items.Add(forms.StackLayoutItem(None, True))  # spacer
            btn_row.Items.Add(forms.StackLayoutItem(btn_ok))
            btn_row.Items.Add(forms.StackLayoutItem(btn_cancel))

            layout = forms.TableLayout()
            layout.Spacing = drawing.Size(8, 6)
            layout.Padding = drawing.Padding(0)

            def row(lbl, ctrl):
                r = forms.TableRow()
                r.Cells.Add(forms.TableCell(lbl,  False))
                r.Cells.Add(forms.TableCell(ctrl, True))
                return r

            layout.Rows.Add(row(lbl_attr,  self.txt_attr))
            layout.Rows.Add(row(lbl_prod,  self.txt_prod))
            layout.Rows.Add(row(lbl_ndis,  self.txt_ndis))
            layout.Rows.Add(row(lbl_vista, rb_group))

            outer = forms.StackLayout()
            outer.Orientation = forms.Orientation.Vertical
            outer.Spacing = 10
            outer.Items.Add(forms.StackLayoutItem(layout, True))
            outer.Items.Add(forms.StackLayoutItem(btn_row))

            self.Content = outer
            self._result = None

        def _on_ok(self, sender, e):
            self._result = {
                "cod_attr":     self.txt_attr.Text.strip(),
                "cod_prod":     self.txt_prod.Text.strip(),
                "n_dis":        self.txt_ndis.Text.strip(),
                "vista_interna": self.rb_interna.Checked,
            }
            self.Close()

        def _on_cancel(self, sender, e):
            self._result = None
            self.Close()

    dlg = MascheraDialog()
    Rhino.UI.EtoExtensions.ShowSemiModal(dlg, Rhino.RhinoDoc.ActiveDoc, Rhino.UI.RhinoEtoApp.MainWindow)
    return dlg._result


# ----------------------------
# TESTI ETICHETTE
# ----------------------------

def _find_point_inside(closed_crv):
    """
    Trova un punto garantito all'interno di una curva chiusa,
    anche se concava o molto allungata.

    Strategia:
    1. Prova il centroide (AreaMassProperties)
    2. Se il centroide cade fuori, campiona punti sul boundary
       e per ognuno lancia un raggio verso il centro dell'AABB:
       il midpoint raggio/boundary e' quasi sempre dentro
    3. Ultimo fallback: offset interno della curva, primo punto valido
    """
    tol = doc_tol()
    plane = Plane.WorldXY

    if not closed_crv or not closed_crv.IsClosed:
        return None

    def is_inside(pt):
        try:
            cont = closed_crv.Contains(pt, plane, tol)
            return cont == Rhino.Geometry.PointContainment.Inside
        except:
            return False

    # 1. Centroide
    try:
        amp = AreaMassProperties.Compute(closed_crv)
        if amp:
            c = amp.Centroid
            if is_inside(c):
                return c
    except:
        pass

    # 2. Campiona punti sul bordo e usa il midpoint verso il centroide AABB
    try:
        bb = closed_crv.GetBoundingBox(True)
        aabb_center = (bb.Min + bb.Max) * 0.5

        dom = closed_crv.Domain
        n_samples = 32
        for i in range(n_samples):
            t = dom.T0 + (dom.T1 - dom.T0) * i / float(n_samples)
            p_border = closed_crv.PointAt(t)
            # normale interna approssimata: direzione verso centro AABB
            v = aabb_center - p_border
            if v.IsTiny():
                continue
            # Prova a vari step (10%, 30%, 50% della distanza)
            dist = v.Length
            v.Unitize()
            for frac in (0.1, 0.3, 0.5, 0.7):
                candidate = p_border + v * (dist * frac)
                if is_inside(candidate):
                    return candidate
    except:
        pass

    # 3. Offset interno piccolo: prendi il primo punto della curva offset
    try:
        offs = closed_crv.Offset(plane, -1.0, tol,
                                  Rhino.Geometry.CurveOffsetCornerStyle.Sharp)
        if offs:
            for oc in offs:
                if oc and oc.IsClosed:
                    p = oc.PointAt(oc.Domain.Mid)
                    if is_inside(p):
                        return p
                    # anche il centroide dell'offset
                    amp2 = AreaMassProperties.Compute(oc)
                    if amp2 and is_inside(amp2.Centroid):
                        return amp2.Centroid
    except:
        pass

    # Fallback finale: restituisce il centroide anche se fuori
    # (almeno il testo compare vicino alla regione)
    try:
        amp3 = AreaMassProperties.Compute(closed_crv)
        if amp3:
            return amp3.Centroid
    except:
        pass

    return closed_crv.PointAt(closed_crv.Domain.Mid)


def add_labels_to_regions(reg_ids, meta):
    """
    Per ogni regione aggiunge 5 righe di testo.
    Ritorna dict {reg_id (str): [text_guid, ...]}
    con associazione esplicita regione -> testi propri.
    """
    import System.Drawing

    result = {}  # {str(reg_id): [guid, ...]}

    if not reg_ids or not meta:
        return result

    n_tot     = len(reg_ids)
    vista_str = "Interna" if meta.get("vista_interna", True) else "Esterna"
    yellow    = System.Drawing.Color.Yellow

    for idx, rid in enumerate(reg_ids):
        crv = rs.coercecurve(rid)
        if not crv:
            continue

        # Punto garantito dentro la regione
        origin = _find_point_inside(crv)
        if origin is None:
            continue

        lines = [
            meta.get("cod_attr", ""),
            meta.get("cod_prod", ""),
            meta.get("n_dis", ""),
            "Vista: {}".format(vista_str),
            "{} / {}".format(idx + 1, n_tot),
        ]

        text_ids_for_region = []
        for i, line in enumerate(lines):
            y_offset = -i * TEXT_LINE_SPACING
            pos = Point3d(origin.X, origin.Y + y_offset, origin.Z)
            tid = _add_single_text(line, pos, yellow)
            if tid:
                text_ids_for_region.append(tid)

        result[str(rid)] = text_ids_for_region

    return result


def _add_single_text(text_str, position, color):
    """Aggiunge una singola TextEntity al documento. Ritorna il GUID o None."""
    import System.Drawing

    try:
        te = Rhino.Geometry.TextEntity()
        te.Plane = Plane(position, Vector3d(0, 0, 1))
        te.PlainText = text_str
        te.TextHeight = TEXT_HEIGHT
        te.Justification = Rhino.Geometry.TextJustification.MiddleCenter

        # Font Arial
        try:
            fnt = Rhino.DocObjects.Font(TEXT_FONT)
            te.Font = fnt
        except:
            pass

        oid = sc.doc.Objects.AddText(te)
        if oid and oid != System.Guid.Empty:
            # Layer
            try:
                rs.ObjectLayer(oid, TEXT_LAYER)
            except:
                pass
            # Colore oggetto
            try:
                obj = sc.doc.Objects.Find(oid)
                if obj:
                    attr = obj.Attributes.Duplicate()
                    attr.ColorSource = Rhino.DocObjects.ObjectColorSource.ColorFromObject
                    attr.ObjectColor = color
                    sc.doc.Objects.ModifyAttributes(obj, attr, True)
            except:
                pass
            return oid  # <-- ritorna il GUID
    except Exception as e:
        if not PROD_MODE:
            print("  _add_single_text fallito:", e)
    return None


# ----------------------------
# EXPORT DXF
# ----------------------------

def export_regions_to_dxf(reg_ids, meta, labels_map=None):
    """
    Esporta ogni regione + i suoi testi in un file DXF separato.
    Struttura:  DXF_ROOT_PATH\<cod_attr>\<cod_attr>-N.dxf

    labels_map: dict {str(reg_id): [text_guid, ...]} da add_labels_to_regions.
    L'associazione regione->testi e' esplicita, non dipende dall'ordine nel layer.
    """
    import os
    import datetime
    import re

    if not reg_ids:
        print("Export DXF: nessuna regione da esportare.")
        return

    cod = (meta.get("cod_attr", "") if meta else "").strip()
    if not cod:
        cod = "Maschera_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    cod_safe   = re.sub(r'[\\/:*?"<>|]', '_', cod)
    export_dir = os.path.join(DXF_ROOT_PATH, cod_safe)

    if not os.path.exists(export_dir):
        try:
            os.makedirs(export_dir)
            print("Export DXF: cartella creata:", export_dir)
        except Exception as e:
            print("Export DXF: impossibile creare cartella:", e)
            return

    lmap = labels_map or {}

    ok_count = 0
    for idx, rid in enumerate(reg_ids):
        piece_num = idx + 1
        filename  = "{}-{}.dxf".format(cod_safe, piece_num)
        filepath  = os.path.join(export_dir, filename)

        # Testi associati esplicitamente a questa regione
        piece_text_ids = lmap.get(str(rid), [])

        # Seleziona curva regione + testi
        try:
            rs.UnselectAllObjects()
            rs.SelectObject(rid)
            for tid in piece_text_ids:
                rs.SelectObject(tid)
        except:
            continue

        # Costruisci comando export
        # Nota: il path va tra virgolette se contiene spazi
        fp_escaped = filepath.replace("\\", "\\")
        cmd = '! _-Export "{}" Scheme="{}" _Enter _Enter'.format(
            fp_escaped, DXF_SCHEME
        )

        ok = Rhino.RhinoApp.RunScript(cmd, False)

        rs.UnselectAllObjects()

        if ok and os.path.exists(filepath):
            ok_count += 1
            print("  Esportato: {}".format(filename))
        else:
            # Fallback: prova senza scheme (export base)
            cmd2 = '! _-Export "{}" Scheme="{}" _Enter _Enter'.format(fp_escaped, DXF_SCHEME)
            ok2 = Rhino.RhinoApp.RunScript(cmd2, False)
            rs.UnselectAllObjects()
            if ok2 and os.path.exists(filepath):
                ok_count += 1
                print("  Esportato (fallback): {}".format(filename))
            else:
                print("  WARN: export fallito per pezzo {}".format(piece_num))

    print("Export DXF completato: {}/{} file in {}".format(ok_count, len(reg_ids), export_dir))
    return export_dir



# ----------------------------
# AUTO-MASCHIATURA TRATTI LUNGHI
# ----------------------------

def _arc_length_on_bordo(bordo_crv, t0, t1):
    """
    Lunghezza dell'arco sul bordo tra i parametri t0 e t1
    (t1 puo' essere < t0 se il tratto attraversa il punto di chiusura).
    """
    dom = bordo_crv.Domain
    total = bordo_crv.GetLength()
    L0 = bordo_crv.GetLength(Rhino.Geometry.Interval(dom.T0, t0))
    L1 = bordo_crv.GetLength(Rhino.Geometry.Interval(dom.T0, t1))
    d = L1 - L0
    if d < 0:
        d += total
    return d


def _midpoint_param_on_bordo(bordo_crv, t0, t1):
    """
    Parametro al punto medio dell'arco tra t0 e t1 sul bordo.
    Gestisce il wrap-around (tratto che attraversa la chiusura).
    """
    dom   = bordo_crv.Domain
    total = bordo_crv.GetLength()
    L0    = bordo_crv.GetLength(Rhino.Geometry.Interval(dom.T0, t0))
    L1    = bordo_crv.GetLength(Rhino.Geometry.Interval(dom.T0, t1))
    seg_len = L1 - L0
    if seg_len < 0:
        seg_len += total
    mid_len = L0 + seg_len / 2.0
    if mid_len > total:
        mid_len -= total
    ok, t_mid = bordo_crv.LengthParameter(mid_len)
    if ok:
        return t_mid
    return (t0 + t1) / 2.0


def auto_add_maschiature_for_long_segments(
        bordo_id, ser_id, division_pts_on_bordo, all_offsets):
    """
    Controlla ogni tratto del bordo tra i punti di divisione esistenti.
    Se un tratto supera MAX_SEGMENT_LENGTH mm, inserisce automaticamente
    una maschiatura nel punto medio del tratto.

    Ripete il controllo finche' non ci sono piu' tratti lunghi
    (gestisce il caso in cui anche meta-tratto sia > 1800mm).

    Aggiorna in-place division_pts_on_bordo e all_offsets.
    Ritorna il numero di maschiature aggiunte.
    """
    bordo_crv = rs.coercecurve(bordo_id)
    if not bordo_crv:
        return 0

    dom   = bordo_crv.Domain
    total = bordo_crv.GetLength()
    added = 0
    max_iters = 50  # safety cap

    for _iteration in range(max_iters):
        # Calcola i parametri t per i punti di divisione correnti,
        # ordinati lungo il bordo
        t_list = []
        for p in division_pts_on_bordo:
            ok, t = bordo_crv.ClosestPoint(p)
            if ok:
                t_list.append(t)

        if len(t_list) < 2:
            break

        # Ordina per lunghezza cumulativa dal t0 del dominio
        def cum_len(t):
            return bordo_crv.GetLength(Rhino.Geometry.Interval(dom.T0, t))

        t_list_sorted = sorted(set(t_list), key=cum_len)
        n = len(t_list_sorted)

        # Controlla tutti i tratti (incluso il wrap-around dall'ultimo al primo)
        found_long = False
        for i in range(n):
            t_a = t_list_sorted[i]
            t_b = t_list_sorted[(i + 1) % n]
            seg_len = _arc_length_on_bordo(bordo_crv, t_a, t_b)

            if seg_len > MAX_SEGMENT_LENGTH:
                # Punto medio del tratto
                t_mid  = _midpoint_param_on_bordo(bordo_crv, t_a, t_b)
                p_mid  = bordo_crv.PointAt(t_mid)

                print("  Tratto lungo {:.0f}mm tra divisori {}/{}: "
                      "aggiunta maschiatura automatica.".format(seg_len, i+1, (i+1)%n+1))

                # Crea la maschiatura nel punto medio
                ids, _ = build_one_maschiatura(bordo_id, ser_id, p_mid)
                if ids:
                    all_offsets.extend(ids)

                # Aggiungi il punto al set dei divisori
                division_pts_on_bordo.append(p_mid)
                added  += 1
                found_long = True
                # Ricomincia il controllo con i nuovi divisori
                break

        if not found_long:
            break

    if added:
        print("Maschiature automatiche aggiunte:", added)
    else:
        print("Nessun tratto supera {}mm.".format(int(MAX_SEGMENT_LENGTH)))

    return added


def _integrate_ganci_small_into_regions(reg_ids, gancio_ids_small):
    """
    Per ogni gancio <45mm (che sporge verso esterno del bordo):
    - Trova la regione piu' vicina
    - Splitta quella regione con la curva del gancio (2 intersezioni)
    - Elimina il pezzo piu' corto (tratto bordo ~180mm)
    - Join pezzo lungo + gancio -> curva chiusa
    - Sostituisce la regione nel documento
    - Elimina la curva gancio separata
    """
    tol = doc_tol()
    n_ok = 0

    for gid in gancio_ids_small:
        g_crv = rs.coercecurve(gid)
        if not g_crv:
            continue

        g_mid = g_crv.PointAt(g_crv.Domain.Mid)

        # Trova regione piu' vicina
        best_rid  = None
        best_dist = float('inf')
        for rid in reg_ids:
            r_crv = rs.coercecurve(rid)
            if not r_crv:
                continue
            ok, t = r_crv.ClosestPoint(g_mid)
            if ok:
                dist = g_mid.DistanceTo(r_crv.PointAt(t))
                if dist < best_dist:
                    best_dist = dist
                    best_rid  = rid

        if best_rid is None:
            if not PROD_MODE:
                print("  Gancio small: nessuna regione trovata")
            continue

        r_crv = rs.coercecurve(best_rid)
        if not r_crv:
            continue

        if not PROD_MODE:
            print("  Integro gancio small in regione (dist bordo={:.1f}mm)".format(best_dist))

        try:
            # Intersezioni gancio / regione
            events = Rhino.Geometry.Intersect.Intersection.CurveCurve(
                g_crv, r_crv, tol, tol * 100
            )
            if events is None or events.Count < 2:
                if not PROD_MODE:
                    print("  WARN: {} intersezioni gancio/regione".format(
                        events.Count if events else 0))
                continue

            # Raccogli parametri sia sul gancio (A) che sulla regione (B)
            ix_pairs = []
            for i in range(events.Count):
                ev = events[i]
                if ev.IsPoint:
                    ix_pairs.append((ev.ParameterA, ev.ParameterB))

            if len(ix_pairs) < 2:
                continue

            # Ordina per parametro sul gancio
            ix_pairs.sort(key=lambda x: x[0])
            tg0, tr0 = ix_pairs[0]
            tg1, tr1 = ix_pairs[-1]

            t_reg = sorted([tr0, tr1])
            t_gan = sorted([tg0, tg1])

            # Splitta la regione nei 2 punti di intersezione
            pieces = r_crv.Split(t_reg)
            if not pieces or len(pieces) < 2:
                if not PROD_MODE:
                    print("  WARN: split regione fallito")
                continue

            # Tieni il pezzo piu' lungo (scarta il tratto bordo ~180mm)
            pieces_by_len = sorted(pieces, key=lambda c: c.GetLength())
            piece_long = pieces_by_len[-1]

            if not PROD_MODE:
                print("  Split regione: tengo {:.1f}mm, scarto {:.1f}mm".format(
                    piece_long.GetLength(), pieces_by_len[0].GetLength()))

            # Trimma il gancio tra i due punti di intersezione con la regione
            # -> prende solo la parte che "attraversa" il bordo
            g_trimmed = g_crv.Trim(t_gan[0], t_gan[1])
            if g_trimmed is None:
                if not PROD_MODE:
                    print("  WARN: trim gancio fallito")
                continue

            if not PROD_MODE:
                print("  Gancio trimmato: {:.1f}mm".format(g_trimmed.GetLength()))

            # Verifica che i punti finali si tocchino
            gap_s = piece_long.PointAtStart.DistanceTo(g_trimmed.PointAtStart)
            gap_e = piece_long.PointAtEnd.DistanceTo(g_trimmed.PointAtEnd)
            gap_se = piece_long.PointAtStart.DistanceTo(g_trimmed.PointAtEnd)
            gap_es = piece_long.PointAtEnd.DistanceTo(g_trimmed.PointAtStart)
            if not PROD_MODE:
                print("  Gap connessione: ss={:.3f} ee={:.3f} se={:.3f} es={:.3f}".format(
                    gap_s, gap_e, gap_se, gap_es))

            # Join pezzo lungo + gancio trimmato
            joined = Rhino.Geometry.Curve.JoinCurves(
                [piece_long, g_trimmed], tol * 10
            )
            if not joined or len(joined) == 0:
                if not PROD_MODE:
                    print("  WARN: join fallito")
                continue

            new_reg = joined[0]

            if not new_reg.IsClosed:
                gap = new_reg.PointAtStart.DistanceTo(new_reg.PointAtEnd)
                if gap < tol * 100:
                    new_reg.MakeClosed(tol * 100)
                else:
                    if not PROD_MODE:
                        print("  WARN: curva non chiusa, gap={:.3f}mm".format(gap))
                    continue

            sc.doc.Objects.Replace(best_rid, new_reg)
            try:
                rs.DeleteObject(gid)
            except:
                pass
            n_ok += 1
            if not PROD_MODE:
                print("  Gancio small integrato OK")

        except Exception as ex:
            if not PROD_MODE:
                print("  Eccezione:", ex)
            import traceback
            traceback.print_exc()

    return n_ok


def ask_ganci_mode():
    """
    Chiede all'utente se vuole posizionare i ganci in modalita'
    automatica (uno per ogni maschiatura) o manuale (click sul bordo).
    Ritorna 'auto' o 'manual'.
    """
    result = rs.MessageBox(
        "Modalita' posizionamento ganci:\n\n"
        "  [Si]  - Automatico (un gancio per ogni maschiatura)\n"
        "  [No]  - Manuale (clicco io dove mettere i ganci)\n",
        4 | 32,   # MB_YESNO | MB_ICONQUESTION
        "Ganci - Modalita'"
    )
    # result: 6=Yes, 7=No
    return 'auto' if result == 6 else 'manual'


def place_ganci_manual(bordo_id, ser_id,
                       gancio_crv, pos1, pos2, dir_pt,
                       gancio_crv_small, pos1_small, pos2_small, dir_pt_small,
                       bool_pts_small):
    """
    Modalita' manuale: l'utente clicca i punti sul bordo dove vuole i ganci.
    Per ogni punto clicca, sceglie automaticamente il profilo in base al GAP.
    Mostra anteprima e chiede conferma prima di procedere.

    Ritorna (gancio_ids_std, gancio_ids_small) come la modalita' automatica.
    """
    gancio_ids_std   = []
    gancio_ids_small = []
    preview_ids      = []  # GUID temporanei per l'anteprima

    while True:
        # Elimina anteprima precedente
        for pid in preview_ids:
            try:
                rs.DeleteObject(pid)
            except:
                pass
        preview_ids = []
        gancio_ids_std   = []
        gancio_ids_small = []

        # Assicura che la viewport sia aggiornata con tutte le maschiature
        # visibili prima di iniziare il click ganci
        rs.EnableRedraw(True)
        sc.doc.Views.Redraw()

        # Chiedi i punti sul bordo
        print("Clicca i punti sul BORDO dove posizionare i ganci (Invio per finire)")
        pts_ganci = []
        while True:
            gp = rs.GetPointOnCurve(
                bordo_id,
                "Clicca punto GANCIO sul bordo (Invio per finire)"
            )
            if gp is None:
                break
            pts_ganci.append(Point3d(gp.X, gp.Y, gp.Z))

        if not pts_ganci:
            print("Nessun punto gancio selezionato.")
            return [], []

        # Posiziona i ganci in anteprima
        rs.EnableRedraw(False)
        try:
            bordo_crv = rs.coercecurve(bordo_id)
            for p_click in pts_ganci:
                # Snappa al bordo
                ok, tb = bordo_crv.ClosestPoint(p_click)
                if not ok:
                    continue
                p_gancio = bordo_crv.PointAt(tb)

                # Calcola GAP per scegliere il profilo
                _, _, gap = gap_from_border_point_to_serigrafia(p_gancio, ser_id)

                use_small = (gap is not None
                             and gap < GANCIO_GAP_THRESHOLD
                             and gancio_crv_small is not None)

                if use_small:
                    crv_use = gancio_crv_small
                    p1_use  = pos1_small
                    p2_use  = pos2_small
                    dir_use = dir_pt_small
                    if not PROD_MODE:
                        print("  Gancio <45mm (gap={:.1f}mm)".format(gap))
                else:
                    crv_use = gancio_crv
                    p1_use  = pos1
                    p2_use  = pos2
                    dir_use = dir_pt
                    if not PROD_MODE and gap is not None:
                        print("  Gancio standard (gap={:.1f}mm)".format(gap))

                gid, _ = place_gancio(
                    crv_use, p1_use, p2_use, dir_use,
                    p_gancio, bordo_id, ser_id
                )
                if gid:
                    preview_ids.append(gid)
                    if use_small:
                        gancio_ids_small.append(gid)
                    else:
                        gancio_ids_std.append(gid)

        finally:
            rs.EnableRedraw(True)
            sc.doc.Views.Redraw()

        print("Ganci posizionati in anteprima: {} (std={}, small={})".format(
            len(preview_ids), len(gancio_ids_std), len(gancio_ids_small)))

        # Chiedi conferma
        conf = rs.MessageBox(
            "Ganci posizionati: {}\n"
            "  Standard:  {}\n"
            "  <45mm:     {}\n\n"
            "[Si] Conferma e procedi\n"
            "[No] Riposiziona i ganci".format(
                len(preview_ids), len(gancio_ids_std), len(gancio_ids_small)
            ),
            4 | 32,
            "Anteprima ganci"
        )

        if conf == 6:  # Yes -> conferma
            # I ganci rimangono nel documento, usciamo
            return gancio_ids_std, gancio_ids_small
        # No -> loop: i ganci vengono eliminati e si riclicca


def _show_all_maschere_layers():
    """
    All'avvio dello script: rende visibili tutti i layer Maschere::*
    che esistono gia' nel documento, in modo che le geometrie del run
    precedente siano visibili durante la nuova sessione di lavoro.
    Vengono poi nascosti normalmente a fine script da _setup_layers_at_end().
    """
    prefix = "Maschere"
    for ln in rs.LayerNames():
        if ln == prefix or ln.startswith(prefix + "::"):
            try:
                rs.LayerVisible(ln, True)
            except:
                pass
    sc.doc.Views.Redraw()


def main():
    ensure_layers()
    _show_all_maschere_layers()

    # --- Dialog metadati maschera ---
    meta = show_maschera_dialog()
    if meta is None:
        print("Operazione annullata dall'utente.")
        return

    bordo_id = get_closed_profile("Seleziona BORDO VETRO (curva chiusa o segmenti unibili)")
    if not bordo_id:
        print("Bordo non valido.")
        return

    ser_id = get_closed_profile("Seleziona SERIGRAFIA (curva chiusa o segmenti unibili)")
    if not ser_id:
        print("Serigrafia non valida.")
        return

    # Applica relief G0 sulla serigrafia (cerchi R10 sugli angoli)
    ser_relief_id, n_relief, g0_pts = apply_g0_relief(ser_id)
    if n_relief > 0:
        # Usa la serigrafia modificata per tutto il resto
        # Metti markers DEBUG sui punti G0 se non in PROD_MODE
        if not PROD_MODE:
            for gp in g0_pts:
                dbg = sc.doc.Objects.AddPoint(gp)
                if dbg:
                    set_layer(dbg, "Maschere::DEBUG")
        ser_id = ser_relief_id
    else:
        print("G0 relief: nessun angolo G0 trovato sulla serigrafia.")

    pts = get_point_on_curve_repeated(bordo_id)
    if not pts:
        print("Nessun punto selezionato.")
        return

    all_offsets = []
    # I punti clic sul bordo (posizione delle maschiature) servono per
    # dividere il bordo in settori e generare i pick points sulle regioni
    division_pts_on_bordo = []

    rs.EnableRedraw(False)
    try:
        for p in pts:
            ids, _ = build_one_maschiatura(bordo_id, ser_id, p)
            if ids:
                all_offsets.extend(ids)
            # Teniamo il punto snappato al bordo come divisore di settore
            bordo_crv = rs.coercecurve(bordo_id)
            if bordo_crv:
                ok, tb = bordo_crv.ClosestPoint(p)
                if ok:
                    division_pts_on_bordo.append(bordo_crv.PointAt(tb))
    finally:
        rs.EnableRedraw(True)
        sc.doc.Views.Redraw()

    print("Fatto. Offset creati:", len(all_offsets))
    print("PROD_MODE:", "ON" if PROD_MODE else "OFF")

    # --- Carica profili gancio dalla libreria (una volta sola) ---
    gancio_crv, pos1, pos2, dir_pt, _bool_std = load_gancio_from_3dm(LIB_3DM_PATH)
    if gancio_crv is None:
        print("WARN: profilo gancio standard non disponibile nel file libreria.")

    gancio_crv_small, pos1_small, pos2_small, dir_pt_small, bool_pts_small = load_gancio_from_3dm(
        LIB_3DM_PATH, GANCIO_LAYER_SMALL
    )
    if gancio_crv_small is None:
        print("WARN: profilo gancio <45mm non disponibile, usero' sempre il profilo standard.")

    # --- Controllo tratti lunghi: auto-maschiatura se > MAX_SEGMENT_LENGTH ---
    # (deve avvenire prima del posizionamento ganci, cosi' anche le maschiature
    #  automatiche riceveranno il loro gancio)
    rs.EnableRedraw(False)
    try:
        auto_add_maschiature_for_long_segments(
            bordo_id, ser_id, division_pts_on_bordo, all_offsets
        )
    finally:
        rs.EnableRedraw(True)
        sc.doc.Views.Redraw()
        # Forza un secondo redraw per assicurarsi che tutte le maschiature
        # (anche quelle automatiche) siano visibili prima del dialog ganci
        sc.doc.Views.Redraw()

    # --- Ganci: chiedi modalita' (auto o manuale) ---
    # Il dialog viene mostrato DOPO che tutte le maschiature sono visibili
    # (sia manuali che automatiche), cosi' l'utente vede la situazione
    # completa prima di scegliere dove mettere i ganci.
    gancio_ids_std    = []
    gancio_ids_small  = []

    if gancio_crv:
        ganci_mode = ask_ganci_mode()

        if ganci_mode == 'auto':
            rs.EnableRedraw(False)
            try:
                if division_pts_on_bordo:
                    gancio_ids_std, gancio_ids_small = auto_place_ganci_from_division_pts(
                        bordo_id, ser_id, division_pts_on_bordo,
                        gancio_crv, pos1, pos2, dir_pt,
                        gancio_crv_small, pos1_small, pos2_small, dir_pt_small,
                        bool_pts_small
                    )
            finally:
                rs.EnableRedraw(True)
                sc.doc.Views.Redraw()

        else:  # manual
            gancio_ids_std, gancio_ids_small = place_ganci_manual(
                bordo_id, ser_id,
                gancio_crv, pos1, pos2, dir_pt,
                gancio_crv_small, pos1_small, pos2_small, dir_pt_small,
                bool_pts_small
            )

    # --- CurveBoolean unica: bordo + serigrafia + offset maschiature + ganci ---
    if DO_REGIONS_BOOLEAN and all_offsets:
        # Crea offset interno del bordo a 5mm verso la serigrafia
        # per generare pick points sicuramente dentro la fascia
        inner_offset_id = _pick_inward_offset_curve(bordo_id, ser_id, 5.0, Plane.WorldXY)

        if inner_offset_id:
            if not PROD_MODE:
                set_layer(inner_offset_id, "Maschere::DEBUG")

            pick_points = _make_pick_points_on_inner_offset(
                bordo_id, inner_offset_id, division_pts_on_bordo
            )

            if not PROD_MODE:
                for pp in pick_points:
                    dbg = sc.doc.Objects.AddPoint(pp)
                    if dbg:
                        set_layer(dbg, "Maschere::DEBUG")

            print("Pick points (offset interno):", len(pick_points))

            # Elimina sempre la curva di offset: e' solo uno strumento di calcolo
            try:
                rs.DeleteObject(inner_offset_id)
            except:
                pass
        else:
            print("WARN: offset interno non disponibile, uso pick points fallback.")
            pick_points = []

        if not pick_points:
            print("Regioni: nessun pick point disponibile.")
        else:
            reg_ids = build_regions_with_pickpoints(
                bordo_id, ser_id, all_offsets + gancio_ids_std, pick_points
            )

            # Integra i ganci <45mm nelle regioni: split + join -> curva chiusa
            if gancio_ids_small and reg_ids:
                n_int = _integrate_ganci_small_into_regions(reg_ids, gancio_ids_small)
                print("Ganci small integrati nelle regioni: {}/{}".format(
                    n_int, len(gancio_ids_small)))

            print("Regioni create:", len(reg_ids))

            # --- Testi etichette ---
            labels_map = {}
            if reg_ids and meta:
                labels_map = add_labels_to_regions(reg_ids, meta)
                print("Etichette aggiunte:", len(labels_map))

            # --- Export DXF ---
            sc.doc.Views.Redraw()
            export_dir = export_regions_to_dxf(reg_ids, meta, labels_map)


    if DELETE_INPUT_PROFILES:
        try:
            rs.DeleteObject(bordo_id)
            rs.DeleteObject(ser_id)
        except:
            pass

    # --- Pulizia layer finale ---
    _setup_layers_at_end()


def _setup_layers_at_end():
    # Al termine: colore bianco su tutti i layer Maschere,
    # spegne tutto, accende solo Maschere::REGIONI, lo imposta come attivo.
    import System.Drawing
    white = System.Drawing.Color.White

    maschere_layers = [
        "Maschere",
        "Maschere::01",
        "Maschere::DEBUG",
        "Maschere::REGIONI",
        GANCIO_LAYER_OUT,
        TEXT_LAYER,
    ]

    for ln in maschere_layers:
        try:
            if not rs.IsLayer(ln):
                continue
            rs.LayerColor(ln, white)
            rs.LayerVisible(ln, False)
        except:
            pass

    # Accendi REGIONI e TESTI (e padre Maschere)
    try:
        rs.LayerVisible("Maschere", True)
        rs.LayerVisible("Maschere::REGIONI", True)
        rs.LayerVisible(TEXT_LAYER, True)
    except:
        pass

    # Colore giallo sul layer testi
    try:
        import System.Drawing
        rs.LayerColor(TEXT_LAYER, System.Drawing.Color.Yellow)
    except:
        pass

    # Layer attivo
    try:
        rs.CurrentLayer("Maschere::REGIONI")
    except:
        pass

    sc.doc.Views.Redraw()


if __name__ == "__main__":
    main()