# -*- coding: utf-8 -*-
# Rhino 7 - Python 2.7
# Esplode una policurva, colora segmenti alternati,
# e aggiunge TextDot G0/G1/G2 colorati sui giunti.

import rhinoscriptsyntax as rs
import scriptcontext as sc
import Rhino
import math

# === COLORI ===
WHITE = (255, 255, 255)
LIGHT_PINK = (255, 0, 191)

RED   = (220,   0,   0)    # G0
GREEN = (  0, 160,   0)    # G1 / G2


def _dist(a, b):
    return a.DistanceTo(b)

def _unit(v):
    if v.IsTiny(): 
        return Rhino.Geometry.Vector3d.Zero
    v.Unitize()
    return v

def _angle_between(v1, v2):
    if v1.IsTiny() or v2.IsTiny():
        return math.pi
    v1u = Rhino.Geometry.Vector3d(v1); v1u.Unitize()
    v2u = Rhino.Geometry.Vector3d(v2); v2u.Unitize()
    dot = max(-1.0, min(1.0, Rhino.Geometry.Vector3d.Multiply(v1u, v2u)))
    return math.acos(dot)

def _curve_domain(crv_id):
    d = rs.CurveDomain(crv_id)
    return d if d else (0.0, 1.0)

def _tangent_at(crv_id, at_end):
    d0, d1 = _curve_domain(crv_id)
    t = d1 if at_end else d0
    tan = rs.CurveTangent(crv_id, t)
    return Rhino.Geometry.Vector3d(tan) if tan else Rhino.Geometry.Vector3d.Zero

def _curvature_vec_at(crv_id, at_end):
    d0, d1 = _curve_domain(crv_id)
    t = d1 if at_end else d0
    c = rs.CurveCurvature(crv_id, t)
    if not c or len(c) < 3 or c[2] is None:
        return Rhino.Geometry.Vector3d.Zero
    return Rhino.Geometry.Vector3d(c[2])

def classify_continuity(segA_id, segB_id, abs_tol, ang_tol_rad):
    pA = Rhino.Geometry.Point3d(rs.CurveEndPoint(segA_id))
    pB = Rhino.Geometry.Point3d(rs.CurveStartPoint(segB_id))

    # G0
    if _dist(pA, pB) > abs_tol:
        return "G0"

    # G1
    ta = _unit(_tangent_at(segA_id, True))
    tb = _unit(_tangent_at(segB_id, False))

    if _angle_between(ta, tb) > ang_tol_rad:
        return "G0"

    # G2
    ka = _curvature_vec_at(segA_id, True)
    kb = _curvature_vec_at(segB_id, False)

    if ka.IsTiny() and kb.IsTiny():
        return "G2"

    if ka.IsTiny() != kb.IsTiny():
        return "G1"

    mag_a = ka.Length
    mag_b = kb.Length
    ang_k = _angle_between(ka, kb)

    mag_tol = max(1e-6, 0.10 * max(mag_a, mag_b))

    if ang_k <= ang_tol_rad and abs(mag_a - mag_b) <= mag_tol:
        return "G2"

    return "G1"

def order_and_orient_segments(seg_ids, abs_tol):
    unused = list(seg_ids)
    ordered = [unused.pop(0)]

    while unused:
        last = ordered[-1]
        last_end = Rhino.Geometry.Point3d(rs.CurveEndPoint(last))

        best_i = None
        best_d = None
        flip = False

        for i, cid in enumerate(unused):
            s = Rhino.Geometry.Point3d(rs.CurveStartPoint(cid))
            e = Rhino.Geometry.Point3d(rs.CurveEndPoint(cid))

            ds = _dist(last_end, s)
            de = _dist(last_end, e)

            dmin = ds
            f = False
            if de < ds:
                dmin = de
                f = True

            if best_d is None or dmin < best_d:
                best_d = dmin
                best_i = i
                flip = f

        nxt = unused.pop(best_i)
        if flip:
            rs.ReverseCurve(nxt)

        ordered.append(nxt)

    return ordered

def main():
    crv_id = rs.GetObject(
        "Seleziona una policurva",
        rs.filter.curve,
        preselect=True
    )
    if not crv_id:
        return

    abs_tol = sc.doc.ModelAbsoluteTolerance
    ang_tol = sc.doc.ModelAngleToleranceRadians

    segs = rs.ExplodeCurves(crv_id, delete_input=True)
    if not segs or len(segs) < 2:
        rs.MessageBox("Policurva non valida.", 0)
        return

    segs = order_and_orient_segments(segs, abs_tol)

    # Colori alternati curve
    for i, s in enumerate(segs):
        rs.ObjectColor(s, WHITE if i % 2 == 0 else LIGHT_PINK)

    dots = []

    for i in range(len(segs) - 1):
        a = segs[i]
        b = segs[i + 1]

        label = classify_continuity(a, b, abs_tol, ang_tol)
        p = rs.CurveEndPoint(a)

        dot = rs.AddTextDot(label, p)
        if dot:
            if label == "G0":
                rs.ObjectColor(dot, RED)
            else:
                rs.ObjectColor(dot, GREEN)
            dots.append(dot)

    rs.SelectObjects(segs + dots)

if __name__ == "__main__":
    main()
