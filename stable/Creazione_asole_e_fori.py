# -*- coding: utf-8 -*-
import Rhino
import Rhino.Geometry as rg
import Rhino.DocObjects as rd
import scriptcontext as sc
import Eto.Forms as forms
import Eto.Drawing as drawing
import System
import math

# -----------------------------
# Visual
# -----------------------------
PREVIEW_COLOR = System.Drawing.Color.FromArgb(220, 180, 0)   # giallo/ambra
FINAL_COLOR   = None  # None = colore layer corrente

TOL = 1e-6

# -----------------------------
# Utility
# -----------------------------
def _safe_doc_redraw():
    if sc.doc:
        sc.doc.Views.Redraw()

def _plane_from_closed_planar_curve(crv):
    if not crv:
        return rg.Plane.WorldXY
    ok, pl = crv.TryGetPlane()
    if ok:
        return pl
    return rg.Plane.WorldXY

def _project_vec_to_plane(v, n):
    return v - rg.Vector3d.Multiply(rg.Vector3d.Multiply(v, n), n)

def _canonical_plane_world_axes(plane):
    """
    Piano canonico: stessa origine + stessa normale del contorno,
    assi X/Y allineati al mondo (proiettati sul piano).
    Così +Y = "verso l'alto" nel mondo.
    """
    n = rg.Vector3d(plane.Normal)
    if not n.Unitize():
        return rg.Plane.WorldXY

    x = _project_vec_to_plane(rg.Vector3d.XAxis, n)
    if x.IsTiny(TOL):
        x = _project_vec_to_plane(rg.Vector3d.YAxis, n)
    if not x.Unitize():
        return rg.Plane.WorldXY

    y = rg.Vector3d.CrossProduct(n, x)
    if not y.Unitize():
        return rg.Plane.WorldXY

    return rg.Plane(plane.Origin, x, y)

def _curve_contains_point_planar(crv, pt, tol=None):
    if tol is None:
        tol = sc.doc.ModelAbsoluteTolerance if sc.doc else 0.01
    ok, pl = crv.TryGetPlane()
    if not ok:
        return False
    rc = crv.Contains(pt, pl, tol)
    return (rc == rg.PointContainment.Inside) or (rc == rg.PointContainment.Coincident)

def _section_y_minmax(crv2d, x):
    """
    Intersezione verticale su curva 2D: ritorna y_low, y_high
    """
    line = rg.LineCurve(rg.Point3d(x, -1e9, 0), rg.Point3d(x,  1e9, 0))
    events = Rhino.Geometry.Intersect.Intersection.CurveCurve(crv2d, line, 1e-6, 1e-6)
    if not events or events.Count < 2:
        return None
    ys = [ev.PointA.Y for ev in events]
    if len(ys) < 2:
        return None
    ys.sort()
    return ys[0], ys[-1]

def _dist_point_to_curve_2d(crv2d, pt):
    ok, t = crv2d.ClosestPoint(pt)
    if not ok:
        return None
    cpt = crv2d.PointAt(t)
    return pt.DistanceTo(cpt)

def _inner_offset_curve_2d(closed_crv2d, offset_dist):
    """
    Ritorna la curva offset interna (2D, su WorldXY).
    Robusta su polilinee e curve lisce:
    - prova +d e -d
    - filtra solo loop chiusi
    - verifica che i loop candidati stiano dentro al contorno originale (multi test)
    - se più candidati: prende quello con area massima
    """
    if not closed_crv2d or not closed_crv2d.IsClosed:
        return None
    if offset_dist <= TOL:
        return closed_crv2d.DuplicateCurve()

    tol = sc.doc.ModelAbsoluteTolerance if sc.doc else 0.01
    plane = rg.Plane.WorldXY

    amp = rg.AreaMassProperties.Compute(closed_crv2d)
    centroid = amp.Centroid if amp else rg.Point3d(0, 0, 0)

    def is_inside_original(pt):
        rc = closed_crv2d.Contains(pt, plane, tol)
        return (rc == rg.PointContainment.Inside) or (rc == rg.PointContainment.Coincident)

    def candidate_ok(c):
        if not c or (not c.IsClosed):
            return False

        # 1) centroide original dentro il candidato
        rc1 = c.Contains(centroid, plane, tol)
        inside_centroid = (rc1 == rg.PointContainment.Inside) or (rc1 == rg.PointContainment.Coincident)

        # 2) più punti del candidato dentro l'originale
        inside_votes = 0
        for tnorm in [0.05, 0.25, 0.5, 0.75, 0.95]:
            t = c.Domain.ParameterAt(tnorm)
            pt = c.PointAt(t)
            if is_inside_original(pt):
                inside_votes += 1

        return inside_centroid and (inside_votes >= 4)

    candidates = []

    # primo tentativo: Sharp
    for d in [offset_dist, -offset_dist]:
        try:
            offs = closed_crv2d.Offset(plane, d, tol, rg.CurveOffsetCornerStyle.Sharp)
        except:
            offs = None
        if not offs:
            continue
        for c in offs:
            if candidate_ok(c):
                candidates.append(c)

    # fallback: Round
    if not candidates:
        for d in [offset_dist, -offset_dist]:
            try:
                offs = closed_crv2d.Offset(plane, d, tol, rg.CurveOffsetCornerStyle.Round)
            except:
                offs = None
            if not offs:
                continue
            for c in offs:
                if candidate_ok(c):
                    candidates.append(c)

    if not candidates:
        return None

    # scegli loop interno "principale": area massima
    best = None
    best_area = -1e99
    for c in candidates:
        a = rg.AreaMassProperties.Compute(c)
        area = a.Area if a else -1e99
        if area > best_area:
            best_area = area
            best = c

    return best


# -----------------------------
# Preview manager
# -----------------------------
class PreviewManager(object):
    def __init__(self):
        self.ids = []

    def clear(self):
        if not sc.doc:
            return
        for oid in self.ids:
            try:
                sc.doc.Objects.Delete(oid, True)
            except:
                pass
        self.ids = []
        _safe_doc_redraw()

    def add_curve(self, crv, color=PREVIEW_COLOR):
        if not sc.doc or not crv:
            return
        attr = rd.ObjectAttributes()
        attr.ColorSource = rd.ObjectColorSource.ColorFromObject
        attr.ObjectColor = color
        oid = sc.doc.Objects.AddCurve(crv, attr)
        if oid != System.Guid.Empty:
            self.ids.append(oid)

# -----------------------------
# Geometry helpers
# -----------------------------
def _fillet_closed_polyline(poly_crv, radius):
    if radius <= 0:
        return poly_crv
    tol = sc.doc.ModelAbsoluteTolerance if sc.doc else 0.01
    ang = sc.doc.ModelAngleToleranceRadians if sc.doc else (math.pi/180.0)
    try:
        fc = rg.Curve.CreateFilletCornersCurve(poly_crv, radius, tol, ang)
        if fc:
            return fc
    except:
        pass
    return poly_crv

def _make_slot_quad_top_follow_2d(curve_for_top, cx, slot_w, y_bottom_const, y_top_const, top_follow, fillet_r):
    """
    Asola 2D:
    - bottom = y_bottom_const (piatto)
    - top:
        * se top_follow=True: top segue il profilo superiore della 'curve_for_top'
          campionato ai lati (x0, x1), usando i massimi y
        * altrimenti top = y_top_const (piatto)
    Nota: curve_for_top è già "inner offset" (se disponibile).
    """
    x0 = cx - slot_w * 0.5
    x1 = cx + slot_w * 0.5

    s0 = _section_y_minmax(curve_for_top, x0)
    s1 = _section_y_minmax(curve_for_top, x1)
    if not s0 or not s1:
        return None

    y0_low, y0_high = s0
    y1_low, y1_high = s1

    yb = y_bottom_const

    if top_follow:
        yt0 = y0_high
        yt1 = y1_high
    else:
        yt0 = y_top_const
        yt1 = y_top_const

    if (yt0 - yb) <= TOL or (yt1 - yb) <= TOL:
        return None

    pts = [
        rg.Point3d(x0, yb, 0),
        rg.Point3d(x1, yb, 0),
        rg.Point3d(x1, yt1, 0),
        rg.Point3d(x0, yt0, 0),
        rg.Point3d(x0, yb, 0)
    ]
    poly = rg.PolylineCurve(rg.Polyline(pts))
    return _fillet_closed_polyline(poly, fillet_r)

# -----------------------------
# Builders
# -----------------------------
def build_slots_mode(boundary_crv, plane_canon,
                     offset_laterale,
                     slot_w,
                     distanza_asole,
                     offset_orz,
                     fillet_r,
                     do_circles,
                     circle_offset_supinf,
                     circle_diam,
                     do_split,
                     split_threshold):
    """
    Modalità ASOLE:
    - offset_orz è un vero offset interno (curve offset)
    - bottom asola piatto (baseline interna bassa)
    - solo top asola segue profilo superiore (inner) per la prima riga della colonna verticale:
        * se split: solo l'ULTIMA (più alta) segue il profilo superiore
    - split verticale se altezza utile >= soglia: crea n asole in colonna con gap = distanza_asole
    - fori circolari sopra/sotto tra asole seguono l'offset interno (stessa inner band)
    """
    geoms = []
    if not boundary_crv:
        return geoms

    to_xy   = rg.Transform.PlaneToPlane(plane_canon, rg.Plane.WorldXY)
    from_xy = rg.Transform.PlaneToPlane(rg.Plane.WorldXY, plane_canon)

    crv2d = boundary_crv.DuplicateCurve()
    crv2d.Transform(to_xy)

    bb = crv2d.GetBoundingBox(True)
    minx, maxx = bb.Min.X, bb.Max.X
    W = maxx - minx

    offset_laterale = float(offset_laterale)
    slot_w = float(slot_w)
    distanza_asole = float(distanza_asole)
    offset_orz = float(offset_orz)
    fillet_r = float(fillet_r)

    # Offset interno reale
    inner2d = _inner_offset_curve_2d(crv2d, offset_orz)
    curve_in = inner2d if inner2d else crv2d

    avail_w = W - 2.0*offset_laterale
    if avail_w <= slot_w + TOL:
        return geoms

    pitch = slot_w + distanza_asole
    if pitch <= TOL:
        return geoms

    # numero asole e centratura sx/dx
    n = int(math.floor((avail_w + distanza_asole) / pitch))
    if n < 1:
        n = 1

    used = n*slot_w + (n-1)*distanza_asole
    extra = avail_w - used
    if extra < 0:
        extra = 0

    start_center_x = minx + offset_laterale + (extra*0.5) + slot_w*0.5

    centers = []
    for i in range(n):
        cx = start_center_x + i*pitch
        if (cx - slot_w*0.5) < (minx + offset_laterale - TOL):
            continue
        if (cx + slot_w*0.5) > (maxx - offset_laterale + TOL):
            break

        sec_mid = _section_y_minmax(curve_in, cx)
        if not sec_mid:
            continue
        y_low_mid, y_high_mid = sec_mid

        # banda interna utile (su inner2d)
        y_bottom = y_low_mid
        y_top_mid = y_high_mid
        height = y_top_mid - y_bottom
        if height <= TOL:
            continue

        # split verticale
        split_threshold = float(split_threshold)
        do_split = bool(do_split) and (split_threshold > TOL)

        if do_split and height >= (split_threshold - TOL):
            nseg = int(math.ceil(height / split_threshold))
            if nseg < 2:
                nseg = 2

            total_gap = (nseg - 1) * distanza_asole
            usable_h = height - total_gap

            if usable_h <= TOL:
                # fallback: singola
                nseg = 1
                usable_h = height

            slot_h = usable_h / float(nseg)

            for si in range(nseg):
                yb = y_bottom + si * (slot_h + distanza_asole)
                yt_const = yb + slot_h

                top_follow = (si == (nseg - 1))  # solo l'ultima segue il profilo sup

                slot2d = _make_slot_quad_top_follow_2d(
                    curve_in, cx, slot_w,
                    yb, yt_const,
                    top_follow, fillet_r
                )
                if not slot2d:
                    continue
                slot2d.Transform(from_xy)
                geoms.append(slot2d)
        else:
            slot2d = _make_slot_quad_top_follow_2d(
                curve_in, cx, slot_w,
                y_bottom, y_top_mid,
                True, fillet_r
            )
            if not slot2d:
                continue
            slot2d.Transform(from_xy)
            geoms.append(slot2d)

        centers.append(cx)

    # fori circolari sopra/sotto tra asole (su inner band)
    if do_circles and len(centers) >= 2:
        circle_offset_supinf = float(circle_offset_supinf)
        r = float(circle_diam) * 0.5

        for i in range(len(centers)-1):
            cxm = 0.5*(centers[i] + centers[i+1])

            secm = _section_y_minmax(curve_in, cxm)
            if not secm:
                continue
            y_low, y_high = secm

            y_bottom_band = y_low
            y_top_band    = y_high

            y_bot = y_bottom_band + circle_offset_supinf
            y_top = y_top_band    - circle_offset_supinf

            for cy in [y_bot, y_top]:
                # deve stare dentro inner
                if (cy - r) <= (y_low + TOL) or (cy + r) >= (y_high - TOL):
                    continue

                # e deve stare dentro l'outer (controllo robusto)
                d_outer = _dist_point_to_curve_2d(crv2d, rg.Point3d(cxm, cy, 0))
                if d_outer is None or d_outer < (r - TOL):
                    continue

                c2d = rg.ArcCurve(rg.Circle(rg.Plane.WorldXY, rg.Point3d(cxm, cy, 0), r))
                c2d.Transform(from_xy)
                geoms.append(c2d)

    return geoms


def build_holes_grid_mode(boundary_crv, plane_canon,
                          offset_laterale,
                          offset_verticale,
                          diametro,
                          passo,
                          diam_min, diam_max):
    """
    Modalità FORI:
    - griglia a passo fisso (X e Y)
    - centrato sx/dx (equidistante)
    - ALLINEATO AL BASSO (y_min + r), NON centrato verticalmente
    - si sale per righe finché:
        * resta dentro y_max (gap alto/basso)
        * dist al contorno >= r
    """
    geoms = []
    if not boundary_crv:
        return geoms

    to_xy   = rg.Transform.PlaneToPlane(plane_canon, rg.Plane.WorldXY)
    from_xy = rg.Transform.PlaneToPlane(rg.Plane.WorldXY, plane_canon)

    crv2d = boundary_crv.DuplicateCurve()
    crv2d.Transform(to_xy)

    bb = crv2d.GetBoundingBox(True)
    minx, maxx = bb.Min.X, bb.Max.X
    W = maxx - minx

    offset_laterale = float(offset_laterale)
    offset_verticale = float(offset_verticale)

    diam_min = float(diam_min)
    diam_max = float(diam_max)
    if diam_max < diam_min:
        diam_min, diam_max = diam_max, diam_min

    D = float(diametro)
    if D < diam_min: D = diam_min
    if D > diam_max: D = diam_max
    r = 0.5*D

    passo = float(passo)
    if passo < D:
        passo = D

    avail_w = W - 2.0*offset_laterale
    if avail_w <= D + TOL:
        return geoms

    ncol = int(math.floor((avail_w - D) / passo)) + 1
    if ncol < 1:
        ncol = 1

    used_w = (ncol-1)*passo + D
    extra = avail_w - used_w
    if extra < 0:
        extra = 0

    start_x = minx + offset_laterale + extra*0.5 + r

    for ci in range(ncol):
        cx = start_x + ci*passo

        sec = _section_y_minmax(crv2d, cx)
        if not sec:
            continue
        y_low, y_high = sec

        y_min = y_low + offset_verticale
        y_max = y_high - offset_verticale
        if (y_max - y_min) <= D + TOL:
            continue

        cy = y_min + r  # baseline bassa

        while (cy + r) <= (y_max + TOL):
            d = _dist_point_to_curve_2d(crv2d, rg.Point3d(cx, cy, 0))
            if d is not None and d >= (r - TOL):
                c2d = rg.ArcCurve(rg.Circle(rg.Plane.WorldXY, rg.Point3d(cx, cy, 0), r))
                c2d.Transform(from_xy)
                geoms.append(c2d)
            cy += passo

    return geoms


# -----------------------------
# Selection
# -----------------------------
def pick_boundary_curve():
    go = Rhino.Input.Custom.GetObject()
    go.SetCommandPrompt("Seleziona il contorno chiuso (Opzione: ClickInterno)")
    opt_click = Rhino.Input.Custom.OptionToggle(False, "No", "Si")
    go.AddOptionToggle("ClickInterno", opt_click)

    go.GeometryFilter = rd.ObjectType.Curve
    go.SubObjectSelect = False
    go.EnablePreSelect(True, True)

    while True:
        res = go.Get()
        if res == Rhino.Input.GetResult.Option:
            continue
        if res == Rhino.Input.GetResult.Object:
            crv = go.Object(0).Curve()
            if not crv or not crv.IsClosed:
                print("La curva selezionata non è chiusa.")
                return None
            ok, _ = crv.TryGetPlane()
            if not ok:
                print("Curva chiusa ma non planare: serve un contorno planare.")
                return None
            return crv.DuplicateCurve()
        if res == Rhino.Input.GetResult.Cancel:
            return None
        if res == Rhino.Input.GetResult.Nothing:
            break

    if not opt_click.CurrentValue:
        return None

    gp = Rhino.Input.Custom.GetPoint()
    gp.SetCommandPrompt("Clicca un punto all'interno del contorno chiuso")
    if gp.Get() != Rhino.Input.GetResult.Point:
        return None
    pt = gp.Point()

    candidates = []
    for obj in sc.doc.Objects:
        if obj.ObjectType != rd.ObjectType.Curve:
            continue
        c = obj.Geometry
        if not isinstance(c, rg.Curve):
            continue
        if not c.IsClosed:
            continue
        ok, _ = c.TryGetPlane()
        if not ok:
            continue
        if _curve_contains_point_planar(c, pt):
            amp = rg.AreaMassProperties.Compute(c)
            area = amp.Area if amp else 1e99
            candidates.append((area, c.DuplicateCurve()))

    if not candidates:
        print("Nessun contorno chiuso planare trovato che contenga il punto.")
        return None

    candidates.sort(key=lambda x: x[0])
    return candidates[0][1]


# -----------------------------
# UI
# -----------------------------
class AsoleForiDialog(forms.Dialog[bool]):
    def __init__(self):
        self.Title = "Asole e Fori"
        self.ClientSize = drawing.Size(860, 600)
        self.Padding = drawing.Padding(10)
        self.Resizable = False

        self.boundary = None
        self.plane = None
        self.plane_canon = None
        self.preview = PreviewManager()

        # Selezione contorno
        self.lbl_boundary = forms.Label(Text="Contorno: NON selezionato")
        self.lbl_boundary.TextColor = drawing.Color.FromArgb(255, 200, 0, 0)
        self.btn_pick = forms.Button(Text="Seleziona contorno")
        self.btn_pick.Click += self.on_pick_boundary

        # mode
        self.rb_fori = forms.RadioButton(Text="Fori")
        self.rb_asole = forms.RadioButton(self.rb_fori, Text="Asole")
        self.rb_fori.Checked = True  # default

        # ----------------
        # DEFAULTS (dallo screenshot)
        # ----------------
        # FORI
        self.f_offset_laterale  = forms.NumericUpDown(MinValue=0, MaxValue=10000, DecimalPlaces=0, Value=300)
        self.f_offset_verticale = forms.NumericUpDown(MinValue=0, MaxValue=10000, DecimalPlaces=0, Value=100)
        self.f_diametro         = forms.NumericUpDown(MinValue=0, MaxValue=10000, DecimalPlaces=0, Value=125)
        self.f_passo            = forms.NumericUpDown(MinValue=0, MaxValue=10000, DecimalPlaces=0, Value=200)
        self.f_diam_min         = forms.NumericUpDown(MinValue=0, MaxValue=10000, DecimalPlaces=0, Value=80)
        self.f_diam_max         = forms.NumericUpDown(MinValue=0, MaxValue=10000, DecimalPlaces=0, Value=250)

        # ASOLE
        self.a_offset_laterale  = forms.NumericUpDown(MinValue=0, MaxValue=10000, DecimalPlaces=0, Value=20)
        self.a_larghezza_asola  = forms.NumericUpDown(MinValue=1, MaxValue=10000, DecimalPlaces=0, Value=60)
        self.a_distanza_asole   = forms.NumericUpDown(MinValue=0, MaxValue=10000, DecimalPlaces=0, Value=35)
        self.a_offset_orizz     = forms.NumericUpDown(MinValue=0, MaxValue=10000, DecimalPlaces=0, Value=30)
        self.a_raggiatura       = forms.NumericUpDown(MinValue=0, MaxValue=10000, DecimalPlaces=0, Value=15)

        self.chk_fori_circolari = forms.CheckBox(Text="Fori Circolari")
        self.chk_fori_circolari.Checked = True
        self.ac_offset_supinf   = forms.NumericUpDown(MinValue=0, MaxValue=10000, DecimalPlaces=0, Value=0)
        self.ac_diam            = forms.NumericUpDown(MinValue=1, MaxValue=10000, DecimalPlaces=0, Value=15)

        # split asole
        self.chk_split = forms.CheckBox(Text="Dividi verticalmente oltre (mm)")
        self.chk_split.Checked = True
        self.split_mm = forms.NumericUpDown(MinValue=1, MaxValue=10000, DecimalPlaces=0, Value=400)

        # buttons
        self.btn_ok = forms.Button(Text="Ok")
        self.btn_cancel = forms.Button(Text="Annulla")
        self.btn_ok.Click += self.on_ok
        self.btn_cancel.Click += self.on_cancel
        self.btn_ok.Enabled = False

        # X = annulla
        self.Closing += self.on_closing

        # events
        controls = [
            self.rb_fori, self.rb_asole,
            self.f_offset_laterale, self.f_offset_verticale, self.f_diametro, self.f_passo, self.f_diam_min, self.f_diam_max,
            self.a_offset_laterale, self.a_larghezza_asola, self.a_distanza_asole, self.a_offset_orizz, self.a_raggiatura,
            self.chk_fori_circolari, self.ac_offset_supinf, self.ac_diam,
            self.chk_split, self.split_mm
        ]
        for ctl in controls:
            try: ctl.ValueChanged += self.on_any_change
            except: pass
            try: ctl.CheckedChanged += self.on_any_change
            except: pass

        # layout
        top_select = forms.TableLayout(Spacing=drawing.Size(10, 6))
        top_select.Rows.Add(forms.TableRow(self.lbl_boundary, self.btn_pick))

        top = forms.TableLayout(Spacing=drawing.Size(10, 6))
        top.Rows.Add(forms.TableRow(self.rb_fori, self._spacer(), self.rb_asole))

        # FORI layout
        fori_grid = forms.TableLayout(Spacing=drawing.Size(10, 6))
        fori_grid.Rows.Add(forms.TableRow(
            forms.Label(Text="Offset laterale (mm):"), self.f_offset_laterale,
            forms.Label(Text="Offset verticale (mm):"), self.f_offset_verticale
        ))
        fori_grid.Rows.Add(forms.TableRow(
            forms.Label(Text="Diametro (mm):"), self.f_diametro,
            forms.Label(Text="Passo (mm):"), self.f_passo
        ))
        fori_grid.Rows.Add(forms.TableRow(
            forms.Label(Text="Diametro min (mm):"), self.f_diam_min,
            forms.Label(Text="Diametro max (mm):"), self.f_diam_max
        ))

        # ASOLE layout
        asole_grid = forms.TableLayout(Spacing=drawing.Size(10, 6))
        asole_grid.Rows.Add(forms.TableRow(
            forms.Label(Text="Offset laterale (mm):"), self.a_offset_laterale,
            forms.Label(Text="Larghezza asola (mm):"), self.a_larghezza_asola,
            forms.Label(Text="Distanza asole (mm):"), self.a_distanza_asole
        ))
        asole_grid.Rows.Add(forms.TableRow(
            forms.Label(Text="Offset orizzontale (mm):"), self.a_offset_orizz,
            forms.Label(Text="Raggiatura asole (mm):"), self.a_raggiatura
        ))

        split_row = forms.TableLayout(Spacing=drawing.Size(10, 6))
        split_row.Rows.Add(forms.TableRow(self.chk_split, self.split_mm, self._spacer(), self._spacer(), self._spacer()))

        ac_row = forms.TableLayout(Spacing=drawing.Size(10, 6))
        ac_row.Rows.Add(forms.TableRow(
            self.chk_fori_circolari,
            forms.Label(Text="Offset bordi sup./inf. (mm):"), self.ac_offset_supinf,
            forms.Label(Text="Diametro (mm):"), self.ac_diam
        ))

        btns = forms.StackLayout(Orientation=forms.Orientation.Horizontal,
                                 Spacing=10, HorizontalContentAlignment=forms.HorizontalAlignment.Center)
        btns.Items.Add(self.btn_ok)
        btns.Items.Add(self.btn_cancel)

        main = forms.TableLayout(Spacing=drawing.Size(10, 12))
        main.Rows.Add(forms.TableRow(top_select))
        main.Rows.Add(forms.TableRow(self._sep()))
        main.Rows.Add(forms.TableRow(top))
        main.Rows.Add(forms.TableRow(self._sep()))
        main.Rows.Add(forms.TableRow(fori_grid))
        main.Rows.Add(forms.TableRow(self._sep()))
        main.Rows.Add(forms.TableRow(asole_grid))
        main.Rows.Add(forms.TableRow(split_row))
        main.Rows.Add(forms.TableRow(ac_row))
        main.Rows.Add(forms.TableRow(self._sep()))
        main.Rows.Add(forms.TableRow(btns))

        self.Content = main
        self._sync_enabled()

    def _spacer(self):
        return forms.Panel(Size=drawing.Size(10, 10))

    def _sep(self):
        line = forms.Panel()
        line.Height = 1
        line.BackgroundColor = drawing.Color.FromArgb(255, 180, 180, 180)
        return line

    def on_closing(self, sender, e):
        # X = Annulla (pulizia preview)
        try:
            self.preview.clear()
        except:
            pass

    def on_pick_boundary(self, sender, e):
        self.Visible = False
        try:
            crv = pick_boundary_curve()
        finally:
            self.Visible = True

        if not crv:
            return

        self.boundary = crv
        self.plane = _plane_from_closed_planar_curve(crv)
        self.plane_canon = _canonical_plane_world_axes(self.plane)

        self.lbl_boundary.Text = "Contorno: SELEZIONATO"
        self.lbl_boundary.TextColor = drawing.Color.FromArgb(255, 0, 160, 0)

        self.btn_ok.Enabled = True
        self.update_preview()

    def on_any_change(self, sender, e):
        self._sync_enabled()
        self.update_preview()

    def _sync_enabled(self):
        is_fori = bool(self.rb_fori.Checked)
        for c in [self.f_offset_laterale, self.f_offset_verticale, self.f_diametro, self.f_passo, self.f_diam_min, self.f_diam_max]:
            c.Enabled = is_fori

        is_asole = bool(self.rb_asole.Checked)
        for v in [self.a_offset_laterale, self.a_larghezza_asola, self.a_distanza_asole,
                  self.a_offset_orizz, self.a_raggiatura, self.chk_fori_circolari,
                  self.chk_split]:
            v.Enabled = is_asole

        split_on = bool(self.chk_split.Checked) and is_asole
        self.split_mm.Enabled = split_on

        fc_on = bool(self.chk_fori_circolari.Checked) and is_asole
        self.ac_offset_supinf.Enabled = fc_on
        self.ac_diam.Enabled = fc_on

    def _compute_geoms(self):
        if not self.boundary or not self.plane_canon:
            return []

        if bool(self.rb_asole.Checked):
            return build_slots_mode(
                self.boundary,
                self.plane_canon,
                float(self.a_offset_laterale.Value),
                float(self.a_larghezza_asola.Value),
                float(self.a_distanza_asole.Value),
                float(self.a_offset_orizz.Value),
                float(self.a_raggiatura.Value),
                bool(self.chk_fori_circolari.Checked),
                float(self.ac_offset_supinf.Value),
                float(self.ac_diam.Value),
                bool(self.chk_split.Checked),
                float(self.split_mm.Value)
            )
        else:
            return build_holes_grid_mode(
                self.boundary,
                self.plane_canon,
                float(self.f_offset_laterale.Value),
                float(self.f_offset_verticale.Value),
                float(self.f_diametro.Value),
                float(self.f_passo.Value),
                float(self.f_diam_min.Value),
                float(self.f_diam_max.Value)
            )

    def update_preview(self):
        self.preview.clear()
        if not self.boundary:
            return
        geoms = self._compute_geoms()
        for g in geoms:
            self.preview.add_curve(g, PREVIEW_COLOR)
        _safe_doc_redraw()

    def on_ok(self, sender, e):
        self.preview.clear()
        geoms = self._compute_geoms()

        for g in geoms:
            if FINAL_COLOR is None:
                sc.doc.Objects.AddCurve(g)
            else:
                attr = rd.ObjectAttributes()
                attr.ColorSource = rd.ObjectColorSource.ColorFromObject
                attr.ObjectColor = FINAL_COLOR
                sc.doc.Objects.AddCurve(g, attr)

        _safe_doc_redraw()
        self.Close(True)

    def on_cancel(self, sender, e):
        self.preview.clear()
        self.Close(False)

# -----------------------------
# Main
# -----------------------------
def main():
    dlg = AsoleForiDialog()
    dlg.Owner = Rhino.UI.RhinoEtoApp.MainWindow
    dlg.ShowModal(dlg.Owner)

main()