# -*- coding: utf-8 -*-
import rhinoscriptsyntax as rs
import scriptcontext as sc
import Rhino
import System.Drawing as SD
import re
import Eto.Forms as forms
import Eto.Drawing as drawing

# =========================================================
# CONFIG
# =========================================================
PARENT_LAYER = "VETRI"

COL_PARENT = SD.Color.Black
COL_F1     = SD.Color.Red
COL_F2     = SD.ColorTranslator.FromHtml("#8B8B00")  # monolitico
COL_1V     = SD.Color.Orange
COL_2V     = SD.Color.LimeGreen
COL_3V     = SD.Color.Indigo
COL_F4     = SD.Color.Cyan
COL_F6     = SD.Color.Violet

STICKY_LAST_COMP = "GLASS_COMP_LAST"

# =========================================================
# LAYER UTILS
# =========================================================
def _find_layer_index(fullpath):
    return sc.doc.Layers.FindByFullPath(fullpath, True)

def _ensure_layer(fullpath, color):
    idx = _find_layer_index(fullpath)
    if idx >= 0:
        return idx

    if "::" in fullpath:
        parent_path, name = fullpath.rsplit("::", 1)
        if _find_layer_index(parent_path) < 0:
            _ensure_layer(parent_path, COL_PARENT)

        pidx = _find_layer_index(parent_path)
        layer = Rhino.DocObjects.Layer()
        layer.Name = name
        layer.Color = color
        layer.ParentLayerId = sc.doc.Layers[pidx].Id
        return sc.doc.Layers.Add(layer)

    layer = Rhino.DocObjects.Layer()
    layer.Name = fullpath
    layer.Color = color
    return sc.doc.Layers.Add(layer)

def _set_obj_layer(obj_id, layer_fullpath):
    rs.ObjectLayer(obj_id, layer_fullpath)

# =========================================================
# COMPOSIZIONE
# =========================================================
def parse_composition(text):
    if not text:
        return []
    text = text.replace(",", ".")
    nums = re.findall(r"\d+(?:\.\d+)?", text)
    return [float(x) for x in nums]

def split_glass_interlayers(seq):
    glass, inter = [], []
    for i, v in enumerate(seq):
        (glass if i % 2 == 0 else inter).append(v)
    return glass, inter

def compute_offsets(glass_thk, inter_thk):
    mid = []
    cum = 0.0
    for i, t in enumerate(glass_thk):
        mid.append(-(cum + t / 2.0))
        cum += t
        if i < len(inter_thk):
            cum += inter_thk[i]

    total = sum(glass_thk) + sum(inter_thk)
    return mid, -total, total

def offset_surface(obj_id, dist):
    try:
        res = rs.OffsetSurface(obj_id, dist, create_solid=False)
    except:
        return []
    if isinstance(res, (list, tuple)):
        return list(res)
    return [res] if res else []

# =========================================================
# LAYER LOGIC
# =========================================================
def color_for_layer(name):
    if name == "F1": return COL_F1
    if name == "F2": return COL_F2
    if name == "1V": return COL_1V
    if name == "2V": return COL_2V
    if name == "3V": return COL_3V
    if name == "F4": return COL_F4
    if name == "F6": return COL_F6
    # fallback robusto
    if name.startswith("F"): return COL_F6
    if name.endswith("V"): return COL_3V
    return SD.Color.Gray

def ensure_needed_layers(glass_count):
    _ensure_layer(PARENT_LAYER, COL_PARENT)

    names = ["F1"]
    for i in range(1, glass_count + 1):
        names.append("{}V".format(i))

    # chiusura: monolitico -> F2, stratificato -> F(2n)
    names.append("F2" if glass_count == 1 else "F{}".format(2 * glass_count))

    for n in names:
        _ensure_layer(PARENT_LAYER + "::" + n, color_for_layer(n))

# =========================================================
# CORE ENGINE
# =========================================================
def apply_composition_to_surface(obj_id, comp_text):
    seq = parse_composition(comp_text)
    if not seq:
        return False, "Composizione non valida."

    glass, inter = split_glass_interlayers(seq)
    if not glass:
        return False, "Manca lo spessore del vetro."

    ensure_needed_layers(len(glass))
    mids, close_off, total = compute_offsets(glass, inter)

    rs.EnableRedraw(False)
    created = []

    # F1 copia
    f1 = rs.CopyObject(obj_id)
    if f1:
        _set_obj_layer(f1, PARENT_LAYER + "::F1")
        created.append(f1)

    # 1V..nV
    for i, off in enumerate(mids, start=1):
        for nid in offset_surface(obj_id, off):
            _set_obj_layer(nid, PARENT_LAYER + "::{}V".format(i))
            created.append(nid)

    # faccia di chiusura: F2 o F(2n)
    fname = "F2" if len(glass) == 1 else "F{}".format(2 * len(glass))
    for nid in offset_surface(obj_id, close_off):
        _set_obj_layer(nid, PARENT_LAYER + "::" + fname)
        created.append(nid)

    rs.EnableRedraw(True)

    if not created:
        return False, "Offset fallito (normale invertita o superficie problematica)."

    rs.SelectObjects(created)
    return True, "OK – creati {} oggetti | spessore {:.2f} mm".format(len(created), total)

def apply_composition_to_selection(comp_text):
    ids = rs.SelectedObjects() or []
    valid = [i for i in ids if rs.IsSurface(i) or rs.IsPolysurface(i)]
    if not valid:
        return False, "Nessuna superficie/polisuperficie selezionata."

    okc = 0
    last = ""
    for o in valid:
        ok, msg = apply_composition_to_surface(o, comp_text)
        last = msg
        if ok:
            okc += 1
    return True, "Applicato a {}/{} superfici. Ultimo: {}".format(okc, len(valid), last)

# =========================================================
# UI – Dialog (chiude per fare pick, poi si riapre)
# =========================================================
class GlassComposeDialog(forms.Dialog[bool]):
    def __init__(self, initial_comp):
        forms.Dialog.__init__(self)
        self.Title = "Creazione composizione multipla"
        self.Padding = drawing.Padding(10)
        self.Resizable = False
        self.ClientSize = drawing.Size(520, 180)

        self.Action = None  # "pick" | "apply_sel" | "close"
        self.Comp = initial_comp or "6/1.52/6"

        self.txtComp = forms.TextBox(Text=self.Comp)
        self.lblStatus = forms.Label(Text="Imposta composizione, poi scegli azione.", Wrap=forms.WrapMode.Word)

        btnPick = forms.Button(Text="Seleziona superficie + Applica")
        btnPick.Click += self.on_pick

        btnSel = forms.Button(Text="Applica a selezione")
        btnSel.Click += self.on_apply_sel

        btnClose = forms.Button(Text="Chiudi")
        btnClose.Click += self.on_close

        layout = forms.DynamicLayout(Spacing=drawing.Size(8, 8))
        layout.AddRow(forms.Label(Text="Composizione:"), self.txtComp)
        layout.AddRow(None)
        layout.AddRow(btnPick, btnSel, None, btnClose)
        layout.AddRow(None)
        layout.AddRow(forms.Label(Text="Note:"), self.lblStatus)

        self.Content = layout

    def _save_comp(self):
        self.Comp = self.txtComp.Text
        sc.sticky[STICKY_LAST_COMP] = self.Comp

    def on_pick(self, s, e):
        self._save_comp()
        self.Action = "pick"
        self.Close(True)

    def on_apply_sel(self, s, e):
        self._save_comp()
        self.Action = "apply_sel"
        self.Close(True)

    def on_close(self, s, e):
        self._save_comp()
        self.Action = "close"
        self.Close(False)

# =========================================================
# RUN LOOP (robusto)
# =========================================================
def run():
    last_comp = sc.sticky.get(STICKY_LAST_COMP, "6/1.52/6")

    while True:
        dlg = GlassComposeDialog(last_comp)
        # Modale = stabile. Sparisce per forza durante la selezione.
        ok = dlg.ShowModal(Rhino.UI.RhinoEtoApp.MainWindow)

        last_comp = dlg.Comp  # persistiamo sempre l'ultimo testo

        if not ok or dlg.Action == "close":
            break

        if dlg.Action == "apply_sel":
            ok2, msg = apply_composition_to_selection(last_comp)
            rs.StatusBarMessage(msg)
            continue

        if dlg.Action == "pick":
            obj_id = rs.GetObject(
                "Seleziona la superficie/polisuperficie (F1)",
                rs.filter.surface | rs.filter.polysurface,
                preselect=True
            )
            if not obj_id:
                rs.StatusBarMessage("Selezione annullata.")
                continue

            ok2, msg = apply_composition_to_surface(obj_id, last_comp)
            rs.StatusBarMessage(msg)
            continue

if __name__ == "__main__":
    run()
