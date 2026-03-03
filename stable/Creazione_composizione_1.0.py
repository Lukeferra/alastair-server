# -*- coding: utf-8 -*-
import rhinoscriptsyntax as rs
import scriptcontext as sc
import Rhino
import System.Drawing as SD
import re

PARENT_LAYER = "VETRI"

# Palette standard
COL_PARENT = SD.Color.Black
COL_F1     = SD.Color.Red
COL_F2     = SD.ColorTranslator.FromHtml("#8B8B00")  # monolitico
COL_1V     = SD.Color.Orange
COL_2V     = SD.Color.LimeGreen
COL_3V     = SD.Color.Indigo
COL_F4     = SD.Color.Cyan
COL_F6     = SD.Color.Violet

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

def parse_composition(text):
    if not text:
        return []
    text = text.replace(",", ".")
    nums = re.findall(r"\d+(?:\.\d+)?", text)
    return [float(x) for x in nums]

def split_glass_interlayers(seq):
    glass = []
    inter = []
    for i, v in enumerate(seq):
        if i % 2 == 0:
            glass.append(v)
        else:
            inter.append(v)
    return glass, inter

def compute_offsets(glass_thk, inter_thk):
    """
    mid_offsets: superfici medie dei vetri
    close_offset: faccia di chiusura (SEMPRE)
    """
    mid = []
    cum = 0.0
    for i, tg in enumerate(glass_thk):
        mid.append(-(cum + tg / 2.0))
        cum += tg
        if i < len(inter_thk):
            cum += inter_thk[i]

    total = sum(glass_thk) + sum(inter_thk)
    close = -total
    return mid, close

def offset_surface(obj_id, dist):
    try:
        res = rs.OffsetSurface(obj_id, dist, create_solid=False)
    except:
        return []

    if isinstance(res, (list, tuple)):
        return list(res)
    return [res] if res else []

def color_for_layer_name(name):
    if name == "F1": return COL_F1
    if name == "F2": return COL_F2
    if name == "1V": return COL_1V
    if name == "2V": return COL_2V
    if name == "3V": return COL_3V
    if name == "F4": return COL_F4
    if name == "F6": return COL_F6

    if name.endswith("V"):
        return COL_3V
    if name.startswith("F"):
        return COL_F6
    return SD.Color.Gray

def ensure_needed_layers(glass_count):
    _ensure_layer(PARENT_LAYER, COL_PARENT)

    needed = ["F1"]

    for i in range(1, glass_count + 1):
        needed.append("{}V".format(i))

    # PATCH MONOLITICO
    if glass_count == 1:
        needed.append("F2")
    else:
        needed.append("F{}".format(2 * glass_count))

    for lname in needed:
        _ensure_layer(PARENT_LAYER + "::" + lname, color_for_layer_name(lname))

def main():
    obj_id = rs.GetObject(
        "Seleziona la superficie/polisuperficie (questa sarà F1)",
        rs.filter.surface | rs.filter.polysurface,
        preselect=True
    )
    if not obj_id:
        return

    comp = rs.GetString(
        "Inserisci composizione (es: 8  oppure 6/1.52/6  oppure 5/1.52interlayer/6/0.76/10)"
    )

    seq = parse_composition(comp)
    if not seq:
        print("Composizione non valida.")
        return

    glass_thk, inter_thk = split_glass_interlayers(seq)
    glass_count = len(glass_thk)
    if glass_count < 1:
        print("Errore composizione.")
        return

    ensure_needed_layers(glass_count)
    mid_offsets, close_offset = compute_offsets(glass_thk, inter_thk)

    rs.EnableRedraw(False)
    created = []

    # F1
    f1 = rs.CopyObject(obj_id)
    if f1:
        _set_obj_layer(f1, PARENT_LAYER + "::F1")
        created.append(f1)

    # 1V..nV
    for i, off in enumerate(mid_offsets, start=1):
        for nid in offset_surface(obj_id, off):
            _set_obj_layer(nid, PARENT_LAYER + "::{}V".format(i))
            created.append(nid)

    # F2 oppure F(2n)
    fname = "F2" if glass_count == 1 else "F{}".format(2 * glass_count)
    for nid in offset_surface(obj_id, close_offset):
        _set_obj_layer(nid, PARENT_LAYER + "::" + fname)
        created.append(nid)

    rs.EnableRedraw(True)

    if created:
        rs.SelectObjects(created)
        print("OK – creati {} oggetti.".format(len(created)))
    else:
        print("Offset fallito.")

if __name__ == "__main__":
    main()
