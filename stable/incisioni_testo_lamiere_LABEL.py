# -*- coding: utf-8 -*-

import rhinoscriptsyntax as rs
import scriptcontext as sc
import Rhino
import os
import Eto.Forms as forms
import Eto.Drawing as drawing
import System.Drawing as sd



class TextInputDialog(forms.Dialog[bool]):

    def __init__(self):
        self.Title = "Insert Text"
        self.Padding = drawing.Padding(10)
        self.Resizable = False

        self.textbox = forms.TextBox()
        self.textbox.KeyDown += self.on_key_down

        ok_button = forms.Button(Text="OK")
        ok_button.Click += self.on_ok

        layout = forms.DynamicLayout()
        layout.Spacing = drawing.Size(5, 5)

        layout.Add(forms.Label(Text="Paste text:"))
        layout.Add(self.textbox)
        layout.Add(ok_button)

        self.Content = layout

    def on_key_down(self, sender, e):
        if e.Key == forms.Keys.Enter:
            self.Close(True)

    def on_ok(self, sender, e):
        self.Close(True)


def _make_group_name_from_text(text):
    """
    Rhino 7 di solito digerisce anche #.
    Puliamo solo caratteri rognosi (a capo, tab) e accorciamo se troppo lungo.
    """
    if text is None:
        return "TXT"

    t = text.strip()
    if not t:
        return "TXT"

    # niente multilinea / tab: lo trasformo in spazi
    t = t.replace("\r\n", " ").replace("\n", " ").replace("\r", " ").replace("\t", " ")
    # comprimi spazi multipli
    while "  " in t:
        t = t.replace("  ", " ")

    # limita lunghezza (nomi troppo lunghi sono solo rogne)
    if len(t) > 40:
        t = t[:40].rstrip()

    return "TXT_" + t


def create_text_from_external_library():

    spacing = 5.0  # spaziatura fissa 5 mm

    # --- percorso hardcoded della libreria ---
    file_path = r"C:\Automazione\FONT_INCISIONI_LIBRERIA.3DM"
    if not os.path.exists(file_path):
        print("File libreria non trovato:", file_path)
        return

    # --- FINESTRA TESTO ---
    dialog = TextInputDialog()
    result = dialog.ShowModal(Rhino.UI.RhinoEtoApp.MainWindow)
    if not result:
        return

    text = dialog.textbox.Text
    if not text:
        return

    # Punto di inserimento
    insert_point = rs.GetPoint("Select insertion point")
    if not insert_point:
        return

    # Layer di destinazione (con colore fissato)
    target_layer = "CODICI"
    layer_color = sd.Color.FromArgb(205, 205, 0)  # <-- scegli qui l'RGB

    if not rs.IsLayer(target_layer):
        rs.AddLayer(target_layer, layer_color)
    else:
        rs.LayerColor(target_layer, layer_color)


    # Carica il modello della libreria
    lib_model = Rhino.FileIO.File3dm.Read(file_path)
    if not lib_model:
        print("Errore nella lettura del file libreria")
        return

    current_x = 0.0
    created_ids = []

    for char in text:

        if char == " ":
            current_x += spacing * 2
            continue

        layer_name = char.upper()

        # Trova l'indice del layer nella libreria
        layer_index = -1
        for i, layer in enumerate(lib_model.Layers):
            if layer.Name == layer_name:
                layer_index = i
                break

        if layer_index == -1:
            print("Layer non trovato nella libreria:", layer_name)
            current_x += spacing
            continue

        # Duplica gli oggetti della lettera
        letter_objects = []
        for obj in lib_model.Objects:
            if obj.Attributes.LayerIndex == layer_index:
                letter_objects.append(obj)

        if not letter_objects:
            current_x += spacing
            continue

        new_ids = []
        for obj in letter_objects:
            geometry = obj.Geometry.Duplicate()
            new_id = sc.doc.Objects.Add(geometry)
            if new_id:
                new_ids.append(new_id)
                created_ids.append(new_id)

        if not new_ids:
            continue

        # Imposta il layer degli oggetti creati
        for nid in new_ids:
            rs.ObjectLayer(nid, target_layer)

        # Bounding box della lettera
        bbox = rs.BoundingBox(new_ids)
        if not bbox:
            continue

        min_x = min(pt.X for pt in bbox)
        max_x = max(pt.X for pt in bbox)
        min_y = min(pt.Y for pt in bbox)

        width = max_x - min_x

        # --- correzione verticale per il trattino ---
        vertical_offset = 0.0
        if char == "-":
            vertical_offset = 7  # regola se vuoi alzarlo di più

        # Sposta la lettera nella posizione corretta
        move_vector = (
            current_x - min_x,
            -min_y + vertical_offset,
            0
        )
        rs.MoveObjects(new_ids, move_vector)

        current_x += width + spacing

    # Sposta tutto il testo al punto scelto
    if created_ids:
        bbox_total = rs.BoundingBox(created_ids)
        if bbox_total:
            total_min_x = min(pt.X for pt in bbox_total)
            total_min_y = min(pt.Y for pt in bbox_total)

            final_move = (
                insert_point.X - total_min_x,
                insert_point.Y - total_min_y,
                insert_point.Z
            )
            rs.MoveObjects(created_ids, final_move)

        # --- RAGGRUPPA TUTTO ---
        group_name = _make_group_name_from_text(text)
        group = rs.AddGroup(group_name)
        if group:
            rs.AddObjectsToGroup(created_ids, group)
            rs.SelectObjects(created_ids)

    sc.doc.Views.Redraw()
    print("Testo creato con successo!")


if __name__ == "__main__":
    create_text_from_external_library()
