

import io
import queue
import threading
import PySimpleGUI as sg
from pathlib import Path
import importlib
from PIL import Image
import os

from pathlib import Path

BASE_DIR = Path(__file__).parent
ASSETS_DIR = BASE_DIR / "assets"
DEFAULT_IMAGE = ASSETS_DIR / "nice.jpg"

try:
    img_path = DEFAULT_IMAGE
    if img_path.exists():
        with Image.open(img_path) as img:
            print(f"✅ Bild OK: {img.format}, {img.size}")
    else:
        print(f"⚠️ Bild nicht gefunden: {img_path}")
except Exception as e:
    print(f"❌ Bild Problem: {e}")


# Importiere dein vorhandenes Script als Modul
backend = importlib.import_module("marxists_to_docx")

# Moderneres Theme
sg.theme("DarkBlue3")

# Definiere Farben
COLOR_BG = "#2b2b2b"
COLOR_FRAME = "#1e1e1e"
COLOR_ACCENT = "#0078d4"
COLOR_ERROR = "#d13438"
COLOR_TEXT = "#ffffff"

# GUI-Layout
layout = [
    [sg.Text("📥 URLs eingeben (eine pro Zeile):", text_color=COLOR_TEXT, font=("Segoe UI", 11, "bold"))],
    [sg.Multiline(key="-URLS-", size=(90, 8), background_color="#3c3c3c", text_color=COLOR_TEXT)],

    [sg.Text("💾 Ausgabe-Dateiname:", text_color=COLOR_TEXT, font=("Segoe UI", 11, "bold"))],
    [sg.Input("marxists.docx", key="-OUTPUT-", size=(45, 1), background_color="#3c3c3c", text_color=COLOR_TEXT)],

    [sg.Frame("⚙️ Formatierung", [
        [sg.Column([
            [sg.Text("Body Font Größe (pt):", text_color=COLOR_TEXT), sg.Input(str(backend.Options().body_font_size), key="-FONTSIZE-", size=(6,1), background_color="#3c3c3c", text_color=COLOR_TEXT)],
            [sg.Text("Zeilenabstand:", text_color=COLOR_TEXT), sg.Input(str(backend.Options().line_spacing), key="-LINE_SPACING-", size=(6,1), background_color="#3c3c3c", text_color=COLOR_TEXT)],
            [sg.Text("Fußnoten-Größe (pt):", text_color=COLOR_TEXT), sg.Input(str(backend.Options().footnote_font_size), key="-FN_SIZE-", size=(6,1), background_color="#3c3c3c", text_color=COLOR_TEXT)],
        ]), sg.Column([
            [sg.Checkbox("🎨 Farben beibehalten", key="-KEEP_COLORS-", default=backend.Options().keep_colors, text_color=COLOR_TEXT)],
            [sg.Checkbox("📖 Zitate kursiv", key="-ITALIC_QUOTES-", default=backend.Options().italic_quotes, text_color=COLOR_TEXT)],
            [sg.Checkbox("✓ Blocksatz", key="-JUSTIFY-", default=backend.Options().justify, text_color=COLOR_TEXT)],
            [sg.Checkbox("📝 Fußnoten kursiv", key="-FN_ITALIC-", default=backend.Options().footnote_italic, text_color=COLOR_TEXT)],
        ])]
    ], title_color=COLOR_TEXT, background_color=COLOR_FRAME)],

    [sg.Text("Fortschritt:", text_color=COLOR_TEXT, font=("Segoe UI", 10, "bold"))],
    [sg.ProgressBar(max_value=100, orientation="h", size=(90, 25), key="-PROGRESS-", bar_color=(COLOR_ACCENT, COLOR_FRAME))],
    [sg.Text("", key="-STATUS-", text_color=COLOR_ACCENT, font=("Segoe UI", 10, "bold"))],

    [sg.Text("Status:", text_color=COLOR_TEXT, font=("Segoe UI", 10, "bold"))],
    [sg.Multiline(key="-LOG-", size=(90, 8), background_color="#3c3c3c", text_color=COLOR_TEXT, autoscroll=True, disabled=True)],

    [sg.Button("▶ Ausführen", key="-RUN-", button_color=(COLOR_TEXT, COLOR_ACCENT), size=(15, 1)),
     sg.Button("⏹ Stop", key="-STOP-", button_color=(COLOR_TEXT, COLOR_ERROR), size=(15, 1)),
     sg.Button("↺ Reset", key="-RESET-", button_color=(COLOR_TEXT, "#5c5c5c"), size=(15, 1)),
     sg.Button("❌ Beenden", key="-EXIT-", button_color=(COLOR_TEXT, COLOR_ERROR), size=(15, 1))],
]

window = sg.Window("📄 Marxists.org → .docx Converter", layout, finalize=True, background_color=COLOR_BG)

task_thread = None
task_queue = queue.Queue()
stop_event = threading.Event()

# Comletion Bild

# Completion-Popup: ersetze vorhandene Funktionen durch diese eine Definition
def show_completion_popup(output_path):
    """Zeigt ein Popup-Fenster mit Erfolgs-Nachricht (Bild wenn möglich)."""
    layout_popup = [
        [sg.Text("✅ Erfolgreich abgeschlossen!", font=("Segoe UI", 14, "bold"), text_color="#107c10")],
        [sg.Text("")],
        [sg.Text("Die Datei wurde erfolgreich erstellt:", font=("Segoe UI", 11))],
        [sg.Text(str(output_path), font=("Segoe UI", 10, "italic"), text_color=COLOR_ACCENT)],
        [sg.Text("")],
    ]

    try:
        img_bytes = None
        if DEFAULT_IMAGE.exists():
            with Image.open(DEFAULT_IMAGE) as im:
                bio = io.BytesIO()
                im.convert("RGBA").save(bio, format="PNG")
                img_bytes = bio.getvalue()
        if img_bytes:
            layout_popup.append([sg.Image(data=img_bytes, size=(600, 600))])
        else:
            layout_popup.append([sg.Text("🎉 🎊 🎉", font=("Segoe UI", 60))])
    except Exception as e:
        print(f"⚠️ Bild konnte nicht geladen werden: {e}")
        layout_popup.append([sg.Text("🎉 🎊 🎉", font=("Segoe UI", 60))])

    layout_popup.append([sg.Text("")])
    layout_popup.append([sg.Button("OK", size=(15, 1), button_color=(COLOR_TEXT, "#107c10"))])

    popup_window = sg.Window("✅ Abgeschlossen", layout_popup, finalize=True, background_color=COLOR_BG, modal=True)
    while True:
        event, values = popup_window.read()
        if event == sg.WINDOW_CLOSED or event == "OK":
            break
    popup_window.close()


def main_task(urls, output_name, opts: backend.Options, out_q: queue.Queue, stop_evt: threading.Event):
    try:
        if not urls:
            out_q.put(("log", "❌ Keine URLs angegeben, Abbruch."))
            return

        output_path = Path(output_name).with_suffix(".docx")
        out_q.put(("log", f"📍 Ziel: {output_path}"))
        out_q.put(("status", "Vorbereitung..."))
        out_q.put(("progress", 0))

        from docx import Document
        from docx.shared import Pt
        from docx.oxml.ns import qn

        document = Document()
        normal = document.styles["Normal"]
        DEFAULT_FONT = getattr(backend, "DEFAULT_FONT", "Times New Roman")
        normal.font.name = DEFAULT_FONT
        normal.font.size = Pt(opts.body_font_size)
        rpr = normal.element.get_or_add_rPr()
        rFonts = rpr.find(qn("w:rFonts"))
        if rFonts is None:
            from docx.oxml import OxmlElement
            rFonts = OxmlElement("w:rFonts")
            rpr.insert(0, rFonts)
        rFonts.set(qn("w:ascii"), DEFAULT_FONT)
        rFonts.set(qn("w:hAnsi"), DEFAULT_FONT)
        rFonts.set(qn("w:cs"), DEFAULT_FONT)

        fn_mgr = backend.FootnoteManager(document)
        seen_top_headings: set[str] = set()
        quote_state = [False]

        for idx, url in enumerate(urls):
            if stop_evt.is_set():
                out_q.put(("log", "⛔ Abbruch angefordert. Stoppe."))
                out_q.put(("status", "Abgebrochen"))
                return

            progress = int((idx / len(urls)) * 100)
            out_q.put(("progress", progress))
            out_q.put(("status", f"Verarbeite {idx + 1}/{len(urls)}: {url[:50]}..."))
            out_q.put(("log", f"🔄 [{idx+1}/{len(urls)}] Lade {url} ..."))

            try:
                soup = backend.fetch(url)
            except Exception as e:
                out_q.put(("log", f"  ⚠️ Fehler beim Laden: {e}"))
                continue

            footnotes = backend.extract_footnotes(soup)
            blocks = backend.find_content_paragraphs(soup)
            out_q.put(("log", f"  ✓ {len(blocks)} Blöcke, {len(footnotes)} Fußnoten gefunden."))

            if idx > 0:
                document.add_page_break()

            used_ids: dict[str, int] = {}
            for block in blocks:
                if stop_evt.is_set():
                    out_q.put(("log", "⛔ Abbruch angefordert. Stoppe."))
                    out_q.put(("status", "Abgebrochen"))
                    return
                if block.name in ("h1", "h2"):
                    key = block.get_text(strip=True).lower()
                    if key in seen_top_headings:
                        continue
                    seen_top_headings.add(key)
                backend.render_element(document, block, fn_mgr, footnotes, used_ids, opts, quote_state)
            quote_state[0] = False

        out_q.put(("progress", 100))
        out_q.put(("status", "Speichern..."))
        document.save(output_path)
        out_q.put(("log", f"✅ Fertig: {output_path.resolve()}"))
        out_q.put(("status", "✅ Erfolgreich abgeschlossen!"))
        out_q.put(("popup", str(output_path.resolve())))
    except Exception as e:
        out_q.put(("log", f"❌ FEHLER: {e}"))
        out_q.put(("status", f"❌ Fehler: {str(e)[:50]}..."))

def start_task(urls_text, output_name, values):
    urls = [u.strip() for u in urls_text.splitlines() if u.strip()]
    try:
        body_font_size = float(values["-FONTSIZE-"])
    except Exception:
        body_font_size = backend.Options().body_font_size
    try:
        line_spacing = float(values["-LINE_SPACING-"])
    except Exception:
        line_spacing = backend.Options().line_spacing
    try:
        fn_size = float(values["-FN_SIZE-"])
    except Exception:
        fn_size = backend.Options().footnote_font_size

    opts = backend.Options(
        keep_colors=values["-KEEP_COLORS-"],
        italic_quotes=values["-ITALIC_QUOTES-"],
        body_font_size=body_font_size,
        line_spacing=line_spacing,
        justify=values["-JUSTIFY-"],
        footnote_font_size=fn_size,
        footnote_italic=values["-FN_ITALIC-"],
    )

    stop_event.clear()
    t = threading.Thread(target=main_task, args=(urls, output_name, opts, task_queue, stop_event), daemon=True)
    return t

# Event loop
while True:
    event, values = window.read(timeout=100)

    # Queue-Nachrichten verarbeiten
    try:
        while True:
            msg_type, msg_data = task_queue.get_nowait()
            if msg_type == "log":
                window["-LOG-"].update(msg_data + "\n", append=True)
            elif msg_type == "progress":
                window["-PROGRESS-"].update(msg_data)
            elif msg_type == "status":
                window["-STATUS-"].update(msg_data)
            elif msg_type == "popup":
                show_completion_popup(msg_data)
    except queue.Empty:
        pass

    if event == sg.WINDOW_CLOSED or event == "-EXIT-":
        if task_thread and task_thread.is_alive():
            stop_event.set()
            task_thread.join(timeout=2)
        break

    if event == "-RUN-":
        if task_thread and task_thread.is_alive():
            window["-LOG-"].update("⚠️ Task läuft bereits.\n", append=True)
            continue
        if not values["-URLS-"].strip():
            window["-LOG-"].update("⚠️ Bitte URLs eingeben!\n", append=True)
            continue
        task_thread = start_task(values["-URLS-"], values["-OUTPUT-"], values)
        task_thread.start()
        window["-LOG-"].update("▶️ Task gestartet...\n", append=True)

    if event == "-STOP-":
        if task_thread and task_thread.is_alive():
            stop_event.set()
            window["-LOG-"].update("⛔ Stop angefordert.\n", append=True)
        else:
            window["-LOG-"].update("⚠️ Kein laufender Task.\n", append=True)

    if event == "-RESET-":
        window["-URLS-"].update("")
        window["-OUTPUT-"].update("marxists.docx")
        window["-FONTSIZE-"].update(str(backend.Options().body_font_size))
        window["-KEEP_COLORS-"].update(backend.Options().keep_colors)
        window["-ITALIC_QUOTES-"].update(backend.Options().italic_quotes)
        window["-LINE_SPACING-"].update(str(backend.Options().line_spacing))
        window["-JUSTIFY-"].update(backend.Options().justify)
        window["-FN_SIZE-"].update(str(backend.Options().footnote_font_size))
        window["-FN_ITALIC-"].update(backend.Options().footnote_italic)
        window["-LOG-"].update("")
        window["-PROGRESS-"].update(0)
        window["-STATUS-"].update("")

window.close()
