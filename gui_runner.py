import threading
import queue
import PySimpleGUI as sg
from pathlib import Path
import importlib

# Importiere dein vorhandenes Script als Modul — passe Namen an falls nötig
backend = importlib.import_module("marxists_to_docx")

# GUI-Layout
sg.theme("SystemDefault")
layout = [
    [sg.Text("URLs (eine pro Zeile):")],
    [sg.Multiline(key="-URLS-", size=(80,8))],
    [sg.Text("Ausgabe-Dateiname:"), sg.Input("marxists.docx", key="-OUTPUT-", size=(30,1))],
    [sg.Frame("Formatierung", [
        [sg.Text("Body font size (pt):"), sg.Input(str(backend.Options().body_font_size), key="-FONTSIZE-", size=(6,1))],
        [sg.Checkbox("Farben beibehalten", key="-KEEP_COLORS-", default=backend.Options().keep_colors)],
        [sg.Checkbox("Zitate kursiv", key="-ITALIC_QUOTES-", default=backend.Options().italic_quotes)],
        [sg.Text("Zeilenabstand:"), sg.Input(str(backend.Options().line_spacing), key="-LINE_SPACING-", size=(6,1))],
        [sg.Checkbox("Blocksatz", key="-JUSTIFY-", default=backend.Options().justify)],
        [sg.Text("Fußnoten-Größe (pt):"), sg.Input(str(backend.Options().footnote_font_size), key="-FN_SIZE-", size=(6,1))],
        [sg.Checkbox("Fußnoten kursiv", key="-FN_ITALIC-", default=backend.Options().footnote_italic)],
    ])],
    [sg.Button("Ausführen", key="-RUN-"), sg.Button("Stop", key="-STOP-"), sg.Button("Reset", key="-RESET-")],
    [sg.Multiline(key="-LOG-", size=(100,20), autoscroll=True, disabled=True)],
]

window = sg.Window("Marxists.org -> .docx", layout, finalize=True)

task_thread = None
task_queue = queue.Queue()
stop_event = threading.Event()

def main_task(urls, output_name, opts: backend.Options, out_q: queue.Queue, stop_evt: threading.Event):
    try:
        # Adapted from your main(): use backend.fetch / extract_footnotes / etc.
        if not urls:
            out_q.put("Keine URLs angegeben, Abbruch.")
            return
        output_path = Path(output_name).with_suffix(".docx")
        out_q.put(f"Ziel: {output_path}")

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
                out_q.put("Abbruch angefordert. Stoppe.")
                return
            out_q.put(f"[{idx+1}/{len(urls)}] Lade {url} ...")
            try:
                soup = backend.fetch(url)
            except Exception as e:
                out_q.put(f"  Fehler beim Laden: {e}")
                continue
            footnotes = backend.extract_footnotes(soup)
            blocks = backend.find_content_paragraphs(soup)
            out_q.put(f"  {len(blocks)} Block-Elemente, {len(footnotes)} Fußnoten gefunden.")
            if idx > 0:
                document.add_page_break()
            used_ids: dict[str, int] = {}
            for block in blocks:
                if stop_evt.is_set():
                    out_q.put("Abbruch angefordert. Stoppe.")
                    return
                if block.name in ("h1", "h2"):
                    key = block.get_text(strip=True).lower()
                    if key in seen_top_headings:
                        continue
                    seen_top_headings.add(key)
                backend.render_element(document, block, fn_mgr, footnotes, used_ids, opts, quote_state)
            quote_state[0] = False

        document.save(output_path)
        out_q.put(f"Fertig: {output_path.resolve()}")
    except Exception as e:
        out_q.put(f"FEHLER: {e}")

def start_task(urls_text, output_name, values):
    urls = [u.strip() for u in urls_text.splitlines() if u.strip()]
    # build Options
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
    try:
        while True:
            msg = task_queue.get_nowait()
            window["-LOG-"].update(msg + "\n", append=True)
    except queue.Empty:
        pass

    if event == sg.WINDOW_CLOSED:
        break

    if event == "-RUN-":
        if task_thread and task_thread.is_alive():
            window["-LOG-"].update("Task läuft bereits.\n", append=True)
            continue
        task_thread = start_task(values["-URLS-"], values["-OUTPUT-"], values)
        task_thread.start()
        window["-LOG-"].update("Task gestartet...\n", append=True)

    if event == "-STOP-":
        if task_thread and task_thread.is_alive():
            stop_event.set()
            window["-LOG-"].update("Stop angefordert.\n", append=True)
        else:
            window["-LOG-"].update("Kein laufender Task.\n", append=True)

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

window.close()
