# marxists-to-docx

Ein kleines Python-Tool, das Texte von [marxists.org](https://www.marxists.org) in eine `.docx`-Datei konvertiert und dabei die oft eigenwillig verlinkten Fußnoten in **echte Word-Fußnoten** umwandelt.

Da Bücher auf marxists.org meist in einzelne Kapitel-Seiten aufgeteilt sind, lassen sich mehrere URLs nacheinander angeben und werden zu einem zusammenhängenden Dokument verbunden.

## Funktionen

- Erkennt das marxists.org-Schema für Fußnoten (`<a href="#noteN">[N]</a>` im Text, `<p class="note">` am Seitenende) und wandelt es in saubere Word-Fußnoten.
- Mehrere URLs (Kapitel) werden in **einer** `.docx` zusammengeführt, getrennt durch Seitenumbrüche.
- Überschriften (h1–h4) bleiben als Word-Überschriftenformate erhalten.
- Inline-Formatierung wie **fett** und *kursiv* wird übernommen.
- Mehrzeilige Fußnoten (z.B. mit Anmerkungen aus mehreren Ausgaben) werden korrekt zusammengeführt.

## Installation

Python 3.10 oder neuer wird benötigt.

### Debian / Ubuntu

Auf modernen Debian-/Ubuntu-Versionen (Debian 12+, Ubuntu 23.04+) ist die system-weite Installation von Python-Paketen via `pip` standardmäßig gesperrt (PEP 668). Daher wird ein virtuelles Environment empfohlen:

```bash
sudo apt update
sudo apt install python3 python3-venv python3-pip git

git clone https://github.com/mgmbmnb/marxists-to-docx.git
cd marxists-to-docx

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Beim nächsten Mal nur noch `source .venv/bin/activate` im Projektordner ausführen, dann `python marxists_to_docx.py` starten.

### macOS

Python 3 ist normalerweise vorinstalliert; ansonsten via [Homebrew](https://brew.sh):

```bash
brew install python git

git clone https://github.com/<DEIN-USERNAME>/marxists-to-docx.git
cd marxists-to-docx

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Windows

Python von [python.org](https://www.python.org/downloads/) installieren (Häkchen bei „Add Python to PATH" setzen). Dann in PowerShell:

```powershell
git clone https://github.com/<DEIN-USERNAME>/marxists-to-docx.git
cd marxists-to-docx

python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Falls PowerShell die Aktivierung blockiert, einmalig:
```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

### Ohne virtuelles Environment

Wenn du kein venv benutzen möchtest (nicht empfohlen, auf Debian 12+ ggf. nicht möglich):

```bash
pip install -r requirements.txt
```

## Benutzung

In VS Code öffnen und mit **F5** / „Run Python File" starten — oder im Terminal:

```bash
python marxists_to_docx.py
```

Das Script fragt dann nach:

1. **URLs** — eine pro Zeile, leere Zeile beendet die Eingabe.
2. **Ausgabe-Dateiname** — Default ist `marxists.docx`.

### Beispiel

```
URL 1: https://www.marxists.org/deutsch/archiv/marx-engels/1848/manifest/1-bourprol.htm
URL 2: https://www.marxists.org/deutsch/archiv/marx-engels/1848/manifest/2-prolkomm.htm
URL 3:

Ausgabe-Dateiname [marxists.docx]: manifest.docx
```

Ergebnis: `manifest.docx` mit beiden Kapiteln und allen Fußnoten.

## Schriftart ändern

Die Schriftart ist fest auf **Times New Roman** eingestellt und wird nicht abgefragt. Sie lässt sich in `marxists_to_docx.py` in **Zeile 34** anpassen:

```python
DEFAULT_FONT = "Times New Roman"
```

Einfach den Wert durch den Namen einer auf deinem System installierten Schriftart ersetzen, z.B. `"Calibri"`, `"Garamond"` oder `"Georgia"`.

## Hinweise

- Manche Fußnoten auf marxists.org werden im selben Text mehrfach referenziert. In diesem Fall wird der Fußnoten-Text wiederholt, da Word das saubere Wiederverwenden einer Fußnote nicht direkt unterstützt.
- Das Script erwartet die typische marxists.org-HTML-Struktur. Bei stark abweichenden Seitenlayouts (z.B. sehr alte oder sehr neue Beiträge) kann die Erkennung der Fußnoten ggf. fehlschlagen.
- Andere Hyperlinks (z.B. Querverweise auf andere Werke) werden als reiner Text übernommen.

## Lizenz

[Unlicense](LICENSE) — Public Domain. Mach damit, was du willst.
