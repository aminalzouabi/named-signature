"""Build the Pages artifact; never write deployment credentials into source files."""
import os
from pathlib import Path
import re
import shutil

ROOT = Path(__file__).resolve().parents[1]


def build(destination, key=""):
    key = key.strip()
    if key and not re.fullmatch(r"bk_[a-fA-F0-9]{64}", key):
        raise ValueError("Invalid Boosthis browser project key; value omitted")
    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=True)
    for name in ("signature.js", "manifest.xml"):
        shutil.copyfile(ROOT / name, destination / name)
    html = (ROOT / "autorun.html").read_text(encoding="utf-8")
    if key:
        # Keep Office.js and signature.js first. Never load this in the JS-only runtime.
        tag = ('    <script defer src="https://www.boosthis.com/kit.js"'
               f' data-key="{key}"'
               ' data-install="4c4d874f-6d93-4f5c-8e0e-c65bbbffdcd9"'
               ' data-project="Outlook automate signature" data-bubble="false"></script>\n')
        html = html.replace("</head>", tag + "</head>", 1)
    (destination / "autorun.html").write_text(html, encoding="utf-8")
    (destination / ".nojekyll").touch()


if __name__ == "__main__":
    build(ROOT / "_site", os.environ.get("BOOSTHIS_PROJECT_KEY", ""))
