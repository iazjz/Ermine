#!/usr/bin/env python3
"""
Package the dashboard as a folder GitHub Pages can serve, so the phone can reach it.

Why this step exists at all: OneDrive and Google Drive both refuse to render an .html
file on a phone - they treat it as a download and the phone does not know what to do with
it. We tried that route and it failed, and there is no other way to get a web page onto a
phone without somewhere that serves web pages. GitHub Pages is free, needs no server, and
turns the same file into a real https address, which is also what "Add to Home Screen"
requires: iOS will not install a file:// page.

What goes in the folder:
  index.html          the dashboard, built by build_dashboard_html.py
  manifest.webmanifest  so the home-screen icon opens it full screen as an app
  icon-180/192/512.png  the icon itself, drawn here rather than fetched
  .nojekyll           GitHub otherwise runs the folder through Jekyll, which quietly
                      ignores files beginning with an underscore
  README.md           the ten minutes of setup, once

    python3 build_site.py
"""
import glob
import os
import shutil
import sys
import unicodedata

try:
    from PIL import Image, ImageDraw
except ImportError:          # the Mac agent does not need an image library
    Image = ImageDraw = None

HERE = os.path.dirname(os.path.abspath(__file__))
SITE = "ermine-site"
BG = (31, 56, 100)          # the workbook's own header blue
FG = (255, 255, 255)


def plan_dir():
    # Only ever used to decide where the built site is written, so ask work_dir,
    # which knows whether this run is allowed to write to the planning folder.
    import ermine_env
    w = ermine_env.work_dir()
    if w:
        return w
    root = os.path.dirname(ermine_env.plan_dir() or "") or "."
    hits = [d for d in os.listdir(root)
            if "تخطيط" in unicodedata.normalize("NFC", d)
            and os.path.isdir(os.path.join(root, d))]
    if not hits:
        raise SystemExit("the planning folder is not connected to this session")
    hits.sort(key=lambda d: len(os.listdir(os.path.join(root, d))), reverse=True)
    return os.path.join(root, hits[0])


def brand_icon(path, size):
    """The crowned ermine on the brand green, copied from brand/icons/.

    Drawn once by a script rather than generated here, because the mark is the real
    logo now and not the letter E this used to draw. Green for the dashboard, purple
    for the checklist - two apps on one home screen need to be told apart at a glance,
    and the only thing different about them at icon size is the colour.
    """
    src = os.path.join(HERE, "brand", "icons", f"dash-{size}.png")
    if os.path.exists(src):
        shutil.copy2(src, path)
        return True
    return False


def icon(path, size):
    """A plain mark, drawn rather than downloaded: a rounded tile with an E on it.

    iOS does not use a manifest icon for the home screen on every version, and it does not
    accept SVG, so these are real PNGs at the three sizes that matter. Keeping the drawing
    here means the icon cannot go missing the way a linked image would.
    """
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    r = int(size * 0.22)
    d.rounded_rectangle([0, 0, size - 1, size - 1], radius=r, fill=BG)
    # the E, built from three bars and a spine so it needs no font file
    m = size * 0.26          # margin
    w = size * 0.085         # bar thickness
    d.rectangle([m, m, m + w, size - m], fill=FG)                      # spine
    for y in (m, size / 2 - w / 2, size - m - w):                      # three arms
        d.rectangle([m, y, size - m - size * 0.06, y + w], fill=FG)
    img.save(path)


MANIFEST = """{
  "name": "Ermine Production",
  "short_name": "Ermine",
  "start_url": "./index.html",
  "scope": "./",
  "display": "standalone",
  "orientation": "portrait",
  "background_color": "#014833",
  "theme_color": "#014833",
  "icons": [
    { "src": "icon-192.png", "sizes": "192x192", "type": "image/png" },
    { "src": "icon-512.png", "sizes": "512x512", "type": "image/png" },
    { "src": "icon-512.png", "sizes": "512x512", "type": "image/png",
      "purpose": "maskable" }
  ]
}
"""

README = """# Ermine — Production dashboard

One page, read from the production workbook. Open it on a phone or a laptop; the layout
changes to suit.

## Putting it online (once, about ten minutes)

1. Create a free account at **github.com** if you do not have one.
2. Click **+** (top right) → **New repository**.
   - Name: `ermine` (any name; it becomes part of the address)
   - Choose **Public**. GitHub Pages needs a paid plan to serve a private repository, and
     the address is unguessable but not secret — see the warning below.
   - Click **Create repository**.
3. On the new repository page click **uploading an existing file**.
4. Drag **everything inside this folder** into the browser window — `index.html`, the
   three icons, `manifest.webmanifest`, `.nojekyll`. Then **Commit changes**.
5. Go to **Settings → Pages**. Under *Source* choose **Deploy from a branch**,
   branch **main**, folder **/ (root)**. **Save**.
6. Wait two or three minutes. The address appears at the top of that page and looks like:

       https://<your-username>.github.io/ermine/

## On the phone

Open that address in **Safari** (iPhone) or **Chrome** (Android), then:

- iPhone: share button ⬆️ → **Add to Home Screen**
- Android: ⋮ → **Add to Home screen**

It becomes an icon that opens full screen with no browser bars.

## Updating it

The page is a snapshot of the workbook, not a live link. To refresh it:

1. Ask Claude to rebuild the dashboard. A new `index.html` lands in this folder.
2. On the repository page: **Add file → Upload files**, drag the new `index.html`,
   **Commit changes**.
3. The phone shows the new figures within a minute or two. Pull down to refresh if not.

The page says at the top how old its reading is, and the line turns amber after eighteen
hours — so a stale page announces itself rather than pretending to be current.

## One warning worth reading

A GitHub Pages site is **public to anyone who has the address**. Nobody can guess it, and
it is not listed anywhere, but it is not protected either. This page carries production
volumes, sales figures and staff counts. If that is not acceptable, say so and we will
find a different route — a password needs more than a static page can do.
"""


def main():
    out = os.path.join(plan_dir(), SITE)
    os.makedirs(out, exist_ok=True)

    src = os.path.join(HERE, "ermine-dashboard.html")
    if not os.path.exists(src):
        raise SystemExit("build the dashboard first: python3 build_dashboard_html.py")

    html = open(src, encoding="utf-8").read()
    # the manifest and the icons have to be announced in the page itself, and iOS reads
    # apple-touch-icon rather than the manifest on several versions
    if "manifest.webmanifest" not in html:
        html = html.replace("</head>",
            '<link rel="manifest" href="manifest.webmanifest">\n'
            '<link rel="apple-touch-icon" href="icon-180.png">\n'
            '<link rel="icon" href="icon-192.png">\n</head>')
    open(os.path.join(out, "index.html"), "w", encoding="utf-8").write(html)

    # A one-line file holding this build's stamp, so the page can tell whether the copy
    # running on a phone is the current one. Added to the site on 20 Oct after a day of
    # fixes that Aziz could not see: the dashboard is installed to his home screen, and a
    # home-screen web app will happily serve a cached page for hours. Every "it still
    # does not work" that day was the old file answering.
    import re as _re
    stamp = (_re.search(r'BUILT = new Date\("([^"]+)"\)', html) or [None, ""])[1]
    os.makedirs(os.path.join(out, "data"), exist_ok=True)
    open(os.path.join(out, "data", "build.txt"), "w", encoding="utf-8").write(stamp)

    for size, name in ((180, "icon-180.png"), (192, "icon-192.png"), (512, "icon-512.png")):

        dst = os.path.join(out, name)

        if os.path.exists(dst):

            continue          # already there; nothing to draw

        if not brand_icon(dst, size) and Image is not None:

            icon(dst, size)   # the drawn E, only if we can draw
    open(os.path.join(out, "manifest.webmanifest"), "w").write(MANIFEST)
    open(os.path.join(out, ".nojekyll"), "w").write("")
    open(os.path.join(out, "README.md"), "w", encoding="utf-8").write(README)

    print(f"saved {out}")
    for f in sorted(os.listdir(out)):
        print(f"  {os.path.getsize(os.path.join(out, f)):>8,}  {f}")
    return out


if __name__ == "__main__":
    main()
