# Ermine — Production dashboard

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
