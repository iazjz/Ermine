#!/usr/bin/env python3
"""
Find Aziz's working files wherever he has put them.

Every script here used to hard-code `OneDrive-Personal/<name>.xlsx`. On 20 October he
moved three of them into a `Work Centre` subfolder - a perfectly reasonable thing to do
with your own filing - and every script stopped finding them at once. The dashboard went
quiet and kept showing the previous day, which is a bad failure: it looked like current
information rather than like an error.

So the name is now what identifies a file, not the path. Search the whole of OneDrive,
skip the backup folder and Excel's lock files, and if more than one copy turns up take
the one most recently written - the live file is by definition the one being used.
"""
import glob
import os
import unicodedata

SKIP_DIRS = {"backup", "archive", "old", ".git"}


def onedrive():
    # One place decides where things are; see ermine_env.
    import ermine_env
    return ermine_env.onedrive()


def find(name, root=None):
    """The newest file called `name` anywhere under OneDrive, or None."""
    root = root or onedrive()
    if not root or not os.path.isdir(root):
        return None
    want = unicodedata.normalize("NFC", name).lower()
    found = []
    denied = []

    # os.walk swallows errors by default. When macOS refused the agent access to
    # OneDrive the walk returned nothing at all, and the build reported "could not
    # find the live workbook" - which reads as "the file was moved" and sent us
    # looking in entirely the wrong place. A refusal is not an absence; say which.
    def refused(err):
        if getattr(err, "errno", None) in (1, 13):
            denied.append(err.filename)

    for base, dirs, files in os.walk(root, onerror=refused):
        dirs[:] = [d for d in dirs if d.lower() not in SKIP_DIRS and not d.startswith(".")]
        for f in files:
            if f.startswith("~$"):
                continue
            if unicodedata.normalize("NFC", f).lower() == want:
                found.append(os.path.join(base, f))
    if denied:
        raise PermissionError(
            1,
            f"macOS refused this process access to {denied[0]}. "
            "A scheduled job has no file permissions of its own: grant Full Disk "
            "Access to /bin/zsh, or run this from Terminal instead.",
        )
    if not found:
        return None
    return max(found, key=os.path.getmtime)


def locked(path):
    """True when Excel has the file open - its lock file sits beside it."""
    if not path:
        return False
    return os.path.exists(os.path.join(os.path.dirname(path), "~$" + os.path.basename(path)))


LIVE      = "Production Plan - LIVE.xlsx"
WAREHOUSE = "Warehouse - Ali Nawaz.xlsx"
MATERIALS = "Raw Materials - Ali Nawaz.xlsx"
TASKS     = "Tasks - Aziz.xlsx"

if __name__ == "__main__":
    for n in (LIVE, WAREHOUSE, MATERIALS, TASKS):
        p = find(n)
        print(f"{n:<34} {'-- not found --' if not p else p}{'  [OPEN IN EXCEL]' if locked(p) else ''}")
