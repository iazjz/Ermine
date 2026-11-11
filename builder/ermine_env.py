#!/usr/bin/env python3
"""
Where things are, on whichever machine this is running.

Until now every script assumed the Linux sandbox and hard-coded /sessions/*/mnt/...
That was fine while the only thing that ever ran them was me. The agent changes it:
the same scripts now run on Aziz's Mac, on a schedule, with nobody watching. So the
two roots are resolved once, here, and everything else asks.

Order of resolution, most explicit first:
  1. the environment - ERMINE_ONEDRIVE and ERMINE_PLAN, set by the agent
  2. the sandbox mounts, when they exist
  3. the usual macOS locations
"""
import glob
import os
import unicodedata


def _first_dir(*paths):
    for p in paths:
        if p and os.path.isdir(p):
            return p
    return None


def onedrive():
    """The OneDrive folder that holds Work Centre."""
    return _first_dir(
        os.environ.get("ERMINE_ONEDRIVE"),
        *sorted(glob.glob("/sessions/*/mnt/OneDrive-Personal")),
        os.path.expanduser("~/Library/CloudStorage/OneDrive-Personal"),
        os.path.expanduser("~/OneDrive"),
    )


def plan_dir():
    """The planning folder - the one with an Arabic name and 'w Claude' inside it."""
    p = os.environ.get("ERMINE_PLAN")
    if p and os.path.isdir(p):
        return p
    roots = sorted(glob.glob("/sessions/*/mnt")) + [
        os.path.expanduser("~/Desktop/2026/Ermine"),
        os.path.expanduser("~/Desktop"),
    ]
    best = None
    for root in roots:
        if not os.path.isdir(root):
            continue
        for d in os.listdir(root):
            full = os.path.join(root, d)
            if not os.path.isdir(full):
                continue
            if "تخطيط" in unicodedata.normalize("NFC", d):
                # more than one can match across roots; the one with most files wins,
                # which is how the sandbox copy was told from a stray empty folder
                n = len(os.listdir(full))
                if best is None or n > best[0]:
                    best = (n, full)
    return best[1] if best else None


def work_dir():
    """Where the build may WRITE - the staged site, mostly.

    Normally the planning folder, beside everything else. But the scheduled agent is
    not allowed to write to the Desktop, so it is given a scratch folder of its own
    and told about it here. Reading Aziz's files and writing our own scratch are two
    different permissions; they deserve two different answers.
    """
    w = os.environ.get("ERMINE_WORK")
    if w:
        os.makedirs(w, exist_ok=True)
        return w
    return plan_dir()


def config_file():
    """github.config.txt, wherever the planning folder is."""
    # The scheduled agent runs without permission to read the Desktop, so install
    # leaves a copy of the token beside the runtime. Prefer it when it is there.
    beside = os.path.join(os.path.dirname(os.path.abspath(__file__)), "github.config.txt")
    if os.path.exists(beside):
        return beside
    p = plan_dir()
    if not p:
        return None
    for name in ("github.config.txt", "github.config"):
        full = os.path.join(p, "w Claude", name)
        if os.path.exists(full):
            return full
    return None


def describe():
    return {"onedrive": onedrive(), "plan": plan_dir(), "config": config_file()}


if __name__ == "__main__":
    for k, v in describe().items():
        print(f"{k:<10}{v or '-- not found --'}")
