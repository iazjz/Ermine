#!/usr/bin/env python3
"""
Rebuild the dashboard and publish it to GitHub Pages in one step.

Before this, every refresh meant Aziz opening the repository in a browser and dragging a
new index.html into it. The page is a snapshot - Excel on OneDrive exposes nothing a web
page can read, so there is no live link to be had - but the UPLOAD did not have to be
manual, and GitHub turns out to be reachable from here.

Credentials are read from a file, never from the command line and never printed. The
token goes in the planning folder, which is Aziz's own machine and is not the folder the
chef or the storekeeper see:

    ملفات تخطيط الإنتاج/w Claude/github.config

    repo = aziz-username/ermine
    token = github_pat_...
    branch = main          # optional, defaults to main

The token wants to be a fine-grained one, scoped to that single repository, with
Contents: Read and write and nothing else. That is the least that can publish a file, and
it can be revoked from GitHub's settings at any moment without touching anything here.

    python3 publish_dashboard.py [--no-build] [--dry-run]
"""
import glob
import os
import re
import shutil
import subprocess
import sys
import unicodedata

HERE = os.path.dirname(os.path.abspath(__file__))
NO_BUILD = "--no-build" in sys.argv
DRY = "--dry-run" in sys.argv


def plan_dir():
    import ermine_env
    p = ermine_env.plan_dir()
    if not p:
        raise SystemExit("the planning folder could not be found")
    return p


def config():
    # On the GitHub runner (the cloud refresh, 11 Nov 2026) there is no config file: the
    # workflow hands over its own short-lived token and the repository name.
    if os.environ.get("ERMINE_GH_TOKEN") and os.environ.get("ERMINE_GH_REPO"):
        return {"repo": os.environ["ERMINE_GH_REPO"], "token": os.environ["ERMINE_GH_TOKEN"],
                "branch": os.environ.get("ERMINE_GH_BRANCH", "main")}
    # github.config.txt first, and .config only as a fallback. macOS has no application
    # registered for a bare .config extension, so double-clicking it does nothing at all -
    # Aziz could not open his own credentials file to rotate the token in it. The name the
    # Finder can open is the one that should win.
    # ermine_env decides, because it knows about the agent's own copy. Working the
    # path out here instead sent the scheduled run straight at the Desktop, which it
    # is not allowed to read: PermissionError on a file sitting beside it in its own
    # folder. Asking twice in two different ways is how that went unnoticed.
    import ermine_env
    path = ermine_env.config_file()
    folder = os.path.join(plan_dir() or "", "w Claude")
    if path is None:
        raise SystemExit(
            "github.config.txt is not there yet. Create it at:\n"
            f"  {os.path.join(folder, 'github.config.txt')}\n"
            "with two lines:\n"
            "  repo = your-username/your-repo\n"
            "  token = github_pat_...")
    conf = {}
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            conf[k.strip().lower()] = v.strip()
    for k in ("repo", "token"):
        if not conf.get(k):
            raise SystemExit(f"github.config has no '{k}' line")
    if not re.fullmatch(r"[\w.-]+/[\w.-]+", conf["repo"]):
        raise SystemExit(f"'repo' should look like owner/name, not {conf['repo']!r}")
    conf.setdefault("branch", "main")
    return conf


BUILDER_CODE = ("build_dashboard_html.py", "build_site.py", "dashboard_template.html",
                "publish_dashboard.py", "ermine_env.py", "ermine_paths.py",
                "history_fixes.json", "sales_flavours.json", "prices.json", "qc_api.txt")


def builder_files():
    """What the cloud refresh needs to build - never the token, logs or built pages."""
    for name in BUILDER_CODE:
        if os.path.isfile(os.path.join(HERE, name)):
            yield name
    for base, dirs, files in os.walk(os.path.join(HERE, "brand")):
        dirs[:] = [d for d in dirs if not d.startswith(".")]
        for f in files:
            if not f.startswith("."):
                yield os.path.relpath(os.path.join(base, f), HERE)


def hide(text, token):
    """Never let the token reach the transcript, not even inside a git error."""
    return text.replace(token, "***") if token else text


def run(cmd, cwd=None, token=""):
    p = subprocess.run(cmd, cwd=cwd, shell=isinstance(cmd, str),
                       capture_output=True, text=True)
    if p.returncode:
        raise SystemExit(hide((p.stderr or p.stdout).strip(), token))
    return p.stdout


def main():
    conf = config()
    repo, token, branch = conf["repo"], conf["token"], conf["branch"]

    # The clone comes FIRST now. The flavour history - the only thing in this system
    # that accumulates rather than being recomputed - lives in the repository, because
    # the repository is the one place both this machine and the GitHub runner can see.
    # The builder has to find it before it builds, not after.
    work = "/tmp/ermine_publish"
    shutil.rmtree(work, ignore_errors=True)
    url = f"https://x-access-token:{token}@github.com/{repo}.git"
    print(f"cloning {repo} ({branch})")
    run(["git", "clone", "--depth", "1", "--branch", branch, url, work], token=token)
    history = os.path.join(work, "data", "flavours.json")

    if not NO_BUILD:
        print("rebuilding the dashboard from the live workbook")
        for script in ("build_dashboard_html.py", "build_site.py"):
            # ulimit -v exists to stop LibreOffice eating the Linux sandbox alive.
            # macOS does not implement it and prints a warning instead, so only ask
            # for it where it means something.
            cap = "ulimit -v 2500000 2>/dev/null; " if sys.platform != "darwin" else ""
            out = run(f'cd "{HERE}" && {cap}'
                      f'ERMINE_HISTORY="{history}" python3 {script}', token=token)
            print("  " + out.strip().splitlines()[-1])

    import ermine_env
    site = os.path.join(ermine_env.work_dir(), "ermine-site")
    if not os.path.isdir(site):
        raise SystemExit("ermine-site is not built yet - run without --no-build")

    # copy the built site over the clone, leaving .git alone
    changed = []
    if os.path.exists(history):
        # written in place by the builder; `git add -A` picks it up
        changed.append("data/flavours.json")
    # The frozen pay-run forecasts live beside it and are written the same way. They
    # are the only figures in the system that must survive a rebuild unchanged, so
    # they go into the repository like the flavour history rather than being
    # recomputed - a forecast you can revise afterwards is not a forecast.
    runs = os.path.join(os.path.dirname(history), "runs.json")
    if os.path.exists(runs):
        changed.append("data/runs.json")
    # Subdirectories used to be skipped outright, which silently left data/build.txt -
    # the file the page reads to notice it is stale - sitting in the build folder and
    # never reaching the site. One level down is enough; the site has no deeper tree.
    def files_to_copy():
        for name in sorted(os.listdir(site)):
            full = os.path.join(site, name)
            if os.path.isfile(full):
                yield name
            elif os.path.isdir(full):
                for sub in sorted(os.listdir(full)):
                    if os.path.isfile(os.path.join(full, sub)):
                        yield os.path.join(name, sub)

    for name in files_to_copy():
        src, dst = os.path.join(site, name), os.path.join(work, name)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        same = (os.path.exists(dst)
                and open(src, "rb").read() == open(dst, "rb").read())
        if not same:
            shutil.copy2(src, dst)
            changed.append(name)

    # The cloud refresh builds from builder/ in the repository. A publish from the Mac
    # carries the Mac's code there, so the two can never drift apart the way the old
    # launchd copy did (frozen on 22 Oct). The runner itself never rewrites builder/.
    if not os.environ.get("GITHUB_ACTIONS"):
        for name in builder_files():
            src, dst = os.path.join(HERE, name), os.path.join(work, "builder", name)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            if not (os.path.exists(dst)
                    and open(src, "rb").read() == open(dst, "rb").read()):
                shutil.copy2(src, dst)
                changed.append("builder/" + name)
    if not changed:
        print("nothing changed since the last publish - the live page is already current")
        return
    print("changed:", ", ".join(changed))

    if DRY:
        print(f"dry run: {work} is staged but nothing was pushed")
        return

    run(["git", "-C", work, "config", "user.email", "aziz@ermine.local"], token=token)
    run(["git", "-C", work, "config", "user.name", "Ermine dashboard"], token=token)
    run(["git", "-C", work, "add", "-A"], token=token)
    run(["git", "-C", work, "commit", "-m",
         "Refresh dashboard from the production workbook"], token=token)
    run(["git", "-C", work, "push", "origin", branch], token=token)
    shutil.rmtree(work, ignore_errors=True)

    owner, name = repo.split("/")
    print("published. GitHub Pages rebuilds in a minute or two:")
    print(f"  https://{owner}.github.io/{name}/")


if __name__ == "__main__":
    main()
