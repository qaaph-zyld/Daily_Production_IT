import os
import sys
import json
from datetime import datetime

# Resolve project root (this script is in scripts/)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(SCRIPT_DIR)

# Ensure we can import pvs_server
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from pvs_server import compute_metrics  # type: ignore


def main() -> None:
    print("[STATIC] Computing metrics via pvs_server.compute_metrics()...")
    res = compute_metrics()
    if not res.get("success"):
        raise SystemExit("compute_metrics() returned success = False")

    snapshot = {
        "success": True,
        "date": res.get("date"),
        "rows": res.get("rows", []),
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "totals": res.get("totals", {}),
        "group_totals": res.get("group_totals", []),
        "olk_totals": res.get("olk_totals", {}),
    }

    tmpl_path = os.path.join(ROOT_DIR, "templates", "pvs_static.html")
    if not os.path.exists(tmpl_path):
        raise SystemExit(f"Static template not found: {tmpl_path}")

    with open(tmpl_path, "r", encoding="utf-8") as f:
        template = f.read()

    json_blob = json.dumps(snapshot, ensure_ascii=False)
    html = template.replace("__PVS_SNAPSHOT_JSON__", json_blob)

    out_dir = os.path.join(ROOT_DIR, "netlify_static")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "index.html")

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"[STATIC] Wrote Netlify snapshot: {out_path}")

    extra_path = r"\\a265m050\\Groups\\Production\\Baza\\HTML\\PVS.html"
    try:
        extra_dir = os.path.dirname(extra_path)
        os.makedirs(extra_dir, exist_ok=True)
        with open(extra_path, "w", encoding="utf-8") as f2:
            f2.write(html)
        print(f"[STATIC] Wrote TV snapshot: {extra_path}")
    except Exception as e:
        print(f"[STATIC] WARNING: Could not write TV snapshot to {extra_path}: {e}")


if __name__ == "__main__":
    main()
