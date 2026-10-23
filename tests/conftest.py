"""Shared fixtures.

The scripts in scripts/ and dashboard/ are plain files, not a package, so they
are loaded by path. Every script only does work inside main(), so importing
one has no side effects.
"""
import importlib.util
import shutil
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def load_script(rel_path):
    """Import a script such as 'scripts/clean_sales_data.py' as a module."""
    path = ROOT / rel_path
    name = "script_" + path.stem
    if name in sys.modules:
        return sys.modules[name]
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="session")
def scripts():
    return {
        "gen": load_script("scripts/generate_synthetic_data.py"),
        "clean": load_script("scripts/clean_sales_data.py"),
        "analysis": load_script("scripts/sales_variance_analysis.py"),
        "insights": load_script("scripts/build_insights.py"),
        "dashboard": load_script("dashboard/build_variance_dashboard.py"),
    }


@pytest.fixture(scope="session")
def pipeline(tmp_path_factory, scripts):
    """Run the whole pipeline once, in a temporary folder, through each script's main().

    Nothing in the repository is written. Returns the paths and exit codes.
    """
    from fpa_lib import config

    base = tmp_path_factory.mktemp("pipeline")
    raw_dir, proc_dir, out_dir, docs_dir = (base / d for d in ("raw", "processed", "outputs", "docs"))
    for d in (raw_dir, proc_dir, out_dir, docs_dir):
        d.mkdir()
    readme = base / "README.md"
    shutil.copyfile(ROOT / "README.md", readme)

    p = {
        "base": base,
        "raw": raw_dir / config.RAW_FILE.name,
        "manifest": raw_dir / config.DEFECT_MANIFEST_FILE.name,
        "clean_csv": proc_dir / config.CLEAN_CSV.name,
        "clean_xlsx": proc_dir / config.CLEAN_XLSX.name,
        "enriched": proc_dir / config.ENRICHED_CSV.name,
        "analysis_xlsx": out_dir / config.ANALYSIS_XLSX.name,
        "insights": docs_dir / config.INSIGHTS_MD.name,
        "metrics": docs_dir / config.KEY_METRICS_JSON.name,
        "readme": readme,
        "html": base / config.DASHBOARD_HTML.name,
    }
    codes = {}
    codes["gen"] = scripts["gen"].main(["--output-dir", str(raw_dir)])
    codes["clean"] = scripts["clean"].main(["--input", str(p["raw"]), "--output-dir", str(proc_dir)])
    codes["analysis"] = scripts["analysis"].main(["--input", str(p["clean_csv"]), "--output-dir", str(proc_dir),
                                                  "--workbook-dir", str(out_dir)])
    codes["insights"] = scripts["insights"].main(["--enriched", str(p["enriched"]), "--workbook", str(p["analysis_xlsx"]),
                                                  "--cleaning-workbook", str(p["clean_xlsx"]),
                                                  "--docs-dir", str(docs_dir), "--readme", str(readme)])
    codes["dashboard"] = scripts["dashboard"].main(["--enriched", str(p["enriched"]), "--workbook", str(p["analysis_xlsx"]),
                                                    "--cleaning-workbook", str(p["clean_xlsx"]),
                                                    "--insights", str(p["insights"]), "--metrics", str(p["metrics"]),
                                                    "--output", str(p["html"])])
    p["codes"] = codes
    return p
