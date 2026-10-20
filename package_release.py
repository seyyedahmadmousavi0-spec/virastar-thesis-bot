"""Create a local source release from an explicit allowlist, excluding runtime data."""
import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
import tempfile
import zipfile

from doc_guard import ensure_distinct_output
from report_artifacts import PRODUCT_VERSION

ROOT_FILES = (
    "checker.py", "academic_checker.py", "analysis_model.py", "bot.py", "bot_runtime.py",
    "doc_guard.py", "document_blocks.py", "en_checker.py", "pdf_handler.py", "protected_content.py",
    "report_artifacts.py", "review_cli.py", "rule_registry.py", "run_benchmark.py", "safe_annotator.py",
    "tums_checker.py", "web_runner.py", "package_release.py", "rules.json", "requirements.txt",
    "README.md", ".env.example", ".dockerignore", "Dockerfile", "virastar-bot.service",
    "run.bat", "run.sh", "run_hidden.vbs", "install_startup.bat", "uninstall_startup.bat",
    "Procfile", "liara.json",
)


def build_release(root, destination):
    root = Path(root).resolve()
    destination = Path(destination).absolute()
    sources = [root / name for name in ROOT_FILES]
    sources += sorted((root / "profiles").glob("*.json"))
    sources += sorted((root / "tests").glob("test_*.py"))
    sources += [root / "tests/benchmark_cases.json", root / "docs/PROFESSIONAL_UPGRADE.md",
                root / "docs/DEPLOYMENT.md",
                root / "docs/benchmark_results.json", root / ".github/workflows/precision_tests.yml"]
    for source in sources:
        if not source.is_file() or not source.resolve().is_relative_to(root):
            raise ValueError("Release source is missing or escapes the project: " + source.name)
        ensure_distinct_output(source, destination)
    example = (root / ".env.example").read_text(encoding="utf-8")
    if re.search(r"^[ \t]*BALE_TOKEN[ \t]*=[ \t]*[^\s#]+", example, re.MULTILINE):
        raise ValueError("The example environment contains a token; refusing to package")
    destination.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary = tempfile.mkstemp(prefix=".release_", suffix=".zip", dir=destination.parent)
    os.close(descriptor)
    try:
        manifest = {"version": PRODUCT_VERSION, "generated_at": datetime.now(timezone.utc).isoformat(),
                    "excludes": ["credentials", "runtime documents", "Git history", "backups", "virtual environments"],
                    "files": {}}
        with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as archive:
            for source in sorted(sources):
                data = source.read_bytes()
                relative = source.relative_to(root).as_posix()
                manifest["files"][relative] = {"bytes": len(data), "sha256": hashlib.sha256(data).hexdigest()}
                archive.writestr(relative, data)
            archive.writestr("release-manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
        # Check the archive before making it the visible deliverable.
        with zipfile.ZipFile(temporary) as archive:
            if archive.testzip() is not None:
                raise ValueError("Release archive failed integrity verification")
        os.replace(temporary, destination)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
    return destination


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", default=f"dist/virastar-professional-{PRODUCT_VERSION}.zip")
    args = parser.parse_args()
    print(build_release(Path(__file__).parent, args.output))
