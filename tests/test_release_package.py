import hashlib
import json
from pathlib import Path
import tempfile
import unittest
import zipfile

from doc_guard import DocumentIntegrityError
from package_release import ROOT_FILES, build_release


class ReleasePackageTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        names = [*ROOT_FILES, "tests/benchmark_cases.json", "tests/kashan_template.py", "docs/PROFESSIONAL_UPGRADE.md",
                 "docs/STUDENT_EXPERIENCE.md",
                 "docs/KASHAN_PROFILE.md",
                 "docs/DEPLOYMENT.md",
                 "docs/benchmark_results.json", ".github/workflows/precision_tests.yml",
                 ".github/workflows/keep_alive.yml"]
        for name in names:
            path = self.root / name
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text("synthetic", encoding="utf-8")
        (self.root / ".env.example").write_text("BALE_TOKEN=\n", encoding="utf-8")

    def test_release_excludes_credentials_runtime_documents_and_backups(self):
        (self.root / ".env").write_text("SYNTHETIC_PRIVATE_TOKEN")
        (self.root / "private.docx").write_bytes(b"private")
        (self.root / "secrets.py").write_text("private")
        backup = self.root / "upgrade_backup"; backup.mkdir()
        (backup / "private.py").write_text("private")
        destination = build_release(self.root, self.root / "release.zip")
        with zipfile.ZipFile(destination) as archive:
            names = set(archive.namelist())
            self.assertFalse(names & {".env", "private.docx", "secrets.py", "upgrade_backup/private.py"})
            manifest = json.loads(archive.read("release-manifest.json"))
            for name, entry in manifest["files"].items():
                self.assertEqual(entry["sha256"], hashlib.sha256(archive.read(name)).hexdigest())

    def test_release_rejects_a_token_in_example_environment(self):
        (self.root / ".env.example").write_text("BALE_TOKEN=SYNTHETIC_SECRET\n")
        with self.assertRaises(ValueError):
            build_release(self.root, self.root / "release.zip")
        self.assertFalse((self.root / "release.zip").exists())

    def test_release_cannot_replace_a_source_file(self):
        with self.assertRaises(DocumentIntegrityError):
            build_release(self.root, self.root / "README.md")
        self.assertEqual("synthetic", (self.root / "README.md").read_text())
