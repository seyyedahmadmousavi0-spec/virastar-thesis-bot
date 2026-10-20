import json
from pathlib import Path
import tempfile
import unittest

from analysis_model import Issue
from doc_guard import DocumentIntegrityError, atomic_write, safe_filename
from report_artifacts import csv_report, html_report, report_data, write_artifacts
from rule_registry import RuleConfigurationError, validate_rules
import checker


class ReportArtifactTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.rules = checker.load_rules()

    def data(self):
        issue = Issue("TEST", message="<script>alert(1)</script>", target_text="=SUM(A1:A2)",
                      char_start=0, char_end=11, location_status="exact", suggestion='</script><img src=x onerror=alert(1)>')
        return report_data([issue], {"words": 5, "private_path": "C:/secret/source.docx"}, self.rules)

    def test_html_escapes_source_and_has_no_external_assets(self):
        result = html_report(self.data())
        self.assertIn("&lt;script&gt;alert(1)&lt;/script&gt;", result)
        self.assertNotIn("<img src=x", result)
        self.assertNotIn("https://", result)
        self.assertIn("connect-src 'none'", result)

    def test_json_excludes_arbitrary_filesystem_metadata(self):
        data = self.data()
        self.assertNotIn("private_path", data["stats"])
        self.assertFalse(data["findings"][0]["auto_applied"])
        self.assertEqual("2.0", data["schema_version"])

    def test_csv_neutralizes_spreadsheet_formula(self):
        self.assertIn("'=SUM(A1:A2)", csv_report(self.data()))

    def test_artifact_publication_rejects_input_collision_before_any_write(self):
        source = self.root / "report.json"
        source.write_bytes(b"original")
        with self.assertRaises(DocumentIntegrityError):
            write_artifacts(self.root, [], {}, self.rules, source_path=source)
        self.assertEqual(b"original", source.read_bytes())
        self.assertFalse((self.root / "report.html").exists())

    def test_artifacts_round_trip_and_stable_finding_id(self):
        issue = Issue("TEST", message="test", block_id="one")
        paths = write_artifacts(self.root, [issue], {}, self.rules, text_report="test")
        self.assertEqual({"html", "csv", "json", "txt"}, set(paths))
        data = json.loads(paths["json"].read_text(encoding="utf-8"))
        again = report_data([issue], {}, self.rules)
        self.assertEqual(data["findings"][0]["finding_id"], again["findings"][0]["finding_id"])

    def test_atomic_failure_preserves_existing_file(self):
        from unittest.mock import patch
        destination = self.root / "existing"
        destination.write_bytes(b"old")
        with patch("os.replace", side_effect=OSError("synthetic")):
            with self.assertRaises(OSError):
                atomic_write(destination, b"new")
        self.assertEqual(b"old", destination.read_bytes())
        self.assertEqual([destination], list(self.root.iterdir()))

    def test_filename_avoids_devices_and_path_traversal(self):
        self.assertEqual("_CON.docx", safe_filename("../../CON.docx"))
        self.assertNotIn("/", safe_filename("../safe.docx"))
        self.assertLessEqual(len(safe_filename("a" * 300 + ".docx")), 176)

    def test_invalid_configuration_fails_explicitly(self):
        import copy
        for mutate in (
            lambda r: r.update(strict="false"),
            lambda r: r.update(protected_terms=[""]),
            lambda r: r["page"].update(width_cm=float("nan")),
            lambda r: r["body"].update(fa_fonts=[]),
            lambda r: r.update(max_pdf_pages=True),
        ):
            rules = copy.deepcopy(self.rules)
            mutate(rules)
            with self.assertRaises(RuleConfigurationError):
                validate_rules(rules)

    def test_actual_report_script_search_filter_and_print_without_browser(self):
        import re
        import shutil
        import subprocess
        node = shutil.which("node")
        if node is None:
            self.skipTest("Node is unavailable; browser script test is optional")
        script = re.search(r"<script>([\s\S]*?)</script>", html_report(self.data())).group(1)
        harness = r'''
const assert=require('node:assert/strict'),vm=require('node:vm');
const input=()=>({value:'',handlers:{},addEventListener(name,fn){this.handlers[name]=fn;}});
const query=input(),level=input(),print=input(),count={},empty={};level.value='default';
const cards=[{dataset:{confidence:'A'},textContent:'alpha target'}, {dataset:{confidence:'B'},textContent:'beta review'}, {dataset:{confidence:'C'},textContent:'gamma ambiguous'}];
const elements={query,level,print,count,empty};let prints=0;
const document={getElementById:id=>elements[id],querySelectorAll:()=>cards};
vm.runInNewContext(SOURCE,{document,window:{print(){prints++;}}});
assert.equal(cards[0].hidden,false);assert.equal(cards[1].hidden,false);assert.equal(cards[2].hidden,true);
query.value='target';query.handlers.input();assert.equal(cards.filter(c=>!c.hidden).length,1);
query.value='';level.value='C';level.handlers.change();assert.equal(cards[2].hidden,false);assert.equal(cards[0].hidden,true);
level.value='all';level.handlers.change();assert.equal(cards.filter(c=>!c.hidden).length,3);
query.value='not-present';query.handlers.input();assert.equal(empty.hidden,false);
print.handlers.click();assert.equal(prints,1);
'''.replace("SOURCE", json.dumps(script))
        result = subprocess.run([node, "-e", harness], capture_output=True, text=True, timeout=10)
        self.assertEqual(0, result.returncode, result.stderr)
