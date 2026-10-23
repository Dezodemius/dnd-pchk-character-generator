"""Validate a self-contained Agent Skill and exercise the real portable export."""
from pathlib import Path
import re
import subprocess
import sys

import yaml

ROOT = Path(__file__).resolve().parents[1]
skill = ROOT / "pchk-character"
source = (skill / "SKILL.md").read_text(encoding="utf-8")
match = re.match(r"\A---\n(.*?)\n---\n", source, re.S)
assert match, "Missing YAML frontmatter"
metadata = yaml.safe_load(match[1])
assert metadata["name"] == "pchk-character"
assert 0 < len(metadata["description"]) <= 1024
assert not re.search(r"(?:[A-Za-z]:[\\/]|/mnt/|/Users/)", source), "Machine-specific path in skill instructions"
for relative in ["scripts/pchk.py", "scripts/portable_pdf.py", "assets/character-sheet.pptx", "assets/example.draft.json", "requirements.txt", "references/draft-format.md"]:
    assert (skill / relative).is_file(), relative
for path in skill.rglob("*.py"):
    compile(path.read_text(encoding="utf-8"), str(path), "exec")
result = subprocess.run([sys.executable, "-X", "utf8", "-m", "unittest", "discover", "-s", "tests", "-v"], cwd=ROOT)
raise SystemExit(result.returncode)
