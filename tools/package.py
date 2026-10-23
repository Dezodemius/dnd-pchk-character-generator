"""Create identical, deterministic ZIP and .skill archives with a single skill root."""
import argparse
import hashlib
import json
from pathlib import Path
import re
import zipfile

ROOT = Path(__file__).resolve().parents[1]
VERSION_PATTERN = r"[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+"


def validate_version(value):
    if not re.fullmatch(VERSION_PATTERN, value):
        raise ValueError("Tag must contain exactly four numeric components, e.g. 0.1.0.0 (no v prefix).")
    return value


def package(destination, version):
    validate_version(version)
    skill = ROOT / "pchk-character"
    files = {}
    for file in sorted(skill.rglob("*")):
        if not file.is_file() or "__pycache__" in file.parts or file.suffix in {".pyc", ".pyo"}:
            continue
        relative = file.relative_to(skill)
        if any(part.startswith(".") for part in relative.parts):
            continue
        files["pchk-character/" + relative.as_posix()] = file.read_bytes()
    for name in ["LICENSE", "NOTICE"]:
        files["pchk-character/" + name] = (ROOT / name).read_bytes()
    files["pchk-character/VERSION"] = (version + "\r\n").encode()
    manifest = {"version": version, "files": {name: hashlib.sha256(data).hexdigest() for name, data in sorted(files.items())}}
    files["pchk-character/MANIFEST.json"] = (json.dumps(manifest, ensure_ascii=False, indent=2) + "\n").replace("\n", "\r\n").encode("utf-8")
    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=True)
    archive = destination / "pchk-character.zip"
    with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as output:
        for name, data in sorted(files.items()):
            info = zipfile.ZipInfo(name, (2026, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o100644 << 16
            output.writestr(info, data)
    data = archive.read_bytes()
    (destination / "pchk-character.skill").write_bytes(data)
    checksums = "".join(hashlib.sha256(data).hexdigest() + "  " + name + "\r\n" for name in ["pchk-character.zip", "pchk-character.skill"])
    (destination / "SHA256SUMS").write_bytes(checksums.encode())
    print(f"Packaged {version}: {len(files)} files, {len(data)} bytes")
    return archive


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("destination", type=Path, nargs="?", default=ROOT / "dist")
    parser.add_argument("--version", required=True)
    args = parser.parse_args()
    try:
        package(args.destination, args.version)
    except ValueError as exc:
        parser.error(str(exc))
