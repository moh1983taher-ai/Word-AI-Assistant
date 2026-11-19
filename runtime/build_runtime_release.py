#!/usr/bin/env python3
"""Repack the verified DuckAI runtime with the repository's recovery scripts."""
from __future__ import annotations

import argparse
import hashlib
import re
import shutil
import tempfile
import zipfile
from pathlib import Path, PurePosixPath

REQUIRED_BASE = {
    "BUILD-INFO.txt",
    "ResearchTools-DuckAgent.exe",
    "ResearchTools-Launcher.vbs",
    "duck2api-main.exe",
    "manifest.xml",
    "uninstall.ps1",
}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def safe_extract(archive_path: Path, destination: Path) -> None:
    with zipfile.ZipFile(archive_path, "r") as archive:
        for member in archive.infolist():
            name = member.filename.replace("\\", "/")
            path = PurePosixPath(name)
            if path.is_absolute() or any(part in ("", ".", "..") for part in path.parts):
                raise ValueError(f"Unsafe ZIP entry: {member.filename!r}")
            target = destination.joinpath(*path.parts)
            resolved = target.resolve()
            if destination.resolve() not in resolved.parents and resolved != destination.resolve():
                raise ValueError(f"ZIP entry escapes staging directory: {member.filename!r}")
        archive.extractall(destination)


def verify_internal_sums(directory: Path, *, require_bootstrap: bool) -> None:
    sums_file = directory / "SHA256SUMS.txt"
    if not sums_file.is_file():
        raise ValueError("Baseline archive is missing SHA256SUMS.txt")

    entries: dict[str, str] = {}
    for line in sums_file.read_text(encoding="utf-8-sig").splitlines():
        if not line.strip():
            continue
        match = re.fullmatch(r"\s*([A-Fa-f0-9]{64})\s+(.+?)\s*", line)
        if not match:
            raise ValueError(f"Malformed SHA256SUMS line: {line!r}")
        digest, name = match.group(1).upper(), match.group(2).strip()
        if Path(name).name != name or "/" in name or "\\" in name:
            raise ValueError(f"Unexpected path in SHA256SUMS: {name!r}")
        if name in entries:
            raise ValueError(f"Duplicate SHA256SUMS entry: {name}")
        entries[name] = digest

    required = set(REQUIRED_BASE)
    if require_bootstrap:
        required.add("DuckAI-Bootstrap.ps1")
    missing = required - entries.keys()
    if missing:
        raise ValueError(f"SHA256SUMS missing required files: {sorted(missing)}")

    for name, expected in entries.items():
        path = directory / name
        if not path.is_file():
            raise ValueError(f"SHA256SUMS references missing file: {name}")
        actual = sha256(path)
        if actual != expected:
            raise ValueError(f"SHA-256 mismatch for {name}: expected {expected}, got {actual}")

    present = {p.name for p in directory.iterdir() if p.is_file()}
    missing_required_files = required - present
    if missing_required_files:
        raise ValueError(f"Archive missing required files: {sorted(missing_required_files)}")


def write_internal_sums(directory: Path) -> None:
    lines = []
    for path in sorted((p for p in directory.iterdir() if p.is_file() and p.name != "SHA256SUMS.txt"), key=lambda p: p.name.casefold()):
        lines.append(f"{sha256(path)}  {path.name}")
    (directory / "SHA256SUMS.txt").write_text("\n".join(lines) + "\n", encoding="ascii", newline="\n")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, type=Path, help="Verified v1.0.0 runtime ZIP")
    parser.add_argument("--runtime-dir", required=True, type=Path, help="Repository runtime/ directory")
    parser.add_argument("--output", required=True, type=Path, help="Output ZIP path")
    args = parser.parse_args()

    for filename in ("ResearchTools-Launcher.vbs", "DuckAI-Bootstrap.ps1"):
        source = args.runtime_dir / filename
        if not source.is_file():
            raise FileNotFoundError(f"Missing repository runtime file: {source}")

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="duckai-runtime-build-") as temp_name:
        stage = Path(temp_name) / "runtime"
        stage.mkdir(parents=True)
        safe_extract(args.input, stage)
        base_files = {p.name for p in stage.iterdir() if p.is_file()}
        if not REQUIRED_BASE.issubset(base_files):
            raise ValueError(f"Baseline ZIP contents unexpected; missing {sorted(REQUIRED_BASE - base_files)}")
        verify_internal_sums(stage, require_bootstrap=False)

        shutil.copy2(args.runtime_dir / "ResearchTools-Launcher.vbs", stage / "ResearchTools-Launcher.vbs")
        shutil.copy2(args.runtime_dir / "DuckAI-Bootstrap.ps1", stage / "DuckAI-Bootstrap.ps1")

        info_path = stage / "BUILD-INFO.txt"
        info = info_path.read_text(encoding="utf-8-sig")
        info = re.sub(
            r"(?m)^Bundle Revision:.*$",
            "Bundle Revision: 2026-11-19-r3 (synchronized bootstrap and launcher checksums)",
            info,
            count=1,
        )
        if "Automatic recovery:" not in info:
            info = info.rstrip() + "\n\nAutomatic recovery:\n"
            info += "The launcher verifies local runtime files before starting Word services.\n"
            info += "Missing or damaged files are restored from the latest stable GitHub Release after SHA-256 verification.\n"
            info += "The stable launcher and bootstrap script are stored outside the DuckAI runtime folder.\n"
        info_path.write_text(info, encoding="utf-8", newline="\n")

        write_internal_sums(stage)
        verify_internal_sums(stage, require_bootstrap=True)

        if args.output.exists():
            args.output.unlink()
        with zipfile.ZipFile(args.output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
            for path in sorted(stage.iterdir(), key=lambda p: p.name.casefold()):
                if path.is_file():
                    archive.write(path, arcname=path.name)

    archive_hash = sha256(args.output)
    sidecar = args.output.with_suffix(args.output.suffix + ".sha256")
    sidecar.write_text(f"{archive_hash}  {args.output.name}\n", encoding="ascii")

    print(f"Created: {args.output}")
    print(f"Size: {args.output.stat().st_size:,} bytes")
    print(f"SHA256: {archive_hash}")
    print(f"Sidecar: {sidecar}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())