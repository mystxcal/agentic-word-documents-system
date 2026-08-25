from __future__ import annotations

import os
import shutil
from pathlib import Path

from .build import build_document
from .current import update_current_from_build
from .diagnostics import DiagnosticBag
from .errors import ReleaseGateError
from .model import BuildMode, GateState, ResolvedDocument
from .reporting import write_json
from .resolver import file_hash
from .source_lock import capture_source_lock, verify_source_lock
from .provenance import collect_runtime_provenance


def _copy_atomic(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_suffix(target.suffix + ".tmp")
    shutil.copy2(source, temporary)
    os.replace(temporary, target)


def _open_gates(resolved: ResolvedDocument) -> list[str]:
    return [
        gate
        for gate in resolved.profile.release_gates
        if resolved.manifest.release.gates.get(gate) not in {GateState.MET, GateState.NOT_APPLICABLE}
    ]


def release_document(
    resolved: ResolvedDocument,
    diagnostics: DiagnosticBag,
    *,
    retain_intermediates: bool = False,
) -> dict:
    open_gates = _open_gates(resolved)
    if open_gates:
        raise ReleaseGateError(
            "Controlled release refused; unresolved gates: " + ", ".join(open_gates)
        )

    build = build_document(
        resolved,
        diagnostics,
        mode=BuildMode.RELEASE,
        update_current=False,
        retain_intermediates=retain_intermediates,
    )
    quality = build.get("quality", {})
    if not quality.get("release_ready", quality.get("passed", False)):
        raise ReleaseGateError(
            f"Release candidate {build['build_id']} was retained as an immutable build but failed automated quality checks"
        )
    if build.get("diagnostics") and any(item.get("severity") == "error" for item in build["diagnostics"]):
        raise ReleaseGateError(
            f"Release candidate {build['build_id']} was retained as an immutable build but contains error diagnostics"
        )

    release_root = (
        resolved.system_root
        / "releases"
        / resolved.manifest.id
        / resolved.manifest.metadata.revision
        / build["build_id"]
    )
    if release_root.exists():
        raise ReleaseGateError(f"Release directory already exists: {release_root}")
    release_root.parent.mkdir(parents=True, exist_ok=True)
    stage = release_root.with_name(f".{release_root.name}.tmp")
    if stage.exists():
        raise ReleaseGateError(f"Release staging directory already exists: {stage}")
    stage.mkdir()
    docx = Path(build["artifacts"]["docx"])
    pdf = Path(build["artifacts"]["pdf"])
    try:
        _copy_atomic(docx, stage / docx.name)
        _copy_atomic(pdf, stage / pdf.name)
        if file_hash(stage / docx.name) != build["artifacts"]["docx_sha256"]:
            raise ReleaseGateError("Release DOCX failed destination hash verification")
        if file_hash(stage / pdf.name) != build["artifacts"]["pdf_sha256"]:
            raise ReleaseGateError("Release PDF failed destination hash verification")
        build_root = Path(build.get("run_directory") or docx.parent)
        resolved_inputs = build_root / "resolved-inputs.json"
        if not resolved_inputs.is_file():
            raise ReleaseGateError(
                f"Release candidate has no complete resolved-input record: {resolved_inputs}"
            )
        _copy_atomic(resolved_inputs, stage / "resolved-inputs.json")
        source_lock = capture_source_lock(
            resolved.system_root,
            stage / "resolved-inputs.json",
            stage / "source-lock.json",
        )
        lock_verification = verify_source_lock(resolved.system_root, stage / "source-lock.json")
        if not lock_verification["valid"]:
            raise ReleaseGateError("Release source lock failed verification")
        source_lock["manifest"] = str(release_root / "source-lock.json")
        source_lock["resolved_inputs"] = str(release_root / "resolved-inputs.json")
        source_lock["verification"] = lock_verification | {
            "manifest": str(release_root / "source-lock.json")
        }
        runtime_path = stage / "runtime-provenance.json"
        write_json(runtime_path, collect_runtime_provenance())
        runtime_hash = file_hash(runtime_path)
        release_report = build | {
            "release_directory": str(release_root),
            "release_artifacts": {
                "docx": str(release_root / docx.name),
                "pdf": str(release_root / pdf.name),
            },
            "source_lock": source_lock,
            "runtime_provenance": {
                "path": str(release_root / "runtime-provenance.json"),
                "sha256": runtime_hash,
            },
        }
        write_json(stage / "release-report.json", release_report)
        os.replace(stage, release_root)
    except Exception:
        shutil.rmtree(stage, ignore_errors=True)
        raise
    try:
        current_update = update_current_from_build(
            build,
            resolved.system_root / "current" / resolved.manifest.id,
            basename=resolved.manifest.outputs.basename,
        )
        release_report["current_updated"] = True
        release_report["current_docx"] = current_update["current_docx"]
        release_report["current_pdf"] = current_update["current_pdf"]
        release_report["current_update"] = current_update
    except Exception as exc:
        release_report["current_updated"] = False
        release_report["current_update_error"] = str(exc)
    write_json(release_root / "release-report.json", release_report)
    return release_report
