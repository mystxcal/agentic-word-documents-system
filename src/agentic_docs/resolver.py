from __future__ import annotations

import hashlib
from pathlib import Path

from pydantic import BaseModel, ValidationError

from .diagnostics import DiagnosticBag
from .errors import ManifestError, ResolutionError
from .jsonc import load_jsonc
from .model import (
    DocumentManifest,
    KitManifest,
    ProfileManifest,
    ProjectManifest,
    ResolvedComponent,
    ResolvedDocument,
)


def file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def _validate(model_type, payload: object, path: Path):
    try:
        return model_type.model_validate(payload)
    except ValidationError as exc:
        details = "; ".join(
            f"{'.'.join(str(part) for part in item['loc'])}: {item['msg']}"
            for item in exc.errors(include_url=False)
        )
        raise ManifestError(f"Invalid manifest {path}: {details}") from exc


def _discover_root(manifest_path: Path) -> Path:
    for parent in [manifest_path.parent, *manifest_path.parents]:
        if (parent / "kits").is_dir() and (parent / "profiles").is_dir() and (parent / "projects").is_dir():
            return parent
    raise ResolutionError(
        f"Could not locate a V2 system root above {manifest_path}; expected kits, profiles, and projects folders"
    )


def _safe_relative(base: Path, value: str, *, allowed_root: Path, label: str) -> Path:
    path = (base / value).resolve()
    try:
        path.relative_to(allowed_root.resolve())
    except ValueError as exc:
        raise ResolutionError(f"{label} path escapes its declared root: {value}") from exc
    return path


def _report_extra(model, diagnostics: DiagnosticBag, location: str) -> None:
    def visit(value, path: str) -> None:
        if isinstance(value, BaseModel):
            extras = getattr(value, "model_extra", None) or {}
            for key in sorted(extras):
                diagnostics.warn(
                    "MANIFEST_UNKNOWN_KEY",
                    f"Unrecognized key {key!r} was retained but is not used by this V2 build",
                    location=f"{location}#{path}" if path else location,
                    hint="Check for a spelling mistake or document the extension before controlled release.",
                )
            for field_name in type(value).model_fields:
                child = getattr(value, field_name)
                visit(child, f"{path}.{field_name}" if path else field_name)
        elif isinstance(value, list):
            for index, child in enumerate(value):
                visit(child, f"{path}[{index}]")
        elif isinstance(value, dict):
            for key, child in value.items():
                visit(child, f"{path}.{key}" if path else str(key))

    visit(model, "")


def resolve_document(
    manifest_path: Path,
    diagnostics: DiagnosticBag | None = None,
    *,
    allow_stale_diagrams: bool = False,
) -> ResolvedDocument:
    diagnostics = diagnostics or DiagnosticBag()
    manifest_path = Path(manifest_path).resolve()
    root = _discover_root(manifest_path)
    document = _validate(DocumentManifest, load_jsonc(manifest_path), manifest_path)
    _report_extra(document, diagnostics, str(manifest_path))

    kit_path = root / "kits" / document.kit / "kit.jsonc"
    profile_path = root / "profiles" / document.profile / "profile.jsonc"
    project_path = root / "projects" / document.project / "project.jsonc"
    for label, path in (("kit", kit_path), ("profile", profile_path), ("project", project_path)):
        if not path.is_file():
            raise ResolutionError(f"Selected {label} manifest does not exist: {path}")

    kit = _validate(KitManifest, load_jsonc(kit_path), kit_path)
    profile = _validate(ProfileManifest, load_jsonc(profile_path), profile_path)
    project = _validate(ProjectManifest, load_jsonc(project_path), project_path)
    for model, path in ((kit, kit_path), (profile, profile_path), (project, project_path)):
        _report_extra(model, diagnostics, str(path))

    profile_root = profile_path.parent.resolve()
    shell_path = _safe_relative(profile_root, profile.shell, allowed_root=profile_root, label="profile shell")
    if not shell_path.is_file():
        raise ResolutionError(f"Profile shell does not exist: {shell_path}")

    document_root = manifest_path.parent.resolve()
    project_root = project_path.parent.resolve()

    def component_source_path(selector: str, *, label: str) -> Path:
        if selector.startswith("project-file:"):
            return _safe_relative(
                project_root,
                selector.split(":", 1)[1],
                allowed_root=project_root,
                label=label,
            )
        source_value = selector.split(":", 1)[1] if selector.startswith("file:") else selector
        return _safe_relative(
            document_root,
            source_value,
            allowed_root=document_root,
            label=label,
        )

    components: dict[str, ResolvedComponent] = {}
    for component_id, declaration in document.components.items():
        source_path = None
        source_hash = None
        related_paths = {}
        related_hashes = {}
        if declaration.source:
            source_path = component_source_path(declaration.source, label=f"component {component_id}")
            if source_path.is_file():
                source_hash = file_hash(source_path)
            elif declaration.required:
                raise ResolutionError(f"Required component source does not exist: {source_path}")
            else:
                diagnostics.warn(
                    "OPTIONAL_SOURCE_MISSING",
                    f"Optional component {component_id!r} is unavailable",
                    location=str(source_path),
                )
        if declaration.type.value == "diagram":
            rendition_selector = declaration.options.rendition
            rendition_path = component_source_path(
                rendition_selector,
                label=f"component {component_id} reviewed rendition",
            )
            if not rendition_path.is_file():
                raise ResolutionError(f"Diagram rendition does not exist: {rendition_path}")
            if rendition_path.suffix.lower() not in {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff"}:
                raise ResolutionError(
                    f"Diagram rendition must be a reviewed raster image for Word publication: {rendition_path}"
                )
            if source_hash != declaration.options.rendition_of_sha256:
                message = (
                    f"Diagram {component_id!r} changed after its reviewed rendition was created. "
                    f"Export and review a fresh image, then run: Document-System.cmd accept-rendition "
                    f"{document.id} --component {component_id} --file <image>"
                )
                if allow_stale_diagrams:
                    diagnostics.warn("DIAGRAM_RENDITION_STALE", message, location=str(source_path))
                else:
                    raise ResolutionError(message)
            related_paths["rendition"] = rendition_path
            related_hashes["rendition"] = file_hash(rendition_path)
        components[component_id] = ResolvedComponent(
            id=component_id,
            declaration=declaration,
            source_path=source_path,
            source_hash=source_hash,
            related_paths=related_paths,
            related_hashes=related_hashes,
        )

    kit_root = kit_path.parent.resolve()

    def selector_path(selector: str | None, *, label: str) -> Path | None:
        if selector is None or selector == "none":
            return None
        if selector.startswith("kit:"):
            name = selector.split(":", 1)[1]
            if name not in kit.components:
                raise ResolutionError(f"{label} selects unknown kit component {name!r}")
            result = _safe_relative(kit_root, kit.components[name], allowed_root=kit_root, label=label)
        elif selector.startswith("document:"):
            name = selector.split(":", 1)[1]
            if name not in components:
                raise ResolutionError(f"{label} selects unknown document component {name!r}")
            result = components[name].source_path
            if result is None:
                raise ResolutionError(f"{label} selects document component {name!r} without a source")
        elif selector.startswith("file:"):
            value = selector.split(":", 1)[1]
            result = _safe_relative(document_root, value, allowed_root=document_root, label=label)
        else:
            raise ResolutionError(
                f"{label} selector must use kit:, document:, file:, or none; got {selector!r}"
            )
        if not result.is_file():
            raise ResolutionError(f"Selected {label} file does not exist: {result}")
        return result

    presentation_paths: dict[str, Path | None] = {
        "styles": selector_path(document.presentation.styles, label="presentation styles"),
        "cover": selector_path(document.presentation.cover, label="presentation cover"),
    }
    for region_id, region in document.presentation.page_regions.items():
        for kind in ("header", "footer"):
            value = getattr(region, kind)
            if isinstance(value, str) or value is None:
                presentation_paths[f"region.{region_id}.{kind}.default"] = selector_path(
                    value, label=f"region {region_id} {kind}"
                )
            else:
                for variant in ("default", "first", "even"):
                    presentation_paths[f"region.{region_id}.{kind}.{variant}"] = selector_path(
                        getattr(value, variant), label=f"region {region_id} {kind} {variant}"
                    )

    return ResolvedDocument(
        manifest_path=manifest_path,
        system_root=root,
        manifest=document,
        kit_path=kit_path,
        kit=kit,
        profile_path=profile_path,
        profile=profile,
        project_path=project_path,
        project=project,
        shell_path=shell_path,
        components=components,
        presentation_paths=presentation_paths,
    )
