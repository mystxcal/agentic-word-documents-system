from __future__ import annotations

from datetime import date
from enum import StrEnum
from pathlib import Path
from typing import Annotated, Any, Literal
import re

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


class OpenModel(BaseModel):
    """Typed known fields while retaining unfamiliar real-world extensions."""

    model_config = ConfigDict(extra="allow", populate_by_name=True)


class BuildMode(StrEnum):
    DRAFT = "draft"
    RELEASE = "release"


class GateState(StrEnum):
    OPEN = "open"
    MET = "met"
    NOT_APPLICABLE = "not_applicable"


SAFE_IDENTIFIER = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*$")


def _require_safe_identifier(value: str) -> str:
    if not SAFE_IDENTIFIER.fullmatch(value):
        raise ValueError("must contain only letters, numbers, dot, underscore, or hyphen and cannot start with punctuation")
    return value


class Ownership(StrEnum):
    SOURCE = "source"
    WORD_FRAGMENT = "word_fragment"
    SNAPSHOT = "snapshot"


class ComponentType(StrEnum):
    COVER = "cover"
    TOC = "toc"
    DOCUMENT = "document"
    TABLE = "table"
    FIGURE = "figure"
    DIAGRAM = "diagram"
    PDF_PAGES = "pdf_pages"
    PAGE_BREAK = "page_break"


class StrictOptions(BaseModel):
    """Typed component controls with one deliberate forward-extension point."""

    model_config = ConfigDict(extra="forbid")
    extensions: dict[str, Any] = Field(default_factory=dict)

    def get(self, key: str, default: Any = None) -> Any:
        """Keep adapters mapping-friendly while options become strongly typed."""

        return getattr(self, key, default)


class EmptyOptions(StrictOptions):
    pass


class DocumentOptions(StrictOptions):
    preserve_sections: bool = False
    whole_document: bool = False
    strict: bool = True
    style_roles: dict[str, str] = Field(default_factory=dict)
    table_style_role: str = "technical"


class ExcelLocator(StrictOptions):
    table: str | None = None
    sheet: str | None = None
    range: str | None = None
    has_headers: bool = True
    headers: list[str] | None = None

    @model_validator(mode="after")
    def exactly_one_locator(self):
        if self.table:
            if self.range:
                raise ValueError("Excel locator.table cannot be combined with locator.range")
            return self
        if not self.sheet or not self.range:
            raise ValueError("Excel locator requires locator.table or both locator.sheet and locator.range")
        return self


class TableColumn(StrictOptions):
    source: str
    heading: str | None = None
    width: float | None = None
    align: Literal["left", "center", "right"] | None = None
    header_align: Literal["left", "center", "right"] | None = None
    format: str | None = None

    @field_validator("width")
    @classmethod
    def positive_width(cls, value: float | None) -> float | None:
        if value is not None and value <= 0:
            raise ValueError("table column width weight must be greater than zero")
        return value


class TableFilter(StrictOptions):
    column: str
    op: Literal["eq", "ne", "in", "not_in", "contains", "not_blank", "blank"] = "eq"
    value: Any = None


class TableSort(StrictOptions):
    column: str
    direction: Literal["asc", "desc"] = "asc"


class TableView(StrictOptions):
    columns: Literal["*"] | list[str | TableColumn] = "*"
    filters: list[TableFilter] = Field(default_factory=list)
    sort: list[TableSort] = Field(default_factory=list)
    group_by: str | None = None
    row_id: str | None = None
    style_role: str = "technical"
    empty_text: str = "No rows selected."


class TableOptions(StrictOptions):
    locator: ExcelLocator
    view: TableView = Field(default_factory=TableView)
    formula_policy: Literal["cached_values", "require_no_formulas", "require_cached_results"] = "cached_values"


class PdfPageRange(StrictOptions):
    start: int
    end: int

    @model_validator(mode="after")
    def valid_range(self):
        if self.start < 1 or self.end < self.start:
            raise ValueError("PDF page range must use positive pages with end greater than or equal to start")
        return self


PdfPageSelector = int | tuple[int, int] | PdfPageRange


class PdfPagesOptions(StrictOptions):
    pages: list[PdfPageSelector]
    dpi: int = 150
    image_width_inches: float | None = None
    alignment: Literal["left", "center", "right"] = "center"
    page_break_between: bool = True
    note: str | None = None

    @field_validator("pages")
    @classmethod
    def nonempty_pages(cls, value: list[PdfPageSelector]) -> list[PdfPageSelector]:
        if not value:
            raise ValueError("pdf_pages options.pages must not be empty")
        for selector in value:
            if isinstance(selector, int) and (isinstance(selector, bool) or selector < 1):
                raise ValueError("PDF page numbers must be positive integers")
            if isinstance(selector, tuple):
                start, end = selector
                if start < 1 or end < start:
                    raise ValueError("PDF page ranges must use positive pages with end greater than or equal to start")
        return value

    @field_validator("dpi")
    @classmethod
    def valid_dpi(cls, value: int) -> int:
        if value < 72 or value > 600:
            raise ValueError("pdf_pages options.dpi must be between 72 and 600")
        return value

    @field_validator("image_width_inches")
    @classmethod
    def positive_image_width(cls, value: float | None) -> float | None:
        if value is not None and value <= 0:
            raise ValueError("pdf_pages image width must be greater than zero")
        return value


class FigureOptions(StrictOptions):
    width_inches: float | None = None

    @field_validator("width_inches")
    @classmethod
    def positive_width(cls, value: float | None) -> float | None:
        if value is not None and value <= 0:
            raise ValueError("figure width must be greater than zero")
        return value


class DiagramOptions(StrictOptions):
    rendition: str
    rendition_of_sha256: str
    width_inches: float | None = None

    @field_validator("rendition_of_sha256")
    @classmethod
    def valid_sha256(cls, value: str) -> str:
        normalized = value.strip().upper()
        if not re.fullmatch(r"[0-9A-F]{64}", normalized):
            raise ValueError("diagram rendition_of_sha256 must be a 64-character SHA-256 value")
        return normalized

    @field_validator("width_inches")
    @classmethod
    def positive_width(cls, value: float | None) -> float | None:
        if value is not None and value <= 0:
            raise ValueError("diagram width must be greater than zero")
        return value


class Reference(OpenModel):
    id: str
    source: str
    location: str | None = None
    note: str | None = None


class DocumentMetadata(OpenModel):
    type: str = "Document"
    title: str
    short_title: str | None = None
    number: str | None = None
    revision: str = "draft"
    revision_display: str | None = None
    date: date
    prepared_by: str | None = None
    status: str | None = None
    subject: str | None = None
    description: str | None = None
    keywords: list[str] = Field(default_factory=list)
    author: str | None = None


class Numbering(OpenModel):
    style: Literal["arabic", "roman_lower", "roman_upper"] = "arabic"
    start: int | None = 1
    page_count_scope: Literal["document", "region"] = "region"

    @field_validator("start")
    @classmethod
    def valid_start(cls, value: int | None) -> int | None:
        if value is not None and value < 0:
            raise ValueError("numbering start must be non-negative or null")
        return value


class HeaderFooterSelection(OpenModel):
    default: str | None = None
    first: str | None = None
    even: str | None = None


class PreviewPolicy(OpenModel):
    """How one component behaves in a lightweight document preview."""

    mode: Literal["include", "placeholder", "omit"]
    label: str | None = None


class PageRegion(OpenModel):
    layout_mode: Literal["managed", "preserve"] = "managed"
    header: str | HeaderFooterSelection | None = None
    footer: str | HeaderFooterSelection | None = None
    numbering: Numbering | None = None
    top_margin_twips: int | None = None
    bottom_margin_twips: int | None = None
    left_margin_twips: int | None = None
    right_margin_twips: int | None = None

    @field_validator(
        "top_margin_twips",
        "bottom_margin_twips",
        "left_margin_twips",
        "right_margin_twips",
    )
    @classmethod
    def nonnegative_margin(cls, value: int | None) -> int | None:
        if value is not None and value < 0:
            raise ValueError("page-region margins must be non-negative or null")
        return value


class Presentation(OpenModel):
    styles: str
    cover: str | None = None
    page_regions: dict[str, PageRegion]


class FieldBinding(OpenModel):
    path: str | None = None
    template: str | None = None
    date_format: str | None = None
    transform: Literal["upper", "lower", "title"] | None = None
    required: bool = False

    @model_validator(mode="after")
    def one_value_source(self):
        if bool(self.path) == bool(self.template):
            raise ValueError("field binding requires exactly one of path or template")
        return self


class SequenceGroup(OpenModel):
    region: str
    items: list[str]


class ComponentBase(OpenModel):
    type: ComponentType
    ownership: Ownership | None = None
    source: str | None = None
    source_tag: str | None = None
    allow_untagged: bool = False
    required: bool = True
    title: str | None = None
    toc_levels: int | None = None
    alignment: Literal["left", "center", "right"] = "left"
    caption: str | bool | None = None
    alt_text: str | None = None
    based_on: list[Reference] = Field(default_factory=list)
    preview: PreviewPolicy | None = None
    slots: dict[str, list[str]] = Field(default_factory=dict)
    options: StrictOptions = Field(default_factory=EmptyOptions)

    @model_validator(mode="after")
    def validate_source_contract(self):
        source_types = {
            ComponentType.COVER,
            ComponentType.DOCUMENT,
            ComponentType.TABLE,
            ComponentType.FIGURE,
            ComponentType.PDF_PAGES,
        }
        if self.type in source_types and not self.source:
            raise ValueError(f"component type {self.type} requires source")
        if self.type == ComponentType.TOC and self.toc_levels is not None and not 1 <= self.toc_levels <= 9:
            raise ValueError("toc_levels must be between 1 and 9")
        if self.ownership is None and self.type not in {ComponentType.TOC, ComponentType.PAGE_BREAK}:
            raise ValueError(f"component type {self.type} requires explicit ownership")
        if self.slots and self.type != ComponentType.DOCUMENT:
            raise ValueError("only document components can own nested component slots")
        if self.slots and (
            not self.source
            or Path(self.source).suffix.lower() not in {".docx", ".md", ".markdown"}
        ):
            raise ValueError("nested component slots require a DOCX- or Markdown-backed document component")
        return self


class CoverComponent(ComponentBase):
    type: Literal[ComponentType.COVER]
    ownership: Ownership
    source: str
    options: EmptyOptions = Field(default_factory=EmptyOptions)


class TocComponent(ComponentBase):
    type: Literal[ComponentType.TOC]
    ownership: None = None
    source: None = None
    options: EmptyOptions = Field(default_factory=EmptyOptions)


class DocumentComponent(ComponentBase):
    type: Literal[ComponentType.DOCUMENT]
    ownership: Ownership
    source: str
    options: DocumentOptions = Field(default_factory=DocumentOptions)

    @model_validator(mode="after")
    def options_match_source(self):
        suffix = Path(self.source).suffix.lower()
        if suffix in {".md", ".markdown"} and (
            self.options.preserve_sections or self.options.whole_document
        ):
            raise ValueError("Markdown components cannot preserve or own a complete Word document package")
        if suffix == ".docx" and (
            self.options.style_roles
            or self.options.table_style_role != "technical"
            or self.options.strict is not True
        ):
            raise ValueError("Markdown rendering options cannot be applied to a DOCX document component")
        if self.options.whole_document:
            if suffix != ".docx":
                raise ValueError("whole_document requires a DOCX-backed document component")
            if not self.options.preserve_sections:
                raise ValueError("whole_document requires preserve_sections so native section geometry is retained")
            if self.source_tag or not self.allow_untagged:
                raise ValueError("whole_document requires whole-body ownership: omit source_tag and set allow_untagged")
        return self


class TableComponent(ComponentBase):
    type: Literal[ComponentType.TABLE]
    ownership: Ownership
    source: str
    options: TableOptions


class FigureComponent(ComponentBase):
    type: Literal[ComponentType.FIGURE]
    ownership: Ownership
    source: str
    options: FigureOptions = Field(default_factory=FigureOptions)


class DiagramComponent(ComponentBase):
    type: Literal[ComponentType.DIAGRAM]
    ownership: Ownership
    source: str
    options: DiagramOptions


class PdfPagesComponent(ComponentBase):
    type: Literal[ComponentType.PDF_PAGES]
    ownership: Ownership
    source: str
    options: PdfPagesOptions


class PageBreakComponent(ComponentBase):
    type: Literal[ComponentType.PAGE_BREAK]
    ownership: None = None
    source: None = None
    options: EmptyOptions = Field(default_factory=EmptyOptions)


Component = Annotated[
    CoverComponent
    | TocComponent
    | DocumentComponent
    | TableComponent
    | FigureComponent
    | DiagramComponent
    | PdfPagesComponent
    | PageBreakComponent,
    Field(discriminator="type"),
]


class OutputNaming(OpenModel):
    basename: str

    @field_validator("basename")
    @classmethod
    def safe_basename(cls, value: str) -> str:
        value = value.strip()
        if not value or any(char in value for char in '<>:"/\\|?*'):
            raise ValueError("basename must be a non-empty Windows-safe filename stem")
        return value


class ReleaseDeclaration(OpenModel):
    gates: dict[str, GateState] = Field(default_factory=dict)
    note: str | None = None


class QualityPolicy(OpenModel):
    allow_blank_pages: bool = False


class DocumentManifest(OpenModel):
    json_schema: str | None = Field(default=None, alias="$schema")
    schema_version: Literal["agentic-document/v2"] = Field(alias="schema")
    id: str
    project: str
    profile: str
    kit: str
    metadata: DocumentMetadata
    presentation: Presentation
    sequence: list[SequenceGroup]
    components: dict[str, Component]
    outputs: OutputNaming
    quality: QualityPolicy = Field(default_factory=QualityPolicy)
    release: ReleaseDeclaration = Field(default_factory=ReleaseDeclaration)

    @field_validator("id", "project", "profile", "kit")
    @classmethod
    def safe_layer_identifiers(cls, value: str) -> str:
        return _require_safe_identifier(value)

    @model_validator(mode="after")
    def validate_graph(self):
        for component_id in self.components:
            _require_safe_identifier(component_id)
        for region_id in self.presentation.page_regions:
            _require_safe_identifier(region_id)
        region_ids = set(self.presentation.page_regions)
        sequence_regions = [group.region for group in self.sequence]
        duplicate_regions = sorted({region for region in sequence_regions if sequence_regions.count(region) > 1})
        if duplicate_regions:
            raise ValueError(
                "each page region must occupy one contiguous sequence group; repeated regions: "
                + ", ".join(duplicate_regions)
            )
        unused_regions = sorted(region_ids - set(sequence_regions))
        if unused_regions:
            raise ValueError("page regions are declared but not sequenced: " + ", ".join(unused_regions))
        top_level: list[str] = []
        for group in self.sequence:
            if group.region not in region_ids:
                raise ValueError(f"sequence references unknown page region {group.region!r}")
            top_level.extend(group.items)
        nested = [
            child
            for component in self.components.values()
            for items in component.slots.values()
            for child in items
        ]
        used = [*top_level, *nested]
        duplicates = sorted({item for item in used if used.count(item) > 1})
        if duplicates:
            raise ValueError(f"sequence contains duplicate component ids: {', '.join(duplicates)}")
        missing = sorted(set(used) - set(self.components))
        if missing:
            raise ValueError(f"sequence references undefined components: {', '.join(missing)}")
        unsequenced = sorted(set(self.components) - set(used))
        if unsequenced:
            raise ValueError(f"components are defined but not sequenced: {', '.join(unsequenced)}")

        whole_documents = [
            component_id
            for component_id, component in self.components.items()
            if component.type == ComponentType.DOCUMENT and component.options.whole_document
        ]
        if len(whole_documents) > 1:
            raise ValueError("only one component may own the complete Word document package")
        if whole_documents:
            owner = whole_documents[0]
            if top_level != [owner] or len(self.sequence) != 1:
                raise ValueError(
                    f"whole-document component {owner!r} must be the only top-level sequence item in one page region; "
                    "additional material belongs in its declared slots"
                )

        def visit(component_id: str, active: tuple[str, ...]) -> None:
            if component_id in active:
                cycle = " -> ".join((*active, component_id))
                raise ValueError(f"component slot graph contains a cycle: {cycle}")
            component = self.components[component_id]
            for slot_tag, children in component.slots.items():
                if not slot_tag.strip():
                    raise ValueError(f"component {component_id!r} contains an empty component slot name")
                for child in children:
                    if self.components[child].type in {ComponentType.COVER, ComponentType.TOC}:
                        raise ValueError(
                            f"component {child!r} cannot be nested because cover/TOC placement is document-level"
                        )
                    visit(child, (*active, component_id))

        for component_id in top_level:
            visit(component_id, ())
        return self


class KitManifest(OpenModel):
    json_schema: str | None = Field(default=None, alias="$schema")
    schema_version: Literal["agentic-kit/v2"] = Field(alias="schema")
    id: str
    components: dict[str, str]
    semantic_styles: dict[str, str] = Field(default_factory=dict)
    table_styles: dict[str, dict[str, Any]] = Field(default_factory=dict)

    @field_validator("id")
    @classmethod
    def safe_id(cls, value: str) -> str:
        return _require_safe_identifier(value)


class RegionStart(OpenModel):
    tag: str
    boundary: Literal["next_page", "continuous"] = "next_page"

    @field_validator("tag")
    @classmethod
    def nonempty_tag(cls, value: str) -> str:
        if not value.strip():
            raise ValueError("region-start tag must not be empty")
        return value.strip()


class ProfileManifest(OpenModel):
    json_schema: str | None = Field(default=None, alias="$schema")
    schema_version: Literal["agentic-profile/v2"] = Field(alias="schema")
    id: str
    shell: str
    body_slot: str
    main_start_tag: str | None = None
    layout_boundary: Literal["next_page", "continuous"] = "next_page"
    region_starts: dict[str, RegionStart] = Field(default_factory=dict)
    field_bindings: dict[str, str | FieldBinding] = Field(default_factory=dict)
    release_gates: list[str] = Field(default_factory=list)

    @field_validator("id")
    @classmethod
    def safe_id(cls, value: str) -> str:
        return _require_safe_identifier(value)

    @field_validator("region_starts")
    @classmethod
    def safe_region_ids(cls, value: dict[str, RegionStart]) -> dict[str, RegionStart]:
        for region_id in value:
            _require_safe_identifier(region_id)
        return value


class ProjectManifest(OpenModel):
    json_schema: str | None = Field(default=None, alias="$schema")
    schema_version: Literal["agentic-project/v2"] = Field(alias="schema")
    id: str
    name: str
    description: str | None = None
    metadata: dict[str, Any] = Field(default_factory=dict)
    sources: dict[str, dict[str, Any]] = Field(default_factory=dict)

    @field_validator("id")
    @classmethod
    def safe_id(cls, value: str) -> str:
        return _require_safe_identifier(value)


class ResolvedComponent(BaseModel):
    model_config = ConfigDict(extra="forbid")
    id: str
    declaration: Component
    source_path: Path | None
    source_hash: str | None
    related_paths: dict[str, Path] = Field(default_factory=dict)
    related_hashes: dict[str, str] = Field(default_factory=dict)


class ResolvedDocument(BaseModel):
    model_config = ConfigDict(extra="forbid", arbitrary_types_allowed=True)
    manifest_path: Path
    system_root: Path
    manifest: DocumentManifest
    kit_path: Path
    kit: KitManifest
    profile_path: Path
    profile: ProfileManifest
    project_path: Path
    project: ProjectManifest
    shell_path: Path
    components: dict[str, ResolvedComponent]
    presentation_paths: dict[str, Path | None]
