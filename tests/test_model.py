import unittest
from datetime import date

from pydantic import ValidationError

from agentic_docs.model import DocumentManifest, GateState, ProfileManifest


def base_manifest() -> dict:
    return {
        "schema": "agentic-document/v2",
        "id": "demo-guide",
        "project": "demo",
        "profile": "guide",
        "kit": "studio",
        "metadata": {
            "type": "Guide",
            "title": "Demo",
            "revision": "R00",
            "date": "2026-08-20",
        },
        "presentation": {
            "styles": "kit:studio",
            "page_regions": {
                "main": {
                    "header": None,
                    "footer": None,
                    "numbering": {"start": 1},
                }
            },
        },
        "sequence": [{"region": "main", "items": ["intro"]}],
        "components": {
            "intro": {
                "type": "document",
                "ownership": "word_fragment",
                "source": "content/intro.docx",
            }
        },
        "outputs": {"basename": "Demo Guide"},
    }


class ModelTests(unittest.TestCase):
    def test_valid_manifest_keeps_unknown_extension_for_diagnostics(self):
        payload = base_manifest()
        payload["operator_note"] = "preserve me"
        result = DocumentManifest.model_validate(payload)
        self.assertEqual(result.metadata.date, date(2026, 8, 20))
        self.assertEqual(result.model_extra["operator_note"], "preserve me")

    def test_sequence_rejects_duplicate_component_use(self):
        payload = base_manifest()
        payload["sequence"][0]["items"] = ["intro", "intro"]
        with self.assertRaisesRegex(ValidationError, "duplicate component ids"):
            DocumentManifest.model_validate(payload)

    def test_sequence_rejects_unsequenced_component(self):
        payload = base_manifest()
        payload["components"]["unused"] = {
            "type": "document",
            "ownership": "word_fragment",
            "source": "content/unused.docx",
        }
        with self.assertRaisesRegex(ValidationError, "defined but not sequenced"):
            DocumentManifest.model_validate(payload)

    def test_windows_unsafe_output_name_is_rejected(self):
        payload = base_manifest()
        payload["outputs"]["basename"] = "bad:name"
        with self.assertRaisesRegex(ValidationError, "Windows-safe"):
            DocumentManifest.model_validate(payload)

    def test_path_like_manifest_identifier_is_rejected(self):
        payload = base_manifest()
        payload["project"] = "../outside"
        with self.assertRaises(ValidationError):
            DocumentManifest.model_validate(payload)

    def test_release_gates_are_typed_but_optional_for_drafts(self):
        payload = base_manifest()
        payload["release"] = {"gates": {"content-review": "open"}}
        manifest = DocumentManifest.model_validate(payload)
        self.assertEqual(manifest.release.gates["content-review"], GateState.OPEN)

    def test_component_may_place_children_in_explicit_word_slots(self):
        payload = base_manifest()
        payload["components"]["intro"]["slots"] = {"AGDOC.SLOT.TABLE": ["data"]}
        payload["components"]["data"] = {
            "type": "table",
            "ownership": "source",
            "source": "content/data.xlsx",
            "options": {"locator": {"sheet": "Data", "range": "A1:B2"}},
        }
        manifest = DocumentManifest.model_validate(payload)
        self.assertEqual(manifest.components["intro"].slots["AGDOC.SLOT.TABLE"], ["data"])

    def test_component_cannot_be_placed_both_top_level_and_in_a_slot(self):
        payload = base_manifest()
        payload["sequence"][0]["items"].append("data")
        payload["components"]["intro"]["slots"] = {"AGDOC.SLOT.TABLE": ["data"]}
        payload["components"]["data"] = {
            "type": "table",
            "ownership": "source",
            "source": "content/data.xlsx",
            "options": {"locator": {"sheet": "Data", "range": "A1:B2"}},
        }
        with self.assertRaisesRegex(ValidationError, "duplicate component ids"):
            DocumentManifest.model_validate(payload)

    def test_markdown_component_may_place_children_in_authored_insert_slots(self):
        payload = base_manifest()
        payload["components"]["intro"] = {
            "type": "document",
            "ownership": "source",
            "source": "content/intro.md",
            "slots": {"reference-table": ["data"]},
        }
        payload["components"]["data"] = {
            "type": "table",
            "ownership": "source",
            "source": "content/data.xlsx",
            "options": {"locator": {"sheet": "Data", "range": "A1:B2"}},
        }
        manifest = DocumentManifest.model_validate(payload)
        self.assertEqual(manifest.components["intro"].slots["reference-table"], ["data"])

    def test_component_options_reject_silent_typo(self):
        payload = base_manifest()
        payload["components"]["intro"]["options"] = {"preserve_section": True}
        with self.assertRaisesRegex(ValidationError, "preserve_section"):
            DocumentManifest.model_validate(payload)

    def test_pdf_component_options_are_typed(self):
        payload = base_manifest()
        payload["sequence"][0]["items"] = ["pages"]
        payload["components"] = {
            "pages": {
                "type": "pdf_pages",
                "ownership": "source",
                "source": "content/guide.pdf",
                "options": {"pages": [1, [3, 5]], "dpi": 180},
            }
        }
        manifest = DocumentManifest.model_validate(payload)
        self.assertEqual(manifest.components["pages"].options.dpi, 180)

    def test_component_preview_policy_is_small_and_explicit(self):
        payload = base_manifest()
        payload["components"]["intro"]["preview"] = {
            "mode": "placeholder",
            "label": "Large native appendix",
        }
        payload["quality"] = {"allow_blank_pages": True}
        manifest = DocumentManifest.model_validate(payload)
        self.assertEqual("placeholder", manifest.components["intro"].preview.mode)
        self.assertEqual("Large native appendix", manifest.components["intro"].preview.label)
        self.assertTrue(manifest.quality.allow_blank_pages)

    def test_diagram_requires_a_reviewed_rendition_bound_to_a_native_hash(self):
        payload = base_manifest()
        payload["sequence"][0]["items"] = ["route-map"]
        payload["components"] = {
            "route-map": {
                "type": "diagram",
                "ownership": "source",
                "source": "content/route-map.vsdx",
                "alt_text": "Walking route map",
                "options": {
                    "rendition": "content/route-map.png",
                    "rendition_of_sha256": "ab" * 32,
                },
            }
        }
        manifest = DocumentManifest.model_validate(payload)
        self.assertEqual(manifest.components["route-map"].options.rendition_of_sha256, "AB" * 32)
        payload["components"]["route-map"]["options"]["rendition_of_sha256"] = "not-a-hash"
        with self.assertRaisesRegex(ValidationError, "64-character SHA-256"):
            DocumentManifest.model_validate(payload)

    def test_page_regions_must_be_contiguous_used_and_have_valid_margins(self):
        payload = base_manifest()
        payload["presentation"]["page_regions"]["annex"] = {"top_margin_twips": 1000}
        with self.assertRaisesRegex(ValidationError, "declared but not sequenced"):
            DocumentManifest.model_validate(payload)

        payload["sequence"].append({"region": "annex", "items": []})
        payload["sequence"].append({"region": "main", "items": []})
        with self.assertRaisesRegex(ValidationError, "repeated regions"):
            DocumentManifest.model_validate(payload)

        payload = base_manifest()
        payload["presentation"]["page_regions"]["main"]["left_margin_twips"] = -1
        with self.assertRaisesRegex(ValidationError, "non-negative"):
            DocumentManifest.model_validate(payload)

    def test_page_region_layout_mode_is_explicit_and_typed(self):
        payload = base_manifest()
        self.assertEqual(
            DocumentManifest.model_validate(payload).presentation.page_regions["main"].layout_mode,
            "managed",
        )
        payload["presentation"]["page_regions"]["main"]["layout_mode"] = "preserve"
        self.assertEqual(
            DocumentManifest.model_validate(payload).presentation.page_regions["main"].layout_mode,
            "preserve",
        )
        payload["presentation"]["page_regions"]["main"]["layout_mode"] = "mystery"
        with self.assertRaises(ValidationError):
            DocumentManifest.model_validate(payload)

    def test_whole_document_is_an_explicit_single_base_package(self):
        payload = base_manifest()
        payload["components"]["intro"].update(
            {
                "source_tag": None,
                "allow_untagged": True,
                "options": {"preserve_sections": True, "whole_document": True},
            }
        )
        manifest = DocumentManifest.model_validate(payload)
        self.assertTrue(manifest.components["intro"].options.whole_document)

        payload["components"]["extra"] = {
            "type": "page_break",
        }
        payload["sequence"][0]["items"].append("extra")
        with self.assertRaisesRegex(ValidationError, "only top-level sequence item"):
            DocumentManifest.model_validate(payload)

    def test_profile_region_start_is_typed(self):
        profile = ProfileManifest.model_validate(
            {
                "schema": "agentic-profile/v2",
                "id": "example",
                "shell": "shell.docx",
                "body_slot": "AGDOC.BODY",
                "region_starts": {
                    "main": {"tag": "AGDOC.REGION.MAIN", "boundary": "next_page"}
                },
            }
        )
        self.assertEqual(profile.region_starts["main"].tag, "AGDOC.REGION.MAIN")


if __name__ == "__main__":
    unittest.main()
