import json
import unittest

from agentic_docs.jsonc import loads_jsonc
from agentic_docs.jsonc_edit import add_component, replace_values
from agentic_docs.model import DocumentManifest
from test_model import base_manifest


class JsoncEditTests(unittest.TestCase):
    def test_replaces_exact_values_without_reformatting_comments(self):
        source = '''{
  // keep this operator note
  "components": {
    "diagram": {
      "options": {
        "rendition": "old.png",
        "rendition_of_sha256": "OLD"
      }
    }
  }
}'''
        result = replace_values(
            source,
            {
                ("components", "diagram", "options", "rendition"): "new.png",
                ("components", "diagram", "options", "rendition_of_sha256"): "NEW",
            },
        )
        self.assertIn("// keep this operator note", result)
        self.assertIn('"rendition": "new.png"', result)
        self.assertIn('"rendition_of_sha256": "NEW"', result)

    def test_adds_nested_component_without_removing_comments(self):
        payload = base_manifest()
        payload["components"]["intro"] = {
            "type": "document",
            "ownership": "source",
            "source": "content/intro.md",
        }
        text = "// keep this authority note\n" + json.dumps(payload, indent=2) + "\n"
        result = add_component(
            text,
            "diagram",
            {
                "type": "figure",
                "ownership": "snapshot",
                "source": "content/diagram.png",
                "alt_text": "Process diagram",
            },
            parent_component="intro",
            slot_name="process-diagram",
        )
        self.assertIn("// keep this authority note", result)
        manifest = DocumentManifest.model_validate(loads_jsonc(result))
        self.assertEqual(manifest.components["intro"].slots["process-diagram"], ["diagram"])

    def test_adds_top_level_component_to_existing_region(self):
        text = json.dumps(base_manifest(), indent=2) + "\n"
        result = add_component(
            text,
            "break-before-appendix",
            {"type": "page_break"},
            region="main",
        )
        manifest = DocumentManifest.model_validate(loads_jsonc(result))
        self.assertEqual(manifest.sequence[0].items, ["intro", "break-before-appendix"])

    def test_adds_another_slot_without_reformatting_existing_parent(self):
        payload = base_manifest()
        payload["components"]["intro"] = {
            "type": "document",
            "ownership": "source",
            "source": "content/intro.md",
            "slots": {"first": ["first-figure"]},
        }
        payload["components"]["first-figure"] = {
            "type": "figure",
            "ownership": "snapshot",
            "source": "content/first.png",
        }
        text = json.dumps(payload, indent=2) + "\n"
        result = add_component(
            text,
            "second-figure",
            {"type": "figure", "ownership": "snapshot", "source": "content/second.png"},
            parent_component="intro",
            slot_name="second",
        )
        manifest = DocumentManifest.model_validate(loads_jsonc(result))
        self.assertEqual(manifest.components["intro"].slots["second"], ["second-figure"])


if __name__ == "__main__":
    unittest.main()
