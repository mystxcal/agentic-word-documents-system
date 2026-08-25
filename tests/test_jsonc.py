import unittest

from agentic_docs.jsonc import loads_jsonc


class JsoncTests(unittest.TestCase):
    def test_preserves_comment_tokens_inside_strings(self):
        value = loads_jsonc(
            r'''
            {
              // an operator comment
              "url": "https://example.test/a//b",
              "literal": "/* keep me */",
              "items": [1, 2,],
            }
            '''
        )
        self.assertEqual(
            value,
            {
                "url": "https://example.test/a//b",
                "literal": "/* keep me */",
                "items": [1, 2],
            },
        )

    def test_preserves_newlines_around_block_comments(self):
        value = loads_jsonc('{\n/* note\n   note */\n"ok": true\n}')
        self.assertIs(value["ok"], True)


if __name__ == "__main__":
    unittest.main()
