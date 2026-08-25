# Command map and recovery

## Orientation

```powershell
agentic-doc documents
agentic-doc workspace DOCUMENT
agentic-doc status DOCUMENT
agentic-doc edit DOCUMENT --show
```

Use `edit --component ID --show` for content and `edit --presentation ROLE --show` for presentation donors.

## Scoped proofs

```powershell
agentic-doc preview DOCUMENT --presentation page-furniture
agentic-doc preview DOCUMENT --component COMPONENT_ID
agentic-doc preview DOCUMENT --component COMPONENT_ID --quick
agentic-doc build DOCUMENT --quick
```

`page-furniture` uses compact synthetic content. `--quick` uses each component's `preview.mode`; PDF-page components default to a visible placeholder. Add `--include-heavy` only when the heavy component must be inspected.

Scoped proofs never update `current` and the publish gate refuses them.

## Complete proof

```powershell
agentic-doc build DOCUMENT
agentic-doc open DOCUMENT --pdf
agentic-doc check DOCUMENT
```

Before delivery, require `content_scope: complete`, no unexpected blank pages, a matching DOCX/PDF pair, and `quality.release_ready: true`.

## Component preview policy

Declare only when a component needs a non-default lightweight behavior:

```json
"preview": {
  "mode": "include | placeholder | omit",
  "label": "Optional human-readable placeholder label"
}
```

PDF-page components default to `placeholder`. Other component types default to `include`.

## Failure recovery

- If Word automation fails, inspect the build diagnostic and confirm no automation-owned `WINWORD.EXE` remains. The engine targets only the Word process it created.
- If a save prompt appears, stop and diagnose the automation lifecycle; do not dismiss it as normal.
- If an even or first-page header/footer vanishes, inspect the donor selections. A real variant in one kind requires the other kind's default to be inherited because Word controls the switch at section level.
- If a blank parity page appears, inspect `compose.page_furniture_inventory.even_and_odd_headers`. Identical default/even donors must not activate odd/even mode.
- If the change concerns only page furniture, do not run a complete build to diagnose it. Re-run the compact presentation proof.
