# Extend PPTX Viewer discovery

Date: 2026-09-13

Package reviewed: `@extend-ai/react-pptx@0.2.0`

## Bottom line

Extend's viewer is a good rendering and inspection foundation, but it is not a mini PowerPoint
editor.

- **Text selection:** likely yes, using normal browser selection. Text is rendered as HTML `span`
  elements, not only as pixels on a canvas.
- **Selected-text events/API:** no public selection callback or selected-range API is exposed. This
  would need a small integration layer using `document.getSelection()` and the rendered DOM.
- **Search and source highlighting:** yes. The controller exposes search results with slide/node
  IDs, character offsets, snippets, and node bounds, plus a built-in node-level highlight method.
- **Text/element editing:** no. There is no documented or shipped `contentEditable` flow,
  `onChange`, drag/move callback, edit controller, or PPTX serializer/exporter.
- **Detecting changes:** only if the application owns the editing model. A `MutationObserver` would
  observe DOM changes, but it would not provide reliable semantic PowerPoint edit history.

## What was inspected

The package was temporarily installed with:

```sh
npm install @extend-ai/react-pptx@0.2.0 @tanstack/react-virtual@^3.13.12
```

The package's shipped declaration file and bundled implementation were inspected locally. The
installation was removed after discovery; this repository should not retain the viewer dependency
just for research.

Primary references:

- [Extend PowerPoint Viewer docs](https://www.extend.ai/ui/docs/components/pptx-viewer)
- [@extend-ai/react-pptx on GitHub](https://github.com/extend-hq/react-pptx)
- [@extend-ai/react-pptx on npm](https://www.npmjs.com/package/@extend-ai/react-pptx)

## Selection findings

The renderer creates a positioned DOM node for each normalized slide node and marks it with a stable
attribute:

```html
<div data-rpv-node-id="..."><div data-rpv-text-body=""><div data-rpv-text-content="">
  <div data-rpv-text-paragraph=""><span>PowerPoint text</span></div>
</div></div></div>
```

The text runs are created with `document.createElement("span")` and `span.textContent = run.text`.
The inspected implementation does not set `user-select: none`, and the slide surface is an HTML
DOM/SVG rendering surface. That means dragging over visible text should produce the browser's normal
selection highlight and copying should work in a normal browser.

The selection is not automatically a clean PowerPoint block selection. A sentence split into several
PowerPoint runs becomes several spans, so a range can cross multiple run elements. The package also
does not attach a public run ID or character-offset metadata to each span.

### How to capture selected text

An application can listen for `selectionchange`, read `document.getSelection()`, and walk the range's
anchor/focus nodes to the nearest `[data-rpv-node-id]` ancestor. `Range.toString()` gives the selected
text. For a citation record, the application would still need to calculate offsets within the
normalized node's concatenated paragraph/run text and record the slide index.

Important caveats:

1. The viewer virtualizes continuous mode, so only slides near the viewport are mounted in the DOM.
   Selection handling should resolve the slide while it is mounted.
2. A selection can span multiple runs or nodes. The integration must normalize whitespace and map
   DOM text nodes back to the package model if exact character offsets matter.
3. The package's `onSlideRendered` callback and `onViewportReady` callback are useful lifecycle hooks,
   but neither is a selection API.

## Search and highlighting findings

`PptxViewerController` exposes:

```ts
search(query, options?)
highlightSearchResult(result, options?)
clearSearchHighlights()
getDocument()
```

`search()` returns results containing:

```ts
{
  slideIndex,
  nodeId,
  nodeType,
  text,
  matchStart,
  matchEnd,
  snippet,
  bounds
}
```

This is useful for a Quarry-style workflow where a known extracted value or citation should jump to
and outline a source shape. The built-in highlight is node-level: it outlines the complete node
rectangle, not the exact characters matched inside the node.

## Editing and change tracking findings

The public API covers parsing, rendering, navigation, zoom, fit mode, thumbnails, search, search
highlighting, diagnostics, and warnings. It does not expose:

- text editing or caret state;
- `contentEditable` or an edit mode;
- element selection, resize handles, or drag/move interactions;
- `onTextChange`, `onElementChange`, `onChange`, or an edit transaction log;
- conversion of a modified model back to `.pptx`.

The package does expose a normalized `PresentationDocument` model containing slide nodes, transforms,
paragraphs, runs, images, tables, charts, and stable node IDs. In theory, an application could clone
that model, apply immutable edits, and pass a new model back to the viewer. That would still require
the application to build the editor state, pointer/keyboard interactions, undo/redo, exact text
layout behavior, and a separate PPTX export path. Passing a changed model to the viewer is a render
update, not built-in change detection or persistence.

For reliable change tracking, use one of these approaches:

1. **Own a slide model and command log.** Record operations such as `replaceText`, `moveNode`, and
   `resizeNode`, then derive the rendered document and a before/after diff.
2. **Snapshot and diff normalized models.** Compare node IDs, transforms, paragraphs/runs, and
   relevant style fields after each application-owned edit.
3. **Use the existing editor model in this repository.** The maintained slide-canvas and export
   pipeline already provide the appropriate place for immutable edits and PowerPoint output; the
   Extend package could be used only as an additional renderer/parser if its model is deliberately
   adapted.

Do not use a DOM `MutationObserver` as the source of truth. It would see renderer lifecycle work and
virtualization mounts as well as any application DOM changes, and it cannot reliably distinguish a
semantic move from a re-render.

## Recommendation

Use Extend if the immediate requirement is:

```text
PPTX/PPT rendering
  -> native text selection
  -> selected text mapped to slide/node evidence
  -> search/citation rectangles
```

Plan a custom overlay/controller for selection capture, run-to-offset mapping, and evidence records.

Do not choose this package alone for:

```text
select element -> drag/resize -> edit text -> undo/redo -> export edited PPTX
```

That workflow needs an application-owned editable slide model and exporter. The installed viewer can
be a useful read-only renderer around that model, but it does not supply the editor behavior.

## Cleanup status

The temporary `@extend-ai/react-pptx` and `@tanstack/react-virtual` dependencies were removed after
inspection. No Extend viewer component existed in the repository before this discovery, so no
existing application component was deleted. The only new artifact from this request is this report.
