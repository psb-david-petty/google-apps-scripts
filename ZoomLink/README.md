## `ZoomLink`

<p style="border-left: thick solid red; margin 0 0 0 2%; font-weight: bold;">This documentation is a <code>DRAFT</code> and needs more work.</p>

This is a [Google Sheets Apps Script](https://developers.google.com/apps-script/guides/sheets) to replace Zoom URIs in a Google Sheet w/ links to them. 

There are currently two ways to embed URI links within a Google Sheet cell:
- Select individual text and use <code>Insert > &#x1f517; Insert link&nbsp;&nbsp;&nbsp;&nbsp;&#x2318;K</code>. This capability to have more than one link in a cell is a relatively new capability for Google Sheets ([May 2020](https://theverge.com/21431536/google-sheets-hyperlink-individual-many-multiple-words-cell)).
- Use the [`HYPERLINK`](https://support.google.com/docs/answer/3093313) function to link an entire cell under programmatic control.

*There is currently no way to link specific text (or more than one link in a cell) under programmatic control.* [`Code.gs`](https://github.com/psb-david-petty/google-apps-scripts/blob/main/ZoomLink/Code.gs) is a solution to that missing capability for Zoom links using Apps Script.

**To use this Apps Script, open <code>Tools > &#x2039;&#x203a; Script Editor</code> from a Google Sheet, copy / paste [`Code.gs`](https://github.com/psb-david-petty/google-apps-scripts/blob/main/ZoomLink/Code.gs) into the editor, and click <code>&#x25b7; Run</code>. (Or make a copy of a spreadsheet with this Apps Script embedded.)**

### Documents

- [Published](https://docs.google.com/spreadsheets/d/e/2PACX-1vT2gwo-SF5JzRwyRdwSviT3a607qNpafRtfGXZpAf1WzJlZPzCNGdDh-eGYao5CQglw87CmZyQfAKKb/pubhtml) spreadsheet `display` sheet w/ active Zoom links.
- [Example](https://docs.google.com/spreadsheets/d/1D5N7W9oFrUXjB9569Qm6inJ0Xr631kPygQz0f742eEU/edit) spreadsheet with two sheets (`display` &amp; `data`).  `data` creates the cell text for `display` &mdash; a selection of which can modified by the `zoom()` function in `Code.gs` through the added menu command `Zoom Menu > Replace Zoom Links`. To test `display`, copy the row containing *Replace Zoom Links above here* cells up to the header row, then use `Zoom Menu > Replace Zoom Links` to modify Zoom links in selected cells.
- [`Code.gs`](https://github.com/psb-david-petty/google-apps-scripts/blob/main/ZoomLink/Code.gs) modifies the selection as follows:
  - Update the rich text for each non-empty cell
  - Replace a Zoom URI with a link to it (match color, no underline)
  - Set BHS schedule conditional formatting rules for background colors *in entire sheet* (not just selection)

### TODO

- The `Code.gs` Apps Script included in the [`sample-schedule`](https://docs.google.com/spreadsheets/d/1D5N7W9oFrUXjB9569Qm6inJ0Xr631kPygQz0f742eEU/edit) example *should* execute `onOpen()` automatically. (If `Zoom Menu` does not appear after `Help`, open <code>Tools > &#x2039;&#x203a; Script Editor</code>.)
- In `Code.gs`, `setConditionalFormattingRules` should set rules for a *range*, not the entire sheet.
- In `Code.gs`, there ought to be another menu item added for link updating *without* `setConditionalFormattingRules`.

[&#128279; permalink](https://psb-david-petty.github.io/google-apps-scripts/ZoomLink/) and [&#128297; repository](https://github.com/psb-david-petty/google-apps-scripts/tree/main/ZoomLink) for this page.
