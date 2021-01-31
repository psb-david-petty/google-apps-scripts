## `ZoomLink`

This is a [Google Sheets Apps Script](https://developers.google.com/apps-script/guides/sheets) to replace Zoom URIs in a Google Sheet w/ links to them. This documentation is a `DRAFT` and needs more work.

- [Example](https://docs.google.com/spreadsheets/d/1D5N7W9oFrUXjB9569Qm6inJ0Xr631kPygQz0f742eEU/edit) spreadsheet with two sheets (`display` &amp; `data`).  `data` creates the cell text for `display` &mdash; a selection of which can modified by the `zoom()` function in `Code.gs` through the added menu command `Zoom Menu > Replace Zoom Links`. To test `display`, copy the row containing *Replace Zoom Links above here* cells up to the header row, then use `Zoom Menu > Replace Zoom Links` to modify Zoom links in selected cells.
- [Published](https://docs.google.com/spreadsheets/d/e/2PACX-1vT2gwo-SF5JzRwyRdwSviT3a607qNpafRtfGXZpAf1WzJlZPzCNGdDh-eGYao5CQglw87CmZyQfAKKb/pubhtml) spreadsheet `display` sheet w/ active Zoom links.
- [`Code.gs`](https://github.com/psb-david-petty/google-apps-scripts/blob/main/ZoomLink/Code.gs) modifies the selection as follows:
  - Update the rich text for each non-empty cell
  - Replace a Zoom URI with a link to it (match color, no underline)
  - Set BHS schedule conditional formatting rules for background colors *in entire sheet* (not just selection)

### TODO

- The `Code.gs` Apps Script inlcuded in the [`sample-schedule`](https://docs.google.com/spreadsheets/d/1D5N7W9oFrUXjB9569Qm6inJ0Xr631kPygQz0f742eEU/edit) example *should* execute `onOpen()` automatically. (If `Zoom Menu` does not appear after `Help`, open `Tools > Script Editor`.)
- In `Code.gs`, `setConditionalFormattingRules` should set rules for a *range*, not the entire sheet.

[&#128279; permalink](https://psb-david-petty.github.io/google-apps-scripts/ZoomLink/) and [&#128297; repository](https://github.com/psb-david-petty/google-apps-scripts/tree/main/ZoomLink) for this page.
