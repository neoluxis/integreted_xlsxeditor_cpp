# XLSXEditor Widget

## Overview

`XLSXEditor` is a Qt `QWidget` that loads a given XLSX file and displays a selected range of picture entries. Each picture entry is rendered as a `DataItem` in a grid layout, preserving the row and column order of the selected range.

The widget supports:

- Loading an XLSX file by path, sheet name, and range.
- Reading images from picture cells and descriptions from the cell below.
- Toggling delete state per entry.
- Saving delete markers and cell background styles back to the XLSX file.
- Restoring UI delete state for edited entries.

## Public API

- `loadXLSX(const QString &filePath, const QString &sheetName, const QString &range)`
  - Loads the XLSX file, reads the specified sheet and range, and rebuilds the UI.
  - `range` supports both `A:C,1:10` and `A1:C10` formats.

## UI Composition

- Toolbar area with `Save` and `Restore` buttons.
- Scrollable grid area that contains `DataItem` widgets.
- Bottom progress bar used during data loading.

## Save Behavior

During save, the widget writes the delete state for all data entries:

- The picture cell is set to `[D]` for deleted entries, or cleared for restored entries.
- The description cell background is set to red for deleted entries, or cleared for restored entries.

This ensures user markings fully override any existing markings in the file.

## Restore Behavior

Restore clears delete flags for modified entries in the UI and resets the description cell background and picture cell value to empty.

## Embedding Example

```cpp
XLSXEditor *editor = new XLSXEditor(parentWidget);
editor->loadXLSX(filePath, sheetName, range);
layout->addWidget(editor);
```

## Notes

- The widget is safe to reuse by calling `loadXLSX` multiple times; it will clear internal state and rebuild the UI.
- Errors during load are reported via message boxes or logs depending on the failure type.
