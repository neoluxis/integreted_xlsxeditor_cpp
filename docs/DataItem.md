# DataItem Widget

## Overview

`DataItem` is a Qt `QWidget` that presents a single XLSX data entry consisting of:

- An image button for previewing the cell picture.
- A read-only description field.
- A toggle button that marks the entry as deleted or restored.

The widget is designed to be embedded in a grid layout by the `XLSXEditor` widget.

## Public API

- `setImage(const QImage &image)`: Sets the image shown on the image button.
- `setDescription(const QString &desc)`: Sets the description text.
- `getDescription() const`: Returns the current description text.
- `setDeleted(bool deleted)`: Updates the delete state and visual style.
- `isDeleted() const`: Returns the current delete state.
- `setRowCol(int row, int col)`: Stores the row and column indices.
- `getRow() const`, `getCol() const`: Return stored row and column indices.

## Signals

- `imageClicked(int row, int col)`: Emitted when the image button is clicked.
- `deleteToggled(bool deleted)`: Emitted when the delete state is toggled.

## UI Behavior

- When deleted, the description background turns red and the button text switches to `Restore`.
- When restored, the background is cleared and the button text switches to `Delete`.
- The description field is read-only to prevent accidental edits.

## Usage Example

```cpp
DataItem *item = new DataItem(parent);
item->setImage(image);
item->setDescription(desc);
item->setRowCol(row, col);
connect(item, &DataItem::deleteToggled, [](bool deleted) {
    // Handle delete state
});
connect(item, &DataItem::imageClicked, [](int row, int col) {
    // Show image preview
});
```
