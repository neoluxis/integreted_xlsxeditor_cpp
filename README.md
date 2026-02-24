# FemApp XLSX Editor

A integrated xlsx editor for femapp zju project

## Usage

### Test Application

```bash
# Dry-run mode (default): only mark deleted items with red background
./XLSXEditor_test <xlsx-file>

# Real-delete mode: actually delete pictures and XML definitions
./XLSXEditor_test <xlsx-file> --real-delete
```

### Widget API

```cpp
// Create editor widget with dry-run mode (default)
XLSXEditor* editor = new XLSXEditor(parent);

// Create editor widget with real-delete mode
XLSXEditor* editor = new XLSXEditor(parent, false);

// Switch mode at runtime
editor->setDryRun(true);  // Enable dry-run mode
editor->setDryRun(false); // Enable real-delete mode
```

## Features

- Load XLSX file with specified sheet and range
- Display images with descriptions in grid layout
- Toggle delete status for each item
- Select/Unselect all items
- Preview mode to hide deleted items
- Save changes:
  - **Dry-run mode**: Mark deleted items with red background
  - **Real-delete mode**: Actually delete pictures, XML definitions, and image files
