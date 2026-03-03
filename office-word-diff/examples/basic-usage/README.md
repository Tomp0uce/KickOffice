# Basic Usage Example

This example demonstrates how to use `office-word-diff` in a Word Add-in.

## Setup

1. Install dependencies:
```bash
npm install
```

2. Build the library:
```bash
npm run build
```

3. Create a Word Add-in manifest that references `taskpane.html`

## How It Works

1. Select text in a Word document
2. Enter the new text you want to apply in the textarea
3. Click "Apply Diff" to apply the changes with tracked changes

The example uses the `OfficeWordDiff` class with automatic cascading fallback:
- Token Map Strategy (word-level precision)
- Sentence Diff Strategy (if token mapping fails)
- Block Replace Strategy (final fallback)

## Features Demonstrated

- Loading current selection
- Preview diff statistics before applying
- Applying diffs with tracked changes
- Error handling and user feedback
- Logging and result reporting

## Note

This is a simplified example. In a production add-in, you would:
- Bundle the library code properly
- Handle edge cases more robustly
- Add loading indicators
- Provide undo/redo functionality
- Handle large documents efficiently
