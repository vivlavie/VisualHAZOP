# VisualHAZOP

A Python-based PDF annotation tool designed specifically for HAZOP (Hazard and Operability) workshops. This application allows you to read P&ID (Process & Instrumentation Diagram) or PFD (Process Flow Diagram) PDF files and create interactive annotations with HAZOP analysis data.

## Features

### PDF Viewing and Annotation
- Load and display PDF files (P&ID/PFD diagrams)
- Navigate between pages
- Overlay custom lines on PDF pages with customizable properties

### Line/Node Management
- **Create Lines**: Press `Ctrl+L` or click "Create Line" button to start drawing
  - Left-click to add points and extend the line
  - Press `ESC` or right-click to finish line creation
- **Line Properties**:
  - Customizable color, thickness, and transparency
  - Optional arrow at the end point
  - Name displayed along the longest segment
  - Adjustable font size for names
- **Edit/Delete**: Right-click on a line to edit properties or delete

### Deviation/Note Management
Each line (node) can have multiple deviations (notes) with the following properties:
- **Deviation**: Description of the deviation
- **Causes**: Multiple causes (add/remove as needed)
- **Consequence**: Description of consequences
- **Safeguards**: Multiple safeguards (add/remove as needed)
- **Recommendations**: Multiple recommendations (add/remove as needed)
- **Comments**: Additional comments

All properties are fully editable through a dedicated editor dialog.

### Spreadsheet View
- Separate window displaying all nodes and their deviations in a structured format
- Nodes are grouped with title rows showing the node name
- Title rows use the node's line color as background
- Deviations are displayed with all properties in columns
- Multiple causes, safeguards, and recommendations are properly grouped

### Data Management
- **Save**: Export analysis data to JSON format (includes link to PDF file)
- **Load**: Import previously saved analysis data
- **Excel Export**: Export the spreadsheet view to Excel (.xlsx) format
  - Proper formatting with merged cells for multi-item fields
  - Color-coded node rows
  - Auto-adjusted column widths

## Installation

1. Install Python 3.8 or higher

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

The required packages are:
- PyMuPDF (fitz) - PDF rendering
- Pillow - Image processing
- openpyxl - Excel export
- numpy - Numerical operations

## Usage

1. Run the application:
```bash
python main.py
```

2. **Open a PDF file**:
   - Click "Open PDF" button or use `Ctrl+O`
   - Select your P&ID or PFD PDF file

3. **Create a line/node**:
   - Press `Ctrl+L` or click "Create Line" button
   - Left-click on the PDF to add points
   - Press `ESC` or right-click to finish

4. **Edit node properties**:
   - Click on a line to select it
   - Right-click and choose "Edit Properties" or click "Edit Node" button
   - Adjust name, color, thickness, transparency, font size, and arrow visibility

5. **Add a deviation**:
   - Select a node
   - Right-click and choose "Add Deviation" or click "Add Deviation" button
   - Fill in all the deviation properties
   - Click "Save" to add it to the node

6. **View spreadsheet**:
   - Click "Spreadsheet" button or use View menu
   - See all nodes and deviations in a structured view
   - Export to Excel using the "Export to Excel" button

7. **Save your work**:
   - Use `Ctrl+S` to save (or "Save" button)
   - Use `Ctrl+Shift+S` to save as a new file
   - Data is saved in JSON format with a link to the PDF file

## Keyboard Shortcuts

- `Ctrl+O`: Open PDF file
- `Ctrl+L`: Start creating a line
- `ESC`: End line creation (when in line creation mode)
- `Ctrl+S`: Save analysis
- `Ctrl+Shift+S`: Save analysis as...

## File Structure

- `main.py`: Main application window and entry point
- `models.py`: Data models for nodes, deviations, and HAZOP data
- `pdf_viewer.py`: PDF viewer with overlay drawing capabilities
- `deviation_editor.py`: Dialog for editing deviations/notes
- `spreadsheet_view.py`: Spreadsheet window for viewing and exporting data
- `requirements.txt`: Python dependencies

## Data Format

The application saves data in JSON format with the following structure:
```json
{
  "pdf_path": "path/to/pdf/file.pdf",
  "nodes": [
    {
      "name": "Line 1",
      "color": "#FF0000",
      "thickness": 2,
      "transparency": 0.7,
      "has_arrow": true,
      "font_size": 12,
      "points": [[x1, y1], [x2, y2], ...],
      "page_number": 0,
      "deviations": [
        {
          "deviation": "High Pressure",
          "causes": ["Cause 1", "Cause 2"],
          "consequence": "Equipment damage",
          "safeguards": ["Safeguard 1"],
          "recommendations": ["Recommendation 1", "Recommendation 2"],
          "comments": "Additional notes",
          "minimized": false
        }
      ]
    }
  ]
}
```

## Notes

- The application maintains separate node data for each page of the PDF
- Line coordinates are stored in PDF coordinate space
- When exporting to Excel, multiple causes/safeguards/recommendations are properly grouped with merged cells
- Node background colors in the spreadsheet are automatically adjusted for text readability

## License

This software is provided as-is for HAZOP workshop use.

