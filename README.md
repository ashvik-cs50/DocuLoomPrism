# Document Graph Studio

`Document Graph Studio` is a desktop Python app that reads files like `PDF`, `Word`, `Excel`, `PowerPoint`, `CSV`, and `TXT`, extracts numeric data, and turns it into clear graphs through a simple Tkinter interface.

The UI is built for quick use:

- upload a document
- choose the detected dataset
- choose a numeric column
- pick a graph style from a dropdown
- generate a labeled chart with clear topic numbering

## Features

- Supports `CSV`, `XLSX`, `XLS`, `XLSM`, `PDF`, `DOCX`, `PPTX`, and `TXT`
- Detects tables and simple label-value pairs inside documents
- Lets the user choose graph type from a dropdown menu
- Generates `Bar`, `Line`, `Horizontal Bar`, `Area`, `Scatter`, `Pie`, and `Histogram` charts
- Adds clear chart titles, topic labels, and numbered items for readability
- Shows a data preview beside the graph output

## Project Files

- `document_graph_studio.py`: main desktop application
- `requirements.txt`: Python dependencies
- `.gitignore`: common Python and editor ignores for GitHub

## Installation

1. Create and activate a virtual environment.
2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Run

```bash
python document_graph_studio.py
```

If `python` is not available in your terminal on Windows, try:

```bash
py document_graph_studio.py
```

## How It Works

1. Choose a supported file from the file picker.
2. The app extracts chart-ready numeric data when possible.
3. Select the dataset and value column you want to graph.
4. Pick a chart type from the dropdown.
5. Generate a graph with a clear title and numbered topics.

## Notes

- `PDF`, `DOCX`, and `PPTX` files work best when they contain tables or obvious label-value pairs.
- `Pie Chart` requires values greater than zero.
- Large datasets are limited for some chart types so labels stay readable.

## GitHub Ready Checklist

- Focused single-purpose app
- Clean dependency list
- Python `.gitignore`
- Updated README for installation and usage
