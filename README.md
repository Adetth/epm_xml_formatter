# 🛠️ Oracle EPM XML Multi-Tool

A powerful, browser-based utility designed to automate, visualize, and manage layout formatting for Oracle EPM/EPBCS XML form exports. 

Oracle EPM's native formatting UI can be notoriously finicky—especially when dealing with Data Validation Rules (DVRs), Tuple mappings, and dynamic runtime segments. This tool bypasses the UI by parsing the raw XML, injecting structural styling rules, and providing a full data dictionary of your form's color palette, all completely client-side.

**Live Application:** [Insert your GitHub Pages Link Here]

## ✨ Key Features

The application is split into 5 core modules:

### 1. 📁 File Manager (RAM Registry)
* **Dual Uploads:** Append individual files or entire folders into the browser's memory without overwriting existing data.
* **Batch Processing:** Format dozens of EPM forms simultaneously and download them packaged in a clean `.zip` file.
* **Zero Server Storage:** Everything runs entirely in your local browser memory. No sensitive financial form data is ever uploaded to a backend server.

### 2. 📊 Grid Visualizer
* **Form Preview:** Renders an accurate visual representation of the XML grid layout directly in the browser.
* **Toggle Views:** Switch between viewing Member Names, Dimension Names, or Row/Col Data Types (e.g., `FORMULA`, `MEMBER`).
* **Zoom Controls:** Dynamically scale the grid for massive enterprise forms.

### 3. 🎨 Color Injector (Data Dictionary)
* **Hex Code Mapping:** Reverse-engineers EPM's complex web of `cellStyles` and `dataValidationRules` to show exactly *where* a color is being used (e.g., `R:2, C:-1` or `Mbr: Room Rent`).
* **Safe Deletion:** Safely clear a color to instantly strip all of its associated formatting rules, styles, and tuples from the XML without crashing the form.
* **Blank Injections:** Add new custom hex codes to the dictionary to inject entirely new styling rules.

### 4. 📝 Raw Text I/O
* **Bypass IT Restrictions:** Built for strict corporate environments that block file uploads. Paste raw XML text and instantly copy the formatted output back to your clipboard.
* **Auto-Sanitizer:** Automatically detects and escapes Oracle EPM Substitution Variables (e.g., `&CurrYear`) to prevent XML parser crashes.

### 5. 💻 Console Log
* **Real-Time Debugging:** The tool features a custom DOM-based console that hijacks standard Python outputs (`sys.stdout` and `sys.stderr`), displaying full execution traces and error logs directly in the UI.

## 🏗️ Architecture & Tech Stack
This application is a **Static Site** powered by [PyScript](https://pyscript.net/), which compiles and runs a full Python engine inside the user's web browser via WebAssembly (Pyodide).
* **Frontend:** HTML5, CSS3, Vanilla JavaScript
* **Backend (In-Browser):** Python 3, `xml.etree.ElementTree`
* **Hosting:** GitHub Pages

## 🚀 Local Development
Because the application fetches external files (like `xml_analyzer.py` and `changelog.txt`), it cannot be opened by simply double-clicking the `index.html` file due to CORS security restrictions. 

To run it locally:
1. Clone the repository.
2. Open a terminal in the project folder.
3. Start a local Python web server:
   ```bash
   python -m http.server 8000
