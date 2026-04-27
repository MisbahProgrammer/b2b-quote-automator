📄 B2B Dynamic Quotation & Invoice Generator (Python)
A robust, pure-Python automation tool designed for B2B service providers (such as Building Maintenance, MEP Technicians, and Contracting Firms) to generate professional, branded Microsoft Word (.docx) quotations instantly.

This script eliminates manual data entry, bypasses traditional XML formatting bugs in MS Word, and programmatically draws dynamic tables with accurate price calculations.

🚀 Why This Exists
Manual invoicing in the field is slow and prone to formatting errors. Traditional template engines often crash due to hidden MS Word spellcheck XML tags breaking the code. This project solves that by using a Hybrid Automation Approach:

It uses docxtpl to safely inject top-level client metadata into a branded template.

It leverages python-docx to programmatically draw and style the line-item tables from scratch, ensuring 100% bug-free table rendering and vertical row stacking.

✨ Key Features
Bypass MS Word Formatting Bugs: Generates tables using raw Python objects instead of relying on fragile in-document Jinja2 loop tags.

Dynamic Table Generation: Automatically scales the invoice table to fit any number of line items.

Auto-Calculations: Automatically calculates the Grand Total from inputted line items.

Custom Branding & Styling: Programmatically applies custom hex colors (e.g., #1F4E79 Dark Blue, #D9D9D9 Grey) to table headers and footer banners to match company branding.

CLI Interface: Easy-to-use Command Line Interface for rapid data entry by field technicians or administrators.

🛠️ Built With
Python 3.x - Core logic and CLI

python-docx - For programmatic document and table construction
🔮 Future Roadmap (AI Integration)
This project serves as the foundational infrastructure for a future AI Engineering deployment. Planned updates include:

LLM Integration: Integrating a local model via Ollama to automatically parse unstructured, messy field notes dictated by technicians into structured JSON line items.

Web UI: Transitioning from a CLI to a Streamlit-based web dashboard.

PDF Conversion: Automated .docx to .pdf conversion pipeline.

🤝 Contributing
Contributions, issues, and feature requests are welcome! Feel free to check the issues page.

📝 License
This project is licensed under the MIT License - see the LICENSE file for details.

docxtpl - For safe template tag rendering ({{ variables }})
