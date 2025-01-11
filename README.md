# ExcelToLabel

## Overview
ExcelToLabel is a Python tool designed to automate the creation of printable shop labels. It reads tabular data from an Excel file, populates placeholders in a Word document template, and generates ready-to-print labels. This project is ideal for retailers looking to streamline label creation for their products.

---

## Features
ExcelToLabel simplifies label creation by:
- Processing Excel data in chunks of up to 8 rows and generating a separate Word document for each chunk.
- Dynamically replacing placeholders in the Word template (e.g., `[Marca]`, `[Modello]`) with data from the Excel file.
- Clearing unused placeholders in partially filled templates for a clean output.
- Formatting numerical values (e.g., prices) according to European conventions (e.g., replacing `.` with `,`).
- Detecting and handling locked files by prompting the user to close them before retrying.

---

## Installation

### Prerequisites
- Python 3.8 or higher

### Steps
1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/ExcelToLabel.git
   cd ExcelToLabel

2. Install dependencies:
``` bash
   pip install -r requirements.txt
```

---

## Usage

### 1. Prepare Your Files
- **Excel File (`example.xlsx`)**: Place your Excel file in the `files/` directory. It should contain columns like:

  | Marca       | Modello   | Descrizione | Codice | Dettaglio      | Prezzo Listino | Prezzo Scontato |
  |-------------|-----------|-------------|--------|----------------|----------------|-----------------|
  | Brand A     | Model X   | Desc 1      | C001   | Detail A       | 200.50         | 150             |
  | Brand B     | Model Y   | Desc 2      | C002   | Detail B       | 300.75         | 250.00          |

- **Word Template (`template.docx`)**: Customise the Word template with placeholders like:
  `[Marca]`, `[Modello]`, `[Descrizione]`, `[Codice]`, `[Prezzo Listino]`, `[Prezzo Scontato]`.

---

### 2. Run the Script
Execute the script with:
```bash
python main.py
```


### 3. Output
Generated Word documents (e.g., `Risultato_part_1.docx`, `Risultato_part_2.docx`) will appear in the files/ directory.
Each document corresponds to a chunk of up to 8 rows from the Excel file.




