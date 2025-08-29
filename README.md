# NFe XML → Excel Converter

A simple Python tool to parse **Electronic invoices (NF-e) XML files** and export key information to **Excel (.xlsx)**.

## ✨ Features
- Reads all `.xml` files in a folder (recursive search)
- Extracts:
  - Invoice **key** and **@Id**
  - **Issuer name**
  - **Recipient name**
  - **Recipient address** (street, number, district, city, state, ZIP, country)
- Exports results to a clean **Excel spreadsheet**
- Command-line interface (CLI) with arguments
- Safe XML parsing (ignores malformed or non-NFe files)
- Verbose mode for debugging

---

## 📦 Requirements
- Python **3.9+**
- Libraries:
  - `xmltodict`
  - `pandas`
  - `openpyxl`

---

## 🚀 Usage
1. Prepare the XML files

Place your .xml files inside a folder, e.g. NFs/.

By default:

- Input folder: ./NFs
- Output file: ./Invoices.xlsx

---

## 📜 License

MIT License – free to use, modify, and distribute.

---

## 🤝 Contributing

Pull requests are welcome! For major changes, please open an issue first to discuss what you would like to improve.

---

## 👨‍💻 Author

Developed by Marcos Vinicius Thibes Kemer

---