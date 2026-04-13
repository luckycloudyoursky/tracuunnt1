# Tra Cuu NNT Batch

Python batch tool for looking up Vietnamese taxpayer information from the GDT taxpayer lookup page and exporting results to Excel.

## Run From Source

```bash
pip install beautifulsoup4 pillow playwright playwright-stealth ddddocr openpyxl
python -m playwright install chromium
python tra_cuu_thue_batch_excel.py 0110016997 -o ket_qua.xlsx
```

You can pass MST values directly, read from a text/CSV/Excel file, or pipe values through stdin.

```bash
python tra_cuu_thue_batch_excel.py -i mst_list.xlsx --input-column MST -o ket_qua.xlsx
python tra_cuu_thue_batch_excel.py 0110016997 0100109106 --delay 1 --retries 15
```

## Main Files

- `api_client_v2.py`: Playwright client for GDT lookup and CAPTCHA OCR.
- `tra_cuu_thue_batch_excel.py`: batch CLI that reads MST values and writes Excel output.

The GDT site may block or time out from some hosting providers. Run the tool from a machine or VPS that can reach `https://tracuunnt.gdt.gov.vn/tcnnt/mstdn.jsp`.
