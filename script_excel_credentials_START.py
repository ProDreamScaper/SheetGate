import os
import shutil
import xlwings as xw
from pathlib import Path

def get_xl(message, ext):
    # 1) lowercase, strip any dots, then re-add a single leading dot
    ext = ext.lower().lstrip('.')
    ext = f'.{ext}'

    while True:
        raw = input(message)
        cleaned = raw.strip().strip('"').strip("'")
        p = Path(cleaned).expanduser().resolve()

        if p.suffix.lower() == ext and p.exists():
            return str(p)
        else:
            print(f"❌ '{raw}' is not a valid {ext} file. Please try again.")

def main():
    xlsm_template = get_xl(
        "Enter the path to the XLSM template file: ", 
        'xlsm'       # you can now pass 'xlsm', '.xlsm', 'XLSM', etc.
    )
    print("Using template file:", xlsm_template)

        
    """ # Ask for XLSM template (the file to copy with macros)
    xlsm_template = input("Enter the path to the XLSM template file: ").strip()
    while not xlsm_template.lower().endswith('.xlsm') or not os.path.exists(xlsm_template):
        xlsm_template = input("Invalid XLSM path. Try again: ").strip()

    # Ask for XLSX file to import sheets from
    xlsx_path = input("Enter the path to the XLSX file to import sheets from: ").strip()
    while not xlsx_path.lower().endswith('.xlsx') or not os.path.exists(xlsx_path):
        xlsx_path = input("Invalid XLSX path. Try again: ").strip() """

    """ # Ask for XLSM template (the file to copy with macros)
    xlsm_template = get_xl("Enter the path to the XLSM template file: ", '.xlsm')
    print("Using template file:", xlsm_template) """

    # Ask for XLSX file to import sheets from
    xlsx_path = get_xl(
        "Enter the path to the XLSX file to import sheets from: ", 
        '.xlsx'
    )
    print("Using XLSX file:", xlsx_path)
    
    # Output XLSM file name based on XLSX filename
    xlsx_name = os.path.splitext(os.path.basename(xlsx_path))[0]
    output_file = os.path.join(os.path.dirname(xlsx_path), f"{xlsx_name}.xlsm")

    # Copy the XLSM template to a new file
    shutil.copyfile(xlsm_template, output_file)
    print(f"\n✅ Copied template to: {output_file}")

    app = xw.App(visible=False)
    try:
        # Open both workbooks
        wb_xlsx = app.books.open(xlsx_path)
        wb_xlsm = app.books.open(output_file)

        # Optionally: clear existing sheets (except first sheet if needed)
        # while len(wb_xlsm.sheets) > 0:
        #     wb_xlsm.sheets[0].delete()

        # Copy each sheet from xlsx to xlsm
        for sheet in wb_xlsx.sheets:
            print(f"📄 Copying sheet: {sheet.name}")
            sheet.api.Copy(Before=wb_xlsm.sheets[0].api)

        # Save and close
        wb_xlsx.close()
        wb_xlsm.save()
        wb_xlsm.close()
        print(f"\n✅ Done! Final XLSM with imported sheets saved as:\n{output_file}")

    finally:
        app.quit()

if __name__ == "__main__":
    main()