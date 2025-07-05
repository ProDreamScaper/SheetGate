import os
import shutil
import xlwings as xw
from pathlib import Path

def get_xlsm(message):
    while True:
        raw = input(message)
        # 1. Remove leading/trailing whitespace and any surrounding single or double quotes
        cleaned = raw.strip().strip('"').strip("'")
        # 2. Build a pathlib.Path, expand ~, normalize separators, etc.
        p = Path(cleaned).expanduser().resolve()
        # 3. Validate extension and existence
        if p.suffix.lower() == '.xlsm' and p.exists():
            return str(p)
        else:
            print(f"❌ '{raw}' is not a valid .xlsm file. Please try again.")

def main():
    '''
    # Ask for XLSM template (the file to copy with macros)
    xlsm_template = input("Enter the path to the XLSM template file: ").strip()
    while not xlsm_template.lower().endswith('.xlsm') or not os.path.exists(xlsm_template):
        xlsm_template = input("Invalid XLSM path. Try again: ").strip()

    # Ask for SOURCE XLSM to import sheets from
    source_xlsm_path = input("Enter the path to the SOURCE XLSM file to import sheets from: ").strip()
    while not source_xlsm_path.lower().endswith('.xlsm') or not os.path.exists(source_xlsm_path):
        source_xlsm_path = input("Invalid SOURCE XLSM path. Try again: ").strip()
    '''
    xlsm_template = get_xlsm("Enter the path to the XLSM template file: ")
    source_xlsm_path = get_xlsm("Enter the path to the SOURCE XLSM file to import sheets from: ")

    # Output XLSM file name based on SOURCE XLSM filename
    source_xlsm_name = os.path.splitext(os.path.basename(source_xlsm_path))[0]
    output_file = os.path.join(os.path.dirname(source_xlsm_path), f"{source_xlsm_name}.xlsm")

    # Backup/rename XLSM Source
    backup_source_xlsm= source_xlsm_path.replace('.xlsm', '_old.xlsm')
    os.rename(source_xlsm_path, backup_source_xlsm)
    print(f"\n✅ Source renamed as {backup_source_xlsm}")

    # Copy the XLSM template to a new file
    shutil.copyfile(xlsm_template, output_file)
    print(f"\n✅ Copied template to: {output_file}")

    tempalate_sheet_names = []
    error = 0

    app = xw.App(visible=False)
    try:
        # Open both workbooks
        wb_source_xlsm = app.books.open(backup_source_xlsm)
        wb_xlsm = app.books.open(output_file)

        # Optionally: clear existing sheets (except first sheet if needed)
        # while len(wb_xlsm.sheets) > 0:
        #     wb_xlsm.sheets[0].delete()

        for sheet_2 in wb_xlsm.sheets:
            tempalate_sheet_names.append(sheet_2.name)

        # Copy each sheet from xlsm_source to new_xlsm
        for sheet in wb_source_xlsm.sheets:
            try:
                if sheet.name not in tempalate_sheet_names:
                    print(f"📄 Copying sheet: {sheet.name}")
                    sheet.api.Copy(Before=wb_xlsm.sheets[0].api)
                else:
                    print(f"📄 Modifying sheet: {sheet.name}")
            except Exception as e_sheet:
                error = 1
                print(f"❌ Error. Failed to copy sheet {sheet.name!r}: {e_sheet} \nℹ️ Don't forget admin admin credentials for both files.\n")
        # Save and close
        if error == 0:
            wb_xlsm.save()
            print(f"\n✅ Done! Final XLSM with imported sheets saved as:\n{output_file}")
    
    except Exception as exception_info_1:
        error = 1
        print(f"\n❌ Error: {exception_info_1}")

    finally:
        wb_source_xlsm.close()
        wb_xlsm.close()
        app.quit()
    if error == 1:
        try:
            os.remove(output_file)
            os.rename(backup_source_xlsm, source_xlsm_path)
            print(f"\n🔄 Returned files back.")
        except Exception as exception_info_2:
            print(f"\n❌ Error. Failed to return files back: {exception_info_2}")
if __name__ == "__main__":
    main()