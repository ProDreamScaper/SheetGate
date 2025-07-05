import os
import shutil
import traceback
import xlwings as xw
import logging

# Configure a simple logger
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%H:%M:%S",
)

def copy_sheets(template_path, source_path):
    """
    Copy sheets from source_path into a fresh copy of template_path.
    Returns the path to the new file on success.
    Raises Exception on fatal errors.
    """
    # Prepare paths
    source_dir = os.path.dirname(source_path)
    base_name = os.path.splitext(os.path.basename(source_path))[0]
    output_path = os.path.join(source_dir, f"{base_name}.xlsm")
    backup_path = os.path.join(source_dir, f"{base_name}_old.xlsm")

    # 1) Backup original
    os.rename(source_path, backup_path)
    logging.info(f"Renamed source to backup: {backup_path!r}")

    try:
        # 2) Copy the template to the new output
        shutil.copyfile(template_path, output_path)
        logging.info(f"Copied template to new workbook: {output_path!r}")

        # 3) Open Excel and the two workbooks
        app = xw.App(visible=False)
        # Suppress native alerts
        app.api.DisplayAlerts = False

        wb_backup = app.books.open(backup_path)
        wb_new    = app.books.open(output_path)

        try:
            # Record template sheet names so we don't duplicate
            template_sheets = {s.name for s in wb_new.sheets}

            # Loop through each source sheet
            for sheet in wb_backup.sheets:
                try:
                    if sheet.name in template_sheets:
                        logging.info(f"Skipping existing template sheet: {sheet.name!r}")
                    else:
                        logging.info(f"Copying sheet: {sheet.name!r}")
                        # COM copy can sometimes fail; this is protected
                        sheet.api.Copy(Before=wb_new.sheets[0].api)

                except Exception as e_sheet:
                    # Log the traceback but continue with next sheet
                    logging.error(f"Failed to copy sheet {sheet.name!r}: {e_sheet}")
                    logging.debug(traceback.format_exc())

            # Save the result
            wb_new.save()
            logging.info(f"All done – saved new workbook: {output_path!r}")

        finally:
            # Always close books and quit Excel
            wb_backup.close()
            wb_new.close()
            app.quit()

        # If we reach here, everything succeeded
        return output_path

    except Exception:
        logging.info(f"ERROR_1")
        # On any error, try to clean up and restore original
        logging.error("Fatal error during copy; restoring original file.")
        logging.debug(traceback.format_exc())

        # Remove the failed new file if it exists
        if os.path.exists(output_path):
            logging.info(f"ERROR_2")
            os.remove(output_path)
            logging.info(f"Deleted failed output: {output_path!r}")

        # Restore original XLSM from backup
        if os.path.exists(backup_path):
            logging.info(f"ERROR_3")
            os.rename(backup_path, source_path)
            logging.info(f"Restored original file: {source_path!r}")

        # Re-raise so caller knows it failed
        raise

def main():
    # 1) Get valid template path
    template = input("Path to XLSM template: ").strip()
    while not template.lower().endswith('.xlsm') or not os.path.isfile(template):
        template = input("Invalid path. Please enter a valid .xlsm template: ").strip()

    # 2) Get valid source path
    source = input("Path to SOURCE XLSM: ").strip()
    while not source.lower().endswith('.xlsm') or not os.path.isfile(source):
        source = input("Invalid path. Please enter a valid .xlsm source file: ").strip()

    try:
        result = copy_sheets(template, source)
        print(f"\n✅ SUCCESS! New workbook created at:\n   {result}")
    except Exception:
        print("\n❌ FAILED – see log above for details. Your original file has been restored.")

if __name__ == "__main__":
    main()
