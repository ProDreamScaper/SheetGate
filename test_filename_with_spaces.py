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
            

xlsm_template = get_xlsm("Enter the path to the XLSM template file: ")
print("Using template file:", xlsm_template)