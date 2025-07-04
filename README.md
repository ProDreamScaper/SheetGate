# 🔐 Excel User-Based Sheet Access System

This project provides a secure, macro-enabled Excel solution that restricts access to sheets based on user credentials. It now also includes Python scripts to automate the creation and update of workbook versions while preserving security macros.

---

## 📁 Project Structure

- `ProtectedWorkbook.xlsm` – Main macro-enabled workbook with login and access control.
- `script_excel_credentials_START.py` – Creates a new secured `.xlsm` by copying user sheets from a `.xlsx` into a macro-enabled template.
- `script_excel_credentials_UPDATE.py` – Updates an existing `.xlsm` file by merging sheets from an old secured version into a new template.

---

## 🔑 Features

✅ Secure login on file open  
✅ Sheet-level access control (users only see their assigned sheet)  
✅ Admin access to all sheets  
✅ All non-visible sheets are **VeryHidden** (cannot be seen or unhidden manually)  
✅ Sheets reset to login-only visibility on file close  
✅ User credentials managed in a hidden `UserAccess` sheet  
✅ Automate workbook updates using Python (via `xlwings`)

---

## 🧠 How It Works

### 🔸 Excel Macros

- **`Workbook_Open()`**
  - Prompts user for login (username + password)
  - Makes only the appropriate sheets visible
  - Admin sees everything
  - Hides `Login` and `UserAccess` sheets

- **`Workbook_BeforeClose()`**
  - On exit, all sheets are hidden except `Login`
  - Ensures next user must go through login

### 🔸 Python Scripts

#### `script_excel_credentials_START.py`
Use when creating a new secured `.xlsm`:
1. Provide a `.xlsm` template with macros.
2. Provide a `.xlsx` file with user-specific sheets.
3. The script copies sheets into a secured `.xlsm` output.

#### `script_excel_credentials_UPDATE.py`
Use when updating an existing `.xlsm` with a new template:
1. Backs up the old `.xlsm` as `_old.xlsm`
2. Copies sheets from the backup into the new `.xlsm`
3. Avoids overwriting existing macro/template sheets
4. Rolls back changes on error

---

## 🧪 Usage Examples

### 🔹 Create New Secured File

```bash
python script_excel_credentials_START.py
# Input:
#   > Path to macro-enabled template (.xlsm)
#   > Path to user sheet source (.xlsx)
