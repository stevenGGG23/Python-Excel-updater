# 🔍 Serial Number Exact Match & Cost Filler

This Python script is designed to extract, identify, and exactly match **serial numbers** from a subset of a main item list with a reference source list. Upon successful matching, it fills in **unit costs** and flags duplicate entries. The script focuses on accuracy by using **exact 1-to-1 serial matching only**—ensuring no mismatches or fuzzy logic is used.

---

## 📦 What It Does

- Extracts **valid serial numbers** from text using a regex (alphanumeric with at least one digit, min 4 chars)
- Only performs **exact serial number matches** (no similarity or partial matching)
- Fills in the **unit cost** from the source file if an exact match is found
- Flags **duplicate items** by name
- Outputs the updated results (rows 250–500 from the item list) to a clean Excel file

---

## 📁 Input Files

The script expects the following CSV files to be in the same folder:

1. `List_Of_Items.csv`  
   - Contains the item data including description, details, and item name  
   - Must include columns like `Item`, `Description`, `Details`

2. `Source.csv`  
   - Contains valid part numbers and corresponding unit costs  
   - Must include columns like `Part Number` and `Unit Cost`

✅ Make sure both files are encoded with `ISO-8859-1`.

---

## 🧠 How It Works

1. **Load and Clean Data**
   - Trims spaces from column headers and text fields
   - Uppercases all text for consistency

2. **Serial Extraction**
   - Uses regex to find serials that are alphanumeric, 4+ characters, and include at least one digit

3. **Matching Logic**
   - For each row (250–500) in `List_Of_Items.csv`:
     - Extract serials from `Description + Details`
     - Check if any match **exactly** with serials in the source
     - If yes → pull unit cost and save serial
     - If no → leave as `"NA"`

4. **Duplicate Check**
   - Flags duplicate `Item` values in a separate column

5. **Output**
   - Saves final result as `Updated_List_of_Items_Exact_Match.xlsx`

---

## 📊 Output Columns

The final Excel file includes:

| Column       | Description                                       |
|--------------|---------------------------------------------------|
| `Vendor`     | Placeholder (default: `"NA"`)                     |
| `UNITCOST`   | Pulled from source if exact match is found        |
| `DUPLICATES` | Flags if the item appears more than once          |
| `Macro`      | The exact matched serial number (if found)        |

---

## ⚙️ Dependencies

This script uses standard and common libraries:

- `pandas`
- `numpy`
- `re` (built-in)
- `difflib` (optional for future enhancements)

Install required packages with:


pip install pandas numpy
## 📝 Output will be saved as:
- Updated_List_of_Items_Exact_Match.xlsx 
```bash
