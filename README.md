# 🔍 Serial Number Exact Match & Cost Filler

This Python script is designed to extract, identify, and exactly match **serial numbers** from a subset of a main item list with a reference source list. Upon successful matching, it fills in **unit costs** and flags duplicate entries. The script focuses on accuracy by using **exact 1-to-1 serial matching only**—ensuring no mismatches or fuzzy logic is used.

---

## 📦 What It Does

- Extracts **valid serial numbers** from text using a regex (alphanumeric with at least one digit, min 4 chars)
- Only performs **exact serial number matches** (no similarity or partial matching)
- Fills in the **unit cost** from the source file if an exact match is found
- Flags **duplicate items** by name
- **Highlights matched rows in yellow** in the final Excel file
- Outputs the updated results (**rows 250–500** from the item list) to a clean Excel file

---

## 📁 Input Files

The script expects the following CSV files to be in the same folder:

1. **`List_Of_Items.csv`**  
   - Contains the item data including description, details, and item name  
   - Must include columns like `Item`, `Description`, `Details`

2. **`Source.csv`**  
   - Contains valid part numbers and corresponding unit costs  
   - Must include columns like `Part Number` and `Unit Cost`

✅ Make sure both files are encoded with `ISO-8859-1`.

---

## 🧠 How It Works

1. **Load and Clean Data**
   - Trims spaces from column headers and text fields
   - Uppercases all text for consistency

2. **Row Selection**
   - The script specifically processes **rows 250 to 500** (Excel-style counting, i.e., line numbers)
   - In code, this is done using:
     ```python
     subset_df = df.iloc[248:500].copy()
     ```
     > Note: `.iloc[248:500]` selects Python index positions 248 to 499 (which corresponds to Excel rows 249–500).

3. **Serial Extraction**
   - Uses regex to find serials that are alphanumeric, 4+ characters, and include at least one digit

4. **Matching Logic**
   - For each row in the subset:
     - Extract serials from `Description + Details`
     - Check if any match **exactly** with serials in the source
     - If yes → pull unit cost and save serial
     - If no → leave as empty

5. **Duplicate Check**
   - Flags duplicate `Item` values in a separate column

6. **Output**
   - Saves final result as `Updated_List_of_Items_Exact_Match.xlsx`
   - **Highlights matched rows in yellow** for easy review

---

## 📊 Output Columns

The final Excel file includes:

| Column       | Description                                       |
|--------------|---------------------------------------------------|
| `Vendor`     | Placeholder (default: empty)                      |
| `UNITCOST`   | Pulled from source if exact match is found        |
| `DUPLICATES` | Flags if the item appears more than once          |
| `Macro`      | The exact matched serial number (if found)        |

---

## ⚙️ Customization

You can adjust the **row range** or the **matching threshold** if you enable fuzzy matching later:

- **Change row range:**
  ```python
  subset_df = df.iloc[START_ROW:END_ROW].copy()
  ```
  Replace `START_ROW` and `END_ROW` with desired zero-based row indices.

- **Enable or adjust fuzzy matching** (currently not used):
  ```python
  find_closest_serial_match(..., threshold=0.85)
  ```

---

## 📦 Dependencies

This script uses standard and common libraries:

- `pandas`
- `numpy`
- `re` (built-in)
- `difflib` (for future enhancements)
- `openpyxl` (for Excel writing and row highlighting)

Install required packages with:

```bash
pip install pandas numpy openpyxl
```

---

## 🚀 Usage

1. Place both `List_Of_Items.csv` and `Source.csv` in the same folder as the script
2. Run the Python script
3. The output will be saved as `Updated_List_of_Items_Exact_Match.xlsx`

---

## 📝 Output

The script will generate:
- **File**: `Updated_List_of_Items_Exact_Match.xlsx`
- **Content**: Processed rows 250-500 from the original item list
- **Highlighting**: Matched rows highlighted in yellow for easy identification
- **Summary**: Console output showing match statistics (e.g., "62 out of 252 rows matched")

---

## 🔧 Technical Notes

- Uses exact string matching for serial numbers (no fuzzy logic)
- Processes a specific subset of rows (250-500) to manage large datasets efficiently
- Handles duplicate detection based on item names
- Preserves original data structure while adding new columns for analysis
