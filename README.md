# Word VBA: Number to Words (Greek & English)

A lightweight and fast Microsoft Word VBA macro collection designed to convert selected numerical values into words with fractions (`/100`), retaining any trailing text within the selection.

This utility is ideal for legal contracts, financial audits, invoices, and accounting documentation.

---

## 🌟 Key Features

* **Multi-Language Support:** Dedicated standalone modules for both Greek and English.
* **Grammar & Gender Accuracy:** Correct grammatical inflections in Greek (e.g., *τρεις/τέσσερις χιλιάδες*, *τρία/τέσσερα εκατομμύρια*, *εκατόν* preceding tens/units) and standard English hyphenation rules (e.g., *twenty-one*).
* **Fraction Formatting:** Formats fractional components as standard `/100` values (e.g., `,5` or `.5` → `50/100`).
* **Safe Number Handling:** Prevents integer overflow errors, easily processing values up to trillions.
* **Trailing Text Retention:** Converts the numerical portion while preserving adjacent units or currency descriptors (e.g., `1,250.50 USD` → `1,250.50 one thousand two hundred fifty and 50/100 USD`).
* **Optimized Execution:** Employs static memory caching for instant execution inside large documents.

---

## 📁 Repository Structure

* `src/NumberToWords_GR.vba`: Macro configured for European/Greek numbering formatting (`.` for thousands, `,` for decimals).
* `src/NumberToWords_EN.vba`: Macro configured for standard English numbering formatting (`,` for thousands, `.` for decimals).

---

## 🔍 Examples

### Greek Module (`NumberToWords_GR.vba`)
* `14.000` ➔ `14.000 δεκατέσσερις χιλιάδες`
* `101.350,2` ➔ `101.350,2 εκατόν μία χιλιάδες τριακόσια πενήντα και 20/100`
* `345,75 Ευρώ` ➔ `345,75 τριακόσια σαράντα πέντε και 75/100 Ευρώ`

### English Module (`NumberToWords_EN.vba`)
* `14,000` ➔ `14,000 fourteen thousand`
* `101,350.2` ➔ `101,350.2 one hundred one thousand three hundred fifty and 20/100`
* `345.75 USD` ➔ `345.75 three hundred forty-five and 75/100 USD`

---

## 🚀 Installation & Setup in Microsoft Word

Follow these steps to make the macros accessible across all Word documents on your machine:

### Step 1: Add the Code to Word
1. Open **Microsoft Word**.
2. Press `ALT + F11` to open the **Visual Basic for Applications (VBA) Editor**.
3. In the left-hand **Project Explorer** panel, right-click on `Normal` (this ensures the macros are available globally in all future documents).
4. Select **Insert** > **Module**.
5. Copy the code from `src/NumberToWords_GR.vba` and paste it into the code window.
6. To add the English version as well, repeat step 4 (**Insert** > **Module**) and paste the code from `src/NumberToWords_EN.vba`.
7. Press `Ctrl + S` to save your `Normal.dotm` template, then close the VBA editor window.

---

### Step 2: Add a Quick Access Toolbar Button (Recommended)
To run either macro with a single click instead of opening dialog menus:

1. In Microsoft Word, go to **File** > **Options** > **Quick Access Toolbar**.
2. Set the dropdown **Choose commands from:** to **Macros**.
3. Locate the macro you wish to add:
   * `Normal.NewMacros.ΜετατροπήΕπιλογήςΣεΛέξεις` (for Greek)
   * `Normal.NewMacros.ConvertSelectionToEnglishWords` (for English)
4. Select the macro and click **Add >>**.
5. Click **Modify...** to select an icon and customize the display label.
6. Click **OK** to save. The button will now appear on your Word window toolbar.

---

## 💡 How to Use

1. Type your number inside a Word document:
   * Use Greek formatting for the Greek macro (e.g., `12.450,80`).
   * Use English formatting for the English macro (e.g., `12,450.80`).
2. Highlight/select the number (with or without any immediately trailing label, such as currency).
3. Click your toolbar button, or press `ALT + F8`, pick the macro name, and click **Run**.

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
Το project διατίθεται υπό την άδεια [MIT License](LICENSE). Μπορείτε ελεύθερα να το χρησιμοποιήσετε, να το τροποποιήσετε ή να το ενσωματώσετε σε δικά σας έργα.
