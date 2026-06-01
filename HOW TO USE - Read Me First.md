# 📋 Job Application Tracker — How to Use

> This little tool reads the resume files in a folder and builds a tidy,
> **color-coded spreadsheet** of every job you've applied to — company, job
> title, date, status, and a notes column. Run it again whenever you add new
> resumes and it updates the sheet **without losing anything you've typed in.**

**You don't need to know any code.** Follow the steps below once, then it's a single double-click from then on. 💪

---

## 📁 What's in this folder

| File | What it's for |
|---|---|
| **`Update Job Report.bat`** | Double-click this to run the tool |
| **`update_job_report.py`** | The program (you edit just **one** line — see Step 2) |
| **`HOW TO USE - Read Me First.md`** | This guide |
| **`examples/`** | A sample resume you can copy and rename |

---

## ⚙️ One-time setup (about 5 minutes)

### Step 1 — Install Python *(only if you don't already have it)*

1. Go to **[python.org/downloads](https://www.python.org/downloads/)**
2. Download and run the installer.
3. ⚠️ **VERY IMPORTANT:** on the first screen, tick the box that says **“Add Python to PATH”**, then click **Install**.

> If you skip that box, the tool can't find Python. Just re-run the installer and tick it.

### Step 2 — Tell the tool your name

1. Right-click **`update_job_report.py`** → **Open with** → **Notepad**.
2. Near the top you'll see this line:
   ```python
   YOUR_NAME = "First Last"
   ```
3. Change it to **your** full name, exactly as it appears at the **start** of your resume file names:
   ```python
   YOUR_NAME = "Jane Doe"
   ```
4. Save (**Ctrl+S**) and close Notepad.

> That's the only edit you'll ever need to make. ✅

### Step 3 — Put your resumes in this folder

- Copy all your resume / application files (PDF or Word) into this same folder, next to the `.bat` file.
- Name them using the pattern in the next section. 👇 *(this part matters!)*

---

## 🏷️ How to name your files *(the important part)*

Name each resume like this — use underscores `_` between the pieces:

```
FirstName_LastName_Company_JobTitle.pdf
```

**Examples** (for "Jane Doe"):

| File name | Reads as |
|---|---|
| `Jane_Doe_Google_ProductManager.pdf` | **Google** · Product Manager |
| `Jane_Doe_Spotify_DataAnalyst.docx` | **Spotify** · Data Analyst |
| `Jane_Doe_Netflix_UXResearcher.docx` | **Netflix** · UX Researcher |

**A few simple rules:**

- 🔹 Your name goes **first** and must match the `YOUR_NAME` line from Step 2.
- 🔹 The **first word after your name** is the Company. Everything after that becomes the Job Title.
- 🔹 **Multi-word companies:** keep them together with no underscore:
  - ✅ `Jane_Doe_SandCherry_Consultant.pdf` → **Sand Cherry** · Consultant
  - ❌ `Jane_Doe_Sand_Cherry_Consultant.pdf` → the tool thinks the company is just "Sand"
- 🔹 **Squished titles are fine** — the tool adds the spaces for you:
  - `...AIProgramManager` → **AI Program Manager**
  - `...SrDataScientist` → **Senior Data Scientist**
- 🔹 **PDF + Word of the same job merge into one row** automatically.
- 🔹 **Cover letters:** just include the word **“Cover”** in the file name (e.g. `Jane_Doe_Google_CoverLetter.pdf`). They won't get their own row — instead the job shows **Cover Letter? = Yes**.

> 💡 Don't stress about getting every name perfect. Anything the tool guesses wrong, you can fix right in the spreadsheet — and it remembers your fix next time.

---

## ▶️ Running it

> **Double-click → `Update Job Report.bat`**

A black window opens for a few seconds, then closes when you press a key. It creates (or updates) two files in this folder:

| File | |
|---|---|
| **`job_applications.xlsx`** | 👈 Your main tracker. **Open this one.** (Excel) |
| `job_applications.csv` | A plain backup copy of the same data |

---

## 📊 Using the spreadsheet (`job_applications.xlsx`)

- 🧮 The **top banner** shows your total applications and a color legend.
- 🎨 **Status column:** click a cell and a dropdown arrow appears. Pick a status (*To Apply, Applied, Phone Screen, Interviewing, Final Round, Offer, Accepted, Rejected, Withdrawn, Ghosted*) and the cell **instantly changes color**. Everything starts as "Applied".
- 📝 **Notes column:** type anything — recruiter names, interview dates, salary info, follow-up reminders. It's saved and kept forever.
- 🔗 **Company names are clickable links** — click one to open that resume PDF.
- ✏️ You can also edit the Company, Job Title, or Date if the tool guessed something wrong. Your edits are kept next time you run it.

---

## 🔁 Adding more jobs later

1. Drop the new resume file(s) into this folder (named as above).
2. **Close** the spreadsheet if it's open *(see note below)*.
3. Double-click **`Update Job Report.bat`** again.

New jobs get added. Everything you typed (statuses, notes, edits) stays put. The previous version is automatically backed up into a `report_backups/` folder, just in case.

> ⚠️ **Close `job_applications.xlsx` in Excel before you run the `.bat`.** Excel locks the file while it's open, which stops the update. (The `.csv` will still update, but the `.xlsx` won't.)

---

## 🛠️ Troubleshooting

| Problem | Fix |
|---|---|
| **"Python is not installed"** | Re-run the installer and make sure you tick **Add Python to PATH** |
| **Company column shows your first name** | The `YOUR_NAME` line (Step 2) doesn't match the start of your file names — fix it to match exactly, then run again |
| **A company got split in half** (e.g. "Sand" instead of "Sand Cherry") | Rename the file so the company is one chunk (`SandCherry`), or just type the correct company into the spreadsheet — it'll be remembered |
| **The spreadsheet didn't update** | It was probably open in Excel. Close it and run the `.bat` again |
| **A row says `[file not found in folder]`** | You moved or renamed that resume after it was first added. The row is kept (so your notes survive). Delete it if you don't want it |

---

<div align="center">

### Good luck on the job hunt — you've got this! ☕

</div>
