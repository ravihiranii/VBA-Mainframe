# 🚚 VBA-Mainframe — Worldport Automation (VBA + MochaSoft + Excel)

> An end-to-end Excel VBA automation that drives the **Worldport** mainframe terminal via **MochaSoft TN3270 emulator** to upload shipments, fetch confirmation details, calculate accurate transit days (excluding weekends & holidays), and audit results — replacing 6–8 hours of daily manual work with a **1-hour, single-operator run** for **500+ shipments per day**.

> 🔒 **Note on source code:** This project was built for personal use. The `.xlsm` source files are **not included** in this public repository. This README serves as a case study and technical overview of the project.

---

## 📌 The Problem

The shipping operations team was spending **6 to 8 hours every day** manually keying shipment records into the **Worldport** mainframe terminal, one screen at a time, and then re-reading every screen to verify what got saved. With **500+ shipments processed daily**, the work was slow, error-prone, and exhausting — and the manual verification step was so tedious that errors regularly slipped through.

## 💡 The Solution

I built an end-to-end automation using **VBA inside Excel** as the controller and **MochaSoft (TN3270 terminal emulator)** as the bridge to the mainframe. Excel becomes the single source of truth: the operator pastes the day's shipments into the workbook, sets up the Worldport screen in MochaSoft, runs the macro, and the automation takes over — navigating screens, typing fields, reading responses, calculating correct transit days, and writing the results back into Excel along with a full audit trail.

## 📊 Results

| Metric | Before Automation | After Automation | Improvement |
|---|---|---|---|
| Daily shipment volume | 500+ | 500+ | — |
| Time required | 6–8 hours | ~1 hour | **~85% faster** |
| People required | Multiple operators | 1 operator | **Headcount freed** |
| Audit / error checking | Manual, sampled | Automated, 100% coverage | **Full coverage** |
| Risk of typing errors | High | Near zero | **Eliminated** |

---

## ✨ Key Features

- **Bulk Shipment Upload** — Reads shipment rows from Excel and pushes each one into Worldport via MochaSoft screen automation. The macro picks up the **active Worldport (WP) window** and works on it automatically.
- **Two-Way Data Sync** — After upload, the macro re-queries Worldport to **fetch confirmation details** and writes them back into Excel next to each row.
- **Smart Transit-Day Calculator** — Computes the correct transit day and final due date while **automatically skipping Saturdays, Sundays, and a configurable holiday list**. Handles edge cases: past due dates, under-3-day transits, and over-9-day transits (full logic below).
- **Separate Audit Automation** — A dedicated audit workbook re-reads the mainframe screens and **flags mismatches** directly in Excel:
  - ❌ Address mismatch between Excel input and Worldport screen
  - ❌ Wrong transit day / due date posted
  - ✅ Clean records pass through silently

---

## 🛠️ Tech Stack

- **Microsoft Excel (.xlsm)** — Input sheets, output, and audit report
- **VBA (Visual Basic for Applications)** — Automation engine and business logic
- **MochaSoft TN3270 Emulator** — Connection layer to the Worldport mainframe
- **Worldport Mainframe (WP)** — Target system for shipment data

---

## 🏗️ Architecture

```
┌─────────────────┐      ┌──────────────┐      ┌────────────────┐      ┌──────────────┐
│  Excel Input    │      │  VBA Macro   │      │   MochaSoft    │      │   Worldport  │
│  (Fetch Data)   │─────▶│  Orchestrator│─────▶│  TN3270 Window │─────▶│   Mainframe  │
│  + Holiday list │      │              │      │   (Active WP)  │      │              │
└─────────────────┘      └──────────────┘      └────────────────┘      └──────────────┘
                                │                                              │
                                │                                              │
                                ▼                                              │
                         ┌──────────────┐                                      │
                         │  Excel       │◀─────────────────────────────────────┘
                         │  Output +    │      (read confirmation screens)
                         │  Audit Sheet │
                         └──────────────┘
```

The two workbooks are intentionally kept separate so they can run independently — upload in the morning, audit later in the day after Worldport has processed the records.

```
VBA-Mainframe/
├── Worldport_Upload.xlsm           # Uploads shipments to Worldport
│   ├── Sheet: Fetch Data
│   └── Sheet: Holiday
│
└── Worldport_Audit.xlsm            # Audits posted records vs. Excel input
    ├── Sheet: Fetch Data
    └── Sheet: Holiday
```

---

## ⚙️ How It Works

1. **Open MochaSoft** and log in to Worldport. **Navigate to the correct screen first** — the macro does not log in or navigate menus; it expects the WP screen to already be ready.
2. Open the relevant `.xlsm` file (upload or audit).
3. Paste the day's shipments into the **`Fetch Data`** sheet.
4. Make sure the **`Holiday`** sheet has the current year's holidays listed in column A.
5. **Run the VBA macro.** The automation picks up the **active Worldport (WP) window** automatically and starts working on it.
6. For each shipment row, the macro:
   - Calculates the correct transit day and final due date (logic below)
   - Types the fields into the Worldport screen
   - Reads back the confirmation
   - Writes results into the output columns in Excel
7. After the upload run, the **audit workbook** is run to verify everything posted correctly.

---

## 📅 Transit Day Logic — The Core Algorithm

The transit-day calculation is the heart of this automation. It counts **working days** between the current date and the fetched due date, skipping **Saturdays, Sundays, and any date listed in the `Holiday` sheet**, then decides the final due date and transit label based on four scenarios.

### Step 1 — Count working days

Starting from `current_date + 1`, walk forward day by day up to the fetched due date. Skip Saturdays, Sundays, and any holiday — increment a `count` for every valid working day.

```vba
count = 0
checkDate = current_date + 1

Do While checkDate <= duedate1
    skipDate = False
    If Weekday(checkDate, vbSunday) <> 1 And Weekday(checkDate, vbSunday) <> 7 Then
        If dateList.Contains(checkDate) Then
            skipDate = True       ' Skip holiday
        End If
        If Not skipDate Then
            count = count + 1     ' Valid working day
        End If
    End If
    checkDate = checkDate + 1
Loop
```

### Step 2 — Apply one of four scenarios

| Scenario | Condition | Action | Transit Label |
|---|---|---|---|
| **1. Past due date** | `fetch_due_date < current_date` | Move forward from the fetched due date until 4 working days land **after** the current date | `"5TH"` |
| **2. Under 3 days** | `count < 3` | Move the due date forward until count reaches 3 working days | `"3RD"` |
| **3. Normal range** | `3 ≤ count ≤ 9` | Keep the fetched due date as-is | `"3RD"` if count = 3, else `"{count}th"` |
| **4. Over 9 days** | `count > 9` | Move the due date **backward** until count drops to 9 working days | `"{count}th"` |

### Example — Handling an over-9-day transit (snippet)

```vba
ElseIf count > 9 Then    ' Too far out — walk the due date backward to the 9th working day
    change_duedate = duedate1
    Do While count > 9
        If Weekday(change_duedate, vbSunday) = 1 _
           Or Weekday(change_duedate, vbSunday) = 7 _
           Or dateList.Contains(change_duedate) Then
            change_duedate = change_duedate - 1    ' Skip weekend / holiday
        Else
            change_duedate = change_duedate - 1
            count = count - 1
        End If
    Loop
    transit = count & "th"
    final_duedate = change_duedate
End If
```

### Why this logic?

- Worldport expects shipments to be tagged with a transit position like `3RD`, `5TH`, `7th`, etc.
- If the math gives an impossible or out-of-range result (past date, too short, too long), the automation **auto-adjusts** the due date to the nearest valid working day rather than failing or producing a bad record.
- The `Holiday` sheet makes this fully configurable — ops can add/remove holidays without touching code.

---

## 🧪 Audit Workflow

The audit is run as a **separate `.xlsm` file** after the upload is complete. It re-reads each shipment's screen from Worldport and:

- Compares the **address** on the WP screen with the Excel input → flags mismatches
- Compares the **transit day / due date** on the WP screen with the expected value → flags incorrect postings
- Marks each row with a clear ✅ / ❌ status so ops can fix only the rows that need attention

This catches both data-entry issues and any cases where Worldport itself didn't accept the record cleanly. It's the safety net that makes 100% verification feasible at 500+ shipments/day.

---

## 🎯 Challenges & Learnings

- **Mainframe screen automation is fragile** — Worldport screen layouts had to be mapped precisely; small shifts would break field positions. Built in retry logic and screen-wait timings to handle lag.
- **Date math is harder than it looks** — Handling weekends, holidays, past due dates, and out-of-range transit windows in one clean algorithm took several iterations. The four-scenario design came out of real edge cases hit in production.
- **Auditing is as important as automating** — The audit workbook turned out to be just as valuable as the upload workbook. Without it, you can't trust the automation at scale.
- **Keep upload and audit separate** — Splitting into two `.xlsm` files made the system more flexible (run them at different times, by different people, with different inputs) and easier to maintain.

---

## 🔐 Notes & Limitations

- The macro depends on the **layout of Worldport screens**. If the screens change, the VBA code will need to be updated.
- The automation **does not log in** to MochaSoft or Worldport — the operator must connect and reach the correct screen manually before running.
- The macro works on whichever **Worldport window is currently active**, so close or minimize any unrelated WP sessions before running.
- The `Holiday` sheet must be kept up to date — outdated holidays will cause incorrect transit calculations.

---

## 📈 Future Improvements

- Merge upload and audit into a single workbook with a mode toggle.
- Replace screen automation with HLLAPI/EHLLAPI calls for faster and more reliable screen reads.
- Add a progress bar UI form during long runs.
- Auto-email a daily audit summary via Outlook automation.
- Move the holiday list to a shared location (SharePoint / network drive) so multiple operators share the same calendar.

---

## 👤 Author

**Ravi Hirani**
Shipping operations automation • Excel/VBA • Mainframe integration

A personal project built to automate Worldport shipment processing. Feel free to connect if you'd like to discuss the approach or similar automation patterns.

---

## 📄 License & Usage

This project was built as a **personal project**. The source code is **not included** in this public repository. This README is shared as a **technical case study and portfolio piece** only. All rights reserved.
