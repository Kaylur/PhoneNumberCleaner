# Phone Number Cleaner for Microsoft Word

A Word VBA macro that opens a pop-up window to clean raw phone number data and insert the results into a document as a bulleted list.

## What it does

Raw phone data often includes extra text such as prefixes, dates, labels, and percentages. This tool removes that extra text and formats each number as `(XXX) XXX-XXXX`.

**Example input**

```
C:210-555-0123 (07/01/2025)
(210) 555-0187 (CT) (M) (90%) [Feedback]
1-830-555-0145
512.555.0199 [Landline]
2105550123
```

**Cleaned phone numbers**

```
(210) 555-0123
(210) 555-0187
(830) 555-0145
(512) 555-0199
```

The last entry in the input is a duplicate of the first, so it appears only once in the results.

## Features

- Removes prefixes, dates, percentages, brackets, and labels
- Recognizes common U.S. formats:
  - `2105550123`
  - `210-555-0123`
  - `210.555.0123`
  - `210 555 0123`
  - `(210) 555-0123`
  - `(210)555-0123`
  - `1-210-555-0123`
  - `+1 210 555 0123`
- Handles one or many numbers per line
- Formats every number as `(XXX) XXX-XXXX` and removes duplicates
- Reports how many lines had no recognizable number
- Inserts the results as a bulleted list anywhere in the document
- Stays open while you work, so you can click into different sections of a report
- Each insert can be undone with a single Ctrl+Z

## Requirements

- Microsoft Word for Windows (desktop version)
- Macros enabled

## Installation

1. Download `PhoneCleaner.bas`.
2. Open Word and press **Alt+F11** to open the VBA editor.
3. In the project list, select **Normal**, then go to **File > Import File** and choose `PhoneCleaner.bas`.
4. In Word, go to **File > Options > Trust Center > Trust Center Settings > Macro Settings** and check **Trust access to the VBA project object model**.
5. Run the `BuildPhoneCleanerForm` macro once. This creates the pop-up window.
6. You can turn off the trust setting from step 4 after the form is built.

Installing into Normal makes the tool available in every document. This tool can be installed alongside the Email Cleaner without conflict.

## Usage

1. Run the `PhoneCleaner` macro. You can add it to the Quick Access Toolbar for one-click access.
2. Paste the raw phone data into the top box.
3. Click **Clean Numbers** to fill the cleaned list.
4. Check the status message at the bottom. If any lines were skipped, review them by hand.
5. Click in the document where the list should go, then click **Insert Phone Numbers at Cursor**.

If the cursor is on a line that already has text, such as a section heading, the list is inserted on the line below it. Lists use Word's built-in List Bullet style.

You can edit the contents of the result box before inserting.

## Buttons

| Button | Action |
| --- | --- |
| Clean Numbers | Extracts and formats phone numbers from the input box |
| Clear All | Clears all boxes |
| Close | Closes the pop-up |
| Insert Phone Numbers at Cursor | Inserts the cleaned numbers as a bulleted list |

## Limitations

- Only 10-digit U.S. (NANP) numbers are recognized. International numbers are skipped.
- Extensions are not kept.
- Numbers with an area code or exchange starting with 0 or 1 are treated as invalid and skipped.
- Review the results and any skipped lines against the source data before finalizing a report.
- The form must be rebuilt if it is deleted. To rebuild, remove `frmPhoneCleaner` in the VBA editor and run `BuildPhoneCleanerForm` again.

## Troubleshooting

**"Word is blocking access to the VBA project"**
Turn on **Trust access to the VBA project object model** (see Installation, step 4) and run `BuildPhoneCleanerForm` again.

**"The Phone Cleaner form has not been built yet"**
Run `BuildPhoneCleanerForm` once before using `PhoneCleaner`.

**Macros are disabled**
Check your macro settings in the Trust Center or contact your IT administrator.
