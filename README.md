# EditCells

This is a Google Apps Script designed to bulk lock down regions of a spreadsheet.  The goal is to allow a group of people to each have a region on the sheet that they can edit without them being able to edit anyone else's or the rest of the sheet.

## Installation

As this script affects the spreadsheet itself, it has to be *bound* to the sheet.  In its current state, this means that it has to be installed on each sheet.

1. load the spreadsheet in Google Sheets in a browser.
2. From `Tools` select `Script Editor`.
3. Copy the code from this project into the code window.
4. Save the code.
5. Reload the spreadsheet.  A new menu item `Protect` should appear.

## Usage

To use this, decide on where you want it to have an effect.  At the start of each row in the range, enter the email address of the person who will be allowed to edit that row.  NB by "start of each row", I mean within the range.  So if the range is to be `B2:E14` then the `B` column corresponds to the start of each row.  Repetitions can be omitted, so if the first three rows are to be editable by the same person then just put their email in the first row.

Now select the range.  From the `Protect` menu, select the `Protect Range` option.  This will whirr away for a bit, and afterwards each user should be able to edit only those rows designated.

## Notes

1. Everything outside the selected range is locked for editing by anyone other than an owner.
2. The users have to have edit rights on the spreadsheet as far as sharing is concerned.  This script will give them those rights.
3. The script starts by clearing all protection from the spreadsheet.  Future versions may play more nicely with existing protection.
4. Once the script is run, the emails can be removed.
