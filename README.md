# RepairFormat for Google Sheets 2.0

## What is it?

This script strips special formatting from a Google Sheet, and then applies new formatting (per a preferences object) to the sheet.

## Why use it?

I use Google Sheets to track my checking account. Every couple weeks, I perform edits on the sheet and part of this includes pasting in confirmation numbers<sup>1</sup> and other data from billing websites. Most of this data is not plaintext -- it comes with bold, color, different font sizes, etc. My sheet looks like a mess after pasting in these pieces of text a bunch of times. So, now, in order to get a consistent look back to my spreadsheet, I can simply click the RepairFormat > Run menu item and have it scrub out those special formatting changes automatically.

## Installation

If you haven't installed this as an add-on, you can use it by clicking 'Tools' > 'Script editor' in Sheets and then pasting the code into the Code.gs document. Save the .gs file, and then reload the tab your spreadsheet is in.

## Usage

You will now see a 'RepairFormat' menu item to the right of the 'Help' menu item. Simply click 'RepairFormat' and then click 'Run'. The script will take a couple seconds to run (most of it is quick, but auto-resizing columns is slow).

To alter the add-on's behavior, simply edit the values in the two preferences objects near the top of the file to your desires settings.

---

1. Yes, I'm aware of right click > 'Paste special' > 'Paste values only', but I am lazy. This requires four actions, for *every* pasted item. With this script, I can reduce this down to three actions *total*.