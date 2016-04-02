# Beautify Sheet 2.0 (for Google Sheets)

## What is it?

This is an add-on for Google Sheets. It strips special formatting from a Sheet, and then applies new formatting (per two preferences objects) to the sheet.

## Why use it?

I use Google Sheets to track my checking account. Every couple weeks, I perform edits on the sheet and part of this includes pasting in confirmation numbers<sup>1</sup> and other data from billing websites. Most of this data is not plaintext -- it comes with bold, color, different font sizes, etc. My sheet looks like a mess after pasting in these pieces of text a bunch of times. So, now, in order to get a consistent look back to my spreadsheet, I can simply click the 'Beautify' > 'Active sheet' menu items and have it scrub out those special formatting changes automatically.

## Installation

If you haven't installed this as an add-on, you can use it by:

1. Copying the code from the 'add-on.js' file to your clipboard
2. In Google Sheets, clicking 'Tools' > 'Script editor'
3. Pasting the code into the Code.gs document
4. Saving the .gs file
5. Reloading the tab your spreadsheet is in

## Usage

You will now see a 'Beautify' menu item to the right of the 'Help' menu item. Simply click 'Beautify' and then click 'Active sheet'. The add-on will take a couple seconds to run (most of it is quick, but auto-resizing columns is slow).

To alter the add-on's behavior, simply edit the values in the two preferences objects near the top of the file to your desired settings.

---

1. Yes, I'm aware of right click > 'Paste special' > 'Paste values only', but I am lazy. This requires four actions, for *every* pasted item. With this add-on, I can reduce this down to three actions *total*.