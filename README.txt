PMW Ticket + Fabrication APP v25

Built from your best app(2).py file.

New:
- Open the Ticket emails drop workbook and check exactly which emails to import
- Imported tickets appear on the Tickets page
- Tickets can be added to Numbering or Fabrication
- The schedule cell gets a hidden .msg link behind it
- Linked cells show a small envelope button
- Base color buttons for cell/text colors

How to test:
1. Put this folder at C:\PMW_APP
2. Run START_APP.bat
3. Login admin/admin123
4. Click Open Ticket Drop / Pick Emails, select Ticket emails drop.xlsx, then check the rows you want
5. Go to Tickets
6. Click Add to Numbering or Add to Fabrication
7. On the schedule, click the envelope icon to open the saved .msg email

This app does not map network drives and does not create Windows users.


V16 changes:
- Sorting now moves the entire row data together:
  visible text, cell color, text color, ticket/email link, and envelope.
- Schedule PDF now prints the cell background colors and text colors.
- Added Snip / Print / Email button.
  Choose start row, end row, and section, then create a smaller PDF snip.


V17 changes:
- Added Print PDF button for full colored schedule printing.
- Browser Print now tries to preserve colors better.
- Snip tool now uses the actual schedule numbers you typed, not hidden Excel row numbers.
  Example: Start # 1 and End # 5 finds rows where the Number column says 1,2,3,4,5.
- Snip PDF always includes the top schedule title/date and section header.


V18 changes:
- Select multiple cells by dragging across cells or Ctrl-clicking cells.
- Apply background color, text color, font size, or bold to all selected cells at once.
- Added selected-word editor: click one cell, click "Edit selected words", highlight words inside the editor, then color/bold/resize selected words.
- Whole-cell font size and bold are stored and included in PDFs.

V20 change:
- Fixed selected-word rich text so it appears directly in the regular schedule cell after saving.


NETWORK TESTING VERSION

This is built from v20, but the server is already set for network access.

How to test on PC:
1. Put the extracted folder at C:\PMW_APP
2. Double-click START_APP.bat
3. Open http://127.0.0.1:5050

How to test on iPhone:
1. Keep START_APP.bat running on the PC
2. Make sure iPhone is on the same Wi-Fi as the PC
3. Look in the black window for IPv4 Address
4. On iPhone Safari open:
   http://YOUR-PC-IP:5050

Example:
   http://192.168.1.115:5050

If iPhone cannot connect:
- Windows Firewall may ask to allow Python. Click Allow.
- Make sure Wi-Fi is the same network.
- Try turning off VPN on the iPhone/PC.


V23 MOBILE LAYOUT:
- Built from stable v20 Network Ready.
- iPhone screens now hide the large desktop toolbar.
- Adds a small mobile color/action strip.
- Adds a bottom mobile action bar.
- Makes rows/cells larger for touch.
- Keeps horizontal swipe scrolling for the full schedule.
- Use Safari at http://YOUR-PC-IP:5050 while on the same Wi-Fi.


V24 MOBILE CHANGES:
- iPhone can pinch zoom again.
- Added Zoom - and Zoom + buttons for the schedule.
- Fixed bottom mobile action buttons so their text is readable.
- Keeps the stable v20/v23 base.


V25 AUTO SAVE:
- Phone and PC cell edits auto-save to the shared database.
- Other devices should refresh the page to see changes.
- Keep the PC running START_APP.bat as the server.
