# Google Sheets Vinyl Collection Tracker
A Google Script project for Google Sheets that can be used to keep an inventory of your vinyl collection with Discogs integration.

# Features
- Automatically builds and maintains the format of spreadsheet structure.
- Automatically imports user collection from Discogs (if Discogs collection is set to public).
- Allows the user to input their purchase statistics for their own records, such as date of purchase, price, tax, shipping, etc.
- Automatically fetches the lowest listed price on Discogs for each item and updates the color of the cell correlating to the profit/loss percentage.
- Automatically calculates total investment costs and current minimum collection value.

# Getting Started
1. Go to Google Sheets and create a new sheet.
2. Name your spreadsheet. You may change the name at any time.
3. In the menu bar, navigate to Extensions -> Apps Script. You should be presented with a text editor for Apps Script in a new window.
4. Copy the code from the vinyl.js file in this repository and paste it into the Apps Script text editor.
