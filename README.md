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
4. Copy the code from the vinyl.js file in this repository and paste it into the Apps Script text editor, replacing the boilerplate code. Then hit the save button.
5. In Apps Script, a dropdown should be available next to the Debug button. Make sure this dropdown has the "updateSpreadsheet" option selected.
6. Click Run. You will be prompted to review permissions. Since you are running this code yourself, it is not sanctioned by Google. 
7. When you reach a screen with the text "Google hasn't verified this app," click the small "Advanced" button and then click "Go to Project (unsafe)".
8. A popup saying "Project wants to access your Google Account." Click "Allow." Permissions are needed to:
   * Let the script have full control over the spreadsheet so it can make changes automatically.
   * Let the script make external web requests to the Discogs API to import user and pricing data.
   * Let the script display errors to you when something goes wrong.
9. You may need to click Run again in Apps Script. Once the script is complete, you can go back to the sheet to see the pre-defined structure.

# Adding To Your Collection - Discogs Import
The easiest way to bring your collection into the spreadsheet is to let the script import your Discogs collection automatically. A pre-requisite is to set your Discogs collection to be public so the script can fetch it. The option to make your collection public can be found [here](https://www.discogs.com/settings/privacy) under Basic privacy settings when logged in. If you do not wish to expose your collection to the public or do not have a collection in Discogs to import, you may use the next section's steps for Manual Input.

When you enter your username into the Info box for "Discogs Username", and then click Run in Apps Script, the script will automatically start adding the Discog IDs for the items in your collection. If you do not see anything get added, then either your collection in Discogs is still set to private, or you have input an invalid username.

Once all the Discogs IDs are added, the script will begin to load Discogs data for each item. See the section below for Loading Discogs Data.

# Adding To Your Collection - Manual Input
There are certain fields that only you would know such as the purchase date or the price you've paid. Columns that require manual input are noted with an (M). If you do not see the notes, they are most likely hidden behind the filter buttons.

# Loading Discogs Data
130 rule
loads (A) stuff - overwrites
