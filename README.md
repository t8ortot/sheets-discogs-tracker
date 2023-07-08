# Discogs Vinyl Collection Tracker for Google Sheets
A Google Script project for Google Sheets that can be used to keep an inventory of your vinyl collection with Discogs integration.

# Main Features
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
10. You are ready to start using the script! Visit back here for updates. If you would like to set up the script to run automatically, please see the Script Scheduling section below.

# Outline of Script Steps
1. Loads current data from the spreadsheet
2. Normalizes the structure of the spreadsheet
3. Loads Discogs Collection (see Automatic Discogs Import)
4. Updates each item with a Discogs ID with data from Discogs (see Loading Discogs Data)
5. Spreadsheet cells with formulas update in real-time.

# Adding To Your Collection
There are two ways to import your collection into the spreadsheet, automatically or manually. A third option is being considered to be able to add a Discogs collection using the Discogs export file but has not yet been developed.

## Automatic Discogs Import
The easiest way to bring your collection into the spreadsheet is to let the script import your Discogs collection automatically. A prerequisite in order to do this is to set your Discogs collection to be public so the script can fetch it. The option to make your collection public can be found in the [Discogs privacy settings](https://www.discogs.com/settings/privacy) when signed in. If you do not wish to expose your collection to the public or do not have a collection in Discogs to import, you may use the next section's steps for Manual Import. Once your collection has been made public, can enter your username into the spreadsheet's Info box for "Discogs Username", and then click Run in Apps Script.

When a username has been added to the spreadsheet, every time the script is run it will automatically add the Discogs IDs for all the items in your Discogs collection. The script ONLY adds IDs if they are not already in the spreadsheet. This behavior requires you to add duplicates manually if they are in your collection. Also, the script NEVER removes items, even if they are not in your Discogs collection. Therefore, items can only be deleted from the spreadsheet manually. If nothing is being added, then either your collection in Discogs is set to private, your collection is empty, all items in your collection have already been added, or you have input an invalid username.

Once all the Discogs IDs are added, the script begins to load Discogs data for each item. See the section below for Loading Discogs Data.

## Manual Import
You can manually add items to the spreadsheet with or without a Discogs ID. The Discogs ID is the number that can be found in the URL of a release, like so: discogs.com/[Discogs ID]-Artist-Album. When an item is added without a Discogs ID, the script will highlight the row the next time it runs. Once a Discogs ID is added, the script will remove the highlight the next time it is run. An example use of this feature is you may add items by artist/album name that you have ordered but have not yet received. Once you receive the item, you can add its Discogs ID to the row.

There are certain fields that always require manual input because only you would know their values. Column headers that require manual input are marked with an (M), which includes the price and purchased date columns. If you do not see the marks, they are most likely hidden behind the filter buttons, so you may have to expand the column widths temporarily.

By default, the Discogs 

# Loading Discogs Data
Every time the script is run, each item that contains a Discogs ID and a Last Reload Date not equal to today will load or calculate all fields that are marked with an (A). This action overwrites any data that was previously populated.

The following data is reloaded in the following way:
- Artist: Populated using the first name mentioned for the release in Discogs, which is usually the main artist/contributor. Artists that have the same name as other artists are denoted with a number in Discogs, but this number is trimmed off when added to the spreadsheet.
- Album: Populated with the album name displayed on the release page.
- Total: Populated with a formula that calculates the sum of the "Price", "Tax", and "Shipping" columns. This field updates in real-time and changes only when one of the three mentioned fields is updated
- Discogs Lowest: Populates with the lowest listed price on Discogs. This is not to be mistaken with the last sold price, which is data that cannot be accessed using the API. 
- Discogs Lowest Color: The color of the cell is also updated to reflect the percentage of profit or loss when compared to the "Total" amount. The color gradient reaches its max/min color at +/- 10% respectively by default. The colors and percentages can be changed in the code to meet your needs.
- Reload Difference: Populates with the change in Discogs Lowest amount 
130 rule
loads (A) stuff - overwrites

# Advanced Setup
There are a few features you may add yourself using the instructions below. These are mostly quality-of-life improvements that cannot be integrated into the script itself.

## Script Scheduling
Add steps for scheduling

## Run button in spreadsheet
Add steps for adding run button in spresheet

# Feature Request List
These are features that have either been thought of or requested. They are considered, but not guaranteed, to be added in the future.
- Ability to switch currency
- Ability to import using Discogs export .csv file.
