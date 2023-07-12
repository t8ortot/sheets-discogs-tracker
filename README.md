# Discogs Vinyl Collection Tracker for Google Sheets
A Google Script project for Google Sheets that can be used to keep an inventory of your vinyl collection with Discogs integration.

# Main Features
- Automatically builds and maintains the format of spreadsheet structure.
- Automatically imports user collection from Discogs.
- Allows the user to input their purchase statistics for their own records, such as date of purchase, price, tax, shipping, etc.
- Automatically fetches the lowest listed price on Discogs for each item and updates color indicators to show profit/loss percentages.
- Automatically calculates total investment costs and current minimum collection value.

# Outline of Script Steps
1. Normalizes the structure of the spreadsheet, which maintains and corrects the structure of the spreadsheet to how it was designed. See [Customizing Structure](#customizing-structure) for instructions to make the spreadsheet fit your needs.
2. Loads Discogs Collection. See [Automatic Discogs Import](#automatic-discogs-import) for instructions to start automatically importing your Discogs collection.
3. Updates each item with a Discogs ID with data from Discogs. See [Loading Discogs Data](#loading-discogs-data) for which fields are populated with Discogs data, and how.
4. Spreadsheet cells with formulas update in real-time. See [Loading Discogs Data](#loading-discogs-data) for which fields have formulas and what they calculate.

# Getting Started
1. Go to Google Sheets and create a new sheet.

![New Blank Spreadsheet](assets/NewBlankSpreadsheet.png)

3. In the menu bar, navigate to Extensions -> Apps Script. You should be presented with a text editor for Apps Script in a new window.

![Extension -> Apps Script](assets/ExtensionAppsScript.png)

4. Copy the code from the vinyl.js file in this repository and paste it into the Apps Script text editor, replacing the boilerplate code. Then hit the save button above the text editor.

![Save Project](assets/SaveProject.png)

5. After you save, a dropdown should be available next to the Debug button. Make sure this dropdown has the "updateSpreadsheet" option selected.

![Select Update Spreadsheet](assets/SelectUpdateSpreadsheet.png)

6. Click Run. You will be prompted to review permissions, which are only needed on the first run. Click the profile that the sheet belongs to if presented with multiple profiles.

![Review Permissions](assets/ReviewPermissions.png)

7. When you reach a screen with the text "Google hasn't verified this app," click the small "Advanced" button and then click "Go to Project (unsafe)". Since you are running this code yourself, it is not verified by Google, thus triggering the warning.

![Advanced](assets/Advanced.png) 
![Go To Project](assets/GoToProject.png)

8. A popup saying "Project wants to access your Google Account." Click "Allow." Permissions are needed to:
   - Let the script have full control over the spreadsheet so it can make changes automatically.
   - Let the script make external web requests to the Discogs API to import user and pricing data.
   - Let the script display errors to you when something goes wrong.

![Allow](assets/Allow.png)

9. You may need to click Run again in Apps Script if you took too long to accept the permissions. Once the script is complete, you can go back to the sheet to see the pre-defined structure.
10. You are ready to start using the script! You can learn how to add your collection to the spreadsheet in the [Adding To Your Collection](#adding-to-your-collection) section below. You can also continue reading to learn how the script works and to use it most effectively. Visit back here for future updates to the script and documentation. If you would like to set up the script to run automatically, please see the [Script Scheduling](#script-scheduling) section below.

# Updating The Script
Optionally, you may update your script to the latest version for new features. Official updates to the script can only be found in this GitHub repository. In order to preserve your data while performing updates to the script, it is recommended that you follow the instructions in the [Getting Started](#getting-started) section again and create a new spreadsheet from scratch, which you would configure with the latest version of the script, and then migrate your data over. Loss of your data is not my responsibility. The structure of the spreadsheet is never guaranteed to remain the same between updates. You are responsible for maintaining your own modifications to the script, if any.

# Explanation of Spreadsheet Structure
Add an explanation for the structure that the spreadsheet builds.

# Adding To Your Collection
There are two ways to import your collection into the spreadsheet, automatically or manually. A third option is being considered to be able to add a Discogs collection using the Discogs export file but has not yet been developed.

## Automatic Discogs Import
The easiest way to bring your collection into the spreadsheet is to let the script import your Discogs collection automatically. You can do this by adding your username into the Info box on the right side for "Discogs Username", and then click Run in Apps Script.

### Prerequisite
A prerequisite in order to do this is to set your Discogs collection to be public so the script can fetch it. The option to make your collection public can be found in the [Discogs privacy settings](https://www.discogs.com/settings/privacy) when signed in. If you do not wish to expose your collection to the public or do not have a collection in Discogs to import, you may use the next section's steps for [Manual Import](#manual-import).

### What Happens
When a username has been added to the spreadsheet, every time the script is run it will automatically add the Discogs IDs for all the items in your Discogs collection. The script ONLY adds IDs if they are not already in the spreadsheet. This behavior requires you to add duplicates manually if they are in your collection. Also, the script NEVER removes items, even if they are not in your Discogs collection. Therefore, items can only be deleted from the spreadsheet manually. If nothing is being added, then either your collection in Discogs is set to private, your collection is empty, all items in your collection have already been added, or you have input an invalid username.

Once all the Discogs IDs are added, the script begins to load Discogs data for each item. See the section below for [Loading Discogs Data](#loading-discogs-data).

## Manual Import
You can manually add items to the spreadsheet with or without a Discogs ID. The Discogs ID is the number that can be found in the URL of a release, like so: discogs.com/[Discogs ID]-Artist-Album. When an item is added without a Discogs ID, the script will highlight the row the next time it runs. Once a Discogs ID is added, the script will remove the highlight the next time it is run. An example use of this feature is you may add items by artist/album name that you have ordered but have not yet received. Once you receive the item, you can add its Discogs ID to the row.

There are certain fields that always require manual input because only you would know their values. Column headers that require manual input are marked with an (M), which includes the price and purchased date columns. If you do not see the marks, they are most likely hidden behind the filter buttons, so you may have to expand the column widths temporarily. 

# Loading Discogs Data
Every time the script is run, each item that contains a Discogs ID and a Last Reload Date not equal to today will load or calculate all fields that are marked with an (A). This action overwrites any data that was previously populated.

**:bangbang: IMPORTANT NOTE :bangbang:**: Due to limitations in how many times you can get data from Discogs API, in combination with Google Apps Script's strict 6-minute timer (unless you run this from a Google Workspace account), the script can only update approximately 135 items per run. Once your collection exceeds this size, the script will have to be run more than once to update all rows. See [Script Scheduling](#script-scheduling) to get the script to keep retrying on an interval until all rows are updated. 

The following data is reloaded in the following way:
- **Artist**: Populated using the first name mentioned for the release in Discogs, which is usually the main artist/contributor. Artists that have the same name as other artists are denoted with a number in Discogs, but this number is trimmed off when added to the spreadsheet.
- **Album**: Populated with the album name displayed on the release page.
- **Total**: Populates with a formula that calculates the sum of the "Price", "Tax", and "Shipping" columns. This field updates in real-time and changes only when one of the three mentioned fields is updated
- **Discogs Lowest**: Populates with the lowest listed price on Discogs. This is not to be mistaken with the last sold price, which is data that cannot be accessed using the API. If nothing is listed, then the value is set to $0.00.
- **Discogs Lowest Color**: The color of the cell is also updated to reflect the percentage of profit or loss when compared to the "Total" amount. The color gradient reaches its max/min color at +/- 10% respectively by default. The colors and percentages can be changed in the code to meet your needs.
- **Reload Difference**: Populates with the change in Discogs Lowest amount since the last time the script ran.
- **Last Reload Date**: Populates with the date the script last updated the row.

## Info Box
The following fields in the Info box, located to the right of the spreadsheet, update automatically in the following ways:
- **Item Investment**: Populates with a formula that calculates the sum of all values in the Price column. Can be used to determine investment before tax and shipping costs.
- **Total Investment**: Populates with a formula that calculates the sum of all values in the Total column. Can be used to determine the total investment including tax and shipping costs. 
- **Total Discogs Lowest**: Populates with a formula that calculates the sum of all values in the Discogs Lowest column. Can be loosely used as a minimum collection value amount.
- **Total Reload Difference**: Populates with a formula that calculates the sum of all values in the Reload Difference column. Can be used to show an increase or decrease in the Total Discogs Lowest since the last run.

# Advanced Setup
There are a few features you may add yourself using the instructions below. These are mostly quality-of-life improvements that have not or cannot be integrated into the script itself.

## Script Scheduling
Add steps for scheduling
note that reload date prevents loading more than once

## Run button in the spreadsheet
Add steps for adding a run button in the spreadsheet

## Customizing Structure
Add steps to customize sheet structure in code. note warnings about changes to structure could cause messes that they would have to clean up.

# Frequently Asked Questions
How does the script get updated?
- Updates to the script will be posted in this repository. Since the setup of this script is run manually as a copy in Apps Script, it does not and cannot have any ties to GitHub. In order to get and apply updates to the script, you will have to come back here to get the latest updates yourself. Updating is not required, but new features may be available that you may find useful. See [Updating The Script](#updating-the-script) for more information on how to update.

Why does the script keep reverting my structural changes to the spreadsheet?
- The script is hard-coded to maintain a certain structure so it knows where to put information properly labeled. Changes to the structure, such as adding or removing columns, can be done by altering the script itself. See [Customizing Structure](#customizing-structure) for more information.


# Feature Request List
These are features that have either been thought of or requested. They are considered, but not guaranteed, to be added in the future.
- Ability to switch currency. Currently only outputs USD.
- Ability to switch timezone for Last Reload Date. Currently only outputs in PST.
- Get the script to add a run button automatically
- Ability to import using Discogs export .csv file.
- Expose more settings in the Info box so the code does not need to be touched as often by the user
- Change the timing mechanism to use the rate-limiting response headers from Discogs to prevent too many request errors.
- Get the script Google-certified as an add-on so users can get and update the script via Google Marketplace instead.
- Build a standalone desktop app that is separate from Google entirely, but would allow authentication to Discogs and grant more benefits from Discogs API.
- Have Discogs Lowest set to a neutral color when there are no listings and lowest gets set to $0.00.
- See if automatic fields can be set to protect so the user doesn't touch them since their changes would be irrelevant.
- Set automatic columns to colors that represent that they are not editable.
- Add a feature to send email reports
- Add reload icon automatically
