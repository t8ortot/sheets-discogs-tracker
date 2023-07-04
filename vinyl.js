//The sheet is the object which represents your Google Sheet document
//The data is a 2D array of all the data in the spreadsheet.
var sheet = SpreadsheetApp.getActiveSheet();
var data = null;

//Below is a list of constant variables that do not change.
//Some of these are considered the "settings" of the application.
const TODAY = new Date();

//This is the currency that the script writes with. Mutli-currency is not supported, therefore all amounts should be converted to the same currency. The currency can be changed below.
const numberToCurrency = new Intl.NumberFormat("en-US", {
    currency: "USD",
    style: "currency"
});

//This is the color that appears when a row does not have a Discogs ID to load data with.
const MISSING_COLOR = [255, 179, 186]; //Red

//These are the RGB color values that you would like to see past the bounds of the threshold margin.
//These colors are also used to calculate the gradient color anywhere between the bound and 0.
const MINIMUM_COLOR = [255, 179, 186]; // Red
const ZERO_COLOR = [255, 255, 186]; // Yellow
const MAXIMUM_COLOR = [186, 255, 201]; // Green

const INFO_BOX_HEADER_COLOR = [169, 169, 169]; //Dark Gray
const INFO_BOX_SUB_HEADER_COLOR = [211, 211, 211]; //Light Gray

//The +/- percentage threshold is used to determine the upper and lower bounds of the calculated gradient color's profit margin.
//Example: 0.10 mean that profits/losses of 10% or more will appear as the MAXIMUM_COLOR/MINIMUM_COLOR. A profit of 4% would show a blended color of 40% MAXIMUM_COLOR and 60% ZERO_COLOR.
const THRESHOLD_PERCENTAGE = 0.10;

//Sortable Column Headers (Denoted with an (A) if it is automatically populated, or (M) if it requires manual input.)
const DISCOGS_ID = "Discogs ID (M)";
const ARTIST = "Artist (A)";
const ALBUM = "Album (A)";
const PURCHASED_DATE = "Purchased Date (M)";
const PRICE = "Price (M)";
const TAX = "Tax (M)";
const SHIPPING = "Shipping (M)";
const TOTAL = "Total (A)";
const NOTES = "Notes (M)";
const DISCOGS_LOWEST = "Discogs Lowest (A)";
const RELOAD_DIFF = "Reload Difference (A)";
const LAST_RELOAD_DATE = "Last Reload Date (A)";

//You may alter which column headers appear and in which order by altering this list.
//WARNING: Altering this list after adding items will mostly likely cause a mess on said itmes that you will need to clean up. It is recommended to take a backup before making changes to the structure and manually fill in the data to their new locations.
const sortableColumnNames = [DISCOGS_ID, ARTIST, ALBUM, PURCHASED_DATE, PRICE, TAX, SHIPPING, TOTAL, DISCOGS_LOWEST, RELOAD_DIFF, LAST_RELOAD_DATE, NOTES];

//Row titles in the Info box.
const USERNAME = "Discogs Username";
const ITEM_INVESTMENT = "Item Investment";
const TOTAL_INVESTMENT = "Total Investment";
const TOTAL_DISCOGS_LOWEST = "Total Discogs Lowest";
const TOTAL_RELOAD_DIFF = "Total Reload Difference";

//You may alter which rows appear in the Info box and in which order they appear by altering this list.
//WARNING: Altering this list after it has been already inititalized will mostly likely cause a mess that you will need to clean up. It is recommended to take a backup before making changes to the structure and manually fill in the data to their new locations.
const infoRows = [ITEM_INVESTMENT,TOTAL_INVESTMENT,TOTAL_DISCOGS_LOWEST,TOTAL_RELOAD_DIFF, USERNAME];
const infoBoxRowOffset = 2;
const infoBoxColumnOffset = sortableColumnNames.length + 2;


//This is the main method that runs the script.
//Google has a strict 6-minute time limit for scripts.
//Due to the usage throttling of the Discogs API, once your collection grows past approximately 120 items, this script will have to run multiple times to update all rows.
function updateSpreadsheet() {
    reloadSpreadsheet();
    normalizeSheetStructure();
    for (var i = 1; i < data.length; i++) {
        Logger.log(DISCOGS_ID + ': ' + data[i][columnIndexFor(DISCOGS_ID)] + '  ' + ARTIST + ': ' + data[i][columnIndexFor(ARTIST)] + '   ' + ALBUM + ': ' + data[i][columnIndexFor(ALBUM)]);
        if (hasDiscogsItemID(i)) {
            if (shouldUpdateRow(data, i)) {
                var oldLowest = data[i][columnIndexFor(DISCOGS_LOWEST)];
                resetRowColor(i + 1);
                updateRowWithDiscogsData(i + 1);
                reloadSpreadsheet();
                updateColor(i + 1);
                updateRefreshDiff(oldLowest, i + 1)
                updateRefreshDate(i + 1);
                SpreadsheetApp.flush();
                Utilities.sleep(2000);
            }
        } else {
            setRowToMissingColor(i + 1);
        }
    }
}

//Returns true if row has a discogs ID to load data with
function hasDiscogsItemID(i) {
    return data[i][columnIndexFor(DISCOGS_ID)] != null && data[i][columnIndexFor(DISCOGS_ID)] != '';
}

//Initializes the column headers on the first row and sets up the filters each time to ensure the expected spreasheet structure remains intact.
function normalizeSheetStructure() {
    for (var i = 1; i <= sortableColumnNames.length; i++) {
        sheet.getRange(1, i).setValue(sortableColumnNames[i - 1]);
    }

    for (var i = 2; i <= data.length; i++){
        initTotalCostFormula(i);
    }

    if (sheet.getFilter() != null) {
        sheet.getFilter().remove();
    }
    sheet.getRange(1, 1, data.length, sortableColumnNames.length).createFilter();

    createInfoBox();
    loadUserCollection()
    sheet.autoResizeColumns(1, sortableColumnNames.length - 1);
    sheet.autoResizeColumns(infoBoxColumnOffset, 2);
}

function createInfoBox(){
  sheet.getRange(infoBoxRowOffset - 1, infoBoxColumnOffset).setValue("Info");
  sheet.getRange(infoBoxRowOffset - 1, infoBoxColumnOffset + 1).setValue("Values");
  sheet.getRange(infoBoxRowOffset - 1, infoBoxColumnOffset, 1, 2).setBackground(rgbToHex(INFO_BOX_HEADER_COLOR[0], INFO_BOX_HEADER_COLOR[1], INFO_BOX_HEADER_COLOR[2]));
  sheet.getRange(infoBoxRowOffset - 1, infoBoxColumnOffset, 1, 2).setBorder(true, true, true, true, false, false, "#000000", null);
  
  for(var i = 0; i < infoRows.length; i++){
    sheet.getRange(i + infoBoxRowOffset, infoBoxColumnOffset).setValue(infoRows[i]);
    sheet.getRange(i + infoBoxRowOffset, infoBoxColumnOffset).setBackground(rgbToHex(INFO_BOX_SUB_HEADER_COLOR[0], INFO_BOX_SUB_HEADER_COLOR[1], INFO_BOX_SUB_HEADER_COLOR[2]));
  }

  sheet.getRange(infoBoxRowOffset, infoBoxColumnOffset, infoRows.length, 2).setBorder(true, true, true, true, true, false, "#000000", null);

  var itemInvestmentCostFormula = "=SUM(" + convertIndexToLetter(columnIndexFor(PRICE) + 1) + ":" + convertIndexToLetter(columnIndexFor(PRICE) + 1) + ")";
  var totalInvestmentCostFormula = "=SUM(" + convertIndexToLetter(columnIndexFor(TOTAL) + 1) + ":" + convertIndexToLetter(columnIndexFor(TOTAL) + 1) + ")";
  var totalDiscogsLowestFormula = "=SUM(" + convertIndexToLetter(columnIndexFor(DISCOGS_LOWEST) + 1) + ":" + convertIndexToLetter(columnIndexFor(DISCOGS_LOWEST) + 1) + ")";
  var totalReloadDiffFormula = "=SUM(" + convertIndexToLetter(columnIndexFor(RELOAD_DIFF) + 1) + ":" + convertIndexToLetter(columnIndexFor(RELOAD_DIFF) + 1) + ")";

  sheet.getRange(infoBoxRowOffset + rowIndexFor(ITEM_INVESTMENT), infoBoxColumnOffset + 1).setFormula(itemInvestmentCostFormula);
  sheet.getRange(infoBoxRowOffset + rowIndexFor(TOTAL_INVESTMENT), infoBoxColumnOffset + 1).setFormula(totalInvestmentCostFormula);
  sheet.getRange(infoBoxRowOffset + rowIndexFor(TOTAL_DISCOGS_LOWEST), infoBoxColumnOffset + 1).setFormula(totalDiscogsLowestFormula);
  sheet.getRange(infoBoxRowOffset + rowIndexFor(TOTAL_RELOAD_DIFF), infoBoxColumnOffset + 1).setFormula(totalReloadDiffFormula);

}

function loadUserCollection(){
  var userName = sheet.getRange(infoBoxRowOffset + rowIndexFor(USERNAME), infoBoxColumnOffset + 1).getValue();
  if (userName != null && userName != ""){
    var url = 'https://api.discogs.com/users/' + userName + '/collection/folders/0/releases';
    
    do{
      var response = UrlFetchApp.fetch(url);
      var json = JSON.parse(response.getContentText());
      var collection = json.releases;
      for (var i = 0; i < collection.length; i++){
        if(!alreadyContainsRelease(collection[i].id)){
          sheet.getRange(data.length + 1, columnIndexFor(DISCOGS_ID) + 1).setValue(collection[i].id);
          reloadSpreadsheet();
        }
      }
      url = json.pagination.urls.next;
    } while (url != null)
  }
}

function alreadyContainsRelease(discogsID){
  for(var i = 1; i < data.length; i++){
      if(data[i][columnIndexFor(DISCOGS_ID)] == discogsID){
        return true;
      }
  }
  return false;
}

//Sets the formula of the total cost column to "=SUM(PRICE,TAX,SHIPPING)""
function initTotalCostFormula(rowNumber){
    var priceCellNumber = convertIndexToLetter(columnIndexFor(PRICE) + 1) + rowNumber;
    var taxCellNumber = convertIndexToLetter(columnIndexFor(TAX) + 1) + rowNumber;
    var shippingCellNumber =  convertIndexToLetter(columnIndexFor(SHIPPING) + 1) + rowNumber;
    var formula = "=SUM(" + priceCellNumber + "," + taxCellNumber + "," + shippingCellNumber + ")";

    sheet.getRange(rowNumber, columnIndexFor(TOTAL) + 1).setFormula(formula);
}

function convertIndexToLetter(i){
  return String.fromCharCode(i + 64);
}

//Sets whole row to MISSING_COLOR
function setRowToMissingColor(rowNumber) {
    sheet.getRange(rowNumber, 1, 1, sortableColumnNames.length).setBackground(rgbToHex(MISSING_COLOR[0], MISSING_COLOR[1], MISSING_COLOR[2]));
}

//Resets entire row color
function resetRowColor(rowNumber) {
    sheet.getRange(rowNumber, 1, 1, sortableColumnNames.length).setBackground("transparent");
}

//Uses the Discogs ID to query the Discogs API. Copies over relevent information into the spreadsheet such as the lowest price, album name, and artist name.
function updateRowWithDiscogsData(rowNumber) {
    var url = 'https://api.discogs.com/releases/' + data[rowNumber - 1][columnIndexFor(DISCOGS_ID)]
    var response = UrlFetchApp.fetch(url);
    var json = JSON.parse(response.getContentText());
    Logger.log('Lowest price: ' + json.lowest_price);
    Logger.log('Discogs Album Name: ' + json.title);
    sheet.getRange(rowNumber, columnIndexFor(DISCOGS_LOWEST) + 1).setValue(numberToCurrency.format(json.lowest_price));
    sheet.getRange(rowNumber, columnIndexFor(ALBUM) + 1).setValue(json.title);
    var artistName = json.artists[0].name.replace(/ \(.*\)/g, '');
    sheet.getRange(rowNumber, columnIndexFor(ARTIST) + 1).setValue(artistName);
}

//Updates the color of the value cell to represent the profit or loss compared to the purchased price.
function updateColor(rowNumber) {
    var total = data[rowNumber - 1][columnIndexFor(TOTAL)];
    var lowest = data[rowNumber - 1][columnIndexFor(DISCOGS_LOWEST)];

    //This will calculate the percentage or profit or loss, keeping it within the +/- THRESHOLD_PERCENTAGE bounds to set static colors past the specified range.
    var diffPercentage = Math.max(-THRESHOLD_PERCENTAGE, Math.min((lowest / total) - 1.00, THRESHOLD_PERCENTAGE));

    //if the diff is negative, we need to calculate the gradiant color between the MINUMUM_COLOR and ZERO_COLOR
    //if the diff is positive, we need to calculate the gradiant color between the ZERO_COLOR and MAXIMUM_COLOR
    if (diffPercentage < 0) {
        var r = calculateGradientColor(ZERO_COLOR[0], MINIMUM_COLOR[0], ZERO_COLOR[0], diffPercentage);
        var g = calculateGradientColor(ZERO_COLOR[1], MINIMUM_COLOR[1], ZERO_COLOR[1], diffPercentage);
        var b = calculateGradientColor(ZERO_COLOR[2], MINIMUM_COLOR[2], ZERO_COLOR[2], diffPercentage);
        sheet.getRange(rowNumber, columnIndexFor(DISCOGS_LOWEST) + 1).setBackground(rgbToHex(r, g, b));
    } else {
        var r = calculateGradientColor(MAXIMUM_COLOR[0], ZERO_COLOR[0], ZERO_COLOR[0], diffPercentage)
        var g = calculateGradientColor(MAXIMUM_COLOR[1], ZERO_COLOR[1], ZERO_COLOR[1], diffPercentage)
        var b = calculateGradientColor(MAXIMUM_COLOR[2], ZERO_COLOR[2], ZERO_COLOR[2], diffPercentage)
        sheet.getRange(rowNumber, columnIndexFor(DISCOGS_LOWEST) + 1).setBackground(rgbToHex(r, g, b));
    }
}

//Method uses the slope-intercept formula to calculate gradiant color values between the bound and 0
function calculateGradientColor(greaterColor, lesserColor, zeroIntercept, margin) {
    return Math.round(((greaterColor - lesserColor) / THRESHOLD_PERCENTAGE) * margin + zeroIntercept);
}

//Calculates and sets the difference in price between now and the last time the data was updated.
function updateRefreshDiff(oldPrice, rowNumber) {
    var diff = (data[rowNumber - 1][columnIndexFor(DISCOGS_LOWEST)] - oldPrice).toFixed(2);
    sheet.getRange(rowNumber, columnIndexFor(RELOAD_DIFF) + 1).setValue(numberToCurrency.format(diff));
}

//Updates the last reload date to be today
function updateRefreshDate(rowNumber) {
    sheet.getRange(rowNumber, columnIndexFor(LAST_RELOAD_DATE) + 1).setValue(TODAY.toLocaleDateString('en-ZA', {
        timeZone: 'PST'
    }));
}

//Converts number to hexadecimal value
function componentToHex(c) {
    var hex = c.toString(16);
    return hex.length == 1 ? "0" + hex : hex;
}

//Converts R, G, and B number values into hexidecimal string components, and then combines them.
function rgbToHex(r, g, b) {
    return "#" + componentToHex(r) + componentToHex(g) + componentToHex(b);
}

//Loads updated data into the data array. Required frequently so that steps can use previously updated data.
function reloadSpreadsheet() {
    var allData = sheet.getRange("A:" + convertIndexToLetter(sortableColumnNames.length - 1)).getValues();
    
    for(var i = 0; i < allData.length; i++){
      if(rowEmpty(allData[i])){
        var minimizedArray = [i];
        for(var j = 0; j < i ; j++){
          minimizedArray[j] = allData[j];
        }
        data = minimizedArray;
        return;
      }
    }
}

function rowEmpty(rowData){
  for(var i = 0; i < rowData.length; i++){
    if(rowData[i] != ""){
      return false;
    }
  }
  return true;
}

//Determines if the row should be updated using the last reload date. This prevents the script from updating rows twice in one day.
function shouldUpdateRow(data, i) {
    return data[i][columnIndexFor(LAST_RELOAD_DATE)] == null || data[i][columnIndexFor(LAST_RELOAD_DATE)] == '' || data[i][columnIndexFor(LAST_RELOAD_DATE)].toLocaleDateString('en-ZA', {
        timeZone: 'PST'
    }) != TODAY.toLocaleDateString('en-ZA', {
        timeZone: 'PST'
    })
}

//Finds the index number in the sortableColumnNames for the column header name.
//This allows the user to move the column headers around on the sortableColumnNames array without any additional code changes.
function columnIndexFor(columnName) {
    return sortableColumnNames.findIndex(c => columnName == c);
}

//Finds the index number in the infoRows for the row header name.
//This allows the user to move the row headers around on the infoRows array without any additional code changes.
function rowIndexFor(rowName) {
    return infoRows.findIndex(c => rowName == c);
}
