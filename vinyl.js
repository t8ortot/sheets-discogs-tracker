//The sheet is the object which represents your Google Sheet document
//The data is a 2D array of all the data in the spreadsheet.
var sheet = SpreadsheetApp.getActiveSheet();
var data = null;

//Below is a list of constant variables that do not change.
//Some of these are considered the "settings" of the application.
const TODAY = new Date().toLocaleDateString('en-ZA', {timeZone: 'PST'});

//This is the currency that the script writes with and Discogs communicates with. Mutli-currency is not supported, therefore all amounts should be converted to the same currency. The currency can be changed below.
const numberToCurrency = new Intl.NumberFormat("en-US", {
    currency: "USD",
    style: "currency"
});

const INFO_BOX_HEADER_COLOR = [169, 169, 169]; //Dark Gray
const INFO_BOX_SUB_HEADER_COLOR = [211, 211, 211]; //Light Gray

//Sortable Column Headers (Denoted with an (A) if it is automatically populated, or (M) if it requires manual input.)
const DISCOGS_ID = createHeader("Discogs ID       ", "The ID for the release in Discogs. It's also a hyperlink to the Discogs page. It's automatically populated once in your Discogs collection.");
const ARTIST = createHeader("Artist", "The name of the artist. It's automatically populated when a Discogs ID is added to the row.");
const ALBUM = createHeader("Album","The name of the album. It's automatically populated when a Discogs ID is added to the row.");
const PURCHASED_DATE = createHeader("Purchased Date       ","The date of purchase for the item. It can only be manually entered. It's not required.");
const PRICE = createHeader("Price","The price of the item, pre-tax/shipping, or the total price if you do not wish to store the costs in detail. It can only be manually entered.");
const TAX = createHeader("Tax","The price of tax for the item. It can only be manually entered.");
const SHIPPING = createHeader("Shipping       ","The price of shipping for the item. It can only be manually entered.");
const TOTAL = createHeader("Total","The sum of the Price, Tax, and Shipping values. It's automatically calculated. It's required to determine the Discogs Lowest profit/loss color.");
const DISCOGS_LOWEST = createHeader("Discogs Lowest       ","The lowest listed price on Discogs. It's automatically populated when a Discogs ID is added to the row. The color is determined by the calculated profit/loss percentage. If there is no listing for the item, it's set to 0.00 and the Not Listed Color.");
const RELOAD_DIFF = createHeader("Reload Difference       ","The difference amount since the Discogs Lowest was last updated. It's automatically calculated by the script.");
const LAST_RELOAD_DATE = createHeader("Last Reload Date        ","The date the script last run on the row. It's automatically populated.");
const NOTES = createHeader("Notes","Your notes about the item. They can only be manually entered and are not overwritten by the script.");

//You may alter which column headers appear and in which order by altering this list.
//WARNING: Altering this list after adding items will mostly likely cause a mess on said itmes that you will need to clean up. It is recommended to take a backup before making changes to the structure and manually fill in the data to their new locations.
const sortableColumnNames = [DISCOGS_ID, ARTIST, ALBUM, PURCHASED_DATE, PRICE, TAX, SHIPPING, TOTAL, DISCOGS_LOWEST, RELOAD_DIFF, LAST_RELOAD_DATE, NOTES];

//Row titles in the Collection Summary
const ITEM_INVESTMENT = createHeader("Item Investment", "The sum of all values in the Price column. If costs are entered in detail, it's the amount paid for your collection pre-tax/shipping.");
const TOTAL_INVESTMENT = createHeader("Total Investment", "The sum of all values in the Total column. It's the total amount paid for your collection, including tax and shipping.");
const TOTAL_DISCOGS_LOWEST = createHeader("Total Discogs Lowest", "The sum of all values in the Discogs Lowest column. It's the minumum amount your collection is currently selling for.");
const TOTAL_RELOAD_DIFF = createHeader("Total Reload Difference", "The sum of all values in the Reload Difference column. It's the total amount your collection's value has changed.");

//You may alter which rows appear in the Summary box and in which order they appear by altering this list.
//WARNING: Altering this list after it has been already inititalized will mostly likely cause a mess that you will need to clean up. It is recommended to take a backup before making changes to the structure and manually fill in the data to their new locations.
const summaryRows = [ITEM_INVESTMENT,TOTAL_INVESTMENT,TOTAL_DISCOGS_LOWEST,TOTAL_RELOAD_DIFF];
const summaryBoxRowOffset = 4;
const summaryBoxColumnOffset = sortableColumnNames.length + 2;

//Row titles in the Settings
const USERNAME = createHeader("Discogs Username", "Your Discogs username. Set it to allow the script to import your collection automatically.");
const PROFIT_LOSS_THRESHOLD = createHeader("Profit/Loss Threshold %", "The percentage used as the upper and lower bound when calculating the Discogs Lowest cell's color. Set to 10% by default. Update to expand or shrink the threshold.");
const LOSS_COLOR = createHeader("Loss Color", "The color that the Discogs Lowest cell is set to when the loss equals or exceeds the Loss Threshold %. Update this cell's color to change the color the script will use.");
const BREAK_EVEN_COLOR = createHeader("Break-Even Color", "The color that the Discogs Lowest cell is set to when the Discogs Lowest is equal to the price paid. Update this cell's color to change the color the script will use.");
const PROFIT_COLOR = createHeader("Profit Color", "The color that the Discogs Lowest cell is set to when the profit equals or exceeds the Profit Threshold %. Update this cell's color to change the color the script will use.");
const NOT_LISTED_COLOR = createHeader("Not Listed Color", "The color that the Discogs Lowest cell is set to when the value is set to 0.00, meaning that there is no lowest listed price. Update this cell's color to change the color the script will use.");
const MISSING_ID_COLOR = createHeader("Missing ID Color", "The color that the row is set to when there is no Discogs ID present. Update this cell's color to change the color the script will use.");

//You may alter which rows appear in the Settings box and in which order they appear by altering this list.
//WARNING: Altering this list after it has been already inititalized will mostly likely cause a mess that you will need to clean up. It is recommended to take a backup before making changes to the structure and manually fill in the data to their new locations.
const settingsRows = [USERNAME, PROFIT_LOSS_THRESHOLD, LOSS_COLOR, BREAK_EVEN_COLOR, PROFIT_COLOR, NOT_LISTED_COLOR, MISSING_ID_COLOR];
const settingsBoxRowOffset = summaryBoxRowOffset + summaryRows.length + 1;
const settingsBoxColumnOffset = sortableColumnNames.length + 2;

const linksBoxRowOffset = settingsBoxRowOffset + settingsRows.length;
const linksBoxColumnOffset = sortableColumnNames.length + 2


//This is the main method that runs the script.
//Google has a strict 6-minute time limit for scripts.
//Due to the usage throttling of the Discogs API, once your collection grows past approximately 120 items, this script will have to run multiple times to update all rows.
function updateSpreadsheet() {
    reloadSpreadsheet();
    normalizeSheetStructure();
    loadUserCollection();
    loadDiscogsData();
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Vinyl Tracker')
      .addItem('Run Script', 'updateSpreadsheet')
      .addItem('Reset Structure', 'normalizeSheetStructure')
      .addItem('Load New Discogs Items', 'loadUserCollection')
      .addItem('Load Discogs Data', 'loadDiscogsData')
      .addToUi();
}

//Returns true if row has a discogs ID to load data with
function hasDiscogsItemID(i) {
    return data[i][columnIndexFor(DISCOGS_ID)] != null && data[i][columnIndexFor(DISCOGS_ID)] != '';
}

//Initializes the column headers on the first row and sets up the filters each time to ensure the expected spreasheet structure remains intact.
function normalizeSheetStructure() {
    reloadSpreadsheet();
    for (var i = 1; i <= sortableColumnNames.length; i++) {
        sheet.getRange(1, i).setValue(sortableColumnNames[i - 1].displayName);
        sheet.getRange(1, i).setNote(sortableColumnNames[i - 1].description);
    }

    for (var i = 2; i <= data.length; i++){
        initTotalCostFormula(i);
    }

    //Reset sticky row and filter
    sheet.setFrozenRows(0);
    if (sheet.getFilter() != null) {sheet.getFilter().remove();}

    //Initialize filter and sticky row.
    sheet.getRange(1, 1, data.length, sortableColumnNames.length).createFilter();
    sheet.setFrozenRows(1);

    createCollectionSummaryBox();
    createSettingsBox();
    createLinks();
    sheet.getRange(convertIndexToLetter(columnIndexFor(ARTIST) + 1) + ":" + convertIndexToLetter(columnIndexFor(ARTIST) + 1)).setHorizontalAlignment("left");
    sheet.getRange(convertIndexToLetter(columnIndexFor(ALBUM) + 1) + ":" + convertIndexToLetter(columnIndexFor(ALBUM) + 1)).setHorizontalAlignment("left");
    sheet.getRange(convertIndexToLetter(columnIndexFor(DISCOGS_ID) + 1) + ":" + convertIndexToLetter(columnIndexFor(DISCOGS_ID) + 1)).setHorizontalAlignment("left");
    sheet.getRange(convertIndexToLetter(summaryBoxColumnOffset) + ":" + convertIndexToLetter(summaryBoxColumnOffset)).setHorizontalAlignment("left");
    sheet.getRange(convertIndexToLetter(summaryBoxColumnOffset + 1) + ":" + convertIndexToLetter(summaryBoxColumnOffset + 1)).setHorizontalAlignment("left");
    sheet.autoResizeColumns(1, sortableColumnNames.length - 1);
}

function createCollectionSummaryBox(){
  sheet.getRange(summaryBoxRowOffset - 1, summaryBoxColumnOffset).setValue("Collection Summary");
  sheet.getRange(summaryBoxRowOffset - 1, summaryBoxColumnOffset + 1).setValue("Values");
  sheet.getRange(summaryBoxRowOffset - 1, summaryBoxColumnOffset, 1, 2).setBackground(rgbToHex(INFO_BOX_HEADER_COLOR[0], INFO_BOX_HEADER_COLOR[1], INFO_BOX_HEADER_COLOR[2]));
  sheet.getRange(summaryBoxRowOffset - 1, summaryBoxColumnOffset, 1, 2).setBorder(true, true, true, true, false, false, "#000000", null);
  
  for(var i = 0; i < summaryRows.length; i++){
    sheet.getRange(i + summaryBoxRowOffset, summaryBoxColumnOffset).setValue(summaryRows[i].displayName);
    sheet.getRange(i + summaryBoxRowOffset, summaryBoxColumnOffset).setNote(summaryRows[i].description);
    sheet.getRange(i + summaryBoxRowOffset, summaryBoxColumnOffset).setBackground(rgbToHex(INFO_BOX_SUB_HEADER_COLOR[0], INFO_BOX_SUB_HEADER_COLOR[1], INFO_BOX_SUB_HEADER_COLOR[2]));
  }

  sheet.getRange(summaryBoxRowOffset, summaryBoxColumnOffset, summaryRows.length, 2).setBorder(true, true, true, true, true, false, "#000000", null);

  var itemInvestmentCostFormula = "=SUM(" + convertIndexToLetter(columnIndexFor(PRICE) + 1) + ":" + convertIndexToLetter(columnIndexFor(PRICE) + 1) + ")";
  var totalInvestmentCostFormula = "=SUM(" + convertIndexToLetter(columnIndexFor(TOTAL) + 1) + ":" + convertIndexToLetter(columnIndexFor(TOTAL) + 1) + ")";
  var totalDiscogsLowestFormula = "=SUM(" + convertIndexToLetter(columnIndexFor(DISCOGS_LOWEST) + 1) + ":" + convertIndexToLetter(columnIndexFor(DISCOGS_LOWEST) + 1) + ")";
  var totalReloadDiffFormula = "=SUM(" + convertIndexToLetter(columnIndexFor(RELOAD_DIFF) + 1) + ":" + convertIndexToLetter(columnIndexFor(RELOAD_DIFF) + 1) + ")";

  sheet.getRange(summaryBoxRowOffset + summaryRowIndexFor(ITEM_INVESTMENT), summaryBoxColumnOffset + 1).setFormula(itemInvestmentCostFormula);
  sheet.getRange(summaryBoxRowOffset + summaryRowIndexFor(TOTAL_INVESTMENT), summaryBoxColumnOffset + 1).setFormula(totalInvestmentCostFormula);
  sheet.getRange(summaryBoxRowOffset + summaryRowIndexFor(TOTAL_DISCOGS_LOWEST), summaryBoxColumnOffset + 1).setFormula(totalDiscogsLowestFormula);
  sheet.getRange(summaryBoxRowOffset + summaryRowIndexFor(TOTAL_RELOAD_DIFF), summaryBoxColumnOffset + 1).setFormula(totalReloadDiffFormula);
}

function createSettingsBox(){
  sheet.getRange(settingsBoxRowOffset - 1, settingsBoxColumnOffset).setValue("Settings");
  sheet.getRange(settingsBoxRowOffset - 1, settingsBoxColumnOffset + 1).setValue("Values");
  sheet.getRange(settingsBoxRowOffset - 1, settingsBoxColumnOffset, 1, 2).setBackground(rgbToHex(INFO_BOX_HEADER_COLOR[0], INFO_BOX_HEADER_COLOR[1], INFO_BOX_HEADER_COLOR[2]));
  sheet.getRange(settingsBoxRowOffset - 1, settingsBoxColumnOffset, 1, 2).setBorder(true, true, true, true, false, false, "#000000", null);
  
  for(var i = 0; i < settingsRows.length; i++){
    sheet.getRange(i + settingsBoxRowOffset, settingsBoxColumnOffset).setValue(settingsRows[i].displayName);
    sheet.getRange(i + settingsBoxRowOffset, settingsBoxColumnOffset).setNote(settingsRows[i].description);
    sheet.getRange(i + settingsBoxRowOffset, settingsBoxColumnOffset).setBackground(rgbToHex(INFO_BOX_SUB_HEADER_COLOR[0], INFO_BOX_SUB_HEADER_COLOR[1], INFO_BOX_SUB_HEADER_COLOR[2]));
  }

  sheet.getRange(settingsBoxRowOffset, settingsBoxColumnOffset, settingsRows.length, 2).setBorder(true, true, true, true, true, false, "#000000", null);

  if(sheet.getRange(settingsBoxRowOffset + settingsRowIndexFor(PROFIT_LOSS_THRESHOLD), settingsBoxColumnOffset + 1).getValue() == ''){
    sheet.getRange(settingsBoxRowOffset + settingsRowIndexFor(PROFIT_LOSS_THRESHOLD), settingsBoxColumnOffset + 1).setValue("10");
  }
  
  setColorDefault(LOSS_COLOR, "#ffb3ba");
  setColorDefault(BREAK_EVEN_COLOR, "#ffffba");
  setColorDefault(PROFIT_COLOR, "#baffc9");
  setColorDefault(NOT_LISTED_COLOR, "#ffffff");
  setColorDefault(MISSING_ID_COLOR, "#ffb3ba");
}

function createLinks(){
  sheet.getRange(linksBoxRowOffset, linksBoxColumnOffset).setFormula('=HYPERLINK("https://github.com/t8ortot/sheets-discogs-tracker", "Check for latest updates and documentation to the script.")');
  sheet.getRange(linksBoxRowOffset + 1, linksBoxColumnOffset).setFormula('=HYPERLINK("https://paypal.me/t8ortot?country.x=US&locale.x=en_US", "Like it? Donate to show appreciation!")');
  sheet.getRange(linksBoxRowOffset + 2, linksBoxColumnOffset).setFormula('=HYPERLINK("https://discord.gg/qsQ2CZ8rcS", "Question, issue, or suggestion? Join my Discord!")');
}

function setColorDefault(colorSetting, defaultColorHex){
  var valueCell = sheet.getRange(settingsBoxRowOffset + settingsRowIndexFor(colorSetting), settingsBoxColumnOffset + 1);
  
  if (valueCell.getValue() != '' && defaultColorHex != valueCell.getBackground()) {
    valueCell.setValue("override");
  } else {
    valueCell.setBackground(defaultColorHex);
    valueCell.setValue("default");
  }
}

function loadUserCollection(){
  var userName = sheet.getRange(settingsBoxRowOffset + settingsRowIndexFor(USERNAME), settingsBoxColumnOffset + 1).getValue();
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

function loadDiscogsData(){
  for (var i = 1; i < data.length; i++) {
        Logger.log(DISCOGS_ID.displayName + ': ' + data[i][columnIndexFor(DISCOGS_ID)] + '  ' + ARTIST.displayName + ': ' + data[i][columnIndexFor(ARTIST)] + '   ' + ALBUM.displayName + ': ' + data[i][columnIndexFor(ALBUM)]);
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
  var missing_id_color = hexToRgb(sheet.getRange(settingsBoxRowOffset + settingsRowIndexFor(MISSING_ID_COLOR), settingsBoxColumnOffset + 1).getBackground());
  sheet.getRange(rowNumber, 1, 1, sortableColumnNames.length).setBackground(rgbToHex(missing_id_color.r, missing_id_color.g, missing_id_color.b));
}

//Resets entire row color
function resetRowColor(rowNumber) {
    sheet.getRange(rowNumber, 1, 1, sortableColumnNames.length).setBackground("transparent");
}

//Uses the Discogs ID to query the Discogs API. Copies over relevent information into the spreadsheet such as the lowest price, album name, and artist name.
function updateRowWithDiscogsData(rowNumber) {
    var api_url = 'https://api.discogs.com/releases/' + data[rowNumber - 1][columnIndexFor(DISCOGS_ID)];
    var web_url = 'https://www.discogs.com/release/' + data[rowNumber - 1][columnIndexFor(DISCOGS_ID)];
    var response = UrlFetchApp.fetch(api_url);
    var json = JSON.parse(response.getContentText());
    Logger.log('Lowest price: ' + json.lowest_price);
    Logger.log('Discogs Album Name: ' + json.title);
    sheet.getRange(rowNumber, columnIndexFor(DISCOGS_ID) + 1).setFormula('=HYPERLINK("' + web_url + '", "' + data[rowNumber - 1][columnIndexFor(DISCOGS_ID)] + '")');
    sheet.getRange(rowNumber, columnIndexFor(DISCOGS_LOWEST) + 1).setValue(numberToCurrency.format(json.lowest_price));
    sheet.getRange(rowNumber, columnIndexFor(ALBUM) + 1).setValue(json.title);
    var artistName = json.artists[0].name.replace(/ \(.*\)/g, '');
    sheet.getRange(rowNumber, columnIndexFor(ARTIST) + 1).setValue(artistName);
}

//Updates the color of the value cell to represent the profit or loss compared to the purchased price.
function updateColor(rowNumber) {
    var total = data[rowNumber - 1][columnIndexFor(TOTAL)];
    var lowest = data[rowNumber - 1][columnIndexFor(DISCOGS_LOWEST)];
    var profit_loss_threshold = sheet.getRange(settingsBoxRowOffset + settingsRowIndexFor(PROFIT_LOSS_THRESHOLD), settingsBoxColumnOffset + 1).getValue() / 100;
    var minimum_color = hexToRgb(sheet.getRange(settingsBoxRowOffset + settingsRowIndexFor(LOSS_COLOR), settingsBoxColumnOffset + 1).getBackground());
    var neutral_color = hexToRgb(sheet.getRange(settingsBoxRowOffset + settingsRowIndexFor(BREAK_EVEN_COLOR), settingsBoxColumnOffset + 1).getBackground());
    var maximum_color = hexToRgb(sheet.getRange(settingsBoxRowOffset + settingsRowIndexFor(PROFIT_COLOR), settingsBoxColumnOffset + 1).getBackground());
    var zero_color = hexToRgb(sheet.getRange(settingsBoxRowOffset + settingsRowIndexFor(NOT_LISTED_COLOR), settingsBoxColumnOffset + 1).getBackground());

    //This will calculate the percentage or profit or loss, keeping it within the +/- THRESHOLD_PERCENTAGE bounds to set static colors past the specified range.
    var diffPercentage = Math.max(-profit_loss_threshold, Math.min((lowest / total) - 1.00, profit_loss_threshold));

    //if the lowest value is 0, meaning there is nothing listed, the color is set to NOT_LISTED_COLOR
    //if the diff is negative, we need to calculate the gradiant color between the MINUMUM_COLOR and BREAK_EVEN_COLOR
    //if the diff is positive, we need to calculate the gradiant color between the BREAK_EVEN_COLOR and PROFIT_COLOR
    if(lowest == 0){
      sheet.getRange(rowNumber, columnIndexFor(DISCOGS_LOWEST) + 1).setBackground(rgbToHex(zero_color.r, zero_color.g, zero_color.b));
    } else if (diffPercentage < 0) {
        var r = calculateGradientColor(neutral_color.r, minimum_color.r, neutral_color.r, diffPercentage);
        var g = calculateGradientColor(neutral_color.g, minimum_color.g, neutral_color.g, diffPercentage);
        var b = calculateGradientColor(neutral_color.b, minimum_color.b, neutral_color.b, diffPercentage);
        sheet.getRange(rowNumber, columnIndexFor(DISCOGS_LOWEST) + 1).setBackground(rgbToHex(r, g, b));
    } else {
        var r = calculateGradientColor(maximum_color.r, neutral_color.r, neutral_color.r, diffPercentage)
        var g = calculateGradientColor(maximum_color.g, neutral_color.g, neutral_color.g, diffPercentage)
        var b = calculateGradientColor(maximum_color.b, neutral_color.b, neutral_color.b, diffPercentage)
        sheet.getRange(rowNumber, columnIndexFor(DISCOGS_LOWEST) + 1).setBackground(rgbToHex(r, g, b));
    }
}

//Method uses the slope-intercept formula to calculate gradiant color values between the bound and 0
function calculateGradientColor(greaterColor, lesserColor, zeroIntercept, margin) {
  var profit_loss_threshold = sheet.getRange(settingsBoxRowOffset + settingsRowIndexFor(PROFIT_LOSS_THRESHOLD), settingsBoxColumnOffset + 1).getValue() / 100;
    return Math.round(((greaterColor - lesserColor) / profit_loss_threshold) * margin + zeroIntercept);
}

//Calculates and sets the difference in price between now and the last time the data was updated.
function updateRefreshDiff(oldPrice, rowNumber) {
    var diff = (data[rowNumber - 1][columnIndexFor(DISCOGS_LOWEST)] - oldPrice).toFixed(2);
    sheet.getRange(rowNumber, columnIndexFor(RELOAD_DIFF) + 1).setValue(numberToCurrency.format(diff));
}

//Updates the last reload date to be today
function updateRefreshDate(rowNumber) {
    sheet.getRange(rowNumber, columnIndexFor(LAST_RELOAD_DATE) + 1).setValue(TODAY);
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

function hexToRgb(hex) {
  var result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? {
    r: parseInt(result[1], 16),
    g: parseInt(result[2], 16),
    b: parseInt(result[3], 16)
  } : null;
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
    return data[i][columnIndexFor(LAST_RELOAD_DATE)] == null || data[i][columnIndexFor(LAST_RELOAD_DATE)] == '' || sheet.getRange(i + 1, columnIndexFor(LAST_RELOAD_DATE) + 1).getDisplayValues() != TODAY;
}

//Finds the index number in the sortableColumnNames for the column header name.
//This allows the user to move the column headers around on the sortableColumnNames array without any additional code changes.
function columnIndexFor(columnName) {
    return sortableColumnNames.findIndex(c => columnName.displayName == c.displayName);
}

//Finds the index number in the summaryRows for the row header name.
//This allows the user to move the row headers around on the summaryRows array without any additional code changes.
function summaryRowIndexFor(rowName) {
    return summaryRows.findIndex(c => rowName.displayName == c.displayName);
}

//Finds the index number in the settingsRows for the row header name.
//This allows the user to move the row headers around on the settingsRows array without any additional code changes.
function settingsRowIndexFor(rowName) {
    return settingsRows.findIndex(c => rowName.displayName == c.displayName);
}

function createHeader(displayName, description) {
  return {
    displayName: displayName,
    description: description
  }
}
