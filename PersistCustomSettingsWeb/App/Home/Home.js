// Declare global variables for the get/set functions and the storage type.
var _getSettings;
var _setSettings;
var _storageMode;

// Add any initialization logic to this function
Office.initialize = function (reason) {

    // Use the Apps for Office property bag to store data for the app.
    // This is the default storage method.
    _setSettings = function (key, value) { StorageLibrary.saveToPropertyBag(key, value) }
    _getSettings = function (key) { return StorageLibrary.getFromPropertyBag(key) }

    $("#footer").hide();

    // TODO: If you wanted to save the settings stored in the app's property 
    // bag before the app is closed - for instance, for saving app state - 
    // add a handler to the Internet Explorer window.onunload event.
    //window.onunload = function () {

    //    if (Office.context.document.settings) {
    //        Office.context.document.settings.saveAsync();
    //    }
    //};
}

// Sets the type of storage to be used for the plug-in for Office.
function setStorageMode() {

    try {

        // Get the selected option fromt the drop-down list.
        var selectionList = document.getElementById("storagetype");
        var index = selectionList.selectedIndex;
        var modeSelected = selectionList.options[index];
        var mode = modeSelected.value;

        switch (mode) {

            case "PropertyBag":
                // Use the app for Office property bag to store and retrieve data.
                _setSettings = function (key, value) { StorageLibrary.saveToPropertyBag(key, value) }
                _getSettings = function (key) { return StorageLibrary.getFromPropertyBag(key) }

                break;

            case "Cookies":
                // Use browser cookies to store and retrieve data.
                if (navigator.cookieEnabled == true) {
                    _setSettings = function (key, value) { StorageLibrary.saveToBrowserCookies(key, value) }
                    _getSettings = function (key) { return StorageLibrary.getFromBrowserCookies(key) }
                }
                else {

                    var browserError = { name: "Error", message: "Browser cookies are disabled. You may want to enable them." }
                    throw browserError;
                }
                break;

            case "LocalStorage":
                // Use Web Storage to store and retrieve data - storage won't expire.
                if (typeof (Storage) !== "undefined") {
                    _setSettings = function (key, value) { StorageLibrary.saveToLocalStorage(key, value) }
                    _getSettings = function (key) { return StorageLibrary.getFromLocalStorage(key) }
                }
                else {
                    var webStorageError = { name: "Error", message: "Browser storage not available in your browser (sorry)." }
                    throw webStorageError;
                }
                break;

            case "SessionStorage":
                // Use Web Storage to store and retrieve data, limited to the lifetime of the browser window.
                if (typeof (Storage) !== "undefined") {
                    _setSettings = function (key, value) { StorageLibrary.saveToSessionStorage(key, value) }
                    _getSettings = function (key) { return StorageLibrary.getFromSessionStorage(key) }
                }
                else {
                    var webStorageError = { name: "Error", message: "Browser storage not available in your browser (sorry)." }
                    throw webStorageError;
                }
                break;

            case "Document":
                // Use a programmatically created, hidden <div> to store and retrieve data.
                _setSettings = function (key, value) { StorageLibrary.saveToDocument(key, value) }
                _getSettings = function (key) { return StorageLibrary.getFromDocument(key) }
                break;

        }
        Toast.showToast("Switched storage", modeSelected.text);
    }
    catch (err) {

        Toast.showToast(err.name, err.message);

    }
}

// Toggles between the Storage and Settings "pages."
function switchPage(node) {

    // Get references to the two "pages" (divs).
    var savePage = document.getElementById("saveui");
    var getPage = document.getElementById("getui");
    var settingsPage = document.getElementById("appsettings");
    var pageTitle = document.getElementById("page");

    // Determine which page to show. Default to the storage page.
    switch (node) {
        case "Settings":
            savePage.setAttribute("class", "hiddenpage");
            getPage.setAttribute("class", "hiddenpage");
            settingsPage.setAttribute("class", "displayedpage");
            //page.innerHTML = "Settings";
            page.innerText = "Settings";
            break;

        default:
            savePage.setAttribute("class", "displayedpage");
            getPage.setAttribute("class", "displayedpage");
            settingsPage.setAttribute("class", "hiddenpage");
            //page.innerHTML = "Storage";
            page.innerText = "Storage";
            break;
    }
}

// Saves the specified settings typed into the textboxes.
function setSettings() {
    try {
        var mySetting = document.getElementById('newSetting');
        var myValue = document.getElementById('newValue');

        _setSettings(mySetting.value, myValue.value);

        changeText(mySetting, true);
        changeText(myValue, true);
    }
    catch (err) {

        Toast.showToast(err.name, err.message);
    }
}

// Gets the saved setting using the name typed into the textbox.
// Results are displayed in a toast at the bottom of the plug-in.
function getSettings() {
    try {
        var settingName = document.getElementById('storedSetting');
        var settingValue = _getSettings(settingName.value);

        if (settingValue == null) {
            Toast.showToast("Error", "No setting by that name.");
        }
        else {
            Toast.showToast("Info Retrieved", "Key: " + settingName.value + ", Value: " + settingValue);
        }

        changeText(settingName, true);
    }
    catch (err) {
        Toast.showToast(err.name, err.message);
    }
}

// Changes the UI text for onfocusin, onfocusout, and save/get events.
function changeText(node, isSave) {
    var output = document.getElementById('settingVal');
    var inputId = node.id;
    var inputText = "Type a setting ";

    switch (inputId) {
        case "newSetting":
            inputText += "name";
            break;
        case "newValue":
            inputText += "value";
            break;
        case "storedSetting":
            inputText += "to get";
            break;
    }

    if (node.value == inputText) { node.value = ""; }
    else if (node.value != "" & !isSave) { }
    else { node.value = inputText; }

}