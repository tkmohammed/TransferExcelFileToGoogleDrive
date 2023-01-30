/// <summary>
/// When particular text message is sent in the Line group, Line bot will call a rest POST method,
/// and the message will be recorded in the particular google sheet
/// </summary>
/// <input="e">The POST event sent from Line by the bot</input>
function doPost(e) {
    // Retrieve the message
    var data = JSON.parse(e.postData.contents);
    var events = data.events;
    var event = events[0];
    var file_name = "ERROR";
    var FOLDER_ID = "ID OF THE FOLDER WHERE YOU WANT TO SAVE THE FILE";
    var ACCESS_TOKEN = "ACCESS TOKEN FROM LINE DEVELOPERS CONSOLE";

    // Sheet information
    var spreadsheetId = "YOUR SPREADSHEET ID";
    var sheetName = "YOUR SHEET'S NAME";
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName(sheetName);

    if (event.type == 'message') {
        if (event.message.type == 'file') {
            var file_name = JSON.stringify(event.message.fileName).slice(1, -1);
            let file_id = event.message.id;
            if (file_name.includes(".xlsx")) {
                if (fileExistsInFolder(file_name, sheet)) {
                    recordDetail(event.source.userId, event.timestamp, file_name, false, sheet);
                }
                else {
                    try {
                        let file_data = getFileData(file_id, file_name, ACCESS_TOKEN);
                        let id = saveFile(file_data, FOLDER_ID);
                        recordDetail(event.source.userId, event.timestamp, file_name, id, sheet);
                    }
                    catch (e) {
                        Console.log(e);
                    }
                }
            }
        }
    }

    return ContentService.createTextOutput(JSON.stringify({ 'content': 'post ok' })).setMimeType(ContentService.MimeType.JSON);
}

/// <summary>
/// Retrieves file content from GET API
/// </summary>
/// <input="file_id">The file id from Line message JSON</input>
/// <input="file_name">The name of the retrieving file</input>
/// <input="ACCESS_TOKEN">The access token to securely make GET API call</input>
/// <returns>binary oject file content</returns>
function getFileData(file_id, file_name, ACCESS_TOKEN) {
    // endpoint for sending and receiving large amounts of data in the LINE platform for Messaging API
    const url = 'https://api-data.line.me/v2/bot/message/' + file_id + '/content';
    const data = UrlFetchApp.fetch(url,
        {
            'headers': { 'Authorization': 'Bearer ' + ACCESS_TOKEN },
            'method': 'get'
        });

    const file_data = data.getBlob().setName(file_name);
    return file_data;
}

/// <summary>
/// Checks if the file already exists in the Google drive folder
/// </summary>
/// <input="file_name">The name of the file that you want to check if it's already in Google drive folder</input>
/// <input="sheet">The active sheet</input>
/// <returns>boolean</returns>
function fileExistsInFolder(file_name, sheet) {
    var col = "B";
    var row = 4;
    var fileExists = false;

    while (!sheet.getRange(col + row).isBlank()) {
        if (sheet.getRange(col + row).getValue() === file_name) {
            fileExists = true;
            break;
        }
        row += 1;
    }

    return fileExists;
}

/// <summary>
/// Saves the file in the Google drive folder
/// </summary>
/// <input="file_data">The file content that you want to save in Google drive folder</input>
/// <input="FOLDER_ID">The id of the Google drive folder where you wish to save the file</input>
/// <returns>int file id</returns>
function saveFile(file_data, FOLDER_ID) {
    try {
        const folder = DriveApp.getFolderById(FOLDER_ID);
        const file = folder.createFile(file_data);
        return file.getId();
    } catch (e) {
        return false;
    }
}

/// <summary>
/// Record if the saving worked with detail information for the future reference
/// </summary>
/// <input="userId">The id of the user who sent the file in the Line chat</input>
/// <input="timestamp">The time when the file was sent in the chat</input>
/// <input="id">The id of the file retrieved from Google drive folder after successful save</input>
function recordDetail(userId, timestamp, file_name, id, sheet) {
    const row = sheet.getLastRow() + 1;
    let today = new Date();
    let year = today.getFullYear();
    let month = today.getMonth();
    let date = today.getDate();
    let error_message = "This file already exists in the folder. Overwrite was prevented.";
    var detail = {
        "date": `${year}-${month + 1}-${date}`,
        "fileName": file_name,
        "file_link": `https://drive.google.com/file/d/${id}`,
        "timestamp": Utilities.formatDate(new Date(timestamp), 'EST', 'yyyy-MM-dd HH:mm'),
        "userId": userId
    }

    if (id === false) {
        detail.file_link = error_message;
    }

    for (let col = 1; col < 6; col++) {
        sheet.getRange(row, col).setValue(Object.values(detail)[col - 1]);
    }

    return;
}