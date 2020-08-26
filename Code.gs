// CHANGE THIS!
var apiKey = "";

// This is okay
var url = "https://api.meraki.com/api/v1";

function onOpen(e) {
  MerakiReport();
}

function MerakiReport() {
  
  // Sheet details
  var sheetName = "Meraki";  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  // Display
  sheet.clear();
  var row = 1;
  var columnNames = ["Organization","Device","Firmware","Device Type","Firmware Version"];
  var columnIndex = 0;
  for(var columnIndex in columnNames) {
    sheet.getRange(row,columnIndex+1).setValue(columnNames[columnIndex]);
    columnIndex += 1;
  }
  row++;
  
  var organizations = fetch("/organizations");
  for(var organizationIndex in organizations) {
    var organization = organizations[organizationIndex];
    var devices = fetch("/organizations/"+organization.id+"/devices");
    
    for(var deviceIndex in devices) {
      var device = devices[deviceIndex];
      var networkId = device.networkId;
      var firmware = device.firmware;
      
      sheet.getRange(row,1).setValue(organization.name);
      var deviceHyperlink = '=HYPERLINK("' + device.url + '", "' + device.name + '")'
      sheet.getRange(row,2).setValue(deviceHyperlink);
      sheet.getRange(row,3).setValue(device.firmware);
      
      if(device != undefined) {
        if(device.firmware != undefined) {
          var splitFirmware = (device.firmware).split("-");
          var deviceType = splitFirmware[0];
          splitFirmware.shift();
          var firmwareVersion = "";
          for(var splitFirmwareIndex in splitFirmware) {
            firmwareVersion = firmwareVersion + "." + splitFirmware[splitFirmwareIndex];
          }
          firmwareVersion = firmwareVersion.substring(1, firmwareVersion.length);
          sheet.getRange(row,4).setValue(deviceType);
          sheet.getRange(row,5).setValue(firmwareVersion);
      
          row++; 
        } 
      }
    } 
  }
}

function fetch(path)
{ 
  var url_path = url + path;
  var options = {
    method: 'get',
    //contentType: "application/json",
    headers: {
      'X-Cisco-Meraki-API-Key': apiKey
    },
    muteHttpExceptions: true,
    followRedirects: true
  };
  
  // Wait a little bit, API rate limit at 5/sec
  Utilities.sleep(250);
  
  // Return data or follow redirects
  var response = UrlFetchApp.fetch(url_path, options); 
  
  if(response.getResponseCode() == 404) {
    Logger.log("HTTP 404 - Not Found - " + response);
    // skip
  }
  else {
    return JSON.parse(response);
  }
  
}
