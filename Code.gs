// CHANGE THIS!
var apiKey = "";

// This is okay
var url = "https://api.meraki.com/api/v0";

function onOpen(e) {
  MerakiReport();
}

function MerakiReport() {
  
  // Sheet details
  var sheetName = "Meraki";  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  // Display
  var row = 1;
  sheet.getRange(row,1).setValue("Organization");
  sheet.getRange(row,2).setValue("Network");
  sheet.getRange(row,3).setValue("Device");
  sheet.getRange(row,4).setValue("First Uplink");
  sheet.getRange(row,5).setValue("Device Type");
  sheet.getRange(row,6).setValue("Firmware");
  row++;
  
  
  sheet.getRange(1,1).setValue("Getting Organizations...");
  var organizations = fetch("/organizations");
  sheet.getRange(1,1).setValue("Organizations");
  
  // organization, network, device
  sheet.getRange(1,2).setValue("Getting Networks...");
  for(var organizationIndex in organizations) {
    var organization = organizations[organizationIndex];
    var networks = fetch("/organizations/"+organization.id+"/networks");
    
    sheet.getRange(1,3).setValue("Getting Devices...");
    for(var networkIndex in networks) {
      var network = networks[networkIndex];
      var devices = fetch("/networks/"+network.id+"/devices");
      
      for(var deviceIndex in devices) {
        var device = devices[deviceIndex];
        
        var uplink = fetch("/networks/"+network.id+"/devices/"+device.serial+"/uplink");
        
        if(organization != undefined) { sheet.getRange(row,1).setValue(organization.name); }
        if(network != undefined) { sheet.getRange(row,2).setValue(network.name); }
        if(device != undefined) { sheet.getRange(row,3).setValue(device.name); }
        if(uplink[0] != undefined) { sheet.getRange(row,4).setValue(uplink[0].status) }
        if(device != undefined) {
          if(device.firmware != "Not running configured version") {
            var splitFirmware = (device.firmware).split("-");
            var deviceType = splitFirmware[0];
            splitFirmware.shift();
            var firmwareVersion = "";
            for(var splitFirmwareIndex in splitFirmware) {
              firmwareVersion = firmwareVersion + "." + splitFirmware[splitFirmwareIndex];
            }
            firmwareVersion = firmwareVersion.substring(1, firmwareVersion.length);
            sheet.getRange(row,5).setValue(deviceType);
            sheet.getRange(row,6).setValue(firmwareVersion);
          }
        }
        
        // Go to the next row
        row++;
      }
    }    
  }
  
  sheet.getRange(1,2).setValue("Networks");
  sheet.getRange(1,3).setValue("Devices");
  
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
  Utilities.sleep(200);
  
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
