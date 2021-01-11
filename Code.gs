function onOpen() {
  Logger.log("Set Gestion automatique transfert inscriptions Menu");
  var menu = SpreadsheetApp.getUi().createMenu("Automatisation");  
  menu.addItem('Démarrer service transfert des inscriptions (commerciaux & régie)', 'startAllInscriptionsService');
  menu.addItem('Arrêter service transfert des inscriptions (commerciaux & régie)', 'stopAllInscriptionsService');
  menu.addSeparator()
  menu.addItem('Transférer manuellement les inscriptions des commerciaux à la régie', 'transfertAllCommercialInscriptionsToRegieInscriptions');
  menu.addItem('Transférer manuellement les inscriptions des régies au général', 'transfertAllRegieInscriptionsToGeneralInscriptions');
  menu.addToUi();    
}

function startAllInscriptionsService() {
  deleteAllInscriptionsTriggers();
  setAllInscriptionsTriggers();
}

function stopAllInscriptionsService() {
  deleteAllInscriptionsTriggers();
}

function setAllInscriptionsTriggers() {
  Logger.log("Set All Inscriptions Triggers");  
  ScriptApp.newTrigger('transfertAllCommercialInscriptionsToRegieInscriptions')
  .timeBased()
  .atHour(19)
  .inTimezone("Europe/Paris")
  .everyDays(1)
  .create();
  ScriptApp.newTrigger('transfertAllRegieInscriptionsToGeneralInscriptions')
  .timeBased()
  .atHour(6)
  .inTimezone("Europe/Paris")
  .everyDays(1)
  .create();
}

function deleteAllInscriptionsTriggers() {
  Logger.log("Delete All Inscriptions Triggers");  
  // Loop over all triggers.
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
}

function transfertAllCommercialInscriptionsToRegieInscriptions() {
  var folders = DriveApp.getFolderById("1DDFWCEEFIEpG6FQObkeCRLZFGnZtH7aX").getFolders();
  var listOfInscriptions = [];
  while (folders.hasNext()) {
    var folder = folders.next();
    if (folder.getName().indexOf("Régie") != -1) {      
      var subFolders = folder.getFoldersByName("Commerciaux inscriptions");
      while (subFolders.hasNext()) {
        var subfolder = subFolders.next();
        var subfolderFiles = subfolder.getFiles();
        while (subfolderFiles.hasNext()) {
          var subfolderFile = subfolderFiles.next();
          if (subfolderFile.getName().toLowerCase().indexOf(" inscription") != -1) {
            listOfInscriptions.push({name: subfolderFile.getName(), id: subfolderFile.getId()});
          }
        }
      }
    }
  }
  var rows = [];
  var now;
  for (var i=0;i<listOfInscriptions.length;i++) {
    now = new Date();
    var result = Trainingmanagementlibrary.transfertCommercialInscriptionsToRegieInscriptions(SpreadsheetApp.openById(listOfInscriptions[i].id));
    rows.push([now, listOfInscriptions[i].name, listOfInscriptions[i].id, result.status, result.description]);
  }
  if (rows.length>0) {
    var reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Commerciaux->Regie (Etape 1->2)");
    if (now.getDay() == 1)
      reportSheet.getRange(2,1,reportSheet.getLastRow()-1,reportSheet.getLastColumn()).clearContent();
    reportSheet.getRange(reportSheet.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
  }
}

function transfertAllRegieInscriptionsToGeneralInscriptions() {
  var folders = DriveApp.getFolderById("1DDFWCEEFIEpG6FQObkeCRLZFGnZtH7aX").getFolders();
  var listOfInscriptions = [];
  while (folders.hasNext()) {
    var folder = folders.next();
    if (folder.getName().indexOf("Régie") != -1) {      
      var files = folder.getFiles();
      while (files.hasNext()) {
      var file = files.next();
        if (file.getName().toLowerCase().indexOf("inscriptions ") != -1) {
          listOfInscriptions.push({name: file.getName(), id: file.getId()});
        }
      }
    }
  }
  var rows = [];
  var now;
  for (var i=0;i<listOfInscriptions.length;i++) {
    now = new Date();
    var result = Trainingmanagementlibrary.transfertRegieInscriptionsToGeneralInscriptions(SpreadsheetApp.openById(listOfInscriptions[i].id));
    rows.push([now, listOfInscriptions[i].name, listOfInscriptions[i].id, result.status, result.description]);
  }
  if (rows.length>0) {
    var reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Regie->General (Etape 2->3)");
    if (now.getDay() == 1)
      reportSheet.getRange(2,1,reportSheet.getLastRow()-1,reportSheet.getLastColumn()).clearContent();
    reportSheet.getRange(reportSheet.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
  }
}