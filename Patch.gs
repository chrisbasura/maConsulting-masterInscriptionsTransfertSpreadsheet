/*
function patchAllCommercialLeads() {

  const folders = DriveApp.getFolderById("1DDFWCEEFIEpG6FQObkeCRLZFGnZtH7aX").getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    if (folder.getName().indexOf("Régie") != -1) { 
      var regie = folder.getName().replace("Régie ","");
      const subFolders = folder.getFoldersByName("Commerciaux Leads");
      while (subFolders.hasNext()) {
        const subfolder = subFolders.next();
        const subfolderFiles = subfolder.getFiles();
        while (subfolderFiles.hasNext()) {
          const subfolderFile = subfolderFiles.next();
          if (subfolderFile.getName().toLowerCase().indexOf(" leads") != -1) {
            var spreadsheet = SpreadsheetApp.openById(subfolderFile.getId());
            spreadsheet.getSheetByName("LEADS").insertColumnsBefore(10, 1);
            spreadsheet.getRange('J1').setValue('Commentaire');
            spreadsheet.getRange('O1').setValue('Détails relance');
            spreadsheet.getSheetByName("Parametres techniques").getRange('D1:D11').setValues([["Liste de status"],["Pas intéressé"],["Pas éligible"],["Doublon"],["Faux numéro"],["A rappeler"],["NRP2"],["NRP3"],["NRP4"],["NRP4+"],["Signé"]]);
            spreadsheet.getSheetByName("Parametres techniques").getRange('D1').setBackground('#fce5cd');
            spreadsheet.getSheetByName("LEADS").getRange('I2:I530').setDataValidation(SpreadsheetApp.newDataValidation()
            .setAllowInvalid(false)
            .setHelpText('Toute autre information est à rajouter dans la colonne Commentaires')
            .requireValueInRange(spreadsheet.getRange('\'Parametres techniques\'!$D$2:$D'), true)
            .build());
          }
        }
      }
    }
  }  
}
*/

function patchAllInscriptions() {
  const folders = DriveApp.getFolderById("1DDFWCEEFIEpG6FQObkeCRLZFGnZtH7aX").getFolders();
  var listOfInscriptions = [];
  while (folders.hasNext()) {
    const folder = folders.next();
    if (folder.getName().indexOf("Régie") != -1) {      
      const subFolders = folder.getFoldersByName("Commerciaux inscriptions");
      while (subFolders.hasNext()) {
        const subfolder = subFolders.next();
        const subfolderFiles = subfolder.getFiles();
        while (subfolderFiles.hasNext()) {
          const subfolderFile = subfolderFiles.next();
          if (subfolderFile.getName().toLowerCase().indexOf(" inscription") != -1) {
            listOfInscriptions.push({name: subfolderFile.getName(), id: subfolderFile.getId()});
          }
        }
      }
    }
  }
  for (var i=0;i<listOfInscriptions.length;i++) {
    SpreadsheetApp.openById(listOfInscriptions[i].id).getSheetByName("Inscriptions").getRange("J1").setValue("Numéro de dossier");
  }  
}

function addOldInscriptions() {
  var now = new Date();
  var couleurPaye = "#4a86e8";
  var couleurAnnule = "#ff0000";
  var couleurFacture = "#00ff00";
  var couleurEtape1et2Uniquement = "#ff9900";

  const sujets = ["Allemand","Anglais","Espagnol","Italien","Portugais","Russe","Multi langue","Français","Excel","Outlook","PowerPoint","Word"];

  var regie = "Elie";
  const moussafSpreadsheetId = "12U9xJAS1HkNuNMhCTCg0vlmwwIQp9Flztr-eZfvtzvE";
  const inscriptionsSheet = SpreadsheetApp.openById(moussafSpreadsheetId).getSheetByName("2020");
  const sheetData = inscriptionsSheet.getDataRange().getValues();
  const sheetBackground = inscriptionsSheet.getDataRange().getBackgrounds();
  var commercialRows = {};
  var regieRows = {};
  var generalRows = [];

  var commercialSpreadsheetID = {};
  var regieSpreadsheetID = {};
  const folders = DriveApp.getFolderById("1DDFWCEEFIEpG6FQObkeCRLZFGnZtH7aX").getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    if (folder.getName().indexOf("Régie") != -1) {  
      var files = folder.getFiles();
      while (files.hasNext()) {
      var file = files.next();
        if (file.getName().toLowerCase().indexOf("inscriptions ") != -1) {
          regieSpreadsheetID[file.getName().replace("Inscriptions ","")] = file.getId();
        }
      }

      const subFolders = folder.getFoldersByName("Commerciaux inscriptions");
      while (subFolders.hasNext()) {
        const subfolder = subFolders.next();
        const subfolderFiles = subfolder.getFiles();
        while (subfolderFiles.hasNext()) {
          const subfolderFile = subfolderFiles.next();
          if (subfolderFile.getName().toLowerCase().indexOf(" inscriptions") != -1) {
            commercialSpreadsheetID[subfolderFile.getName().replace(" inscriptions","")] = subfolderFile.getId();
          }
        }
      }
    }
  }

  for (let i=1;i<sheetData.length;i++) {

    const couleurFond = sheetBackground[i][0];    
    if (couleurFond != couleurAnnule) {

      const dateAppel = sheetData[i][0];
      const prenom = Trainingmanagementlibrary.capitalizeFirstLetter(sheetData[i][1]);
      const nom = sheetData[i][2].toUpperCase();
        if (nom != "" && prenom != "") {
        const email = sheetData[i][3];
        const phone = sheetData[i][4];
        var sujetFormation = sheetData[i][5].trim();
        if (sujets.indexOf(sujetFormation) == -1) {
          sujetFormation += " (A CORRIGER)";
        }
        const status = sheetData[i][6];
        const dateDebutSouhaitee = sheetData[i][7];
        const montant = sheetData[i][8];
        const nomCommercial = sheetData[i][9].trim();

        if (commercialRows[nomCommercial] == null) {
          commercialRows[nomCommercial] = [];
        }
        var commercialRow = [dateAppel, prenom, nom, email, phone, sujetFormation, status, dateDebutSouhaitee, montant, "", "X"];
        commercialRows[nomCommercial].push(commercialRow);

        if (regieRows[regie] == null) {
          regieRows[regie] = [];
        }
        var regieRow = [dateAppel, prenom, nom, email, phone, sujetFormation, status, dateDebutSouhaitee, montant, "", nomCommercial, "", now];
        regieRows[regie].push(regieRow);
        if (couleurFond != couleurEtape1et2Uniquement) {
          var generalRow = [dateAppel, prenom, nom, email, phone, sujetFormation, status, dateDebutSouhaitee, montant, "", nomCommercial, regie, "X", "Re-integration anciennes inscriptions", "", "", "", "", "", "", "", (couleurFond == couleurFacture || couleurFond == couleurPaye)?"X":"", (couleurFond == couleurPaye)?"X":""];
          generalRows.push(generalRow);
        }
      }
    }
  }

  //Logger.log(regieRows);

  for(var commercial in commercialRows)
  {
    if(commercialRows.hasOwnProperty(commercial)) {
      if (commercialSpreadsheetID[commercial]) {
        const commercialInscriptionSheet = SpreadsheetApp.openById(commercialSpreadsheetID[commercial]).getSheetByName("Inscriptions");
        const rowsToInsert = commercialRows[commercial];
        //commercialInscriptionSheet.getRange(commercialInscriptionSheet.getLastRow()+1,1,rowsToInsert.length,rowsToInsert[0].length).setValues(rowsToInsert);
        Logger.log("Fichier commercial mis à jour pour "+commercial);
      }
      else {
        Logger.log("Fichier commercial non trouvé pour "+commercial);
      }
    }
  }
  for(var reg in regieRows)
  {
    if(regieRows.hasOwnProperty(reg)) {
      if (regieSpreadsheetID[reg]) {
        const regieInscriptionSheet = SpreadsheetApp.openById(regieSpreadsheetID[regie]).getSheetByName("Inscriptions");
        const rowsToInsert = regieRows[reg];
        //regieInscriptionSheet.getRange(regieInscriptionSheet.getLastRow()+1,1,rowsToInsert.length,rowsToInsert[0].length).setValues(rowsToInsert);
        Logger.log("Fichier regie mis à jour pour "+reg);
      }
      else {
        Logger.log("Fichier regie non trouvé pour "+reg);
      }
    }
  }
  const generalInscriptionSheet = SpreadsheetApp.openById("18D5_FEyvHmYzm2lFfs7wwa1ZQrNhERFg-tx-fL4TWOg").getSheetByName("Inscriptions");
  //generalInscriptionSheet.getRange(generalInscriptionSheet.getLastRow()+1,1,generalRows.length,generalRows[0].length).setValues(generalRows);
}