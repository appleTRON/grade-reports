function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('AutoFill Docs');
    menu.addItem('Write Recs', 'writerecs');
    menu.addToUi();
  }
  
  function writerecs() {
    
    const googleDocTemplate = DriveApp.getFileById('1RrceuBXuHgK3R4JLo2Hg_ogmYvvb_y1h3zHB4dvOcq4');
    const destinationFolder = DriveApp.getFolderById('0B-L2_p1eAGDqQlBxcU9LRGltbnM')
    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Sheet1')
    const rows = sheet.getDataRange().getValues();
  
    rows.forEach(function(row, index){
      if (index === 0) return;
  
      if (row[3] == "M") {
        var pronsubu = "He";
        var pronsubl = "he";
        var pronposu = "His";
        var pronposl = "his";
        var tobeconj = "is";
        var regconj = "s";
      } 
      else if (row[3] == "F") {
        var pronsubu = "She";
        var pronsubl = "she";
        var pronposu = "Her";
        var pronposl = "her";
        var tobeconj = "is";
        var regconj = "s"
      }
      else {
        var pronsubu = "They";
        var pronsubl = "they";
        var pronposu = "Their";
        var pronposl = "their";
        var tobeconj = "are";
        var regconj = "";
      }
  
      if (row[13]) return;
  
      const copy = googleDocTemplate.makeCopy(`${row[1]} ${row[2]} Recommendation Letter - Appleton` , destinationFolder)
      const doc = DocumentApp.openById(copy.getId())
      const body = doc.getBody();
    
      body.replaceText('{{studentfirst}}', row[1]);
      body.replaceText('{{studentlast}}', row[2]);
      body.replaceText('{{class}}', row[0]);
      body.replaceText('{{word1}}', row[4]);
      body.replaceText('{{word2}}', row[5]);
      body.replaceText('{{word3}}', row[6]);
      body.replaceText('{{overalladj}}', row[7]);
      body.replaceText('{{strength1}}', row[8]);
      body.replaceText('{{strength2}}', row[9]);
      body.replaceText('{{workethic}}', row[10]);
      body.replaceText('{{problemsolving}}', row[11]);
      body.replaceText('{{recommendation}}', row[12]);
      body.replaceText('{{pronposl}}', pronposl);
      body.replaceText('{{pronposu}}', pronposu);
      body.replaceText('{{pronsubl}}', pronsubl);
      body.replaceText('{{pronsubu}}', pronsubu);
      body.replaceText('{{tobeconj}}', tobeconj);
      body.replaceText('{{regconj}}', regconj);
  
      
      doc.saveAndClose();
     
      const url = doc.getUrl();
      sheet.getRange(index + 1, 14).setValue(url)
      
    })
  
  }