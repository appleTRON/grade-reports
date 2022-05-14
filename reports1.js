function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('AutoFill Docs');
    menu.addItem('Create Reports', 'createreports');
    menu.addToUi();
  }
  
  function createreports() {
    
    const googleDocTemplate = DriveApp.getFileById('1MBhcWkAGb6Wz85KdtE_KpSRlseBOutuaICrX557Wm6Q');
    const destinationFolder = DriveApp.getFolderById('1QmWFcH5Lndo8VazyXHWd-cV_U98aiEmr')
    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Sheet1')
    const rows = sheet.getDataRange().getValues();
  
    var madeof = ['made of','composed of','comprised of','derived from'];
    var anumberof = ['a number of','a few','many','a bunch of'];
    var components = ['components','parts','pieces','factors'];
    var reflected = ['reflected','shown','demonstrated','recorded'];
    var totals = ['adds up to','sums to','makes','coalesces in'];
  
    var qualaplus = ['exceptional','stellar','superb','outstanding'];
    var freqaplus = "always";
  
    var quala = ['excellent','wonderful','great','exemplary'];
    var freqa = "almost always";
  
    var qualaminus = ['very good','distinguished','favorable','reputable'];
    var freqaminus = "regularly";
  
    var qualbplus = ['good','quality','admirable','skillful'];
    var freqbplus = "usually";
  
    var qualb = ['positive','commendable','proficient','accomplished'];
    var freqb = "frequently";
  
    var qualbminus = ['decent','acceptable','solid','respectable'];
    var freqbminus = "often";
  
    var qualcplus = ['satisfactory','adequate','sufficient','capable'];
    var freqcplus = "sometimes";
  
    var qualc = ['fair','competent','average','moderate'];
    var freqc = "irregularly";
  
    var qualcminus = ['lackluster','uninspired','mediocre	','second-rate'];
    var freqcminus = "infrequently"
  
    var qualdplus = ['less than satisfactory','unfortunate','flat','flimsy'];
    var freqdplus = "rarely"
  
    var quald = ['disappointing','inadequate','weak','insubstantial'];
    var freqd = "seldom";
  
    var qualdminus = ['insufficient','poor','unsatisfactory','lacking'];
    var freqdminus = "almost never";
  
    var qualf = ['incompetent','deficient','unacceptable','inferior'];
    var freqf = "never";
  
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
  
      if (row[4] == "A+") {
        var classquality = qualaplus[Math.floor(Math.random()*qualaplus.length)];
      } 
      else if (row[4] == "A") {
        var classquality = quala[Math.floor(Math.random()*quala.length)];
      }
      else if (row[4] == "A-") {
        var classquality = qualaminus[Math.floor(Math.random()*qualaminus.length)];
      }
      else if (row[4] == "B+") {
        var classquality = qualbplus[Math.floor(Math.random()*qualbplus.length)];
      }
      else if (row[4] == "B") {
        var classquality = qualb[Math.floor(Math.random()*qualb.length)];
      }
      else if (row[4] == "B-") {
        var classquality = qualbminus[Math.floor(Math.random()*qualbminus.length)];
      }
      else if (row[4] == "C+") {
        var classquality = qualcplus[Math.floor(Math.random()*qualcplus.length)];
      }
      else if (row[4] == "C") {
        var classquality = qualc[Math.floor(Math.random()*qualc.length)];
      }
      else if (row[4] == "C-") {
        var classquality = qualcminus[Math.floor(Math.random()*qualcminus.length)];
      }
      else if (row[4] == "D+") {
        var classquality = qualdplus[Math.floor(Math.random()*qualdplus.length)];
      }
      else if (row[4] == "D") {
        var classquality = quald[Math.floor(Math.random()*quald.length)];
      }
      else if (row[4] == "D-") {
        var classquality = qualdminus[Math.floor(Math.random()*qualdminus.length)];
      }
      else {
        var classquality = qualf[Math.floor(Math.random()*qualf.length)];
      }
  
      if (row[5] == "A+") {
        var hwquality = qualaplus[Math.floor(Math.random()*qualaplus.length)];
      } 
      else if (row[5] == "A") {
        var hwquality = quala[Math.floor(Math.random()*quala.length)];
      }
      else if (row[5] == "A-") {
        var hwquality = qualaminus[Math.floor(Math.random()*qualaminus.length)];
      }
      else if (row[5] == "B+") {
        var hwquality = qualbplus[Math.floor(Math.random()*qualbplus.length)];
      }
      else if (row[5] == "B") {
        var hwquality = qualb[Math.floor(Math.random()*qualb.length)];
      }
      else if (row[5] == "B-") {
        var hwquality = qualbminus[Math.floor(Math.random()*qualbminus.length)];
      }
      else if (row[5] == "C+") {
        var hwquality = qualcplus[Math.floor(Math.random()*qualcplus.length)];
      }
      else if (row[5] == "C") {
        var hwquality = qualc[Math.floor(Math.random()*qualc.length)];
      }
      else if (row[5] == "C-") {
        var hwquality = qualcminus[Math.floor(Math.random()*qualcminus.length)];
      }
      else if (row[5] == "D+") {
        var hwquality = qualdplus[Math.floor(Math.random()*qualdplus.length)];
      }
      else if (row[5] == "D") {
        var hwquality = quald[Math.floor(Math.random()*quald.length)];
      }
      else if (row[5] == "D-") {
        var hwquality = qualdminus[Math.floor(Math.random()*qualdminus.length)];
      }
      else {
        var hwquality = qualf[Math.floor(Math.random()*qualf.length)];
      }
  
      if (row[8] == "A+") {
        var prepfreq = freqaplus;
      } 
      else if (row[8] == "A") {
        var prepfreq = freqa;
      }
      else if (row[8] == "A-") {
        var prepfreq = freqaminus;
      }
      else if (row[8] == "B+") {
        var prepfreq = freqbplus;
      }
      else if (row[8] == "B") {
        var prepfreq = freqb;
      }
      else if (row[8] == "B-") {
        var prepfreq = freqbminus;
      }
      else if (row[8] == "C+") {
        var prepfreq = freqcplus;
      }
      else if (row[8] == "C") {
        var prepfreq = freqc;
      }
      else if (row[8] == "C-") {
        var prepfreq = freqcminus;
      }
      else if (row[8] == "D+") {
        var prepfreq = freqdplus;
      }
      else if (row[8] == "D") {
        var prepfreq = freqd;
      }
      else if (row[8] == "D-") {
        var prepfreq = freqdminus;
      }
      else {
        var prepfreq = freqf;
      }
      
      if (row[9] == "A+") {
        var partfreq = freqaplus;
      } 
      else if (row[9] == "A") {
        var partfreq = freqa;
      }
      else if (row[9] == "A-") {
        var partfreq = freqaminus;
      }
      else if (row[9] == "B+") {
        var partfreq = freqbplus;
      }
      else if (row[9] == "B") {
        var partfreq = freqb;
      }
      else if (row[9] == "B-") {
        var partfreq = freqbminus;
      }
      else if (row[9] == "C+") {
        var partfreq = freqcplus;
      }
      else if (row[9] == "C") {
        var partfreq = freqc;
      }
      else if (row[9] == "C-") {
        var partfreq = freqcminus;
      }
      else if (row[9] == "D+") {
        var partfreq = freqdplus;
      }
      else if (row[9] == "D") {
        var partfreq = freqd;
      }
      else if (row[9] == "D-") {
        var partfreq = freqdminus;
      }
      else {
        var partfreq = freqf;
      }
  
      if (row[10] == "A+") {
        var posfreq = freqaplus;
      } 
      else if (row[10] == "A") {
        var posfreq = freqa;
      }
      else if (row[10] == "A-") {
        var posfreq = freqaminus;
      }
      else if (row[10] == "B+") {
        var posfreq = freqbplus;
      }
      else if (row[10] == "B") {
        var posfreq = freqb;
      }
      else if (row[10] == "B-") {
        var posfreq = freqbminus;
      }
      else if (row[10] == "C+") {
        var posfreq = freqcplus;
      }
      else if (row[10] == "C") {
        var posfreq = freqc;
      }
      else if (row[10] == "C-") {
        var posfreq = freqcminus;
      }
      else if (row[10] == "D+") {
        var posfreq = freqdplus;
      }
      else if (row[10] == "D") {
        var posfreq = freqd;
      }
      else if (row[10] == "D-") {
        var posfreq = freqdminus;
      }
      else {
        var posfreq = freqf;
      }
      
      if (row[11] == "A+") {
        var persfreq = freqaplus;
      } 
      else if (row[11] == "A") {
        var persfreq = freqa;
      }
      else if (row[11] == "A-") {
        var persfreq = freqaminus;
      }
      else if (row[11] == "B+") {
        var persfreq = freqbplus;
      }
      else if (row[11] == "B") {
        var persfreq = freqb;
      }
      else if (row[11] == "B-") {
        var persfreq = freqbminus;
      }
      else if (row[11] == "C+") {
        var persfreq = freqcplus;
      }
      else if (row[11] == "C") {
        var persfreq = freqc;
      }
      else if (row[11] == "C-") {
        var persfreq = freqcminus;
      }
      else if (row[11] == "D+") {
        var persfreq = freqdplus;
      }
      else if (row[11] == "D") {
        var persfreq = freqd;
      }
      else if (row[11] == "D-") {
        var persfreq = freqdminus;
      }
      else {
        var persfreq = freqf;
      }
      
      if (row[12] == "A+") {
        var termquality = qualaplus[Math.floor(Math.random()*qualaplus.length)];
      } 
      else if (row[12] == "A") {
        var termquality = quala[Math.floor(Math.random()*quala.length)];
      }
      else if (row[12] == "A-") {
        var termquality = qualaminus[Math.floor(Math.random()*qualaminus.length)];
      }
      else if (row[12] == "B+") {
        var termquality = qualbplus[Math.floor(Math.random()*qualbplus.length)];
      }
      else if (row[12] == "B") {
        var termquality = qualb[Math.floor(Math.random()*qualb.length)];
      }
      else if (row[12] == "B-") {
        var termquality = qualbminus[Math.floor(Math.random()*qualbminus.length)];
      }
      else if (row[12] == "C+") {
        var termquality = qualcplus[Math.floor(Math.random()*qualcplus.length)];
      }
      else if (row[12] == "C") {
        var termquality = qualc[Math.floor(Math.random()*qualc.length)];
      }
      else if (row[12] == "C-") {
        var termquality = qualcminus[Math.floor(Math.random()*qualcminus.length)];
      }
      else if (row[12] == "D+") {
        var termquality = qualdplus[Math.floor(Math.random()*qualdplus.length)];
      }
      else if (row[12] == "D") {
        var termquality = quald[Math.floor(Math.random()*quald.length)];
      }
      else if (row[12] == "D-") {
        var termquality = qualdminus[Math.floor(Math.random()*qualdminus.length)];
      }
      else {
        var termquality = qualf[Math.floor(Math.random()*qualf.length)];
      }  
  
    if (row[13]) return;
  
      const copy = googleDocTemplate.makeCopy(`${row[0]} ${row[2]} Progress Report` , destinationFolder)
      const doc = DocumentApp.openById(copy.getId())
      const body = doc.getBody();
    
      body.replaceText('{{studentfirst}}', row[1]);
      body.replaceText('{{synmadeof1}}', madeof[Math.floor(Math.random()*madeof.length)]);
      body.replaceText('{{synmadeof2}}', madeof[Math.floor(Math.random()*madeof.length)]);
      body.replaceText('{{synanumberof1}}', anumberof[Math.floor(Math.random()*anumberof.length)]);
      body.replaceText('{{synanumberof2}}', anumberof[Math.floor(Math.random()*anumberof.length)]);
      body.replaceText('{{syncomponents1}}', components[Math.floor(Math.random()*components.length)]);
      body.replaceText('{{syncomponents2}}', components[Math.floor(Math.random()*components.length)]);
      body.replaceText('{{synreflected}}', reflected[Math.floor(Math.random()*reflected.length)]);
      body.replaceText('{{syntotals}}', totals[Math.floor(Math.random()*totals.length)]);
      body.replaceText('{{pronposl}}', pronposl);
      body.replaceText('{{pronposu}}', pronposu);
      body.replaceText('{{pronsubl}}', pronsubl);
      body.replaceText('{{pronsubu}}', pronsubu);
      body.replaceText('{{tobeconj}}', tobeconj);
      body.replaceText('{{regconj}}', regconj);
      body.replaceText('{{classquality}}', classquality);
      body.replaceText('{{classgrade}}', row[4]);
      body.replaceText('{{hwquality}}', hwquality);
      body.replaceText('{{hwgrade}}', row[5]);
      body.replaceText('{{quizgrade}}', row[6]);
      body.replaceText('{{testgrade}}', row[7]);
      body.replaceText('{{prepfreq}}', prepfreq);
      body.replaceText('{{partfreq}}', partfreq);
      body.replaceText('{{posfreq}}', posfreq);
      body.replaceText('{{persfreq}}', persfreq);
      body.replaceText('{{termquality}}', termquality);
      
      doc.saveAndClose();
     
      const url = doc.getUrl();
      sheet.getRange(index + 1, 14).setValue(url)
      
    })
  
  }