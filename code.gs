function onInstall(e){
  onOpen(e);
}

function onOpen() {
  SpreadsheetApp.getUi().createAddonMenu()
         .addItem('Create Exam', 'showSidebar') 
         .addToUi();
}

// Show menu
function showSidebar(){
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
       .setTitle('Create exam');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function createExam() {
  SpreadsheetApp.getActiveSpreadsheet().toast("Creating exam...");
  
  var general = "general!";
    
  var title = SpreadsheetApp.getActiveSheet().getRange(general+"C2").getValue();
  var description = SpreadsheetApp.getActiveSheet().getRange(general+"C3").getValue();
  var date = SpreadsheetApp.getActiveSheet().getRange(general+"G2").getValue();
  var curse = SpreadsheetApp.getActiveSheet().getRange(general+"J2").getValue();
  var range = SpreadsheetApp.getActiveSheet().getRange(general+"I3").getValue();
  var range = "\'"  + range; //We need add ' because the spreadsheet delete it
  
  var questions = SpreadsheetApp.getActiveSheet().getRange(range).getNumRows();
  var answers = SpreadsheetApp.getActiveSheet().getRange(range).getNumColumns();
  var values = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(range).getValues();
  
  //Check answer range (max is 8) 9 - questionId, 10 - title, 11 - type
  if(answers>11){  
    var uiFormError = SpreadsheetApp.getUi();
    var responseFormError = uiFormError.alert('ERROR','The limit of answers is 8.', uiFormError.ButtonSet.OK);
    return;
  };
  
  //Creating form
  var form = FormApp.create(title);
  var formDescription = form.setDescription(description);
  var formId = form.getId();
  var formEditUrl = form.getEditUrl();
  var formPublishedUrl = form.getPublishedUrl();
  
  //todo: add curse and date to description
  
  var questionTitle = 1; // question title colum
  var questionType = 2;  // question type colum
  
  // Creating questions
  for(var i=0;i<questions;i++){
    var type = values[i][questionType];
    
    switch(type){
      case "RC":  //Respuesta corta - Short answer
        form.addTextItem()
        .setTitle(values[i][questionTitle]); 
        break;
      case "P":   //Párrafo - Paragraph
        form.addParagraphTextItem()
        .setTitle(values[i][questionTitle])
        break;
      case "SM":  //Selección múltiple - Multiple choice selection
        
        var answersValues = [];
        var item = form.addMultipleChoiceItem();
        
        for(var j=3; j<answers; j++){
          if(values[i][j] != ""){
            answersValues.push(item.createChoice(values[i][j]));
          }
        }
        
        item
          .setTitle(values[i][questionTitle])
          .setChoices(answersValues);
        
        break;
      case "CV":  //Casilla de verificación - CheckBox question
        
        var answersValues = [];
        var item = form.addCheckboxItem();
        
        for(var j=3; j<answers; j++){
          if(values[i][j] != ""){
            answersValues.push(item.createChoice(values[i][j]));
          }
        }
        
        item
          .setTitle(values[i][questionTitle])
          .setChoices(answersValues);

        break;
      case "D":   //Desplegable - Dropdown list
        
        var answersValues = [];
        var item = form.addListItem();
        
        for(var j=3; j<answers; j++){
          if(values[i][j] != ""){
            answersValues.push(item.createChoice(values[i][j]));
          }
        }
        
        item
          .setTitle(values[i][questionTitle])
          .setChoices(answersValues);
        
        break;
      case "SA":  //Subir archivo - Upload File
        SpreadsheetApp.getActiveSpreadsheet().toast("Upload file has not been implemented");
            
        break;
      case "EL":  //Escala lineal - Linear Scale
        if(values[i][3] != "" && values[i][4] != "" && values[i][4] >= 3 && values[i][4] <= 10){
          item = form.addScaleItem()
          .setTitle(values[i][questionTitle])
          .setBounds(values[i][3], values[i][4])
        
          if(values[i][5] != "" && values[i][6] != ""){
            item.setLabels(values[i][5], values[i][6])
          }          
        } else {
         SpreadsheetApp.getActiveSpreadsheet().toast("Linear scale has not been implemented"); 
        }
          
        break;
      case "CVO":  // Cuadricula de varias opciones - Multiple choice grid question
        SpreadsheetApp.getActiveSpreadsheet().toast("Multiple choice grid questions has not been implemented");
        break;
      case "F":   //Fecha - Date
        form.addDateItem().setTitle(values[i][questionTitle]);
        break;
      case "H":   //Hora - Time
        form.addTimeItem().setTitle(values[i][questionTitle]);  
        break;
      default:
        SpreadsheetApp.getActiveSpreadsheet().toast(type + " is not defined");
        break;        
    }
    
  }
  
  //Formulario creado
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Exam created: '+ formEditUrl, ui.ButtonSet.OK);
  
   
}

