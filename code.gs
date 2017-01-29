function onInstall(e){
  onOpen(e);
}


// Añade la función 'Generar test' al complemento 'Generador de tests'
function onOpen() {
  SpreadsheetApp.getUi().createAddonMenu().addItem('Generar test', 'showSidebar').addToUi();
}


// Muestra la ventana para confirmar la generación del test
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Generador de tests');
  SpreadsheetApp.getUi().showSidebar(ui);
}


// Función que procesa los parámetros para generar el test
function generateTest() {
   
  SpreadsheetApp.getActiveSpreadsheet().toast("Generando el test...");
  
  var general = "general!";
  var DIFFERENCE = 10;
  var testNumber = SpreadsheetApp.getActiveSheet().getRange(general + "C2").getValue();
  
  var title = SpreadsheetApp.getActiveSheet().getRange(general + "B" + (testNumber+DIFFERENCE)).getValue();  
  var degree = SpreadsheetApp.getActiveSheet().getRange(general + "C1").getValue();
  var subject = SpreadsheetApp.getActiveSheet().getRange(general + "G1").getValue();
  
  var description = SpreadsheetApp.getActiveSheet().getRange(general + "C" + (testNumber+DIFFERENCE)).getValue();
  if ((new Date().getMonth()+1) < 10 && (new Date().getMinutes()) < 10)
    var dateAndTime = (new Date().getUTCDate()) + "/0" + (new Date().getMonth()+1) + "/" + (new Date().getFullYear()) + " " + (new Date().getHours()) + ":0" + (new Date().getMinutes());
  else if ((new Date().getMonth()+1) < 10 && (new Date().getMinutes()) >= 10)
    var dateAndTime = (new Date().getUTCDate()) + "/0" + (new Date().getMonth()+1) + "/" + (new Date().getFullYear()) + " " + (new Date().getHours()) + ":" + (new Date().getMinutes());
  else if ((new Date().getMonth()+1) >= 10 && (new Date().getMinutes()) < 10)
    var dateAndTime = (new Date().getUTCDate()) + "/" + (new Date().getMonth()+1) + "/" + (new Date().getFullYear()) + " " + (new Date().getHours()) + ":0" + (new Date().getMinutes());
  else
    var dateAndTime = (new Date().getUTCDate()) + "/" + (new Date().getMonth()+1) + "/" + (new Date().getFullYear()) + " " + (new Date().getHours()) + ":" + (new Date().getMinutes());
  var course = SpreadsheetApp.getActiveSheet().getRange(general + "I2").getValue();
  
  var range = "\'" + title + "\'" + "!A2:K101"; // Hay que añadir ' porque spreadsheet lo borra
  var numQuestions = SpreadsheetApp.getActiveSheet().getRange(range).getNumRows();
  var numAnswers = SpreadsheetApp.getActiveSheet().getRange(range).getNumColumns();
  var values = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(range).getValues();
  
  // Comprueba el rango de las respuestas (el máximo es 8). 9 - idPregunta, 10 - título, 11 - tipo
  if (numAnswers > 11) {  
    var uiFormError = SpreadsheetApp.getUi();
    var responseFormError = uiFormError.alert("ERROR: El límite de respuestas es 8.", uiFormError.ButtonSet.OK);
    return;
  }
  
  // Genera el formulario del test
  var form = FormApp.create(title);
  var formDescription = form.setDescription(description + "\n\n\nASIGNATURA: " + subject + "\n\nTITULACIÓN: " + degree + "\n\nCURSO: " + course);
  var formId = form.getId();
  var formEditUrl = form.getEditUrl();
  var formPublishedUrl = form.getPublishedUrl();
  
  var questionTitle = 1; // Columna de los títulos
  var questionType = 2;  // Columna de los tipos
  var endQuestion = false;
  
  // Genera las preguntas del formulario del test
  for (var i = 0; i < numQuestions && !endQuestion; i++) {
    var type = values[i][questionType];
    
    switch(type) {
      case "RC":  //Respuesta corta - Short answer
        form.addTextItem().setTitle(values[i][questionTitle]); 
      break;
        
      case "P":   //Párrafo - Paragraph
        form.addParagraphTextItem().setTitle(values[i][questionTitle]);
      break;
        
      case "SM":  //Selección múltiple - Multiple choice selection
        var answersValues = [];
        var item = form.addMultipleChoiceItem();
        for (var j = 3; j < numAnswers; j++)
          if (values[i][j] != "")
            answersValues.push(item.createChoice(values[i][j]));
        item.setTitle(values[i][questionTitle]).setChoices(answersValues);
      break;
        
      case "CV":  //Casilla de verificación - CheckBox question  
        var answersValues = [];
        var item = form.addCheckboxItem();
        for (var j = 3; j < numAnswers; j++)
          if (values[i][j] != "")
            answersValues.push(item.createChoice(values[i][j]));
        item.setTitle(values[i][questionTitle]).setChoices(answersValues);
      break;

      case "D":   //Desplegable - Dropdown list
        var answersValues = [];
        var item = form.addListItem();        
        for (var j = 3; j < numAnswers; j++)
          if (values[i][j] != "")
            answersValues.push(item.createChoice(values[i][j]));
        item.setTitle(values[i][questionTitle]).setChoices(answersValues);
      break;
        
      case "SA":  //Subir archivo - Upload File
        SpreadsheetApp.getActiveSpreadsheet().toast("ERROR: La función de subir fichero no está implementada.");
      break;
      
      case "EL":  //Escala lineal - Linear Scale
        if (values[i][3] != "" && values[i][4] != "" && values[i][4] >= 3 && values[i][4] <= 10) {
          item = form.addScaleItem().setTitle(values[i][questionTitle]).setBounds(values[i][3], values[i][4]);
          if(values[i][5] != "" && values[i][6] != "")
            item.setLabels(values[i][5], values[i][6])
        }
        else
          SpreadsheetApp.getActiveSpreadsheet().toast("ERROR: Los parámetros para el tipo 'Escala lineal' no están bien definidos."); 
      break;
      
      case "CVO":  // Cuadrícula de varias opciones - Multiple choice grid question
        SpreadsheetApp.getActiveSpreadsheet().toast("ERROR: La función de 'Cuadrícula de varias opciones' no está implementada."); 
      break;
      
      case "F":   //Fecha - Date
        form.addDateItem().setTitle(values[i][questionTitle]);
      break;
      
      case "H":   //Hora - Time
        form.addTimeItem().setTitle(values[i][questionTitle]);  
      break;
      
      default:
        endQuestion = true;
      break;        
    }
  }
  
  // Formulario generado 
  SpreadsheetApp.getActiveSheet().getRange(general + "C5").setValue(formEditUrl);
  SpreadsheetApp.getActiveSheet().getRange(general + "C6").setValue(formPublishedUrl);
  SpreadsheetApp.getActiveSheet().getRange(general + "C7").setValue(title);
  SpreadsheetApp.getActiveSheet().getRange(general + "I7").setValue(dateAndTime);
  
  SpreadsheetApp.getActiveSheet().getRange(general + "G"+(testNumber+DIFFERENCE)).setValue(formEditUrl);
  SpreadsheetApp.getActiveSheet().getRange(general + "H"+(testNumber+DIFFERENCE)).setValue(formPublishedUrl);
  SpreadsheetApp.getActiveSheet().getRange(general + "I"+(testNumber+DIFFERENCE)).setValue(dateAndTime);
  SpreadsheetApp.getActiveSheet().getRange(general + "J"+(testNumber+DIFFERENCE)).setValue(course);  
  
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Test '" + title + "' generado. El test se encuentra en la carpeta 'Mi unidad' de Google Drive.\n\nURL de edición:\n" 
                          + formEditUrl + "\n\nURL para compartir:\n" 
                          + formPublishedUrl, ui.ButtonSet.OK);  
}
