/********NEW*************/
function triggerEmail(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var range = sheet.getActiveRange();
  var values =range.getValues();
  
  var sheetTrigger = SpreadsheetApp.getActiveSheet();
  var rangeData= sheetTrigger.getRange(1, 1, 1, sheetTrigger.getLastColumn());
  var newData= rangeData.getValues();
  var rowEmail=getRowEmail(newData);
  
  var emailData= getEmail(values,rowEmail);
  range.setNote(emailData);
  addNoiseFinalTrigger(emailData);
  //var sheetEmail = activeSpreadsheet.getSheetByName(form.email);
  //var rangeEmail = sheetEmail.getRange(1, 1, 1, sheetEmail.getLastColumn()-1);
}

function getEmail(newData,rowEmail){
  for(var i=0; i<newData.length;i++)//Row 
  {
    for (var j=0; j <newData[0].length;j++)//Column 
    {
      if(j+1==rowEmail)
      {
        return newData[i][j];
      }
    }
  }
  return -1;
}

function getRowEmail(newData){
  var string="Email Address";
  for(var i=0; i<newData.length;i++)//Row 
  {
    for (var j=0; j <newData[0].length;j++)//Column 
    {
      if(string.localeCompare(newData[i][j])==0)
      {
        return j+1;
      }
    }
  }
  return -1;
}

function saveInformation(form)
{
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var name= activeSpreadsheet.getActiveSheet().getName();//PROBAR
  var newSheet=activeSpreadsheet.getSheetByName("Information");
  if(newSheet==null){newSheet =activeSpreadsheet.insertSheet("Information");}
  else{newSheet.clearContents();}
  var string="undefined";
  newSheet.getRange(1, 1).setValue(name);
  newSheet.getRange(1, 2).setValue(form.name);
  newSheet.getRange(1, 3).setValue(form.row);
  newSheet.getRange(1, 4).setValue(form.colum);
  if(string.localeCompare(form.negative)==0){newSheet.getRange(1, 5).setValue(0);}
  else{newSheet.getRange(1, 5).setValue(1);}
  if(string.localeCompare(form.decimal)==0){newSheet.getRange(1, 6).setValue(0);}
  else{newSheet.getRange(1, 6).setValue(1);}
  newSheet.getRange(1, 7).setValue(form.type);
  newSheet.getRange(1, 8).setValue(form.message);
  newSheet.getRange(1, 9).setValue(form.question);
  newSheet.protect().setWarningOnly(true);
  
  //DEBO PREGUNTAR DE ACUERDO AL FORM 
  //RECORRER TODOS LOS EMAILS
  var stringShare="2";
  if(stringShare.localeCompare(form.share)==0)
  {
      var emailSheet = activeSpreadsheet.getSheetByName(form.email);
      var rangeEmail = emailSheet.getRange(1, 1, 1, emailSheet.getLastColumn());
      var valuesEmail =rangeEmail.getValues();
      var emailData= getRowEmail(valuesEmail);
      
      var rangoEmail= emailSheet.getRange(2, 1, emailSheet.getLastRow()-1, emailSheet.getLastColumn());
      var valoresEmail= rangoEmail.getValues();
      for(var i=0; i< valoresEmail.length ;i++) //fila
      {
      
        addNoiseFinalTrigger(valoresEmail[i][emailData-1]);
      }
   }
}

//PONER EL ITEM EN EL INDEX
//VALIDAR LOS CORREOS


//This function send the number of emails and validate de data(data with only numbers) emails (the mails are in the correct format)
function validateInformation(form)
{
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheetEmail = activeSpreadsheet.getSheetByName(form.email);
  var rangeEmail = sheetEmail.getRange(1, 1, 1, sheetEmail.getLastColumn());
  var valuesEmail =rangeEmail.getValues();
  var emailData= getRowEmail(valuesEmail);
  if(emailData==-1){return -7;}
  
  var rangoEmail= sheetEmail.getRange(2, 1, sheetEmail.getLastRow()-1, sheetEmail.getLastColumn());
  var valoresEmail= rangoEmail.getValues();
  var flagEmails= validarEmails(valoresEmail,emailData);//Validate the whole spreedsheet(EMAIL)
  if(typeof flagEmails === 'string' || flagEmails instanceof String)
  {
    if(flagEmails.localeCompare("a")!=0){console.info("IF IF"); return -9;}//Invalidate mails
  }
  else{
    console.info("ELSE"); 
    var res= "";
    res="b,"+flagEmails;  
    return -9;//Mails or sheet incorrect
  }
  
  //DATA
  var rangoData = activeSpreadsheet.getRange(form.name);//Rango de datos
  var valoresData= rangoData.getValues();
  
  var rangeData= activeSpreadsheet.getRange(form.name);
  var newData= rangeData.getValues();
  //Add the noise depending of the value distribution
  switch (parseInt(form.type)) 
  {
    case 1:
      newData=matrix(valoresData,newData,form.negative,form.decimal);
      break;
    case 2:
      newData=column(valoresData,newData,form.negative,form.decimal);
      break;
    case 3:
      newData=row(valoresData,newData,form.negative,form.decimal);
  } 
  //If the original spreadsheet has letters in the data, the will send a messege of the mistake
  if(newData==-1)
  {
    return -2;//e-data with letters
  } 
  
  //FILE-Validate that the spreadsheet is not in the folder root
  var carpeta="My Drive";
  var idSource =activeSpreadsheet.getId();
  var fileSource = DriveApp.getFileById(idSource);
  var folders =fileSource.getParents();
  if(folders.hasNext())
  {  
    while (folders.hasNext()) 
    {
      var folder = folders.next();
    }
    if(carpeta.localeCompare(folder.getName())==0)
    {
      return -3;//the file in the root
    }
  }
  //Information spreadsheet
  var sheetInfo= activeSpreadsheet.getSheetByName("Information");
  if(!sheetInfo==null){
    activeSpreadsheet.deleteSheet(sheetInfo);
  }
  
  //Questions sheet
  var questionsSheet = activeSpreadsheet.getSheetByName(form.question);
  if(questionsSheet==null){return -4;}//No existe
  //var contador=0;
  var rangeQues= questionsSheet.getRange(2, 2, questionsSheet.getLastRow()-1,1);
  var formulas= rangeQues.getFormulas();
  
  var flagQuestiones= validationQuestions(formulas);
  if(flagQuestiones)
  {
    return -5;//Bad Questions
  } 
  return 1;
}

function addNoiseFinalTrigger(email)
{
  console.info("NOISE"); 
  //var numEmails2=form.numEmails2;
  var activesheet = SpreadsheetApp.getActiveSpreadsheet();
  var aSpreadsheet= activesheet.getSheetByName("Information");
  //Get Data
  var nameaSpreadsheet= aSpreadsheet.getRange(1, 1).getValue();
  var nameValues=aSpreadsheet.getRange(1, 2).getValue();
  var nameRows=aSpreadsheet.getRange(1, 3).getValue();
  var nameColumns=aSpreadsheet.getRange(1, 4).getValue();
  var nameNegative=aSpreadsheet.getRange(1, 5).getValue();
  var nameDecimal=aSpreadsheet.getRange(1, 6).getValue();
  var nameType=aSpreadsheet.getRange(1, 7).getValue();
  var nameMessage=aSpreadsheet.getRange(1, 8).getValue();
  var nameQuestion=aSpreadsheet.getRange(1, 9).getValue();
  
  //get the data from the index
  var activeSpreadsheet= activesheet.getSheetByName(nameaSpreadsheet);
  var rangoRow = activeSpreadsheet.getRange(nameRows);//Rango de filas
  var valoresRow= rangoRow.getValues();
  var rangoData = activeSpreadsheet.getRange(nameValues);//Rango de datos
  var valoresData= rangoData.getValues();

  //get the format of the headers
  var string2="";
  if(string2.localeCompare(nameColumns)!=0)
  {  
    var rangoColumn = activeSpreadsheet.getRange(nameColumns);//Rango de columnas
    var valoresColumn= rangoColumn.getValues();
    var sBG = activeSpreadsheet.getRange(nameColumns).getBackgrounds();
    var sFC = activeSpreadsheet.getRange(nameColumns).getFontColors();
    var sFF = activeSpreadsheet.getRange(nameColumns).getFontFamilies();
    var sFL = activeSpreadsheet.getRange(nameColumns).getFontLines();
    var sFFa = activeSpreadsheet.getRange(nameColumns).getFontFamilies();
    var sFSz = activeSpreadsheet.getRange(nameColumns).getFontSizes();
    var sFSt = activeSpreadsheet.getRange(nameColumns).getFontStyles();
    var sFW = activeSpreadsheet.getRange(nameColumns).getFontWeights();
    var sHA = activeSpreadsheet.getRange(nameColumns).getHorizontalAlignments();
    var sVA = activeSpreadsheet.getRange(nameColumns).getVerticalAlignments();
    var sNF = activeSpreadsheet.getRange(nameColumns).getNumberFormats();
    var sWR = activeSpreadsheet.getRange(nameColumns).getWraps();
  }  
  //get the format of the ROW
  var sBG1 = activeSpreadsheet.getRange(nameRows).getBackgrounds();
  var sFC1 = activeSpreadsheet.getRange(nameRows).getFontColors();
  var sFF1 = activeSpreadsheet.getRange(nameRows).getFontFamilies();
  var sFL1 = activeSpreadsheet.getRange(nameRows).getFontLines();
  var sFFa1 = activeSpreadsheet.getRange(nameRows).getFontFamilies();
  var sFSz1 = activeSpreadsheet.getRange(nameRows).getFontSizes();
  var sFSt1 = activeSpreadsheet.getRange(nameRows).getFontStyles();
  var sFW1 = activeSpreadsheet.getRange(nameRows).getFontWeights();
  var sHA1 = activeSpreadsheet.getRange(nameRows).getHorizontalAlignments();
  var sVA1 = activeSpreadsheet.getRange(nameRows).getVerticalAlignments();
  var sNF1 = activeSpreadsheet.getRange(nameRows).getNumberFormats();
  var sWR1 = activeSpreadsheet.getRange(nameRows).getWraps();
  

  if(validarEmail(email))
  {
    var nameFile = email +" - "+ activeSpreadsheet.getName(); //nombre apellido
    
    var carpeta="My Drive";
    var idSource =activesheet.getId();
    var fileSource = DriveApp.getFileById(idSource);
    var folders =fileSource.getParents();
    
    while (folders.hasNext()) 
    {
      var folder = folders.next();
    }
    if(carpeta.localeCompare(folder.getName())==0)
    {
      return "3-"+email+ " " +email;//nombre apellido
    }
    else
    {
      var sheet = SpreadsheetApp.create(nameFile);
      var spreadSheet= sheet.getSheetByName("Sheet1");
      spreadSheet.setName(activeSpreadsheet.getName());
      spreadSheet.protect().setWarningOnly(true);
      var id = sheet.getId();
      var file = DriveApp.getFileById(id);
      var folsStudenst= folder.getFoldersByName("Students");
      var folderStudenst =null;
      while (folsStudenst.hasNext()) 
      {var folderStudenst = folsStudenst.next();}
      if(folderStudenst==null){folderStudenst=folder.createFolder("Students");}
      folderStudenst.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    }
        
    //Agrego las columnas y filas al nuevo libro creado
    sheet.getRange(nameRows).setValues(valoresRow);
    //Put the format of the headers
    if(string2.localeCompare(nameColumns)!=0)
    {  
      sheet.getRange(nameColumns).setValues(valoresColumn);
      sheet.getRange(nameColumns)
      .setBackgrounds(sBG)
      .setFontColors(sFC)
      .setFontFamilies(sFF)
      .setFontLines(sFL)
      .setFontFamilies(sFFa)
      .setFontSizes(sFSz)
      .setFontFamilies(sFFa)
      .setFontSizes(sFSz)
      .setFontStyles(sFSt)
      .setFontWeights(sFW)
      .setHorizontalAlignments(sHA)
      .setVerticalAlignments(sVA)
      .setNumberFormats(sNF)
      .setWraps(sWR);
    }
    //Put the format of the headers
    sheet.getRange(nameRows)
    .setBackgrounds(sBG1)
    .setFontColors(sFC1)
    .setFontFamilies(sFF1)
    .setFontLines(sFL1)
    .setFontFamilies(sFFa1)
    .setFontSizes(sFSz1)
    .setFontFamilies(sFFa1)
    .setFontSizes(sFSz1)
    .setFontStyles(sFSt1)
    .setFontWeights(sFW1)
    .setHorizontalAlignments(sHA1)
    .setVerticalAlignments(sVA1)
    .setNumberFormats(sNF1)
    .setWraps(sWR1);
      
      
    var rangeData= sheet.getRange(nameValues);
    var newData= rangeData.getValues();
    //Add the noise depending of the value distribution
    switch (parseInt(nameType)) 
    {
      case 1:
        newData=matrix(valoresData,newData,nameNegative,nameDecimal);
        break;
      case 2:
        newData=column(valoresData,newData,nameNegative,nameDecimal);
        break;
      case 3:
        newData=row(valoresData,newData,nameNegative,nameDecimal);
    } 
    //If the original spreadsheet has letters in the data, the will send a messege of the mistake
    if(newData==-1)
    {
      folder.removeFile(file);
      return "2-"+email + " " +email;//nombre apellido
    }  
      
      rangeData.setValues(newData); 
      putAnswersTrigger(sheet, activesheet.getSheetByName(nameQuestion),activesheet,email);
     
      //share and send email
      var url = sheet.getUrl();
      var email = email;
      var subject = sheet.getName();
      var body ="Dear student \n" + nameMessage + "\n" +url;//nombre apellido
      //IVAN
      try 
      {
        sheet.addEditor(email);
        MailApp.sendEmail(email, subject, body);
        return "1-"+email + " " +email;//nombre apellido
      } 
      catch(e) 
      {
        folder.removeFile(file);
        return "0-"+email+ " " +email;//nombre apellido
      }
  }
  else
  {
    return "0-"+email + " " +email;//nombre apellido
  }  
  return "0-"+email+ " " +email;  //nombre apellido
}

/******MENU*******/
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
  .addItem('Add Noise', 'showSidebar')
  .addToUi();  
}
function putAnswersTrigger(spreadStudent,questionsSheet,activeSpreadsheet,email)
{
    //Get formulas     
    var rangeQues= questionsSheet.getRange(2, 2, questionsSheet.getLastRow()-1,1);
    var formulas= rangeQues.getFormulas();
    //abrir la hoja
    var sheetStudent = spreadStudent.insertSheet("Questions");
    //Poner las respuestas
    var rangeQues1= sheetStudent.getRange(2,2,formulas.length,1);//change
    var nData1 =rangeQues1.getValues();
    nData1=answersData(formulas,nData1,2,"nombre");
    rangeQues1.setValues(nData1);
    //Crear la master spreadsheet
    var bandera=true;
    var idAux = activeSpreadsheet.getId();
    var fileSourceAux = DriveApp.getFileById(idAux);
    var foldersAux =fileSourceAux.getParents();
    var sheet=null;
    while(foldersAux.hasNext())
    {
      var folderAux = foldersAux.next();
      var sheet4Ite=folderAux.getFilesByName("MasterAnswers");
      
      while(sheet4Ite.hasNext())//Ask if master file exists
      {  
        var page= sheet4Ite.next()
        var idMA= page.getId();
        sheet= SpreadsheetApp.openById(idMA)
        bandera=false;
      }
    }
  
    if(bandera)
    {  
      //Copy to the folder
      sheet = SpreadsheetApp.create("MasterAnswers");
      var id = sheet.getId();
      var idSource =activeSpreadsheet.getId();
      var file = DriveApp.getFileById(id);
      var fileSource = DriveApp.getFileById(idSource);
      var folders =fileSource.getParents();
      while (folders.hasNext()) 
      {
        var folder = folders.next();
      }
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    } 
    
    //Put the questions in the Mastersheet 
    var rangeQuestions= questionsSheet.getRange(2, 1, questionsSheet.getLastRow()-1,1);
    var valuesQuestions= rangeQuestions.getValues();
    var spreadSheet= sheet.getSheetByName("Sheet1");
    
    var rangeAnswer= spreadSheet.getRange(1, 1, 1, questionsSheet.getLastRow()+1);
    var newData= rangeAnswer.getValues();
    newData=questionsData(valuesQuestions,newData);
    rangeAnswer.setValues(newData); 
    
    var rangeQues1= sheetStudent.getRange(2,2,sheetStudent.getLastRow(),1);//change
    var formulas1= rangeQues1.getValues();
    
    var ranAnswer1= spreadSheet.getRange(spreadSheet.getLastRow()+1, 1, 1, questionsSheet.getLastRow()+1);//IVAN
    var nData2 =ranAnswer1.getValues();
    nData2=answersData(formulas1,nData2,1,"nombre",email);//IVAN
    ranAnswer1.setValues(nData2);

    //Borrar las respuesta de la hoja del estudiante
    spreadStudent.deleteSheet(sheetStudent);
}
/********NEW*************/
function onInstall(e) {onOpen(e);}

function showAnswer() {
  var ui = HtmlService.createHtmlOutputFromFile('answer').setTitle('Answers');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('index').setTitle('Noise Data');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('faqs').setWidth(450).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Read before using the "Noisy Sheets" add-on');
}

function showAbout() {
  var html = HtmlService.createHtmlOutputFromFile('about').setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'About');
}
/******MENU*******/

/*****INDEX******/
/*These functions are called from the index.html*/
//This function put names of the sheet in the select tag of email (index)
function celdas()
{
  var respuesta="";
  var res="Selecciones";
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  respuesta =activeSpreadsheet.getActiveRange().getDataSourceUrl();
  var num = respuesta.search("range");
  res = respuesta.substring(num+6, respuesta.length);
  //var name =SpreadsheetApp.getActiveSheet().getName();
  //res= name+"!"+res;
  return res; 
}
//This function put names of the sheet in the select tag in Answers.hmtl
function comboBox()
{
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var numFiles = activeSpreadsheet.getNumSheets();
  var respuesta= "";
  for(var i=0; i< numFiles ;i++) //fila
  {
    var sheet = activeSpreadsheet.getSheets()[i];
    respuesta= respuesta + sheet.getSheetName() + "-";
  }
  return respuesta;
}

/*****INDEX******/

/*****VALIDATION******/
//This function validate the format of the emails
function validarEmails(valoresEmail,columna)
{
  var numEmails= valoresEmail.length;//Numero de correos
  var names="a";
  for(var s= 0; s<numEmails;s++ )//NUMERO DE CORREOS
  { 
    console.info("Email:"+valoresEmail[s][columna-1]); 
    if(validarEmail(valoresEmail[s][columna-1]))
    {
      var testSheet = SpreadsheetApp.openById("1cbwGKwZooexFKVM31qK4QXkByY2BbQ9gqsxcXG7v5Fs");
      try{
        testSheet.addViewer(valoresEmail[s][columna-1]);
        testSheet.removeViewer(valoresEmail[s][columna-1]);
      } 
      catch(e) 
      {
        names=names+","+valoresEmail[s][0]+" "+valoresEmail[s][1];
      }
    }
    else{
     return s+1;
    }
  }
  return names;
}
//This function validate the format of the emails
function validarEmail(email)
{
  var res= /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  return res.test(email);
}
//This function validate is cell has numbers
function isDigit(c)
{
  return !isNaN(parseFloat(c)) && isFinite(c);
}
//This function validate if the questions sheet has formulas(answer of the questions)
function validationQuestions(formulas)
{
  var string="";
  for(var i=0; i<formulas.length;i++) 
  {
    for (var j=0; j <formulas[0].length;j++) 
    {
      if(string.localeCompare(formulas[i][j])==0)
      {
        return true;
      }
    }
  }
  return false;
}
/*****VALIDATION******/

/******MENU*******/
function questionsData(valuesQuestions,newData)
{
  newData[0][0]="Emails";//IVAN
  newData[0][1]="Students";//IVAN
  for(var i=0; i< valuesQuestions.length ;i++) //fila
  {
    
    for(var j=0; j < valuesQuestions[0].length;j++)//columna
    {      
      newData[j][i+1]=  valuesQuestions[i][j];//IVAN
    }
  }  
  return newData;
}
function answersData(formulas,nData,num,name,email)//IVAN
{
  if(num==1)
  {
    nData[0][0]=email;//IVAN
    nData[0][1]=name;//IVAN
    for(var i=0; i<nData.length;i++) 
    { 
      for (var j=0; j <(nData[0].length-2);j++) //IVAN
      {
        nData[i][j+1]=formulas[j][i];//IVAN CHANGE
      }
    }
  }
  else{
    console.info("Nombre: " +name); 
    for(var i=0; i<nData.length;i++) //fila
    {
      for (var j=0; j <nData[0].length;j++) //Columna
      { 
        nData[i][j]=formulas[i][j];
      }
    }
  }
  return nData;
}

/*Function of Values distribution*/
function row(valoresData,newData,bandera,flagdecimal)
{
  
  for(var i=0; i< valoresData.length ;i++) //fila
  {
    var string="";
    var newArray = [];
    var cont=0;
    for(var j=0; j < valoresData[0].length;j++)//columna
    {
      if((string.localeCompare(valoresData[i][j]))!=0)
      {  
        if(isDigit(valoresData[i][j]))
        {  
          newArray[cont]=parseFloat(valoresData[i][j]);
          cont++;
        }
        else
        {
          return -1;
        }
        
      }  
    }
    var variance2= getVariance(newArray,5);
    
    for(var l=0; l < newData[0].length;l++)//columna
    {     
      var valor=valoresData[i][l];
      if((string.localeCompare(valor))!=0)
      {  
        var ruido=((Math.random()*2)-1)* Math.sqrt(variance2);
        var respuesta = ruido+valor;
        if(!bandera&& respuesta<0)
        {
          ruido=(Math.random())* Math.sqrt(variance2);
          respuesta = ruido+valor;
        }
        if(!flagdecimal){respuesta=parseInt(respuesta);} 
        newData[i][l]=respuesta;
      } 
      else
      {
        newData[i][l]=valor;  
      } 
    }
  }
  return newData;
}
function column(valoresData,newData,bandera,flagdecimal)
{
  for(var i=0; i< valoresData[0].length ;i++) //column
  {
    var string="";
    var newArray = [];
    var cont=0;
    for(var j=0; j < valoresData.length;j++)//row
    {
      if((string.localeCompare(valoresData[j][i]))!=0)
      {           
        if(isDigit(valoresData[j][i]))
        {  
          newArray[cont]=parseFloat(valoresData[j][i]);
          cont++;
        }
        else{return -1;}
      }  
    }
    var variance2= getVariance(newArray,5);
    
    for(var l=0; l < newData.length;l++)//row
    {     
      var valor=valoresData[l][i];
      if((string.localeCompare(valor))!=0)
      {  
        var ruido=((Math.random()*2)-1)* Math.sqrt(variance2);
        var respuesta = ruido+valor;
        if(!bandera&& respuesta<0)
        {
          ruido=(Math.random())* Math.sqrt(variance2);
          respuesta = ruido+valor;
        }
        if(!flagdecimal){respuesta=parseInt(respuesta);} 
        newData[l][i]=respuesta;
      } 
      else
      {
        newData[l][i]=valor;  
      } 
    }
  }
  return newData;
  
}
function matrix(valoresData,newData,bandera,flagdecimal)
{
  var string="";
  var newArray = [];
  var cont=0;
  for(var i=0; i< valoresData.length ;i++) //fila
  {
    for(var j=0; j < valoresData[0].length;j++)//columna
    {
      if((string.localeCompare(valoresData[i][j]))!=0)
      {
        if(isDigit(valoresData[i][j]))
        {  
          newArray[cont]=parseFloat(valoresData[i][j]);
          cont++;
        }
        else{return -1;}
      }  
    }
  }
  var variance2= getVariance(newArray,5);
  for(var k=0; k< newData.length ;k++) //fila
  {
    for(var l=0; l < newData[0].length;l++)//columna
    {     
      var valor=valoresData[k][l];
      if((string.localeCompare(valor))!=0)
      {  
        var ruido=((Math.random()*2)-1)* Math.sqrt(variance2);
        var respuesta = ruido+valor;
        if(!bandera&& respuesta<0)
        {
          ruido=(Math.random())* Math.sqrt(variance2);
          respuesta = ruido+valor;
        }
        if(!flagdecimal){respuesta=parseInt(respuesta);} 
        newData[k][l]=respuesta;
      } 
      else
      {
        newData[k][l]=valor;  
      } 
    }
  }
  return newData;
}
/*Function of Values distribution*/

/*Statistics function*/
function mean(numbers) 
{
  var total = 0,i;
  for (i = 0; i < numbers.length; i += 1) 
  {
    total += numbers[i];
  }
  return total / numbers.length;
}

function getNumWithSetDec( num, numOfDec)
{
  var pow10s = Math.pow( 10, numOfDec || 0 );
  return ( numOfDec ) ? Math.round( pow10s * num ) / pow10s : num;
}
function getAverageFromNumArr( numArr, numOfDec )
{
  var i = numArr.length, 
      sum = 0;
  while( i-- ){
    sum += numArr[ i ];
  }
  return (sum / numArr.length );
}
function getVariance(numArr, numOfDec)
{
  var avg = getAverageFromNumArr( numArr, numOfDec ), 
      i = numArr.length,
      v = 0;
  
  while( i-- ){
    v += Math.pow( (numArr[ i ] - avg), 2 );
  }
  v /= numArr.length;
  return v;
}
function standardDeviation(values)
{
  var avg = average(values);
  var squareDiffs = values.map(function(value)
                               {
                                 var diff = value - avg;
                                 var sqrDiff = diff * diff;
                                 return sqrDiff;
                               });
  var avgSquareDiff = average(squareDiffs);
  var stdDev = Math.sqrt(avgSquareDiff);
  return stdDev;
}
function average(data)
{
  var sum = data.reduce(function(sum, value)
                        {
                          return sum + value;
                        }, 0);
  
  var avg = sum / data.length;
  return avg;
}
/*Statistics function*/
