function FormulaCreator() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var SheetForm = ss.getSheetByName('Forumals');
  var SheetKey = ss.getSheetByName('Key');
  var row = SheetForm.getRange(2,2,1,7);         //Getting the range of the Formula table. Change the 7 if more gets added.
  var NumberOfWords = row.getValues();           //See how many words there is per key 
  IfEmptyCells();
  for(var i = 0 ; i < 7; i ++)
  {
    //Put a row into an array
    var value = NumberOfWords[0][i]-2;
    var TempArray = [];
    //Putting each coloumn of words into a temp Array.
    for(var j = 0 ; j < value; j ++)
    {
      TempArray[j] = SheetKey.getRange(j+2,i+1,NumberOfWords[0][i]).getValue();
    }
    //Translates that array into the formula
    var TempForm = "";
    for(var k = 0 ; k < value ; k++)
    {
      var Form = '((ISNUMBER(FIND("'+TempArray[k]+'",Data!E2:E617))))';
      TempForm+=Form + "+";
    }
    //Sets the formula into an cell.
    var Form = TempForm.length;
    var AddFormulaToSheet = SheetForm.getRange(3,i+2);
    AddFormulaToSheet.setFormula("=SUMPRODUCT(--"+TempForm.substring(0,Form-5)+")))))");
  }

}


function IfEmptyCells()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var SheetForm = ss.getSheetByName('Forumals');
  var row = SheetForm.getRange(3,2,1,7);

  if(row.getValue() != "")
  {
   for (var i = 0 ; i < 7; i ++)// change 7 if row gets changed in myFunction
   {
     row.setValue("");
   }
  }
    
}
