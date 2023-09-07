function openForm(e)
{
  populateQuestions();
}

function populateQuestions() {
  var form = FormApp.getActiveForm();
  var googleSheetsQuestions = getQuestionValues();
  var itemsArray = form.getItems();
  itemsArray.forEach(function(item){
    googleSheetsQuestions[0].forEach(function(header_value, header_index) {
      if(header_value == item.getTitle())
      {
        var choiceArray = [];
        for(j = 1; j < googleSheetsQuestions.length; j++)
        {
          (googleSheetsQuestions[j][header_index] != '') ? choiceArray.push(googleSheetsQuestions[j][header_index]) : null;
        }
        item.asListItem().setChoiceValues(choiceArray);
        // If using Dropdown Questions use line below instead of line above.
        //item.asListItem().setChoiceValues(choiceArray);
      }
    });     
  });
}

function getQuestionValues() {
  var ss= SpreadsheetApp.openById('1LJl3f2ZYcIFjesLR3xuKyfSAzwbtIEi8dSldJRbGUz0');
  var questionSheet = ss.getSheetByName('Dropdowns');
  var returnData = questionSheet.getDataRange().getValues();
  return returnData;
}
