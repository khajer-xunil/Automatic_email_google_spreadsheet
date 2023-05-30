//New item "email" in menu 
function onOpen (){
 const ui = SpreadsheetApp.getUi();
 const menu = ui.createMenu("Email");
 menu.addItem("Send Update","sendMail")
 menu.addToUi();

}

function myFunction() {
  const list = SpreadsheetApp.getActive().getSheetByName("Shipments");
  list.activate;
  const today = Utilities.formatDate(new Date(), "GMT","dd-MM-YYYY");
  const testDate = new Date();
  const secondDate = new Date();
  const tomorrow = Utilities.formatDate(new Date(secondDate.setDate(testDate.getDate()+1)), "GMT", "dd-MM-YYYY");
  var count_today = 0;
  var count_tomorrow = 0;
  var row = list.getLastRow();
  for (var counter = 2; counter <= row; counter = counter + 1 ){
    var check_cell = list.getRange(counter,28).getValue();
    var check_cell_2 = list.getRange(counter,11).getValue();
      //validation of date    
      if (check_cell == today && check_cell_2 == "")
          count_today = count_today + 1;
        else if (check_cell == tomorrow)  
          count_tomorrow = count_tomorrow + 1;   
      else 
          count_today = count_today + 0;
};

    // HTML template
    if (count_today>0){
      const list2 = SpreadsheetApp.getActive().getSheetByName("Summary");
      const data = list2.getRange(3,2).setValue(count_today);
      const data_2 = list2.getRange(3,3).setValue(count_tomorrow);

      const header = list2.getRange("B2").getValue();
      const header1 = list2.getRange("C2").getValue();

      const to_be_shipped = list2.getRange("A3").getValue(); 
      const ship_today = list2.getRange("B3").getValue();
      const ship_tomorrow = list2.getRange("C3").getValue();

      const lr = list2.getLastRow();
      const tableRangeValues = list2.getRange(9,2,lr-3,2).getValues();

      const htmlTemplate = HtmlService.createTemplateFromFile("tabulka");
      htmlTemplate.header = header;
      htmlTemplate.header1 = header1;
      htmlTemplate.to_be_shipped = to_be_shipped;
      htmlTemplate.ship_today = ship_today;
      htmlTemplate.ship_tomorrow = ship_tomorrow;
      htmlTemplate.tableRangeValues = tableRangeValues;

      const htmlForEmail = htmlTemplate.evaluate().getContent();
      console.log(htmlForEmail);  
      GmailApp.sendEmail("customs.brno@amtech.cz","Casambi Overview"+ " " + today, "OPEN with a client",{ htmlBody: htmlForEmail });
    }
    // HTML template if count date is "0"
    else{
      const list2 = SpreadsheetApp.getActive().getSheetByName("Summary");
      const data = list2.getRange(3,2).setValue(count_today);
      const data_2 = list2.getRange(3,3).setValue(count_tomorrow);

      const header = list2.getRange("B2").getValue();
      const header1 = list2.getRange("C2").getValue();

      const to_be_shipped = list2.getRange("A3").getValue(); 
      const ship_today = list2.getRange("B3").getValue();
      const ship_tomorrow = list2.getRange("C3").getValue();

      const htmlTemplate = HtmlService.createTemplateFromFile("tabulka1");
      htmlTemplate.header = header;
      htmlTemplate.header1 = header1;
      htmlTemplate.to_be_shipped = to_be_shipped;
      htmlTemplate.ship_today = ship_today;
      htmlTemplate.ship_tomorrow = ship_tomorrow;

      const htmlForEmail = htmlTemplate.evaluate().getContent();
      console.log(htmlForEmail);  
      GmailApp.sendEmail("email","Email Subject"+ " " + today, "OPEN with a client",{ htmlBody: htmlForEmail });
    }    
 };
