function main() {
var newSpreadSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/10joMWl7QzPi6qO9O4qg_rrI_nC05tMMHQtg831k_FqE/edit#gid=0");
  //komentarz
var numOfSheets = newSpreadSheet.getSheets(); 
 if(numOfSheets.length>0){
 for(var i=1,len=numOfSheets.length;i<len;i++)
 newSpreadSheet.deleteSheet(numOfSheets[i]);
 }
 newSpreadSheet.getActiveSheet().clear().setName("Last 1 Day");

 var allHeaders = ["Conversions"];
 var headerIndexes = {};
 var headerString = "";
 var daysForMonth = 28;
 var sheet = newSpreadSheet.getActiveSheet();

 var tempIndex = 3;

 for(var i=0,len=allHeaders.length;i<len;i++){
 sheet.activate();
 sheet.getRange(1,tempIndex).setValue(allHeaders[i]);

 headerIndexes[allHeaders[i]]=tempIndex;
 headerString=headerString+allHeaders[i]+", ";
 tempIndex = tempIndex+7;
 }

 sheet.appendRow(["Account Name","Account Id"]);

 newSpreadSheet.duplicateActiveSheet().setName("Last 7 Days");
 newSpreadSheet.duplicateActiveSheet().setName("Last 30 Days");

 var allSheets = newSpreadSheet.getSheets();

 headerString = headerString.substring(0,headerString.length-2);

 Logger.log(headerIndexes);
 var datesForFirst=[];var dateForSecond=[];var dateForThird = [];
 var currentDate = new Date();var prevDate = new Date();var anotherPrevDate = new Date();
 var sheetIndex =0;

 //insert dates in the next rows 
 //for 1st sheet
 var fixedDate = new Date(Utilities.formatDate(new Date(),AdWordsApp.currentAccount()
     .getTimeZone(), "MMM dd,yyyy HH:mm:ss")); 

 var time = fixedDate.getTime() -(1 * 24 * 60 * 60 * 1000);
 fixedDate = new Date(time);



 appendDates(sheetIndex, datesForFirst, 0); 

 //for second sheet
 sheetIndex = 1;
 time = fixedDate.getTime() -(5 * 24 * 60 * 60 * 1000);
 currentDate = new Date(time);
 time = fixedDate.getTime() -(14 * 24 * 60 * 60 * 1000);
 prevDate = new Date(time);
 time = currentDate.getTime() -(30 * 24 * 60 * 60 * 1000);
 anotherPrevDate = new Date(time);
 dateForSecond.push(currentDate);


 //dates inserted

 //get accounts and data respectively
 var accounts_iterator = MccApp.accounts().withCondition("Impressions>0").forDateRange("LAST_30_DAYS").get();
 var current_mccaccount = AdWordsApp.currentAccount();
 var all_accounts=[];
 while(accounts_iterator.hasNext()){
 all_accounts.push(accounts_iterator.next());
 }

 Logger.log("no of accounts"+all_accounts.length);
 for(var i=0,len=all_accounts.length;i<len;i++){
 MccApp.select(all_accounts[i]);

 appendData(1,dateForSecond,7);


 for(var j=0;j<3;j++){
 var sheetCurrent = allSheets[j];

 sheetCurrent.activate();
 var lRow = sheetCurrent.getLastRow();
 for(var key in headerIndexes){
 var index = headerIndexes[key];
 var positiveColor = "green";
 var negativeColor = "red";
 if(key=="Cost"){
 positiveColor="red";
 negativeColor="green";
 }

 var firstVal = sheetCurrent.getRange(lRow,index).getValue();
 var secondVal = sheetCurrent.getRange(lRow,index+1).getValue();

 var diff = firstVal-secondVal;
 sheetCurrent.getRange(lRow, index+2).setValue(diff);
 var pcent = (diff/secondVal)*100;
 if(secondVal==0)
 pcent=firstVal*100;
 if(diff==0)
 pcent=0;


 secondVal = sheetCurrent.getRange(lRow,index+4).getValue();
 var diff = firstVal-secondVal;
 sheetCurrent.getRange(lRow, index+5).setValue(diff);
 var pcent = (diff/secondVal)*100;
 if(diff==0)
 pcent=0;
 if(secondVal==0)
 pcent=firstVal*100;
 sheetCurrent.getRange(lRow, index+6).setValue(pcent+"%");

 }
 }
 }



 function appendData(indexForSheet,dateArray,days){

 currentSheet = allSheets[indexForSheet];
 currentSheet.activate();
 currentRow = currentSheet.getLastRow()+1;
 var date_range = "";
 var fieldGap = [0,1,4];
 for(var i=0,len=dateArray.length;i<len;i++){
 var toDate = dateArray[i];
 if(indexForSheet!=0){
 tempDate = dateArray[i].getTime()+(days * 24 * 60 * 60 * 1000);
 toDate = new Date(tempDate);

 }
 date_range = ""+Utilities.formatDate(dateArray[i], "CST", "yyyyMMdd")+","+Utilities.formatDate(toDate, "CST", "yyyyMMdd");

 var report = AdWordsApp.report("SELECT "+headerString+ 
 " FROM ACCOUNT_PERFORMANCE_REPORT "+
 "DURING "+date_range);
 var rows = report.rows();
 while(rows.hasNext()){
 var row = rows.next();

 var currentIndex = 0;
 currentSheet.getRange(currentRow, 1).setValue(AdWordsApp.currentAccount().getName());
 currentSheet.getRange(currentRow, 2).setValue(AdWordsApp.currentAccount().getCustomerId());
 for(var key in headerIndexes){
 var index = headerIndexes[key];
 currentSheet.getRange(currentRow, index+fieldGap[i]).setValue(row[key]);
 }
 }
 }

 }

 function appendDates(sheetIndex, dateArray, days){
 currentSheet = allSheets[sheetIndex];
 currentSheet.activate();
 currentRow = currentSheet.getLastRow()+1;

 var date_range = [];

 for(var i=0,len=dateArray.length;i<len;i++){
 var toDate = dateArray[i];
 if(sheetIndex!=0){
 tempDate = dateArray[i].getTime()+(days * 24 * 60 * 60 * 1000);
 toDate = new Date(tempDate);
 date_range.push(""+Utilities.formatDate(dateArray[i], "CST", "MM/dd/yyyy")+" - "+Utilities.formatDate(toDate, "CST", "MM/dd/yyyy"));
 }
 else{
 date_range.push(Utilities.formatDate(dateArray[i], "PST", "MM/dd/yyyy"));
 }
 }
 
 }

}