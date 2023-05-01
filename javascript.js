function CallAPI() {
//Get Google Sheet Active
  const sheet = SpreadsheetApp.getActiveSheet();
//Configure Google Sheet header
  headerConfig(sheet);
  var fetchedcontentRow=2;
//Call Azure Price API and Declare API Parameters
  let response = UrlFetchApp.fetch(APIParameters());
  let data = JSON.parse(response);
//Get Azure Price API Content
  for ( RowID = 0; RowID < data.Items.length; RowID++) {
//Write API Content into Google Sheet
    Value=fetchedcontentRowConfig(RowID,data);
    for( ValueColumn = 0 ; ValueColumn < Value.length ; ValueColumn++){
      sheet.getRange(fetchedcontentRow,ValueColumn+1).setValue(Value[ValueColumn]); 
    }
    fetchedcontentRow++; 
  }
//Call API for remaining content
  console.log(data.NextPageLink);
  response = UrlFetchApp.fetch(data.NextPageLink);
  data = JSON.parse(response);
  console.log(data.NextPageLink);
  while ("NextPageLink" in data) {
//Get API Call Content
    for ( RowID = 0; RowID < data.Items.length; RowID++) {
//Write API Content into Google Sheet
      Value=fetchedcontentRowConfig(RowID,data);
      for( ValueColumn = 0 ; ValueColumn < Value.length ; ValueColumn++){
        sheet.getRange(fetchedcontentRow,ValueColumn+1).setValue(Value[ValueColumn]); 
      }
      fetchedcontentRow++;
    }
//if remaining content does not exist
    if(data.NextPageLink == null){
      //API Calls Finished
      break;
    }
//if the remaining content exists
    response = UrlFetchApp.fetch(data.NextPageLink);
    data = JSON.parse(response);
    console.log(data.NextPageLink);
  }
}

function APIParameters(){
  const AzurePriceLink = "https://prices.azure.com/api/retail/prices?api-version=2023-01-01-preview&currencyCode='USD'&$filter="
  const serviceFamily = "serviceName eq 'Virtual Machines'"
  const armRegionName = "armRegionName eq 'eastasia'"
//priceType :[ DevTestConsumption , Consumption , Reservation ]
  const priceType = "priceType eq 'DevTestConsumption'"   
  const meterName = "contains(meterName, 'Spot') eq false"
  const APIParametersresult = AzurePriceLink.concat([serviceFamily,armRegionName,priceType,meterName].join(' and '))
  console.log(APIParametersresult);
  return APIParametersresult;
}
function headerConfig(sheet){
  const header =1
  const headerlist=[
    "productName",
    "skuName",
    "armSkuName",
    "meterName",
    "armRegionName",
    "location",
    "retailPrice",
    "unitPrice",
    "currencyCode" ,
    "unitOfMeasure",
    "type",
    "reservationTerm",
    "savingsPlan_unitPrice_3_Years",
    "savingsPlan_retailPrice_3_Years",
    "savingsPlan_term_3_Years",
    "savingsPlan_unitPrice_1_Years",
    "savingsPlan_retailPrice_1_Years",
    "savingsPlan_term_1_Years"       
    ]
  for( headerColumn = 0 ; headerColumn < headerlist.length ; headerColumn++){
     sheet.getRange(header,headerColumn+1).setValue(headerlist[headerColumn]); 
  }
}

function fetchedcontentRowConfig(RowID,data){
//Service Name
  const productName = data.Items[RowID].productName;
  const skuName = data.Items[RowID].skuName;
  const armSkuName = data.Items[RowID].armSkuName;
  const meterName = data.Items[RowID].meterName;
//region & location
  const armRegionName = data.Items[RowID].armRegionName;
  const location = data.Items[RowID].location;
// Service retail Price
  const retailPrice = data.Items[RowID].retailPrice;
  const unitPrice = data.Items[RowID].unitPrice;
  const unitOfMeasure = data.Items[RowID].unitOfMeasure;
//other
  const type = data.Items[RowID].type;
  const reservationTerm = data.Items[RowID].reservationTerm;
  const currencyCode = data.Items[RowID].currencyCode;
//Saving Plan
if("savingsPlan" in data.Items[RowID]){
    savingsPlan_unitPrice_3_Years = data.Items[RowID].savingsPlan[0].unitPrice;
    savingsPlan_retailPrice_3_Years=data.Items[RowID].savingsPlan[0].retailPrice
    savingsPlan_term_3_Years=data.Items[RowID].savingsPlan[0].term;
    savingsPlan_unitPrice_1_Years=data.Items[RowID].savingsPlan[1].unitPrice;
    savingsPlan_retailPrice_1_Years=data.Items[RowID].savingsPlan[1].retailPrice;
    savingsPlan_term_1_Years =data.Items[RowID].savingsPlan[1].term;
}else{
  savingsPlan_unitPrice_3_Years =" ";
  savingsPlan_retailPrice_3_Years=" ";
  savingsPlan_term_3_Years=" ";
  savingsPlan_unitPrice_1_Years=" ";
  savingsPlan_retailPrice_1_Years=" ";
  savingsPlan_term_1_Years =" ";
}
  return [
    productName,
    skuName,
    armSkuName,
    meterName,
    armRegionName,
    location,
    retailPrice,
    unitPrice,
    currencyCode,
    unitOfMeasure,
    type,
    reservationTerm,
    savingsPlan_unitPrice_3_Years,
    savingsPlan_retailPrice_3_Years,
    savingsPlan_term_3_Years,
    savingsPlan_unitPrice_1_Years,
    savingsPlan_retailPrice_1_Years,
    savingsPlan_term_1_Years 
    ]
}
