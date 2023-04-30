function CallAPI() {
  let response = UrlFetchApp.fetch(APIParameters());
  let data = JSON.parse(response);
  const sheet = SpreadsheetApp.getActiveSheet();
  headerConfig(sheet);
  var fetchedcontentRow=2;
  for ( RowID = 0; RowID < data.Items.length; RowID++) {
    Value=fetchedcontentRowConfig(RowID,data);
    for( ValueColumn = 0 ; ValueColumn < Value.length ; ValueColumn++){
      sheet.getRange(fetchedcontentRow,ValueColumn+1).setValue(Value[ValueColumn]); 
    }
    fetchedcontentRow++; 
  }
  console.log(data.NextPageLink);
  response = UrlFetchApp.fetch(data.NextPageLink);
  data = JSON.parse(response);
  console.log(data.NextPageLink);
  while ("NextPageLink" in data) {
    for ( RowID = 0; RowID < data.Items.length; RowID++) {
      Value=fetchedcontentRowConfig(RowID,data);
      for( ValueColumn = 0 ; ValueColumn < Value.length ; ValueColumn++){
        sheet.getRange(fetchedcontentRow,ValueColumn+1).setValue(Value[ValueColumn]); 
      }
      fetchedcontentRow++;
    }
    response = UrlFetchApp.fetch(data.NextPageLink);
    data = JSON.parse(response);
    console.log(data.NextPageLink);
  }
}

function APIParameters(){
  const AzurePriceLink = "https://prices.azure.com/api/retail/prices?currencyCode='USD'&$filter="
  const serviceFamily = "serviceName eq 'Virtual Machines'"
  const armRegionName = "armRegionName eq 'eastasia'"
  const APIParametersresult = AzurePriceLink.concat([serviceFamily,armRegionName].join(' and '))
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
    "reservationTerm"
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
    reservationTerm
    ]
}
