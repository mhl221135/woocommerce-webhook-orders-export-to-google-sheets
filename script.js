
//this is a function that fires when the webapp receives a POST request
function doPost(e) {
    var myData             = JSON.parse([e.postData.contents]);
    var order_created      = myData.date_created;//date of order creation
    var order_modified      = myData.date_modified;//date of order modified like status changed etc.
    var billing_email      = myData.billing.email;
    var order_total        = myData.total;
    var order_number       = myData.number;
    var order_status       = myData.status;
    var billing_first_name = myData.billing.first_name;
    var billing_last_name  = myData.billing.last_name;
    var payment_method      = myData.payment_method;//payment method title woocommerce
    var ip_adr=myData.customer_ip_address;
    var metta = myData.meta_data;//meta for custom checkout meta field 
    var lineitems = myData.line_items;//meta of order like SKU or Variation selected


var all_ordrs = SpreadsheetApp.getActive().getSheetByName('all_orders');
var paid_ordrs = SpreadsheetApp.getActive().getSheetByName('paid');
var cancelled = SpreadsheetApp.getActive().getSheetByName('cancelled');
const keynm = 'WHA_CONTACT';


for (var i in metta)  {
  if(metta[i].key==keynm){
    var meta_value_phone = metta[i].value;
    break;
  }
}

for (var k in lineitems)  {
  if(lineitems[k].sku){
    var sku = lineitems[k].sku;
  }
  for (var z in lineitems[k].meta_data)  {
  if(lineitems[k].meta_data[z].display_value){
    var pack = lineitems[k].meta_data[z].display_value;
  }
  }
}

//all orders add row to spreadsheet all_orders
all_ordrs.appendRow([order_created,order_modified,"#"+order_number+" "+billing_first_name+" "+billing_last_name+" "+billing_email+" "+meta_value_phone+" "+sku+" "+pack,order_number,billing_first_name,billing_last_name,billing_email,meta_value_phone,sku,pack,order_total,payment_method,order_status,ip_adr]);

//completed orders add row to spreadsheet completed
if(order_status=='completed'){
//paid orders
paid_ordrs.appendRow([order_created,order_modified,"#"+order_number+" "+billing_first_name+" "+billing_last_name+" "+billing_email+" "+meta_value_phone+" "+sku+" "+pack,order_number,billing_first_name,billing_last_name,billing_email,meta_value_phone,sku,pack,order_total,payment_method,order_status,ip_adr]);
 }

 //cancelled orders add row to spreadsheet cancelled
 if(order_status=='cancelled'){
//paid orders
cancelled.appendRow([order_created,order_modified,"#"+order_number+" "+billing_first_name+" "+billing_last_name+" "+billing_email+" "+meta_value_phone+" "+sku+" "+pack,order_number,billing_first_name,billing_last_name,billing_email,meta_value_phone,sku,pack,order_total,payment_method,order_status,ip_adr]);
 }

  //sort rows by column 1 in all spreadsheets below
 all_ordrs.getDataRange().sort({column: 1, ascending: true})
 paid_ordrs.getDataRange().sort({column: 1, ascending: true})
 cancelled.getDataRange().sort({column: 1, ascending: true})

  //call function to remove duplicates by sspreadsheet list
removedublicates("all_orders");
removedublicates("paid");
removedublicates("cancelled");

}


//function to remove duplicates
function removedublicates(name){

 var ssheet = SpreadsheetApp.getActive().getSheetByName(name);
  

 ssheet.getDataRange().sort({column: 2, ascending: false}) // replace 1 with your date column
  var ddata = ssheet.getDataRange().getValues();
  var newData = [];

  for (var i in ddata) {
    var row = ddata[i];
    var duplicate = false;
    for (var j in newData) {
      if(row[3] == newData[j][3]){//colum number- important that massif index starting from 0
             duplicate = true;
      }
    }
    if (!duplicate) {
      newData.push(row);
    }
  }

  ssheet.clearContents();
  ssheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
  ssheet.getDataRange().sort({column: 1, ascending: true}) 

}

