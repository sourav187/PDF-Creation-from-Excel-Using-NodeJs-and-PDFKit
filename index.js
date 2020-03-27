const PDFDocument = require('pdfkit');
const XLSX = require('xlsx');
var workbook = XLSX.readFile('test.xlsx');
const fs = require('fs');
//Excel codes
var first_sheet_name = workbook.SheetNames[0];
/* Get worksheet */
var worksheet = workbook.Sheets[first_sheet_name];
for (let index = 2; 100; index++) {
  var VendorNo_cell = 'I' + index;
  desired_cell = worksheet[VendorNo_cell];
  var VendorNo_cell_value = desired_cell ? desired_cell.v : '';
  if (VendorNo_cell_value === '') {
    break;
  }
  var Field1_cell = 'A' + index;
  var Field2_cell = 'B' + index;
  var Field3_cell = 'C' + index;
  var Phone_cell = 'D' + index;
  var Fax_cell = 'E' + index;
  var Email_cell = 'F' + index;
  var HC_cell = 'G' + index;
  var Field4_cell = 'H' + index;
  var InvoiceDate_cell = 'J' + index;
  var Reference_cell = 'K' + index;
  var Item_cell = 'L' + index;
  var Quantity_cell = 'M' + index;
  var UOM_cell = 'N' + index;
  var PricePerUOM_cell = 'O' + index;
  var Amount_cell = 'P' + index;
  var currency_cell = 'Q' + index;
  var TaxCode_cell = 'R' + index;
  var TaxPercentage_cell = 'S' + index;
  var TaxAmount_cell = 'T' + index;
  var grossAmount='W'+index;
  var AccountHolder_cell = 'U' + index;
  var IBANF_cell = 'V' + index;
  
  /* Find desired cell */
  var desired_cell = worksheet[Field1_cell];
  var Field1_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[Field2_cell];
  var Field2_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[Field3_cell];
  var Field3_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[Phone_cell];
  var Phone_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[Fax_cell];
  var Fax_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[Email_cell];
  var Email_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[HC_cell];
  var HC_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[Field4_cell];
  var Field4_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[InvoiceDate_cell];
  var InvoiceDate_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[Reference_cell];
  var Reference_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[Item_cell];
  var Item_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[Quantity_cell];
  var Quantity_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[UOM_cell];
  var UOM_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[PricePerUOM_cell];
  var PricePerUOM_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[Amount_cell];
  var Amount_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[currency_cell];
  var currency_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[TaxCode_cell];
  var TaxCode_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[TaxPercentage_cell];
  var TaxPercentage_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[TaxAmount_cell];
  var TaxAmount_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[grossAmount];
  var grossAmount_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[AccountHolder_cell];
  var AccountHolder_cell_value = desired_cell ? desired_cell.v : '';
  desired_cell = worksheet[IBANF_cell];
  var IBANF_cell_value = desired_cell ? desired_cell.v : '';

  /* Get the value */

  // Create a document
  const doc = new PDFDocument();

  // Pipe its output somewhere, like to a file or HTTP response
  // See below for browser usage
  doc.pipe(
    fs.createWriteStream(
      'Invoices ' +
        Reference_cell_value +
        ' Vendor ' +
        VendorNo_cell_value +
        '.pdf'
    )
  );
  doc.fontSize(9).text(Field1_cell_value, 340, 10);
  doc.text(Field2_cell_value, 340, 20);
  doc.text(Field3_cell_value, 340, 30);
  doc.text('Phone:' + Phone_cell_value, 340, 50);
  doc.text('Fax:' + Fax_cell_value, 340, 70);
  doc.text('Email:' + Email_cell_value, 310, 90);
  doc.text(HC_cell_value, 50, 150);
  doc.text(Field4_cell_value, 50, 170);
  doc.text('Vendor no:' + VendorNo_cell_value, 180, 170);
  doc.text('Invoice Date:' + InvoiceDate_cell_value, 50, 190);
  doc.text('Reference:' + Reference_cell_value, 180, 190);
  doc
    .moveTo(50, 220)
    .lineTo(550, 220)
    .stroke();
  doc.text(Item_cell_value, 70, 255);
  doc.text('Quantity', 300, 240);
  doc.text('Price', 380, 240);
  doc.text('per', 410, 240);
  doc.text('Amount (' + currency_cell_value + ')', 450, 240);
  doc.text(Quantity_cell_value + ' ' + UOM_cell_value, 300, 255);
  doc.text(PricePerUOM_cell_value + ' ' + currency_cell_value, 380, 255);
  //doc.text('per' + ' ' + UOM_cell_value, 350, 230);
  doc.text(Amount_cell_value, 450, 255);
  doc.text('Total', 300, 450);
  doc.text('Tax', 300, 470);
  doc.text(TaxCode_cell_value, 320, 470);
  doc.text(TaxPercentage_cell_value+'%', 340, 470);
  doc.text(Amount_cell_value,450, 450);
  doc.text(TaxAmount_cell_value,450, 470);
  doc
    .moveTo(50, 500)
    .lineTo(550, 500)
    .stroke();
  doc.text('Total', 60, 520);
  doc.text(Quantity_cell_value + ' ' + UOM_cell_value, 300, 520);
  doc.text(grossAmount_cell_value, 450, 520);
  doc.text('Account Holder:' + AccountHolder_cell_value, 50, 550);
  doc.text('IBAN:' + IBANF_cell_value, 50, 560);
  // Finalize PDF file
  doc.end();
}
