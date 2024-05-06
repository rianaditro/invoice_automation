const TEMPLATE_FILE_ID = '1CLsjcH5H9KoXWACiZDi8jvse1fEXeJ9rSpN6H7tdltQ';
const DESTINATION_FOLDER_ID = '1WB5-UcratCA0eVZt58ZKH8x8w0ZqNfut';
const DESTINATION_FOLDER_PDF_ID = '1MoAEflNjuY7eGs8lp8yKTkAowUtqe-H9';
const CURRENCY_SIGN = 'Rp';

// Converts a float to a string value in the desired currency format
function toCurrency(num) {
    var fmt = Number(num).toFixed(2).replace(".", ",").replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
    return `${CURRENCY_SIGN} ${fmt}`;
}

// Format datetimes to: DD-MM-YYYY
function toDateFmt(dt_string) {
    var millis = Date.parse(dt_string);
    var date = new Date(millis);
    var year = date.getFullYear();
    var month = ("0" + (date.getMonth() + 1)).slice(-2);
    var day = ("0" + date.getDate()).slice(-2);

    return `${day}-${month}-${year}`;
}

// Parse and extract the data submitted through the form.
function parseFormData(values, header) {
    // Set temporary variables to hold prices and data.
    var subtotal = 0;
    var ongkir = 0;
    var quantities = [];
    var unitPrices = [];
    var discounts = [];

    var response_data = {};

    // Iterate through all of our response data and add the keys (headers)
    // and values (data) to the response dictionary object.
    for (var i = 0; i < values.length; i++) {
        // Extract the key and value
        var key = header[i];
        var value = values[i];

        // Format dates
        if (key.includes("Invoice Date")) {
            response_data[key] = toDateFmt(value);
            continue;
        }

        // Append & format harga satuan
        if (key.includes("Harga Satuan") && !isempty(value)) {
            response_data[key] = toCurrency(value);
            unitPrices.push(Number(value));
            continue;
        }
        if (key.includes("Harga Satuan") && isempty(value)) {
            response_data[key] = value;
            unitPrices.push(Number(value));
            continue;
        }

        // Append quantity
        if (key.includes("Quantity") && !isempty(value)) {
            response_data[key] = value;
            quantities.push(Number(value));
            continue;
        }
        if (key.includes("Quantity") && isempty(value)) {
            response_data[key] = value;
            quantities.push(Number(value));
            continue;
        }

        // Append discount
        if (key.includes("Diskon") && !isempty(value)) {
            response_data[key] = value;
            discounts.push(Number(value));
            continue;
        }
        if (key.includes("Diskon") && isempty(value)) {
            response_data[key] = value;
            discounts.push(Number(value));
            continue;
        }

        // Format ongkir currency
        if (key.includes("Ongkir") && !isempty(value)) {
            response_data[key] = toCurrency(value);
            ongkir += value
            continue;
        }

        response_data[key] = value;
    }

    // Once all data is added, we'll adjust the subtotal and total
    if (quantities.length === unitPrices.length) {
        for (var i = 0; i < quantities.length; i++) {
            totalValue = unitPrices[i] * quantities[i];
            discountInValue = (discounts[i] / 100) * totalValue;
            if(unitPrices[i] == 0 && quantities[i] == 0){
              response_data["price" + (i + 1).toString()] = "";
            } else {
              response_data["price" + (i + 1).toString()] = toCurrency((totalValue - discountInValue));
            }
            subtotal += totalValue - discountInValue;
        }
    }

    // discount += (Number(discountInPercentage) / 100 * Number(subtotal));
    // response_data["Discount"] = toCurrency(discount);
    var ppn = subtotal * (11 / 100);
    response_data["sub_total_price"] = toCurrency(subtotal);
    response_data["ppn"] = toCurrency(subtotal * (11 / 100));
    response_data["total_price"] = toCurrency((subtotal + ppn + ongkir));

    // Logger.log("Parsed data: " + response_data);
    return response_data;
}

// Helper function to inject data into the template
function populateTemplate(document, response_data) {

    // Get the document header and body (which contains the text we'll be replacing).
    var document_header = document.getHeader();
    var document_body = document.getBody();

    // Replace variables in the header
    for (var key in response_data) {
        var match_text = `{{${key}}}`;
        var value = response_data[key];
        // Logger.log(match_text+":"+value);

        // Replace our template with the final values
        document_body.replaceText(match_text, value);
    }
}

// Function to populate the template form
function createDocFromForm() {

    // Get active sheet and tab of our response data spreadsheet.
    var sheet = SpreadsheetApp.getActiveSheet();
    var last_row = sheet.getLastRow() - 1;

    // Get the data from the spreadsheet.
    var range = sheet.getDataRange();

    // Identify the most recent entry and save the data in a variable.
    var data = range.getValues()[last_row];

    // Extract the headers of the response data to automate string replacement in our template.
    var headers = range.getValues()[0];

    // Parse the form data.
    var response_data = parseFormData(data, headers);

    // Retreive the template file and destination folder.
    var template_file = DriveApp.getFileById(TEMPLATE_FILE_ID);
    var target_folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    var pdf_folder = DriveApp.getFolderById(DESTINATION_FOLDER_PDF_ID);    

    // Copy the template file so we can populate it with our data.
    // The name of the file will be the marketing name and the invoice number in the format: DATE_COMPANY_NUMBER
    var filename = `${response_data["Marketing Name"]}/${response_data["Invoice Date"]}/${response_data["Invoice Number"]}`;
    var document_copy = template_file.makeCopy(filename, target_folder);

    // Open the copy.
    var document = DocumentApp.openById(document_copy.getId());
    
    // Populate the template with our form responses and save the file.
    populateTemplate(document, response_data);
    document.saveAndClose();

    // save as pdf
    var pdfBlob = document.getAs('application/pdf');
    var pdfName = document.getName() + ".pdf";

    pdf_folder.createFile(pdfBlob).setName(pdfName);


}

function isempty(entry) {
    if (entry == undefined) {
        return true;
    }

    if (entry == null) {
        return true;
    }
    var tempstr = entry.toString();

    tempstr = tempstr.replace(/[\r\n\t\s]+$/, "");
    tempstr = tempstr.replace(/^[\r\n\t\s]+/, "");
    if (tempstr.length == 0) {
        return true;
    }

    return false;
}
