// INVOICE HANDLER SCRIPT BY AMEENAH RIFFIN
var script = document.createElement('script');
var version = '1.9.1'
script.type = 'module';
script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.2.67/pdf.min.mjs';
script.onload = function() {
    /* Set the path to the PDF worker script */
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.2.67/pdf.worker.min.mjs';

    /* Create sidebar container */
    var sidebar = document.createElement('div');
    sidebar.id = 'invoiceSidebar';
    sidebar.style.position = 'fixed';
    sidebar.style.top = '0';
    sidebar.style.right = '0';
    sidebar.style.width = '300px';
    sidebar.style.height = '100%';
    sidebar.style.background = 'black';
    sidebar.style.color = 'lightgrey';
    sidebar.style.zIndex = '999999';

    /* Create title */
    var title = document.createElement('div');
    title.textContent = 'Invoice Handler';
    title.style.padding = '10px';
    title.style.fontSize = '20px';
    title.style.textAlign = 'center';
    title.style.borderBottom = '1px solid lightgrey';
    sidebar.appendChild(title);

    /* Show current version */
    var versionDiv = document.createElement('div');
    versionDiv.textContent = 'Version ' + version;
    versionDiv.style.padding = '10px';
    versionDiv.style.fontSize = '14px';
    versionDiv.style.textAlign = 'center';
    sidebar.appendChild(versionDiv);

    /* Create buttons */
    var buttonsDiv = document.createElement('div');
    buttonsDiv.style.padding = '10px';

    var newPOButton = document.createElement('button');
    newPOButton.textContent = 'New Purchase Order';
    newPOButton.style.marginBottom = '5px';
    newPOButton.onclick = function () {
        window.open('https://www.yardiaspuk12.com/74757populo/pages/po.aspx');
    };
    buttonsDiv.appendChild(newPOButton);

    var viewPOButton = document.createElement('button');
    viewPOButton.textContent = 'View Purchase Order';
    viewPOButton.style.marginBottom = '5px';
    viewPOButton.onclick = function () {
        var num = prompt('Enter a purchase order number:');
        if (num !== null) {
            window.open('https://www.yardiaspuk12.com/74757populo/Pages/PO.aspx?Id=' + num);
        }
    };
    buttonsDiv.appendChild(viewPOButton);

    var poEmailerButton = document.createElement('button');
    poEmailerButton.textContent = 'PO Emailer';
    poEmailerButton.onclick = function () {
        var PO_NUMBER = document.getElementById('CommonWorkflow_txtRecordId').value || '';
        var CO_NUMBER_element = document.querySelector('[id^="Ysilinklist2:LinkList:DataTable:row0:Value_anchor"]');
        var CO_NUMBER = CO_NUMBER_element ? CO_NUMBER_element.textContent.trim() : '';
        var VENDOR = document.getElementById('nametxt_Label').textContent.trim();
        var DESC = document.getElementById('Desctxt_TextBox').value || '';
        var TOTAL_AMOUNT = document.getElementById('TotalAmt_TextBox').value || '';
        var POST_AMOUNT = prompt('How much to post?', '');
        var postText = POST_AMOUNT ? 'Please post £' + POST_AMOUNT + '.' : '';
        var subject = 'Invoice - Purchase Order ' + PO_NUMBER + ' - ' + VENDOR;
        var body = 'Hi,\n\nPlease process the below:\n\n' +
            'Purchase Order: ' + PO_NUMBER + '\n' +
            (CO_NUMBER ? 'Change Order: ' + CO_NUMBER + '\n' : '') +
            postText + '\n\n' +
            'Description: ' + DESC + '\n' +
            'Amount: ' + TOTAL_AMOUNT;
        var mailtoLink = 'mailto:invoices@populoliving.co.uk' +
            '?subject=' + encodeURIComponent(subject) +
            '&body=' + encodeURIComponent(body);
        window.location.href = mailtoLink;
    };
    buttonsDiv.appendChild(poEmailerButton);

    sidebar.appendChild(buttonsDiv);


    /* Create settings dropdowns */
    var settings = document.createElement('div');
    settings.style.padding = '10px';

    var companyLabel = document.createElement('label');
    companyLabel.textContent = 'Company:';
    companyLabel.style.display = 'block';
    companyLabel.style.marginBottom = '5px';
    settings.appendChild(companyLabel);

    var companySelect = document.createElement('select');
    companySelect.id = 'companySelect';
    companySelect.style.width = '100%';
    companySelect.style.padding = '5px';
    companySelect.innerHTML = `
        <option value="british_gas">British Gas</option>
        <option value="thames_water">Thames Water</option>
        <option value="base_furnishings">Base Furnishings</option>
        <option value="southern_heat_pump">Southern Heat Pump</option>
        <option value="switch_2">Switch2 (WIP)</option>
        <option value="tlt">TLT</option>
        <option value="southern_heat_pump">CreditLadder (WIP)</option>
        <option value="southern_heat_pump">Adiuvo (WIP)</option>
        <option value="southern_heat_pump">AA4 (WIP)</option>
        <option value="southern_heat_pump">Clearway (WIP)</option>
        <option value="southern_heat_pump">Chigwell</option>
        <option value="homelet">Homelet (WIP)</option>
        <option value="southern_heat_pump">KONE (WIP)</option>
        <option value="ctax_newhamcouncil">Council Tax - Newham</option>
        <option value="ctax_towerhamlets">Council Tax - Tower Hamlets</option>
        <option value="southern_heat_pump">nPower (WIP)</option>
        <option value="southern_heat_pump">OnTheMarket (WIP)</option>
        <option value="southern_heat_pump">RightMove (WIP)</option>
    `;
    settings.appendChild(companySelect);

    sidebar.appendChild(settings);

    /* Create tab contents */
    var tabContents = document.createElement('div');
    tabContents.id = 'parsedDataDiv';
    tabContents.style.padding = '10px';
    sidebar.appendChild(tabContents);

    /* Create file input */
    var fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.pdf';
    fileInput.style.display = 'none';
    sidebar.appendChild(fileInput);

    /* Event listener for Upload Invoice Tab */
    var uploadInvoiceTab = document.createElement('div');
    uploadInvoiceTab.textContent = 'Upload Invoice';
    uploadInvoiceTab.style.cursor = 'pointer';
    uploadInvoiceTab.style.padding = '10px';
    uploadInvoiceTab.style.marginRight = '10px';
    uploadInvoiceTab.style.borderRadius = '5px';
    uploadInvoiceTab.style.background = '#333';
    uploadInvoiceTab.style.textAlign = 'center';
    sidebar.appendChild(uploadInvoiceTab);

    /* Event listener for Upload Invoice Tab */
    uploadInvoiceTab.addEventListener('click', function() {
        fileInput.click();
    });

    /* Event listener for file input change */
    fileInput.addEventListener('change', function(event) {
        var file = event.target.files[0];
        if (file) {
            var fileReader = new FileReader();
            fileReader.onload = function() {
                var typedarray = new Uint8Array(this.result);
                pdfjsLib.getDocument(typedarray).promise.then(function(pdf) {
                    pdf.getPage(1).then(function(page) {
                        page.getTextContent().then(function(textContent) {
                            var textItems = textContent.items;
                            var finalText = '';
                            for (var i = 0; i < textItems.length; i++) {
                                finalText += textItems[i].str + ' ';
                            }
                            // Parse the invoice data based on the selected company
                            var selectedCompany = companySelect.value;
                            window.parsedData = parseInvoice(finalText, selectedCompany); // Store parsed data globally

                            // Update the UI with parsed data
                            var parsedDataDiv = document.getElementById('parsedDataDiv');
                            parsedDataDiv.innerHTML = ''; // Clear previous data
                            for (var key in window.parsedData) {
                                var p = document.createElement('p');
                                p.textContent = `${key}: ${window.parsedData[key]}`;
                                parsedDataDiv.appendChild(p);
                            }

                            // Generate and execute bookmarklet code with parsed data
                            var bookmarkletCode = generateBookmarkletCode(window.parsedData);
                            executeBookmarklet(bookmarkletCode);
                        });
                    });
                });
            };
            fileReader.readAsArrayBuffer(file);
        }
    });

    /* Create close button */
    var closeButton = document.createElement('div');
    closeButton.textContent = 'x';
    closeButton.style.position = 'absolute';
    closeButton.style.top = '10px';
    closeButton.style.right = '10px';
    closeButton.style.cursor = 'pointer';
    closeButton.style.fontSize = '20px';
    closeButton.style.padding = '5px';
    closeButton.style.borderRadius = '50%';
    closeButton.style.background = '#333';
    closeButton.style.width = '30px';
    closeButton.style.height = '30px';
    closeButton.style.textAlign = 'center';
    closeButton.style.lineHeight = '30px';
    sidebar.appendChild(closeButton);

    /* Event listener for close button */
    closeButton.addEventListener('click', function() {
        sidebar.remove();
    });

    /* Append sidebar to the body */
    document.body.appendChild(sidebar);
};
document.head.appendChild(script);

const propertyCodes = {
    "firemans": "firemans",
    "nelson street": "nelsonst",
    "tanneries": "thetanne",
    "stratford": "thetanne",
    "gregory": "gregory",
    "didsbury close": "didslar",
    "wellington": "didsbury",
    "melbourne": "didslar",
    "grange road": "grange",
    "warmington": "doherty",
    "newton": "barking",
    "baxter": "baxter",
    "chargeable": "chargeab",
    "cheviot": "cheviot",
    "eves": "manorrd",
    "old fire": "oldfires",
    "stracey": "stracey",
    "solent": "grange",
    "wordsworth": "wordswor",
    "valetta": "plvg",
    "valletta": "plvg",
    "200 plaistow road": "plvg",
    "hall": "townhall",
    "brickyard": "brickyar",
    "alnwick": "baxter",
    "eve road": "manorrd",
    "station road" : "stracey"
};

const entities = {
    "firemans": "PL",
    "nelsonst" : "PL",
    "thetanne" : "PL",
    "gregory": "PL",
    "didslar" : "PH",
    "didsbury" : "PL",
    "grange" : "PH",
    "doherty" : "PH",
    "barking" : "PL",
    "baxter" : "PH",
    "chargeab" : "PH",
    "cheviot" : "PL",
    "manorrd" : "PL",
    "oldfires" : "PH",
    "stracey" : "PH",
    "wordswor": "PH",
    "plvg" : "PL",
    "londonro" : "PL",
    "townhall" : "PL",
    "brickyar" : "PL",
    "bricklar" : "PH",
};

const expenseType = [
    "Carpenters",
    "Carpenters - Scheme",
    "Chigwell Construction",
    "Communications",
    "Construction - General",
    "Construction - Scheme",
    "Customer Services",
    "Customer Services Plnd Maint",
    "Development - General",
    "Development - Scheme",
    "Legal",
    "Non-Carpenters",
    "Service Charge",
    "VendorCafe"
];

const workflow = [
    "Finance_WF",
    "Customer Services_WF",
    "Executive_WF",
    "Legal_WF",
    "Non Carpenters",
    "CS_P|anned_WF",
    "Airspace_Capex",
    "Carpenters",
    "Communications_WF",
    "Chigwell Construction_WF",
    "Procurement_WF"
];

function parseInvoice(text, company) {
    // Define parsing methods for each company
    var parsers = {
        british_gas: parseBritishGasInvoice,
        thames_water: parseThamesWaterInvoice,
        southern_heat_pump: parseSouthernHeatPumpInvoice,
        base_furnishings: parseBaseFurnishingsInvoice,
        switch_2: parseSwitchtwoInvoice,
        homelet: parseHomeletInvoice,
        ctax_towerhamlets: parseCtaxTowerHamlets,
        ctax_newhamcouncil: parseCTAXNewhamCouncil,
        tlt: parseTltInvoices,
    };

    // Execute the appropriate parsing method based on the selected company
    if (parsers.hasOwnProperty(company)) {
        return parsers[company](text);
    } else {
        return 'Parsing method not implemented for this company.';
    }
}

function parseBritishGasInvoice(text) {
    // Regex patterns to extract account number, owned amount, billing period, supply address, and bill date
    const accountNumberRegex = /Account number[^\S\n]*([\s\S]+?)Contact us/;
    const ownedAmountRegex = /exc VAT[^\S\n]*([\s\S]+?)VAT/;
    const billingPeriodRegex = /Billing period:[^\S\n]*([\s\S]+?)Your account/;
    const supplyAddressRegex = /Site address:[^\S\n]*([\s\S]+?)Your account/;
    const billDateRegex = /Bill date:[^\S\n]*([\s\S]+?)This is a VAT invoice/;
    const billnumRegex = /Bill number[^\S\n]*([\s\S]+?)This/;
    const accountNumberMatch = text.match(accountNumberRegex);
    const ownedAmountMatch = text.match(ownedAmountRegex);
    const billingPeriodMatch = text.match(billingPeriodRegex);
    const supplyAddressMatch = text.match(supplyAddressRegex);
    const billDateMatch = text.match(billDateRegex);
    const billnumMatch = text.match(billnumRegex);
    const accountNumber = accountNumberMatch ? accountNumberMatch[1].trim() : "";
    let ownedAmount = ownedAmountMatch ? ownedAmountMatch[1].trim() : "";
    const billingPeriod = billingPeriodMatch ? billingPeriodMatch[1].trim() : "";
    const supplyAddress = supplyAddressMatch ? supplyAddressMatch[1].trim() : "";
    const billDate = billDateMatch ? billDateMatch[1].trim() : "";
    const billnum = billnumMatch ? billnumMatch[1].trim() : "";

    let supplyTypeRegex = /Gas Charges|Electricity Charges/;
    let gl_code = "5100-1020"
    if ((supplyAddress.indexOf("Flat") === -1)) { // if this isn't a flat... it's communal
        if(supplyAddress.indexOf("Gas Charges") === -1) {
            gl_code = "5100-4020";
        } else {
            gl_code = "5100-4030";
        }
    } else {
        gl_code = "5100-1020";
    }



    if(ownedAmount) {
        ownedAmount = ownedAmount.replace(/£/, "");
    }
    // Return the parsed data
    return {
        accountNumber: accountNumber,
        invoicenum: billnum,
        ownedAmount: ownedAmount,
        billingPeriod: billingPeriod,
        supplyAddress: supplyAddress,
        billDate: billDate,
        vendor: "v0000632",
        gl_code: gl_code,
        taxrate: 3
    };
}

function parseHomeletInvoice(text) {
    const accountNumberRegex = /Account number[^\S\n]*([\s\S]+?)(?:For help|$)/;
    const ownedAmountRegex = /e £[^\S\n]*([\s\S]+?)(?:Please|$)/;
    const billingPeriodRegex = /Billing period[^\S\n]*([\s\S]+?)(?:Supply|$)/;
    const supplyAddressRegex = /Supply address[^\S\n]*([\s\S]+?)(?:What|$)/;
    const unitRegex = /Flat (\d+)/;

    const billDateRegex = /Bill date[^\S\n]*([\s\S]+?)(?:Billing|$)/;
    const accountNumberMatch = text.match(accountNumberRegex);
    const ownedAmountMatch = text.match(ownedAmountRegex);
    const billingPeriodMatch = text.match(billingPeriodRegex);
    const supplyAddressMatch = text.match(supplyAddressRegex);
    const billDateMatch = text.match(billDateRegex);
    const unitMatch = text.match(unitRegex);

    const accountNumber = accountNumberMatch ? accountNumberMatch[1] : "";
    const ownedAmount = ownedAmountMatch ? ownedAmountMatch[1] : "";
    const billingPeriod = billingPeriodMatch ? billingPeriodMatch[1] : "";
    const supplyAddress = supplyAddressMatch ? supplyAddressMatch[1].trim() : "";
    const billDate = billDateMatch ? billDateMatch[1].trim() : "";
    const billnum = accountNumber;
    const unit = unitMatch ? unitMatch[1].trim() : "";

    // Return the parsed data
    return {
        accountNumber: accountNumber,
        invoicenum: billnum,
        ownedAmount: ownedAmount,
        billingPeriod: billingPeriod,
        supplyAddress: supplyAddress,
        billDate: billDate,
        vendor: "v0000308",
        gl_code: "5100-1020",
        unit: unit,
        taxrate: 3
    };
}



function parseThamesWaterInvoice(text) {
    const accountNumberRegex = /Account number[^\S\n]*([\s\S]+?)(?:For help|$)/;
    const ownedAmountRegex = /e £[^\S\n]*([\s\S]+?)(?:Please|$)/;
    const billingPeriodRegex = /Billing period[^\S\n]*([\s\S]+?)(?:Supply|$)/;
    const supplyAddressRegex = /Supply address[^\S\n]*([\s\S]+?)(?:What|$)/;
    const unitRegex = /Flat (\d+)/;

    const billDateRegex = /Bill date[^\S\n]*([\s\S]+?)(?:Billing|$)/;
    const accountNumberMatch = text.match(accountNumberRegex);
    const ownedAmountMatch = text.match(ownedAmountRegex);
    const billingPeriodMatch = text.match(billingPeriodRegex);
    const supplyAddressMatch = text.match(supplyAddressRegex);
    const billDateMatch = text.match(billDateRegex);
    const unitMatch = text.match(unitRegex);

    const accountNumber = accountNumberMatch ? accountNumberMatch[1] : "";
    const ownedAmount = ownedAmountMatch ? ownedAmountMatch[1] : "";
    const billingPeriod = billingPeriodMatch ? billingPeriodMatch[1] : "";
    const supplyAddress = supplyAddressMatch ? supplyAddressMatch[1].trim() : "";
    const billDate = billDateMatch ? billDateMatch[1].trim() : "";
    const billnum = accountNumber;
    const unit = unitMatch ? unitMatch[1].trim() : "";

    // Return the parsed data
    return {
        accountNumber: accountNumber,
        invoicenum: billnum,
        ownedAmount: ownedAmount,
        billingPeriod: billingPeriod,
        supplyAddress: supplyAddress,
        billDate: billDate,
        vendor: "v0000309",
        gl_code: "5100-1020",
        unit: unit,
        taxrate: 3
    };
}


function parseSouthernHeatPumpInvoice(text) {
    const accountNumberRegex = /Number[^\S\n]*([\s\S]+?)(?:Amount|Reference|$)/;
    const ownedAmountRegex = /Subtotal[^\S\n]*([\s\S]+?)(?:TOTAL|$)/;
    const supplyAddressRegex = /Reference[^\S\n]*([\s\S]+?)(?:VAT|$)/;
    const descRegex = /Due Date:[^\S\n]*([\s\S]+?)(?:Please|$)/;
    const billDateRegex = /Invoice Date[^\S\n]*([\s\S]+?)(?:Invoice Number|$)/;
    const unitRegex = /Reference[^\S\n]*([\s\S]+?)(?: |$)/;

    const accountNumberMatch = text.match(accountNumberRegex);
    const ownedAmountMatch = text.match(ownedAmountRegex);
    const supplyAddressMatch = text.match(supplyAddressRegex);
    const billDateMatch = text.match(billDateRegex);
    const descMatch = text.match(descRegex);
    const unitMatch = text.match(unitRegex);

    const accountNumber = accountNumberMatch ? accountNumberMatch[1] : "";
    const ownedAmount = ownedAmountMatch ? ownedAmountMatch[1] : "";
    const supplyAddress = supplyAddressMatch ? supplyAddressMatch[1].trim() : "";
    const billDate = billDateMatch ? billDateMatch[1].trim() : "";
    const desc = descMatch ? descMatch[1].trim() : "";
    const unit = unitMatch ? unitMatch[1].trim() : "";

    return {
        accountNumber: accountNumber,
        invoicenum: accountNumber,
        ownedAmount: ownedAmount,
        billingPeriod: `Due: ${desc}`,
        supplyAddress: supplyAddress,
        billDate: billDate,
        vendor: "v0000822",
        gl_code: "5100-3160",
        unit: unit,
        taxrate: 3
    };
}

function parseCtaxTowerHamlets(inputData) {
    const invoiceNumberRegex = /Account Ref: [^\S\n]*([\s\S]+?)(?: BILL|$)/;
    const ownedAmountRegex = /BILL TOTAL £[^\S\n]*([\s\S]+?)(?: BILL | T|$)/;
    const PODescRegex = /fund adult social care[^\S\n]*([\s\S]+?)(?:£|$)/;
    const supplyAddressRegex = /Property Address: [^\S\n]*([\s\S]+?)(?: £|$)/;
    const billDateRegex = /Date: [^\S\n]*([\s\S]+?)(?:Reason|$)/;

    const invoiceNumberMatch = inputData.match(invoiceNumberRegex);
    const ownedAmountMatch = inputData.match(ownedAmountRegex);
    const PODescMatch = inputData.match(PODescRegex);
    const supplyAddressMatch = inputData.match(supplyAddressRegex);
    const billDateMatch = inputData.match(billDateRegex);

    const invoiceNumber = invoiceNumberMatch ? invoiceNumberMatch[1] : "";
    const ownedAmount = ownedAmountMatch ? ownedAmountMatch[1] : "";
    const PODesc = PODescMatch ? PODescMatch[1] : "";
    const supplyAddress = supplyAddressMatch ? supplyAddressMatch[1].trim() : "";
    const billDate = billDateMatch ? billDateMatch[1].trim() : "";

    return {
        accountNumber: invoiceNumber,
        invoicenum: invoiceNumber,
        ownedAmount: ownedAmount,
        billingPeriod: PODesc,
        supplyAddress: supplyAddress,
        billDate: billDate,
        vendor: "v0000273",
        gl_code: "5100-1050",
        unit: "",
        taxrate: 1
    };
}

function parseCTAXNewhamCouncil(inputData) {
    const invoiceNumberRegex = /(\d{9}_\d{8}_CTMA01)/;
    const ownedAmountRegex = /(\d+\.\d+)\s*to be paid/i;
    const PODescRegex = /Reason for bill([^0-9]+)/;
    const supplyAddressRegex = /FLAT ([^C]+)/;
    const billDateRegex = /E15 4QZ([^C]+)/;

    const invoiceNumberMatch = inputData.match(invoiceNumberRegex);
    const ownedAmountMatch = inputData.match(ownedAmountRegex);
    const PODescMatch = inputData.match(PODescRegex);
    const supplyAddressMatch = inputData.match(supplyAddressRegex);
    const billDateMatch = inputData.match(billDateRegex);

    const invoiceNumber = invoiceNumberMatch ? invoiceNumberMatch[1] : "";
    const ownedAmount = ownedAmountMatch ? ownedAmountMatch[1] : "";
    const PODesc = PODescMatch ? PODescMatch[1] : "";
    const supplyAddress = supplyAddressMatch ? supplyAddressMatch[1].trim() : "";
    const billDate = billDateMatch ? billDateMatch[1].trim() : "";

    return {
        accountNumber: invoiceNumber,
        invoicenum: invoiceNumber,
        ownedAmount: ownedAmount,
        billingPeriod: PODesc,
        supplyAddress: supplyAddress,
        billDate: billDate,
        vendor: "v0000096",
        gl_code: "5100-1050",
        unit: "",
        taxrate: 1
    };
}

function parseBaseFurnishingsInvoice(inputData) {
    const invoiceNumberRegex = /Invoice No: [^\S\n]*([\s\S]+?)(?: London|$)/;
    const ownedAmountRegex = /Net Total [^\S\n]*([\s\S]+?)(?:VAT| T|$)/;
    const PODescRegex = /Total Cost([\s\S]*?)Sub Total/;
    const supplyAddressRegex = /Ref: [^\S\n]*([\s\S]+?)(?:Description|$)/;
    const billDateRegex = /Dated:[^\S\n]*([\s\S]+?)(?:E|$)/;

    const invoiceNumberMatch = inputData.match(invoiceNumberRegex);
    const ownedAmountMatch = inputData.match(ownedAmountRegex);
    const PODescMatch = inputData.match(PODescRegex);
    const supplyAddressMatch = inputData.match(supplyAddressRegex);
    const billDateMatch = inputData.match(billDateRegex);

    const invoiceNumber = invoiceNumberMatch ? invoiceNumberMatch[1] : "";
    const ownedAmount = ownedAmountMatch ? ownedAmountMatch[1] : "";
    let PODesc = PODescMatch ? PODescMatch[1] : "";
    const supplyAddress = supplyAddressMatch ? supplyAddressMatch[1].trim() : "";
    const billDate = billDateMatch ? billDateMatch[1].trim() : "";

    PODesc = PODesc.replace(/[\d£.]/g, '').trim();

    return {
        accountNumber: invoiceNumber,
        invoicenum: invoiceNumber,
        ownedAmount: ownedAmount,
        billingPeriod: "",
        supplyAddress: supplyAddress,
        billDate: billDate,
        vendor: "v0000414",
        gl_code: "1600-0020",
        unit: "",
        taxrate: 1,
        po_desc: PODesc,
    };    
}

function parseSwitchtwoInvoice(inputData) {
    const invoiceNumberRegex = /Invoice No: [^\S\n]*([\s\S]+?)(?: London|$)/;
    const ownedAmountRegex = /Net Total [^\S\n]*([\s\S]+?)(?:VAT| T|$)/;
    const PODescRegex = /Total Cost([\s\S]*?)Sub Total/;
    const supplyAddressRegex = /Ref: [^\S\n]*([\s\S]+?)(?:Description|$)/;
    const billDateRegex = /Dated:[^\S\n]*([\s\S]+?)(?:E|$)/;

    const invoiceNumberMatch = inputData.match(invoiceNumberRegex);
    const ownedAmountMatch = inputData.match(ownedAmountRegex);
    const PODescMatch = inputData.match(PODescRegex);
    const supplyAddressMatch = inputData.match(supplyAddressRegex);
    const billDateMatch = inputData.match(billDateRegex);

    const invoiceNumber = invoiceNumberMatch ? invoiceNumberMatch[1] : "";
    const ownedAmount = ownedAmountMatch ? ownedAmountMatch[1] : "";
    let PODesc = PODescMatch ? PODescMatch[1] : "";
    const supplyAddress = supplyAddressMatch ? supplyAddressMatch[1].trim() : "";
    const billDate = billDateMatch ? billDateMatch[1].trim() : "";

    PODesc = PODesc.replace(/[\d£.]/g, '').trim();

    return {
        accountNumber: invoiceNumber,
        invoicenum: invoiceNumber,
        ownedAmount: ownedAmount,
        billingPeriod: "",
        supplyAddress: supplyAddress,
        billDate: billDate,
        vendor: "v0000414",
        gl_code: "1600-0020",
        unit: "",
        taxrate: 1,
        po_desc: PODesc,
    };  
}

function parseTltInvoices(inputData) {
    const invoiceNumberRegex = /Invoice Number [^\S\n]*([\s\S]+?)(?: VAT|$)/;
    const ownedAmountRegex = /Sub-totals [^\S\n]*([\s\S]+?)(?:VAT| T|$)/;
    const PODescRegex = /Your Ref ([\s\S]*?)Sub Totals/;
    const supplyAddressRegex = /Matter: ([\s\S]*?)Sub Totals/;
    const billDateRegex = /Date[^\S\n]*([\s\S]+?)(?:Our Ref|$)/;

    const invoiceNumberMatch = inputData.match(invoiceNumberRegex);
    const ownedAmountMatch = inputData.match(ownedAmountRegex);
    const PODescMatch = inputData.match(PODescRegex);
    const supplyAddressMatch = inputData.match(supplyAddressRegex);
    const billDateMatch = inputData.match(billDateRegex);

    const invoiceNumber = invoiceNumberMatch ? invoiceNumberMatch[1] : "";
    const ownedAmount = ownedAmountMatch ? ownedAmountMatch[1] : "";
    let PODesc = PODescMatch ? PODescMatch[1] : "";
    const supplyAddress = supplyAddressMatch ? supplyAddressMatch[1].trim() : "";
    const billDate = billDateMatch ? billDateMatch[1].trim() : "";

    PODesc = PODesc.replace(/[\d£.]/g, '').trim();

    return {
        accountNumber: invoiceNumber,
        invoicenum: invoiceNumber,
        ownedAmount: ownedAmount,
        billingPeriod: "",
        supplyAddress: supplyAddress,
        billDate: billDate,
        vendor: "v0000352",
        gl_code: "5100-2011",
        unit: "",
        taxrate: 1,
        po_desc: PODesc,
    };  
}

function getpropertycodebyaddress(supplyAddress) {
    let supplyAddressWords = null;
    if (supplyAddress) {
        supplyAddressWords = supplyAddress.toLowerCase().replace(/[,\.]/g, '').split(/\s+/); // Convert to lowercase and remove commas and dots
        for (const word of supplyAddressWords) {
            const lowerCaseWord = word.toLowerCase(); // Convert to lowercase
            if (propertyCodes[lowerCaseWord]) {
                return propertyCodes[lowerCaseWord];
            }
        }
    }
}

function getentitybypropcode(propertyId) {
    return entities[propertyId]
}

function generateBookmarkletCode(parsedData) {
    const { accountNumber, invoicenum, billingPeriod, ownedAmount, supplyAddress, billDate, vendor, gl_code, unit, taxrate, po_desc} = parsedData;

    let propertyId = "";
    let propentity = "";
 // let communal = false;
    let recoverability = 11;

    propertyId = getpropertycodebyaddress(supplyAddress);

    if(propertyId) {
        propentity = getentitybypropcode(propertyId)
    }

    if((propentity == "PH") && (supplyAddress.indexOf("Flat") === -1)) {
//      communal = true;
        recoverability = 21;
    }

    let totalamount = ownedAmount.replace(/[^0-9.]/g, '').replace(/\s/g, '');

    let vatAmt = taxrate

    // Note: To add:
    // document.getElementById('DetailsGrid_DataTable_row0_Segment3Lookup_Segment3').value = '${unit}';
    // document.getElementById('ExpTypeList_DropDownList').selectedIndex = 7;

    return `
        
        document.getElementById('CommonWorkflow_WorkFlowID_DropDownList').selectedIndex = 2;
        document.getElementById('DetailsGrid_DataTable_row0_VatRateId').selectedIndex = '${vatAmt}';
        document.getElementById('DetailsGrid_DataTable_row0_PropIdLookup_PropId').value = '${propertyId}';
        document.getElementById('VendorLookup_LookupCode').value = '${vendor}';
        document.getElementById('DetailsGrid_DataTable_row0_podetdesc').value = '${invoicenum} ${billingPeriod ? ` - ${billingPeriod}` : ""} ${po_desc ? ` - ${po_desc}` : ""} ';
        document.getElementById('DetailsGrid_DataTable_row0_AcctIdLookup_AcctId').value = '${gl_code}'
        document.getElementById('DetailsGrid_DataTable_row0_trantotal').value = '${totalamount}';
        document.getElementById('DetailsGrid_DataTable_row0_Segment1Lookup_Segment1').value = '${recoverability}';
        document.getElementById('Desctxt_TextBox').value = '${supplyAddress} - (${propentity}) - Invoice Number: ${invoicenum} - ${po_desc ? `${po_desc}` : ""} ${accountNumber ? `Account Number: ${accountNumber}` : ""} ${billingPeriod ? ` - Billing Period: ${billingPeriod}` : ""} ${billDate ? ` - Date: ${billDate}` : ""}';
  
        document.getElementById('DetailsGrid_DataTable_row0_Segment5Lookup_Segment5').value = 'Customer Services';
    `;
}

function executeBookmarklet(code) {
    // Create a script element with the bookmarklet code
    var script = document.createElement('script');
    script.innerHTML = code;

    // Append the script to the document body to execute it
    document.body.appendChild(script);
}
