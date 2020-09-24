(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            
            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("Choose a color or keep the default color.");
                $('#button-text').text("Default");
                $('#button-desc').text("Choose a color or keep the default color");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

            $("#template-description").text("Choose a color or keep the default color.");
            $('#button-text').text("Change");
            $('#button-desc').text("Choose a color or keep the default color");
                
            //loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(hightlightHighestValue);
        });
    };


    function hightlightHighestValue() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the selected range and load its properties
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();

            sheet.getRanges().format.fill.color = "white";
            sheet.getRanges().format.font.name = "Calibri";
            sheet.getRanges().format.font.size = 11;
            sheet.getRanges().format.rowHeight = 16.5;
            sheet.getRange("b:b").format.columnWidth = 12;
            sheet.getRange("h:h").format.columnWidth = 12;
            sheet.getRange("i:i").format.columnWidth = 12;
            sheet.getRange("c:c").format.columnWidth = 95;
            sheet.getRange("d:d").format.columnWidth = 95;
            sheet.getRange("e:e").format.columnWidth = 105;
            sheet.getRange("f:f").format.columnWidth = 80;
            sheet.getRange("g:g").format.columnWidth = 80;
            sheet.getRange("15:15").format.rowHeight = 28;
            sheet.getRange("18:18").format.rowHeight = 28;
            sheet.getRange("17:17").format.rowHeight = 21;
            sheet.getRange("32:32").format.rowHeight = 1.2;

            //constant variables
            const invoice = sheet.getRange("e1");
            const total = sheet.getRange("f34");
            const salutation = sheet.getRange("c39");
            //cells
            const lines = sheet.getRanges("c15:g15");
            const sublines = sheet.getRanges("c18:f18");
            const totalbox = sheet.getRanges("g34");
            const border = sheet.getRanges("g32");
            const color = document.getElementById("picker").value;

            //invoice text
            const invoicetext = sheet.getRange("e1:g1");
            invoicetext.merge();
            invoicetext.format.horizontalAlignment = "Right";
            invoicetext.values = "INVOICE";
            invoicetext.format.font.size = 60;
            invoicetext.format.font.bold = true;
            invoicetext.format.rowHeight = 95;

            //variables
            var compname = sheet.getRange("c3");
            var address = sheet.getRange("c4:d4");
            var city = sheet.getRange("c5:d5");
            var phone = sheet.getRange("c6:d6");
            var email = sheet.getRange("c7:d7");
            var invoiceNo = sheet.getRange("e4");
            var date = sheet.getRange("e5");
            var customerId = sheet.getRange("e6");
            var customername = sheet.getRange("c9:d9");
            var sub = sheet.getRange("c10:d10");
            var payment = sheet.getRange("c33:d33");
            var card = sheet.getRange("c34");
            var check = sheet.getRange("c37");
            //info
            compname.format.font.bold = true;
            compname.format.horizontalAlignment = "Left";
            compname.values = "Company Name";
            compname.format.font.size = 11;

            address.merge();
            address.format.font.bold = false;
            address.format.horizontalAlignment = "Left";
            address.values = "Address";
            address.format.font.size = 11;

            city.merge();
            city.format.font.bold = false;
            city.format.horizontalAlignment = "Left";
            city.values = "City, State, Zip";
            city.format.font.size = 11;

            phone.merge();
            phone.format.font.bold = false;
            phone.format.horizontalAlignment = "Left";
            phone.values = "Phone";
            phone.format.font.size = 11;

            email.merge();
            email.format.font.bold = false;
            email.format.horizontalAlignment = "Left";
            email.values = "Email";
            email.format.font.size = 11;

            invoiceNo.merge();
            invoiceNo.format.font.bold = true;
            invoiceNo.format.horizontalAlignment = "Right";
            invoiceNo.values = "Invoice No :";
            invoiceNo.format.font.size = 11;

            date.merge();
            date.format.font.bold = true;
            date.format.horizontalAlignment = "Right";
            date.values = "Date :";
            date.format.font.size = 11;

            customerId.merge();
            customerId.format.font.bold = true;
            customerId.format.horizontalAlignment = "Right";
            customerId.values = "Customer ID :";
            customerId.format.font.size = 11;

            customername.merge();
            customername.format.font.bold = true;
            customername.format.horizontalAlignment = "Left";
            customername.values = "Customer";
            customername.format.font.size = 16;

            sub.merge();
            sub.format.font.bold = false;
            sub.format.horizontalAlignment = "Left";
            sub.values = "Sub Line";
            sub.format.font.size = 14;

            var job = sheet.getRange("c18")
            job.format.font.bold = true;
            job.format.horizontalAlignment = "Left";
            job.format.verticalAlignment = "Center";
            job.values = "Job";
            job.format.font.size = 12;

            var desc = sheet.getRange("d18:e18")
            desc.merge();
            desc.format.font.bold = true;
            desc.format.horizontalAlignment = "Left";
            desc.format.verticalAlignment = "Center";
            desc.values = "Description";
            desc.format.font.size = 12;

            var linetotal = sheet.getRange("f18")
            linetotal.format.font.bold = true;
            linetotal.format.horizontalAlignment = "Left";
            linetotal.format.verticalAlignment = "Center";
            linetotal.values = "Line Total";
            linetotal.format.font.size = 12;

            payment.merge();
            payment.format.font.bold = true;
            payment.format.horizontalAlignment = "Left";
            payment.values = "Payment Methods";
            payment.format.font.size = 12;

            card.format.font.bold = true;
            card.format.horizontalAlignment = "Left";
            card.values = "Card";
            card.format.font.size = 10;

            check.format.font.bold = true;
            check.format.horizontalAlignment = "Left";
            check.values = "Check";
            check.format.font.size = 10;


            var thanks = sheet.getRange("c39:g39");
            thanks.merge();
            thanks.format.font.bold = true;
            thanks.format.horizontalAlignment = "Center";
            thanks.format.font.size = 14;
            thanks.format.font.name = "Calibri";
            thanks.values = "Thank You For Your Business";

            //formulas
            sheet.getRange("f4").values = '="["&TODAY()&"]"';
            sheet.getRange("f5").values = new Date().toLocaleDateString("en-US", 'numeric', 'long', 'numeric');
            sheet.getRange("f5").format.horizontalAlignment = "left";
            sheet.getRange("g34").values = '=' + '"$"&' + 'SUM(F19:F31)';

            //fill colors
            invoice.format.font.color = color;
            total.format.font.color = color;
            salutation.format.font.color = color;

            lines.format.fill.color = color;
            sublines.format.fill.color = color;

            totalbox.format.fill.color = color;
            totalbox.format.font.color = "white";
            totalbox.format.font.size = 16;
            totalbox.format.horizontalAlignment = "Right"

            total.format.font.bold = true;
            total.format.font.size = 16;
            total.format.horizontalAlignment = "left";
            total.format.font.name = "Cambria";
            total.values = "Total";

            border.format.fill.color = color;
            sheet.tabColor = color;

            //lightgrey
            sheet.getRange("C16:F16").format.fill.color = "#E6E6E6";
            sheet.getRange("C20:E20").format.fill.color = "#E6E6E6";
            sheet.getRange("C22:E22").format.fill.color = "#E6E6E6";
            sheet.getRange("C24:E24").format.fill.color = "#E6E6E6";
            sheet.getRange("C26:E26").format.fill.color = "#E6E6E6";
            sheet.getRange("C28:E28").format.fill.color = "#E6E6E6";
            sheet.getRange("C30:e30").format.fill.color = "#E6E6E6";
            //dark grey
            sheet.getRange("g16").format.fill.color = "#CFCFCF";
            sheet.getRange("f20").format.fill.color = "#CFCFCF";
            sheet.getRange("f22").format.fill.color = "#CFCFCF";
            sheet.getRange("f24").format.fill.color = "#CFCFCF";
            sheet.getRange("f26").format.fill.color = "#CFCFCF";
            sheet.getRange("f28").format.fill.color = "#CFCFCF";
            sheet.getRange("f30").format.fill.color = "#CFCFCF";

            sheet.getRange("c4:d7").format.fill.clear();
            sheet.getRange("f4:f7").format.fill.clear();
            sheet.getRange("c9:d10").format.fill.clear();
            sheet.getRange("f34:f35").format.fill.clear();
          
            return ctx.sync();
        })
        .catch(errorHandler);
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
