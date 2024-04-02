import sql from "mssql";
import { dbConfig } from "../config/dbConfig.js";
import ExcelJS from "exceljs";
import { emailRecipients } from "../config/emailConfig.js";
import { sendEmailWithAttachment } from "../services/emailService.js";

async function invoiceReportMTD_CH() {
  try {
    await sql.connect(dbConfig);

    const invCHOrderInvoicesresult = await sql.query`SELECT
                ol.CompanyName as 'CustomerName',
                abf.Funder,
                oi.InvoiceTo,
                oi.InvoiceNumber,
                CONVERT(varchar, oi.InvoiceDate, 103) as 'InvoiceDate',
                oi.Invoice as 'RecordType',
                oi.InvoiceTypeName,
                FORMAT(oi.InvoiceNetAmount, 'c', 'en-GB') as 'InvoiceNetAmount',
                ISNULL(FORMAT(oi.InvoiceVAT, 'c', 'en-GB'), FORMAT(0, 'c', 'en-GB')) as 'InvoiceVAT',
                FORMAT(oi.InvoiceTotal, 'c', 'en-GB') as 'InvoiceTotal',
                oi.InvoiceDescription,
                dm.VehicleRegistration,
              dm.DepartmentName,
              dm.Channels

                FROM [AMTMAASPROD].[dbo].[OrderInvoices] oi
                LEFT JOIN [AMTMAASPROD].[dbo].[Orders] o ON oi.OrderId = o.OrderId
                LEFT JOIN [AMTMAASPROD].[dbo].[FleetSelector] fs ON o.FleetSelectorId = fs.FleetSelectorID
                LEFT JOIN [AMTMAASPROD].[dbo].[OpportunityLead] ol ON o.OrderLeadId = ol.LeadId
                LEFT JOIN [AMTMAASPROD].[dbo].[AcquisitionBrokerFunder] abf ON o.AcquisitionId = abf.AcquisitionId
                LEFT JOIN [AMTMAASPROD].[dbo].[VM_OrderList] dm ON o.OrderId = dm.OrderId

                WHERE MONTH(oi.InvoiceDate) = MONTH(GETDATE())-1
                AND YEAR(oi.InvoiceDate) = YEAR(GETDATE())
                AND oi.InvoiceNumber <> ''
                AND dm.Channels not in ('Own Book',' Own Book')`;

    console.log(invCHOrderInvoicesresult.length);

    // Create and populate an Excel file with the data
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Invoices");

    // To get this data I need these queries from these tables:
    // OrderInvoices
    // Orders
    // FleetSelector - RegistrationNumber
    // OpportunityLead - CompanyName
    // Acquisition - AcquisitionID
    // AcquisitionBrokerFunder - Funder
    // Define columns in the worksheet
    sheet.columns = [
      { header: "Customer Name", key: "companyName", width: 20 },
      { header: "Funder", key: "Funder", width: 20 },
      { header: "Invoice To", key: "InvoiceTo", width: 20 },
      { header: "Invoice Number", key: "InvoiceNumber", width: 15 },
      { header: "Invoice Date", key: "InvoiceDate", width: 15 },
      { header: "RecordType", key: "Invoice", width: 15 },
      { header: "Invoice Type", key: "InvoiceTypeName", width: 15 },
      { header: "Value", key: "InvoiceNetAmount", width: 15 },
      { header: "Vat", key: "InvoiceVAT", width: 15 },
      { header: "Total Value", key: "InvoiceTotal", width: 15 },
      { header: "Invoice Description", key: "InvoiceDescription", width: 25 },
      { header: "Registration Number", key: "RegistrationNumber", width: 20 },
    ];

    // Add rows to the worksheet
    invCHOrderInvoicesresult.recordset.forEach((record) => {
      sheet.addRow({
        companyName: record.CustomerName,
        Funder: record.Funder,
        InvoiceTo: record.InvoiceTo,
        InvoiceNumber: record.InvoiceNumber,
        InvoiceDate: record.InvoiceDate,
        Invoice: record.RecordType,
        InvoiceTypeName: record.InvoiceTypeName,
        InvoiceNetAmount: record.InvoiceNetAmount,
        InvoiceVAT: record.InvoiceVAT,
        InvoiceTotal: record.InvoiceTotal,
        InvoiceDescription: record.InvoiceDescription,
        RegistrationNumber: record.VehicleRegistration,
      });
    });

    // Write to a file
    const fileName = `CH_Invoices_${new Date().toISOString().replace(/:/g, "-")}.xlsx`;
    const buffer = await workbook.xlsx.writeBuffer();
    console.log(`Report generated: ${fileName} (${buffer.length} bytes)`);

    // Retrieve recipients from the config for the CHInvoiceReport
    const recipients = emailRecipients.reports["CHInvoiceReport"];
    const subject = "CH Invoice Report";
    const textContent = "Please find attached the CH Invoice Report.";

    await sendEmailWithAttachment(buffer, fileName, recipients, subject, textContent);

    return console.log(
      `Email sent to ${recipients} with subject ${subject} and attached ${fileName}`,
    ); // Return the path to the created Excel file
  } catch (err) {
    console.error("SQL error", err);
    throw err;
  }
}

export { invoiceReportMTD_CH };
