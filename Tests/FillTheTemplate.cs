using docbuilder_net;

using CValue = docbuilder_net.CDocBuilderValue;
using CContext = docbuilder_net.CDocBuilderContext;
using CContextScope = docbuilder_net.CDocBuilderContextScope;

using static Helpers.Methods;
using static Tests.Tests;

namespace Tests
{
    public class FillTheTemplate
    {
        public static void FillInvoice(string dataSheetPath, string invoiceTemplatePath, CDocBuilder oBuilder)
        {
            // get data from the spreadsheet
            var doctype = "xlsx";
            OpenFile(oBuilder, doctype, templatesPath + dataSheetPath);
            CContext oContext = oBuilder.GetContext();
            CContextScope oScope = oContext.CreateScope();
            CValue oApi = GetApi(oContext);
            CValue oWorksheet = oApi.Call("GetActiveSheet");

            CValue oRangeValues = oWorksheet.Call("GetUsedRange").Call("GetValue");
            int lastRowIndex = (int)oRangeValues.GetLength() - 1;
            CValue invoiceData = oRangeValues[lastRowIndex];
            CValue invoiceDataHeaders = oRangeValues[0];
            Dictionary<string, string> invoice = createInvoiceDict(invoiceData, invoiceDataHeaders);
            oBuilder.CloseFile();

            // fill invoice template and save as a new document
            doctype = "docx";
            OpenFile(oBuilder, doctype, templatesPath + invoiceTemplatePath);
            oContext = oBuilder.GetContext();
            oScope = oContext.CreateScope();
            oApi = GetApi(oContext);
            
            CValue oDocument = oApi.Call("GetDocument");
            FillInvoice(oDocument, invoice);

            string filePath = GenerateInvoiceFilePath(invoice);
            SaveAndCloseFile(oBuilder, filePath, doctype);
        }
        public static Dictionary<string, string> createInvoiceDict(CValue data, CValue headers)
        {
            Dictionary<string, string> invoice = new Dictionary<string, string>();
            int colIndex = 0;
            while (headers[colIndex].ToString() != "")
            {
                invoice.Add(headers[colIndex].ToString(), data[colIndex].ToString());
                colIndex++;   
            }
            return invoice;
        }
        public static void FillInvoice(CValue oDocument, Dictionary<string, string> invoice)
        {
            CValue oAllTables = oDocument.Call("GetAllTables");
            FillCell(oAllTables, 0, 2, 4, invoice["Date"]);
            FillCell(oAllTables, 0, 3, 4, invoice["Id"]);

            FillCell(oAllTables, 1, 1, 0, invoice["Contact Name"]);
            FillCell(oAllTables, 1, 2, 0, invoice["Client Company"]);
            FillCell(oAllTables, 1, 3, 0, invoice["Address"]);
            FillCell(oAllTables, 1, 4, 0, invoice["Phone"]);
            FillCell(oAllTables, 1, 5, 0, invoice["Email"]);

            FillCell(oAllTables, 1, 1, 2, invoice["Dept"]);
            FillCell(oAllTables, 1, 2, 2, invoice["Client Company"]);
            FillCell(oAllTables, 1, 3, 2, invoice["Address"]);
            FillCell(oAllTables, 1, 4, 2, invoice["Phone"]);

            FillCell(oAllTables, 2, 1, 0, invoice["Description"]);
            FillCell(oAllTables, 2, 1, 1, invoice["Qty"]);
            FillCell(oAllTables, 2, 1, 2, "$ " + invoice["Unit Price"]);
            FillCell(oAllTables, 2, 1, 3, "$ " + invoice["Total Price"]);

            FillCell(oAllTables, 3, 0, 2, "$ " + invoice["Total Price"]);
            FillCell(oAllTables, 3, 1, 2, "$ " + invoice["Discount"]);
            FillCell(oAllTables, 3, 2, 2, "$ " + invoice["Price With Discount"]);
            FillCell(oAllTables, 3, 3, 2, invoice["Tax"] + "%");
            FillCell(oAllTables, 3, 4, 2, "$ " + invoice["Total Tax"]);
            FillCell(oAllTables, 3, 5, 2, "$ " + invoice["Shipping"]);
            FillCell(oAllTables, 3, 6, 2, "$ " + invoice["Balance Due"]);
        }
        public static void FillCell(CValue oAllTables, int table, int row, int cell, string value)
        {
            oAllTables[table].Call("GetRow", row).Call("GetCell", cell).Call("GetContent").Call("GetElement", 0).Call("AddText", value);
        }
        public static string GenerateInvoiceFilePath(Dictionary<string, string> invoice)
        {
            string filePath = filesPath + "invoice_" + invoice["Id"] + "_" + invoice["Client Company"] + ".docx";
            return filePath;
        }
    }
}