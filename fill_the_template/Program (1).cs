using docbuilder_net;
using OfficeFileTypes = docbuilder_net.FileTypes;
using CValue = docbuilder_net.CDocBuilderValue;
using CContext = docbuilder_net.CDocBuilderContext;
using CContextScope = docbuilder_net.CDocBuilderContextScope;
using System.Collections.Generic;
using static FillTheTemplate.Program;

namespace FillTheTemplate
{
    public class Program
    {
        public static void Main(string[] args)
        {
            string workDirectory = "C:/Program Files/ONLYOFFICE/DocumentBuilder";
            string listPath = "../../../invoices-list.xlsx";
            string templatePath = "../../../invoice-template.docx";
            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);
            // Replace with dinamic value
            int invoiceRowNumber = 2;
            FillTheTemplate(workDirectory, listPath, templatePath, invoiceRowNumber);
        }

        public static void FillTheTemplate(string workDirectory, string listPath, string templatePath, int invoiceRowNumber)
        {
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder oBuilder = new CDocBuilder();
            oBuilder.OpenFile(listPath, "xlsx");
            CContext oContext = oBuilder.GetContext();
            CContextScope oScope = oContext.CreateScope();
            CValue oGlobal = oContext.GetGlobal();
            CValue oApi = oGlobal["Api"];
            CValue oWorksheet = oApi.Call("GetActiveSheet");
            CValue oRange = oWorksheet.Call("GetRange", "A"+invoiceRowNumber+":N"+invoiceRowNumber);
            CValue invoiceData = oRange.Call("GetValue")[0];
            Dictionary<string, string> invoice = createInvoiceDict(invoiceData);
            oBuilder.CloseFile();

            oBuilder.OpenFile(templatePath, "docx");
            oContext = oBuilder.GetContext();
            oScope = oContext.CreateScope();
            oGlobal = oContext.GetGlobal();
            oApi = oGlobal["Api"];
            CValue oDocument = oApi.Call("GetDocument");
            FillInvoice(oDocument, invoice);
            var doctype = (int)OfficeFileTypes.Document.DOCX;
            string filePath = "../../../" + "invoice_" + invoice["Id"] + "_" + invoice["ClientCompany"] + ".docx";
            oBuilder.SaveFile(doctype, filePath);
            oBuilder.CloseFile();
            CDocBuilder.Destroy();
        }
        public static Dictionary<string, string> createInvoiceDict(CValue invoiceData)
        {
            Dictionary<string, string> invoice = new Dictionary<string, string>();
            invoice.Add("Id", invoiceData[0].ToString());
            invoice.Add("Date", invoiceData[1].ToString());
            invoice.Add("ContactName", invoiceData[2].ToString());
            invoice.Add("ClientCompany", invoiceData[3].ToString());
            invoice.Add("Address", invoiceData[4].ToString());
            invoice.Add("Phone", invoiceData[5].ToString());
            invoice.Add("Email", invoiceData[6].ToString());
            invoice.Add("Dept", invoiceData[7].ToString());
            invoice.Add("Description", invoiceData[8].ToString());
            invoice.Add("Qty", invoiceData[9].ToString());
            invoice.Add("UnitPrice", invoiceData[10].ToString());
            invoice.Add("Discount", invoiceData[11].ToString());
            invoice.Add("Tax", invoiceData[12].ToString());
            invoice.Add("Shipping", invoiceData[13].ToString());
            int totalPrice = int.Parse(invoice["Qty"]) * int.Parse(invoice["UnitPrice"]);
            int priceWithDiscount = totalPrice - int.Parse(invoice["Discount"]);
            int totalTax = priceWithDiscount / 100 * int.Parse(invoice["Tax"]);
            int balanceDue = priceWithDiscount + totalTax + int.Parse(invoice["Shipping"]);
            invoice.Add("TotalPrice", totalPrice.ToString());
            invoice.Add("PriceWithDiscount", priceWithDiscount.ToString());
            invoice.Add("TotalTax", totalTax.ToString());
            invoice.Add("BalanceDue", balanceDue.ToString());
            return invoice;
        }
        public static void FillInvoice(CValue oDocument, Dictionary<string, string> invoice)
        {
            CValue oAllTables = oDocument.Call("GetAllTables");
            FillCell(oAllTables, 0, 2, 4, invoice["Date"]);
            FillCell(oAllTables, 0, 3, 4, invoice["Id"]);

            FillCell(oAllTables, 1, 1, 0, invoice["ContactName"]);
            FillCell(oAllTables, 1, 2, 0, invoice["ClientCompany"]);
            FillCell(oAllTables, 1, 3, 0, invoice["Address"]);
            FillCell(oAllTables, 1, 4, 0, invoice["Phone"]);
            FillCell(oAllTables, 1, 5, 0, invoice["Email"]);

            FillCell(oAllTables, 1, 1, 2, invoice["Dept"]);
            FillCell(oAllTables, 1, 2, 2, invoice["ClientCompany"]);
            FillCell(oAllTables, 1, 3, 2, invoice["Address"]);
            FillCell(oAllTables, 1, 4, 2, invoice["Phone"]);

            FillCell(oAllTables, 2, 1, 0, invoice["Description"]);
            FillCell(oAllTables, 2, 1, 1, invoice["Qty"]);
            FillCell(oAllTables, 2, 1, 2, "$ " + invoice["UnitPrice"]);
            FillCell(oAllTables, 2, 1, 3, "$ " + invoice["TotalPrice"]);

            FillCell(oAllTables, 3, 0, 2, "$ " + invoice["TotalPrice"]);
            FillCell(oAllTables, 3, 1, 2, "$ " + invoice["Discount"]);
            FillCell(oAllTables, 3, 2, 2, "$ " + invoice["PriceWithDiscount"]);
            FillCell(oAllTables, 3, 3, 2, invoice["Tax"] + "%");
            FillCell(oAllTables, 3, 4, 2, "$ " + invoice["TotalTax"]);
            FillCell(oAllTables, 3, 5, 2, "$ " + invoice["Shipping"]);
            FillCell(oAllTables, 3, 6, 2, "$ " + invoice["BalanceDue"]);
            }
        public static void FillCell(CValue oAllTables, int table, int row, int cell, string value)
        {
            oAllTables[table].Call("GetRow", row).Call("GetCell", cell).Call("GetContent").Call("GetElement", 0).Call("AddText", value);
        }
    }
}