using static Helpers.Methods;

using static Tests.HelloWorld;
using static Tests.CommentErrors;
using static Tests.PassExternalData;
using static Tests.FillTheTemplate;
using static Tests.SpreadsheetToPresentation;

namespace Tests
{
    public class Tests
    {
        public static string filesPath = "../../../files/";
        public static string templatesPath = "../../../templates/";
        public static object[,] sampleTwoDimArray = {
                { "Id", "Product", "Price", "Available"},
                { 1001, "Item A", 12.2, true },
                { 1002, "Item B", 18.8, true },
                { 1003, "Item C", 70.1, false },
                { 1004, "Item D", 60.0, true },
                { 1005, "Item E", 32.6, true },
                { 1006, "Item F", 28.0, false },
                { 1007, "Item G", 11.1, false },
                { 1008, "Item H", 41.4, true }
            };
        
        public static void Main(string[] args)
        {
            SetEnvironment(); // Helpers function: add Docbuilder dlls in path

            // Hello World
            // CreateSimpleDocument("hello_world.docx");

            // Comment errors
            // CommentSpreadsheetErrors("spreadsheet_with_errors.xlsx");

            // External data to a spreadsheet
            // DataToXlsx("external_data.xlsx", sampleTwoDimArray);

            // Fill the template
            FillInvoice("invoices-list.xlsx", "invoice-template.docx");

            // Presentation from the spreadsheet
            // ReadSpreadsheetData("chart_data.xlsx", "presentation.pptx");
        }
    }
}
