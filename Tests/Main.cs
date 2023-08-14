using docbuilder_net;

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
            // Set Environment
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            RunTests();
        }

        public static void RunTests()
        {
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder oBuilder = new CDocBuilder();

            // Hello World
            CreateSimpleDocument("hello_world.docx", oBuilder);

            // Comment errors
            CommentSpreadsheetErrors("spreadsheet_with_errors.xlsx", oBuilder);

            // External data to a spreadsheet
            DataToXlsx("external_data.xlsx", sampleTwoDimArray, oBuilder);

            // Fill the template
            FillInvoice("invoices-list.xlsx", "invoice-template.docx", oBuilder);

            // Presentation from the spreadsheet
            SpreadsheetDataToPresentation("chart_data.xlsx", "presentation.pptx", oBuilder);

            CDocBuilder.Destroy();
        }
    }
}
