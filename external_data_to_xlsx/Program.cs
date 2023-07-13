using docbuilder_net;

using OfficeFileTypes = docbuilder_net.FileTypes;
using CValue = docbuilder_net.CDocBuilderValue;
using CContext = docbuilder_net.CDocBuilderContext;
using CContextScope = docbuilder_net.CDocBuilderContextScope;

namespace ExternalDataToXlsx
{
    public class Program
    {
        public static void Main(string[] args)
        {
            string workDirectory = "C:/Program Files/ONLYOFFICE/DocumentBuilder";
            string resultPath = "../../../data.xlsx";

            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            // Replace sample data with your data stored in 2d array of objects convertable to CValue 
            // (string, int, float, bool etc.)
            object[,] data = {
                { "Id", "Product", "Price", "Available"},
                { 1001, "Item A", 12.2, true }, 
                { 1002, "Item B", 18.8, true }, 
                { 1003, "Item C", 70.1, false }, 
                { 1004, "Item D", 60.0, true } 
            };

            DataToXlsx(workDirectory, resultPath, data);
        }
        public static void DataToXlsx (string workDirectory, string resultPath, object[,] data)
        {
            var doctype = (int)OfficeFileTypes.Spreadsheet.XLSX;
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder oBuilder = new CDocBuilder();
            oBuilder.CreateFile(doctype);

            CContext oContext = oBuilder.GetContext();
            CContextScope oScope = oContext.CreateScope();
            CValue oGlobal = oContext.GetGlobal();
            CValue oApi = oGlobal["Api"];
            CValue oWorksheet = oApi.Call("GetActiveSheet");

            // Passing the whole array works for .docbuilder scripts, but won't work with the lib
            // 2d array ("object[,] data") can't be transformed into CValue
            // oWorksheet.Call("GetRange", "A1:D5").Call("SetValue", data);

            // get 2d array height with Array.GetLength(0)
            for (int row = 0; row < data.GetLength(0); row++)
            {
                // get 2d array width with Array.GetLength(1)
                for (int col = 0; col < data.GetLength(1); col++)
                {
                    // xlsx row and col count starts from 1, not 0
                    oWorksheet.Call("GetCells", row + 1, col + 1).Call("SetValue", data[row, col].ToString());
                }
            }

            oBuilder.SaveFile(doctype, resultPath);
            oBuilder.CloseFile();
            CDocBuilder.Destroy();
        }
    }
}