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
                { 1004, "Item D", 60.0, true },
                { 1005, "Item E", 32.6, true },
                { 1006, "Item F", 28.0, false },
                { 1007, "Item G", 11.1, false },
                { 1008, "Item H", 41.4, true }
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

            // Arrays can't be passed to DocBuilder API methods
            // Array can't be transformed into CValue directly
            // arrayToCValue function below should be used instead
            CValue oArray = arrayToCValue(data, oContext);
            oWorksheet.Call("GetRange", "A:Z").Call("SetValue", oArray);

            oBuilder.SaveFile(doctype, resultPath);
            oBuilder.CloseFile();
            CDocBuilder.Destroy();
        }
        public static CValue arrayToCValue(object[,] data, CContext oContext)
        {
            int rowsLen = data.GetLength(0);
            int colsLen = data.GetLength(1);
            CValue oArray = oContext.CreateArray(colsLen);

            for (int col = 0; col < colsLen; col++)
            {
                CValue oArrayRow = oContext.CreateArray(rowsLen);

                for (int row = 0; row < rowsLen; row++)
                {
                    oArrayRow[row] = data[row, col].ToString();
                }
                oArray[col] = oArrayRow;
            }
            return oArray;
        }
    }
}
