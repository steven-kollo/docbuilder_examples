using System;
using docbuilder_net;
using OfficeFileTypes = docbuilder_net.FileTypes;
using CValue = docbuilder_net.CDocBuilderValue;
using CContext = docbuilder_net.CDocBuilderContext;
using CContextScope = docbuilder_net.CDocBuilderContextScope;


namespace CommentErrors
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // replace filePath with your file path
            string filePath = "../../../data.xlsx";
            // replace range with your data range
            string dataRange = "A1:H12000";
            // replace sheetName with your sheet name
            string sheetName = "Sheet1";

            // add Docbuilder dlls in path
            string workDirectory = "C:/Program Files/ONLYOFFICE/DocumentBuilder";
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);
            
            CommentErrors(workDirectory, filePath, dataRange, sheetName);
        }
        public static void CommentErrors(string workDirectory, string filePath, string dataRange, string sheetName)
        {
            var doctype = (int)OfficeFileTypes.Spreadsheet.XLSX;

            CDocBuilder.Initialize(workDirectory);
            CDocBuilder oBuilder = new CDocBuilder();
            oBuilder.OpenFile(filePath, "xlsx");

            CContext oContext = oBuilder.GetContext();
            CContextScope oScope = oContext.CreateScope();

            CValue oGlobal = oContext.GetGlobal();
            CValue oApi = oGlobal["Api"];
            CValue oWorksheet = oApi.Call("GetSheet", sheetName);
            CValue oRange = oWorksheet.Call("GetRange", dataRange);
            var data = oRange.Call("GetValue");

            for (int row = 0; row < data.GetLength(); row++)
            {
                for (int col = 0; col < data[0].GetLength(); col++)
                {
                    var cell = data[row][col].ToString();
                    if (cell.Contains("#")) 
                    {
                        string comment = "Error" + cell;
                        CValue errorCell = oWorksheet.Call("GetRangeByNumber", row, col);
                        errorCell.Call("AddComment", comment);
                    }        
                }
            }
            oBuilder.SaveFile(doctype, filePath);
            oBuilder.CloseFile();
            CDocBuilder.Destroy();
        }
    }
    
}