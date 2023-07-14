using docbuilder_net;

using OfficeFileTypes = docbuilder_net.FileTypes;
using CValue = docbuilder_net.CDocBuilderValue;
using CContext = docbuilder_net.CDocBuilderContext;
using CContextScope = docbuilder_net.CDocBuilderContextScope;

namespace SpreadsheetToPresentation
{
    public class Program
    {
        public static void Main(string[] args)
        {
            string workDirectory = "C:/Program Files/ONLYOFFICE/DocumentBuilder";
            string spreadsheetPath = "../../../data.xlsx";
            string presentationPath = "../../../presentation.pptx";
            string sheetName = "Sheet1";
            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);   
            ReadSpreadsheetData(workDirectory, spreadsheetPath, presentationPath, sheetName);
        }   
       
        public static void ReadSpreadsheetData(string workDirectory, string spreadsheetPath, string presentationPath, string sheetName)
        {
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder oBuilder = new CDocBuilder();
            oBuilder.OpenFile(spreadsheetPath, "xlsx");
            CContext oContext = oBuilder.GetContext();
            CContextScope oScope = oContext.CreateScope();
            CValue oGlobal = oContext.GetGlobal();
            CValue oApi = oGlobal["Api"];
            CValue oWorksheet = oApi.Call("GetActiveSheet");
            CValue oRange = oWorksheet.Call("GetUsedRange").Call("GetValue");
            // Can't pass CValues between files, contexts is destroyed with oBuilder.CloseFile()
            // Pass 2d array to transform it to CValue inside the presentation context
            object[,] array = oRangeTo2dArray(oRange, oContext);
            oBuilder.CloseFile();

            var doctype = (int)OfficeFileTypes.Presentation.PPTX;
            oBuilder.CreateFile(doctype);
            oContext = oBuilder.GetContext();
            oScope = oContext.CreateScope();
            oGlobal = oContext.GetGlobal();
            oApi = oGlobal["Api"];
            CValue oPresentation = oApi.Call("GetPresentation");
            CValue oSlide = oPresentation.Call("GetSlideByIndex", 0);
            oSlide.Call("RemoveAllObjects");

            // Transform 2d array into cols names, rows names and data
            CValue array_cols = colsFromArray(array, oContext);
            CValue array_rows = rowsFromArray(array, oContext);
            CValue array_data = dataFromArray(array, oContext);
            // Pass CValue data to the CreateChart method
            CValue oChart = oApi.Call("CreateChart", "lineStacked", array_data, array_cols, array_rows);
            oChart.Call("SetSize", 180 * 36000, 100 * 36000);
            oChart.Call("SetPosition", 20 * 36000, 50 * 36000);
            oChart.Call("SetVertAxisLabelsFontSize", 16);
            oChart.Call("SetHorAxisLabelsFontSize", 16);
            oChart.Call("SetLegendFontSize", 16);

            // LegendPos and ChartStyle are not working with the lib
            oChart.Call("SetLegendPos", "Top");
            oChart.Call("ApplyChartStyle", 24);

            oSlide.Call("AddObject", oChart);
            oBuilder.SaveFile(doctype, presentationPath);
            oBuilder.CloseFile();
            CDocBuilder.Destroy();
        }
        public static CValue colsFromArray(object[,] array, CContext oContext)
        {
            int colsLen = array.GetLength(1) - 1;
            CValue cols = oContext.CreateArray(colsLen);
            for(int col = 1; col <= colsLen; col++)
            {
                cols[col-1] = array[0, col].ToString();
            }
            return cols;
        }
        public static CValue rowsFromArray(object[,] array, CContext oContext)
        {
            int rowsLen = array.GetLength(0) - 1;
            CValue rows = oContext.CreateArray(rowsLen);
            for (int row = 1; row <= rowsLen; row++)
            {
                rows[row - 1] = array[row,0].ToString();
            }
            return rows;
        }

        public static CValue dataFromArray(object[,] array, CContext oContext)
        {
            int colsLen = array.GetLength(0) - 1;
            int rowsLen = array.GetLength(1) - 1;
            CValue data = oContext.CreateArray(rowsLen);
            for (int row = 1; row <= rowsLen; row++)
            {
                CValue row_data = oContext.CreateArray(colsLen);
                for (int col = 1; col <= colsLen; col++)
                {
                    row_data[col - 1] = array[col, row].ToString();
                }
                data[row - 1] = row_data;
            }
            return data;
        }
        public static object[,] oRangeTo2dArray(CValue oRange, CContext oContext)
        {
            
            int rowsLen = (int)oRange.GetLength();
            int colsLen = (int)oRange[0].GetLength();
            object[,] oArray = new object[rowsLen, colsLen];

            for (int col = 0; col < colsLen; col++)
            {
                CValue oArrayRow = oContext.CreateArray(rowsLen);

                for (int row = 0; row < rowsLen; row++)
                {
                    oArray[row, col] = oRange[row][col].ToString();
                }
            }
            return oArray;
            
        }
    }
}
