using docbuilder_net;

using CValue = docbuilder_net.CDocBuilderValue;
using CContext = docbuilder_net.CDocBuilderContext;

using static Helpers.Methods;
using static Tests.Tests;

namespace Tests
{
    public class SpreadsheetToPresentation
    {
        public static void SpreadsheetDataToPresentation(string spreadsheetName, string presentationName)
        {
            CDocBuilder oBuilder = InitDocBuilder();

            // create file
            var doctype = "xlsx";
            OpenFile(oBuilder, doctype, templatesPath + spreadsheetName);
            CContext oContext = GetFileContext(oBuilder);
            CValue oApi = GetApi(oContext);
            CValue oWorksheet = oApi.Call("GetActiveSheet");
            CValue oRange = oWorksheet.Call("GetUsedRange").Call("GetValue");
    
            object[,] array = oRangeTo2dArray(oRange, oContext);
            oBuilder.CloseFile();

            doctype = "pptx";
            CreateFile(oBuilder, doctype);
            oContext = GetFileContext(oBuilder);
            oApi = GetApi(oContext);
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
            oSlide.Call("AddObject", oChart);

            SaveAndCloseFile(oBuilder, filesPath + presentationName, doctype);
            CDocBuilder.Destroy();
        }
        public static CValue colsFromArray(object[,] array, CContext oContext)
        {
            int colsLen = array.GetLength(1) - 1;
            CValue cols = oContext.CreateArray(colsLen);
            for (int col = 1; col <= colsLen; col++)
            {
                cols[col - 1] = array[0, col].ToString();
            }
            return cols;
        }
        public static CValue rowsFromArray(object[,] array, CContext oContext)
        {
            int rowsLen = array.GetLength(0) - 1;
            CValue rows = oContext.CreateArray(rowsLen);
            for (int row = 1; row <= rowsLen; row++)
            {
                rows[row - 1] = array[row, 0].ToString();
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
    }
}
