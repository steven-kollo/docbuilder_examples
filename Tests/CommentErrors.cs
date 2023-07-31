using docbuilder_net;

using CValue = docbuilder_net.CDocBuilderValue;
using CContext = docbuilder_net.CDocBuilderContext;

using static Helpers.Methods;
using static Tests.Tests;

namespace Tests
{
    public class CommentErrors
    {
        public static void CommentSpreadsheetErrors(string templateName)
        {
            // init builder
            CDocBuilder oBuilder = InitDocBuilder();

            // open file
            var doctype = "xlsx";
            OpenFile(oBuilder, doctype, templatesPath + templateName);
            CContext oContext = GetFileContext(oBuilder);
            CValue oApi = GetApi(oContext);

            // comment errors
            CValue oWorksheet = oApi.Call("GetActiveSheet");
            CValue oRange = oWorksheet.Call("GetUsedRange");
            var data = oRange.Call("GetValue");

            for (int row = 0; row < data.GetLength(); row++)
            {
                for (int col = 0; col < data[0].GetLength(); col++)
                {
                    CheckCell(oWorksheet, data[row][col].ToString(), row, col);
                }
            }

            // save and close file
            SaveAndCloseFile(oBuilder, filesPath + templateName, doctype);
            CDocBuilder.Destroy();
        }
        public static void CheckCell(CValue oWorksheet, string cell, int row, int col)
        {
            if (cell.Contains("#"))
            {
                string comment = "Error" + cell;
                CValue errorCell = oWorksheet.Call("GetRangeByNumber", row, col);
                errorCell.Call("AddComment", comment);
            }
        }
    }
}
