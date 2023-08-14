using docbuilder_net;

using OfficeFileTypes = docbuilder_net.FileTypes;
using CValue = docbuilder_net.CDocBuilderValue;
using CContext = docbuilder_net.CDocBuilderContext;


namespace Helpers
{
    public class Methods
    {
        public static string workDirectory = "C:/Program Files/ONLYOFFICE/DocumentBuilder";
        public static Dictionary<string, int> doctypes;
        
        static Methods()
        {
            doctypes = new Dictionary<string, int>();
            doctypes.Add("docx", (int)OfficeFileTypes.Document.DOCX);
            doctypes.Add("xlsx", (int)OfficeFileTypes.Spreadsheet.XLSX);
            doctypes.Add("pptx", (int)OfficeFileTypes.Presentation.PPTX);
        }
        public static void CreateFile(CDocBuilder oBuilder, string doctype)
        {
            oBuilder.CreateFile(doctypes[doctype]);
        }
        public static void OpenFile(CDocBuilder oBuilder, string doctype, string filePath)
        {
            oBuilder.OpenFile(filePath, "xlsx");
        }
        public static CValue GetApi(CContext oContext)
        {
            CValue oGlobal = oContext.GetGlobal();
            return oGlobal["Api"];
        }
        public static void SaveAndCloseFile(CDocBuilder oBuilder, string resultPath, string doctype)
        {
            oBuilder.SaveFile(doctypes[doctype], resultPath);
            oBuilder.CloseFile();
        }
        public static CValue TwoDimArrayToCValue(object[,] data, CContext oContext)
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
