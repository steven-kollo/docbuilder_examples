using docbuilder_net;

using CValue = docbuilder_net.CDocBuilderValue;
using CContext = docbuilder_net.CDocBuilderContext;
using CContextScope = docbuilder_net.CDocBuilderContextScope;

using static Helpers.Methods;
using static Tests.Tests;

namespace Tests
{
    public class PassExternalData
    {
        
        public static void DataToXlsx(string fileName, object[,] data, CDocBuilder oBuilder)
        {
            // create file
            var doctype = "xlsx";
            CreateFile(oBuilder, doctype);
            CContext oContext = oBuilder.GetContext();
            CContextScope oScope = oContext.CreateScope();
            CValue oApi = GetApi(oContext);
            CValue oWorksheet = oApi.Call("GetActiveSheet");
            
            // pass data
            CValue oArray = TwoDimArrayToCValue(data, oContext);
            oWorksheet.Call("GetRange", "A:Z").Call("SetValue", oArray);

            SaveAndCloseFile(oBuilder, filesPath + fileName, doctype);
        }
    }
}