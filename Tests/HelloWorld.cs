using docbuilder_net;

using CValue = docbuilder_net.CDocBuilderValue;
using CContext = docbuilder_net.CDocBuilderContext;
using CContextScope = docbuilder_net.CDocBuilderContextScope;

using static Helpers.Methods;
using static Tests.Tests;

namespace Tests
{
    public class HelloWorld
    {
        public static void CreateSimpleDocument(string fileName, CDocBuilder oBuilder)
        {            
            // create file
            var doctype = "docx";
            CreateFile(oBuilder, doctype);
            CContext oContext = oBuilder.GetContext();
            CContextScope oScope = oContext.CreateScope();
            CValue oApi = GetApi(oContext);

            // edit file
            CValue oDocument = oApi.Call("GetDocument");
            CValue oParagraph = oApi.Call("CreateParagraph");
            CValue oContent = oContext.CreateArray(1);
            oParagraph.Call("SetSpacingAfter", 1000, false);
            oParagraph.Call("AddText", "Hello from .net!");
            oContent[0] = oParagraph;
            oDocument.Call("InsertContent", oContent);

            // save and close file
            SaveAndCloseFile(oBuilder, filesPath + fileName, doctype);
        }
    }

}