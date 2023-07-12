using System;
using System.IO;
using docbuilder_net;

using OfficeFileTypes = docbuilder_net.FileTypes;
using CValue = docbuilder_net.CDocBuilderValue;
using CContext = docbuilder_net.CDocBuilderContext;
using CContextScope = docbuilder_net.CDocBuilderContextScope;
using static System.Net.Mime.MediaTypeNames;

namespace Test
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
            CValue oRange = oWorksheet.Call("GetUsedRange");

            // Spreadsheet values to CValue type arrays (only works with xlsx ranges)
            CValue cols = oWorksheet.Call("GetRange", "B1:D1").Call("GetValue");
            CValue rows = oWorksheet.Call("GetRange", "A2:A8").Call("GetValue");
            CValue data = oWorksheet.Call("GetRange", "B2:D8").Call("GetValue");

            // Spreadsheet values to arrays int[,] and string[]
            var values = oRange.Call("GetValue");
            var spreadsheetData = ReadValues(values);

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

            // 1) Can't pass int[] or string[], the Call method requires only types convertable to CValue 
            CValue oChart = oApi.Call("CreateChart", "bar3D", spreadsheetData.Data, spreadsheetData.Measures, spreadsheetData.Facts);

            // 2) Can't pass more than 6 params, the Call method takes 6 params max 
            CValue oChartCValue = oApi.Call("CreateChart", "bar3D", data, cols, rows, 7200000, 3600000, 24);

            // 2) Can run CreateChart method without additional params and set them later.
            // It will work with .docbuilder script, but not with the lib
            CValue oChartCValueValuesOnly = oApi.Call("CreateChart", "bar3D", data, cols, rows);
            oChartCValueValuesOnly.Call("SetSize", 300 * 36000, 130 * 36000);

            // Example of .docbuilder script passing additional params after creating a chart:

            // var oChart = Api.CreateChart("bar3D", [
            //  [200, 240, 280],
            //  [250, 260, 280]
            // ], ["Projected Revenue", "Estimated Costs"], [2014, 2015, 2016]);
            // oChart.SetSize(300 * 36000, 130 * 36000);
            // oSlide.AddObject(oChart);

            oSlide.Call("AddObject", oChartCValueValuesOnly);

            oBuilder.SaveFile(doctype, presentationPath);
            oBuilder.CloseFile();
            CDocBuilder.Destroy();
        }
        public static SpreadsheetData ReadValues(CValue values)
        {
            int[,] data = new int[values[0].GetLength(), values.GetLength()];
            for (int i = 1; i < values.GetLength(); i++)
                for (int j = 1; j < values[0].GetLength(); j++)
                    data[j - 1, i - 1] = int.Parse(values[i][j].ToString());

            string[] measures = new string[values[0].GetLength()];
            for (int i = 0; i < values[0].GetLength(); i++)
                measures[i] = values[0][i].ToString();

            string[] facts = new string[values.GetLength()];
            for (int i = 1; i < values.GetLength(); i++)
                facts[i - 1] = values[i][0].ToString();

            return new SpreadsheetData
            {
                Data = data,
                Measures = measures,
                Facts = facts
            };
        }
        public class SpreadsheetData
        {
            public int[,] Data;
            public string[] Measures;
            public string[] Facts;
        }
    }
}
