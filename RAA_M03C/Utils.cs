using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
//

using System.Linq;
using System.Text;
using System.IO;
using Forms = System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using View = Autodesk.Revit.DB.View;

namespace RAA_M03C
{
    internal static class Utils
    {

        public static List<string[]> GetDataFromExcel(string excelFile, string wsName, int numColumns)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);

            Excel.Worksheet excelWs = GetExcelWorksheetByName(excelWb, wsName);
            Excel.Range excelRng = excelWs.UsedRange as Excel.Range;

            int rowCount = excelRng.Rows.Count;

            List<string[]> data = new List<string[]>();

            for (int i = 1; i <= rowCount; i++)
            {
                string[] rowData = new string[numColumns];

                for (int j = 1; j <= numColumns; j++)
                {
                    Excel.Range cellData = excelWs.Cells[i, j];
                    rowData[j - 1] = cellData.Value.ToString();
                }

                data.Add(rowData);
            }

            excelWb.Close();
            excelApp.Quit();

            return data;
        }

        public static Excel.Worksheet GetExcelWorksheetByName(Excel.Workbook excelWb, string wsName)
        {
            foreach (Excel.Worksheet sheet in excelWb.Worksheets)
            {
                if (sheet.Name == wsName)
                    return sheet;
            }

            return null;
        }

        public static List<SpatialElement> GetAllRooms(Document doc)
        {
            List<SpatialElement> returnList = new List<SpatialElement>();

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Rooms);
            collector.WhereElementIsNotElementType();

            foreach (Element curElem in collector)
            {
                SpatialElement curRoom = curElem as SpatialElement;
                returnList.Add(curRoom);
            }

            return returnList;
        }

        public static string GetParamValue(Element curElem, string paramName)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                Debug.Print(curParam.Definition.Name);
                if (curParam.Definition.Name == paramName)
                {
                    Debug.Print(curParam.AsString());
                    return curParam.AsString();
                }

            }

            return null;
        }

        public static void SetParamValueAsInt(Element curElem, string paramName, int paramValue)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                {
                    curParam.Set(paramValue);
                }
            }
        }

        public static FurnData GetFamilyInfo(string furnName, List<FurnData> furnDataList)
        {
            foreach (FurnData furn in furnDataList)
                if (furn.furnName == furnName)
                    return furn;

            return null;
        }

    }

}
