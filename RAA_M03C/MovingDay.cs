#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Forms = System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

#endregion

namespace RAA_M03C
{
    [Transaction(TransactionMode.Manual)]
    public class MovingDay : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;


            // Get excel file
            int counter = 0;
            string excelFile = "";

            Forms.OpenFileDialog ofd = new Forms.OpenFileDialog();
            ofd.Title = "Select Furniture Excel File";
            ofd.InitialDirectory = @"C:\";
            ofd.Filter = "Excel files (*.xlsx)|*.xlsx";

            if (ofd.ShowDialog() != Forms.DialogResult.OK)
                return Result.Failed;

            excelFile = ofd.FileName;

            //Manage data from excel
            List<string[]> excelFurnSetData = Utils.GetDataFromExcel(excelFile, "Furniture sets", 3);
            List<string[]> excelFurnData = Utils.GetDataFromExcel(excelFile, "Furniture types", 3);

            excelFurnSetData.RemoveAt(0);
            excelFurnData.RemoveAt(0);

            List<FurnSet> furnSetList = new List<FurnSet>();
            List<FurnData> furnDataList = new List<FurnData>();

            foreach (string[] curRow in excelFurnSetData)
            {
                FurnSet tmpFurnSet = new FurnSet(curRow[0].Trim(), curRow[1].Trim(), curRow[2].Trim());
                furnSetList.Add(tmpFurnSet);
            }

            foreach (string[] curRow in excelFurnData)
            {
                FurnData tmpFurnData = new FurnData(doc, curRow[0].Trim(), curRow[1].Trim(), curRow[2].Trim());
                furnDataList.Add(tmpFurnData);
            }

            //Get all rooms
            List<SpatialElement> roomList = Utils.GetAllRooms(doc);

            //Modify the model
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Insert Furniture");

                foreach (SpatialElement room in roomList)
                {
                    string curFurnSet = Utils.GetParamValue(room, "Furniture Set");

                    LocationPoint roomPt = room.Location as LocationPoint;
                    XYZ insPoint = roomPt.Point;

                    foreach (FurnSet tmpFurnSet in furnSetList)
                    {
                        if (tmpFurnSet.setType == curFurnSet)
                        {
                            foreach (string curFurn in tmpFurnSet.furnList)
                            {
                                string tmpFurn = curFurn.Trim();
                                FurnData fd = Utils.GetFamilyInfo(tmpFurn, furnDataList);

                                if (fd != null)
                                {
                                    fd.familySymbol.Activate();

                                    FamilyInstance newFamInst = doc.Create.NewFamilyInstance(insPoint, fd.familySymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                    counter++;
                                }
                            }
                        }
                        Utils.SetParamValueAsInt(room, "Furniture Count", tmpFurnSet.FurnitureCount());
                    }
                }
                t.Commit();
            }

            //Alert user
            TaskDialog.Show("Moving Day", $"{counter} families have been inserted.");

            return Result.Succeeded;
        }
       
    }
}