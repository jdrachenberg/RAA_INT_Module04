#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using Forms = System.Windows.Forms;

#endregion

namespace RAA_INT_Module04
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // 1. Get all links
            FilteredElementCollector linkCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkType));

            // 2. Loop through links and get doc if loaded
            Document linkedDoc = null;
            RevitLinkInstance link = null;

            //foreach (RevitLinkType rvtLink in linkCollector)
            //{
            //    if (rvtLink.GetLinkedFileStatus() == LinkedFileStatus.Loaded)
            //    {
            //        link = new FilteredElementCollector(doc)
            //            .OfCategory(BuiltInCategory.OST_RvtLinks)
            //            .OfClass(typeof(RevitLinkInstance))
            //            .Where(x => x.GetTypeId() == rvtLink.Id).First() as RevitLinkInstance;

            //        linkedDoc = link.GetLinkDocument();
            //    }
            //}

            //FilteredElementCollector collectorA = new FilteredElementCollector(linkedDoc)
            //    .OfCategory(BuiltInCategory.OST_Rooms);

            //TaskDialog.Show("Test Method A", $"There are {collectorA.Count()} rooms in the linked model.");

            //1. Prompt user to select RVT file
            // NOTE: be sure to add reference to System.Windows.Forms
            //string revitFile = "";

            //Forms.OpenFileDialog ofd = new Forms.OpenFileDialog();
            //ofd.Title = "Select Revit File";
            //ofd.InitialDirectory = @"C:\";
            //ofd.Filter = "Revit Files (*.rvt)|*.rvt";
            //ofd.RestoreDirectory = true;

            //if (ofd.ShowDialog() != Forms.DialogResult.OK)
            //    return Result.Failed;

            //revitFile = ofd.FileName;

            //// 2. Open selected file in background
            //UIDocument closedUIDoc = uiapp.OpenAndActivateDocument(revitFile);
            //Document closedDoc = closedUIDoc.Document;

            //FilteredElementCollector collectorB = new FilteredElementCollector(closedDoc)
            //    .OfCategory(BuiltInCategory.OST_Rooms);
            
            //TaskDialog.Show("Test Method B", $"There are {collectorB.Count()} rooms in the model I justed opened.");
            
            // Make other document active then close document
            //uiapp.OpenAndActivateDocument(doc.PathName);
            //closedDoc.Close(false);

            // Method C: get open file with specific name
            // 1. Create document variable
            Document openDoc = null;

            // 2. Loop through open documents and look for match

            foreach (Document curDoc in uiapp.Application.Documents)
            {
                if (curDoc.PathName.Contains("rac"))
                {
                    openDoc = curDoc;
                    break;
                }
            }

            // Create space from room
            // Use LINQ to get a room
            Room curRoom = new FilteredElementCollector(openDoc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .First();

            // Get level from current view
            Level curLevel = doc.ActiveView.GenLevel;

            //Get room data
            string roomName = curRoom.Name;
            string roomNum = curRoom.Number;
            string roomComents = curRoom.LookupParameter("Comments").AsString();

            // Get room location point
            LocationPoint roomPoint = curRoom.Location as LocationPoint;
            
            using(Transaction t = new Transaction(doc))
            {
                t.Start("Create space");
                // Create space and transfer properties
                SpatialElement newSpace = doc.Create.NewSpace(curLevel, new UV(roomPoint.Point.X, roomPoint.Point.Y));
                newSpace.Name = roomName;
                newSpace.Number = roomNum;
                newSpace.LookupParameter("Comments").Set(roomComents);
                t.Commit();
            }

            // Inserting groups
            // Get group type by name
            string groupName = "Group 1";

            GroupType curGroup = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_IOSModelGroups)
                .WhereElementIsElementType()
                .Where(r => r.Name == groupName)
                .Cast<GroupType>().First();

            // Insert group
            XYZ insPoint = new XYZ();

            using(Transaction t = new Transaction(doc))
            {
                t.Start("Insert Group");
                doc.Create.PlaceGroup(insPoint, curGroup);
                t.Commit();
            }

            // Copy elements from other doc
            FilteredElementCollector wallCollector = new FilteredElementCollector(openDoc)
                .OfCategory(BuiltInCategory.OST_Walls)
                .WhereElementIsNotElementType();

            // Get list of element IDs
            List<ElementId> elemIdList = wallCollector.Select(elem => elem.Id).ToList();

            // Copy elements
            CopyPasteOptions options = new CopyPasteOptions();

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Copy elements");
                ElementTransformUtils.CopyElements(openDoc, elemIdList, doc, null, options);
                t.Commit();
            }

            return Result.Succeeded;
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
