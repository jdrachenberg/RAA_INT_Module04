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
    public class Command2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Prompt user to select RVT file
            string revitFile = "";

            Forms.OpenFileDialog ofd = new Forms.OpenFileDialog();
            ofd.Title = "Select Revit File";
            ofd.InitialDirectory = @"C:\";
            ofd.Filter = "Revit Files|*.rvt";
            ofd.RestoreDirectory = true;

            if (ofd.ShowDialog() != Forms.DialogResult.OK)
                return Result.Failed;

            revitFile = ofd.FileName;

            // COPY MODEL GROUP TYPES

            // Open selected file in background
            UIDocument closedUIDoc = uiapp.OpenAndActivateDocument(revitFile);
            Document closedDoc = closedUIDoc.Document;

            // Get model groups
            FilteredElementCollector modelGroupCollector = new FilteredElementCollector(closedDoc)
                .OfCategory(BuiltInCategory.OST_IOSModelGroups);

            //Get list of model group Ids
            List<ElementId> elemIdList = modelGroupCollector.Select(elem => elem.Id).ToList();
            int groupCounter = elemIdList.Count;

            // Copy model group types
            CopyPasteOptions options = new CopyPasteOptions();

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Copy model groups");
                ElementTransformUtils.CopyElements(closedDoc, elemIdList, doc, null, options);
                t.Commit();
            }

            // Close the documents that groups were copied from
            uiapp.OpenAndActivateDocument(doc.PathName);
            closedDoc.Close(false);



            // CREATE AREAS AND MODEL GROUPS FROM ROOMS

            // Get all links
            FilteredElementCollector linkCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkType));

            // Loop through links and get doc if loaded
            Document linkedDoc = null;
            RevitLinkInstance link = null;

            foreach (RevitLinkType rvtLink in linkCollector)
            {
                if (rvtLink.GetLinkedFileStatus() == LinkedFileStatus.Loaded)
                {
                    link = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_RvtLinks)
                        .OfClass(typeof(RevitLinkInstance))
                        .Where(x => x.GetTypeId() == rvtLink.Id).First() as RevitLinkInstance;

                    linkedDoc = link.GetLinkDocument();
                }
            }
            int spaceCounter = 0;

            // Get rooms from linked model document
            FilteredElementCollector collectorA = new FilteredElementCollector(linkedDoc)
                .OfCategory(BuiltInCategory.OST_Rooms);

            //Get level associated with active view
            Level curLevel = doc.ActiveView.GenLevel;


            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create spaces and place model groups");
                foreach (Room curRoom in collectorA)
                {
                    string roomName = curRoom.Name;
                    string roomNum = curRoom.Number;
                    string roomComments = curRoom.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString();

                    // Get room location point
                    LocationPoint roomPoint = curRoom.Location as LocationPoint;

                    // Create space and transfer properties
                    SpatialElement newSpace = doc.Create.NewSpace(curLevel, new UV(roomPoint.Point.X, roomPoint.Point.Y));
                    newSpace.Name = roomName;
                    newSpace.Number = roomNum;
                    newSpace.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(roomComments);

                    //Create model groups
                    PlaceGroup(doc, roomComments, new XYZ(roomPoint.Point.X, roomPoint.Point.Y, roomPoint.Point.Z));

                    spaceCounter++;
                }
                t.Commit();
            }

            // COPY WALLS AND GENERIC MODELS FROM LINKED MODEL

            // Create document variable
            Document openDoc = null;

            // Loop through open documents and look for match

            foreach (Document curDoc in uiapp.Application.Documents)
            {
                if (curDoc.PathName.Contains("Sample 03"))
                {
                    openDoc = curDoc;
                    break;
                }
            }

            ElementMulticategoryFilter filter = new ElementMulticategoryFilter(new List<BuiltInCategory> { BuiltInCategory.OST_Walls, BuiltInCategory.OST_GenericModel });

            // Get model groups
            FilteredElementCollector collectorC = new FilteredElementCollector(openDoc)
                .WhereElementIsNotElementType()
                .WherePasses(filter);

            //Get list of model group Ids
            List<ElementId> collectorCIdList = collectorC.Select(elem => elem.Id).ToList();
            int elemCounter = collectorCIdList.Count;

            // Copy model group types

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Copy walls and generic models");
                ElementTransformUtils.CopyElements(openDoc, collectorCIdList, doc, null, options);
                t.Commit();
            }

            // Task dialog that shows counters
            TaskDialog.Show("Created Elements Report", $"There are \n {spaceCounter} new spaces created, \n {groupCounter} new groups, and \n {elemCounter} walls and grids added.");

            


            return Result.Succeeded;
        }

        //Place a model group with the given name and the given location
        private void PlaceGroup(Document doc, string groupName, XYZ insPoint)
        {
            // Get group type by name
            GroupType curGroup = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_IOSModelGroups)
                .WhereElementIsElementType()
                .Where(r => r.Name == groupName)
                .Cast<GroupType>().First();

            // Insert group
            doc.Create.PlaceGroup(insPoint, curGroup);
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
