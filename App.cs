#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.Versioning;
using System.Windows.Markup;

#endregion

namespace RAA_INT_Module04
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication app)
        {
            //// 1. Create ribbon tab
            //try
            //{
            //    app.CreateRibbonTab("My First Revit Add-in");
            //}
            //catch (Exception)
            //{
            //    Debug.Print("Tab already exists.");
            //}

            //// 2. Create ribbon panel 
            //RibbonPanel panel = Utils.CreateRibbonPanel(app, "My First Revit Add-in", "Revit Tools");

            //// 3. Create button data instances
            //ButtonDataClass myButtonData1 = new ButtonDataClass("btnArchSmarter_Addin_Template_2023_1", "Button 1", Command1.GetMethod(), Properties.Resources.Blue_32, Properties.Resources.Blue_16, "Tooltip 1");
            //ButtonDataClass myButtonData2 = new ButtonDataClass("btnArchSmarter_Addin_Template_2023_2", "Button 2", Command2.GetMethod(), Properties.Resources.Red_32, Properties.Resources.Red_16, "Tooltip 2");

            //// 4. Create buttons
            //PushButton myButton1 = panel.AddItem(myButtonData1.Data) as PushButton;
            //PushButton myButton2 = panel.AddItem(myButtonData2.Data) as PushButton;

            // NOTE:
            // To create a new tool, copy lines 35 and 39 and rename the variables to "myButtonData3" and "myButton3". 
            // Change the name of the tool in the arguments of line 

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }


    }
}
