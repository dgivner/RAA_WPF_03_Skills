#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Documents;
using Autodesk.Revit.Exceptions;
using Microsoft.Win32;
using FileDialog = Autodesk.Revit.UI.FileDialog;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Microsoft.Office.Interop.Excel;

#endregion

namespace RAA_WPF_03_Skills
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        private List<string> titleblockList;
        private List<View> viewsList;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            FilteredElementCollector tblockCollector = new FilteredElementCollector(doc);
            tblockCollector.OfCategory(BuiltInCategory.OST_TitleBlocks);
            ElementId tblockId = tblockCollector.FirstElementId();

            

            //List<SheetNumberNameData> sheetNumberNameData = new List<SheetNumberNameData>();

            

            //FilteredElementCollector viewTemplateCollector = new FilteredElementCollector(doc);
            //viewTemplateCollector.OfClass(typeof(ViewFamilyType));

            FilteredElementCollector vftCollector = new FilteredElementCollector(doc);
            vftCollector.OfClass(typeof(ViewFamilyType));

            ViewFamilyType planVFT = null;
            ViewFamilyType rcpVFT = null;
            ViewFamilyType areaVFT = null;


            foreach (ViewFamilyType vft in vftCollector)
            {
                if (vft.ViewFamily == ViewFamily.FloorPlan) planVFT = vft;

                if (vft.ViewFamily == ViewFamily.CeilingPlan) rcpVFT = vft;

                if (vft.ViewFamily == ViewFamily.AreaPlan) areaVFT = vft;
            }

            // open form
            MyForm currentForm = new MyForm(doc, viewsList, titleblockList)
            {
                Width = 800,
                Height = 450,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
            };

            currentForm.ShowDialog();

            return Result.Succeeded;
        }
        // put any code needed for the form here
        


        

        //Get Levels in Revit model
        private List<Level> GetLevels(Document doc)
        {
            FilteredElementCollector levelCollector = new FilteredElementCollector(doc);
            ICollection<Element> levels = levelCollector.OfClass(typeof(Level)).ToElements();


            return (List<Level>) levels;
        }
        public static List<View> GetAllViews(Document curDoc)
        {
            FilteredElementCollector allViews = new FilteredElementCollector(curDoc);
            allViews.OfCategory(BuiltInCategory.OST_Views);

            List<View> multiViews = new List<View>();
            foreach (View av in allViews.ToElements())
            {
                multiViews.Add(av);
            }

            return multiViews;
        }
        private ViewFamilyType GetViewFamilyTypeByName(Document doc, string vftName, ViewFamily vf)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewFamilyType));

            foreach (ViewFamilyType currentVFT in collector)
            {
                if (currentVFT.Name == vftName && currentVFT.ViewFamily == vf)
                {
                    return currentVFT;
                }
            }

            return null;
        }
        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
      
    }
}
