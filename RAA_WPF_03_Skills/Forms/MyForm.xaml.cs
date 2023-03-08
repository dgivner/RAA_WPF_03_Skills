using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using FileDialog = Autodesk.Revit.UI.FileDialog;


namespace RAA_WPF_03_Skills
{
    /// <summary>
    /// Interaction logic for Window.xaml
    /// </summary>
    public partial class MyForm : System.Windows.Window
    {
        private List<string> allTitleBlocks;
        private List<View> allViews;
        private ObservableCollection<FamilySymbol> titleblockCollection { get; set; }
        //private ObservableCollection<View> allViewsCollection { get; set; }

        //private ObservableCollection<SheetNumberNameData> sheetData =
        //    new ObservableCollection<SheetNumberNameData>();
        //private ObservableCollection<string> sheetNumItems;
        //private ObservableCollection<string> sheetNameItems;
        private Document doc;
        //private List<View> viewsList;
        //private List<FamilySymbol> titleblockList;
        private IEnumerable<SheetData> sheetDataList;

        public MyForm(Document doc, List<View> views,List<string> titleBlockList)
        {
            InitializeComponent();

            //Populate the combobox with title blocks
            allTitleBlocks = GetAllTitleblocks(doc);
            //dataGrid.ItemsSource = allViewsCollection;
            titleBlockItem.ItemsSource = allTitleBlocks;
            allViews = GetAllViews(doc);
            viewItem.ItemsSource = allViews;
            

            SheetNumberNameData.doc = doc;

        }

        //public MyForm(Document doc, List<View> viewsList, List<FamilySymbol> titleblockList)
        //{
        //    this.doc = doc;
        //    this.viewsList = viewsList;
        //    this.titleblockList = titleblockList;
        //}

        private List<string> GetAllTitleblocks(Document doc)
        {
            List<string> returnList = new List<string>();

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.WhereElementIsElementType();
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);

            foreach (FamilySymbol curTB in collector)
            {
                returnList.Add(curTB.Name);
            }

            return returnList;
        }
        public static List<View> GetAllViews(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Views);

            List<View> m_views = new List<View>();
            foreach (View x in collector.ToElements())
            {
                m_views.Add(x);
            }

            return m_views;
        }
        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            Element tblock = GetTitleBlockByName(doc, "");

            // get form data and do something
            Transaction t2 = new Transaction(doc);
            t2.Start("Create sheets");

            foreach (SheetData currentSheetData in sheetDataList)
            {
                ViewSheet currentSheet = ViewSheet.Create(doc, tblock.Id);

                currentSheet.SheetNumber = currentSheetData.SheetNumber;
                currentSheet.Name = currentSheetData.SheetName;

                View currentView = GetViewByName(doc, currentSheet.Name);

                XYZ insPoint = new XYZ(1.5, 1, 0);
                //XYZ insPoint = GetSheetCenterPoint(currentSheet);

                Viewport currentVP = Viewport.Create(doc, currentSheet.Id, currentView.Id, insPoint);
            }

            t2.Commit();
            t2.Dispose();
            this.DialogResult = true;
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void btnRemove_Click(object sender, RoutedEventArgs e)
        {
            //try
            //{
            //    foreach (FamilySymbol curRow in titleblockCollection)
            //    {
            //        if (dataGrid.SelectedItem == curRow)
            //            titleblockCollection.Remove(curRow);
            //    }
            //}
            //catch (Exception)
            //{ }
            //this.Close();
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            //SheetNumberNameData newData = new SheetNumberNameData(sheetNumItems,sheetNameItems);
            //sheetData.Add(newData);
            //dataGrid.Items.Refresh();
            //ObservableCollection<string> titleblockCollection = new ObservableCollection<string>();
            //this.Close();
        }

        private void btnLoad_Click(object sender, RoutedEventArgs e)
        {
            

            string sheetPath = OpenFile();
            //var fileName = OpenFile();
            //Read Sheets excel file for data
            List<SheetData> sheetDataList = new List<SheetData>();

            string[] sheetArray = File.ReadAllLines(sheetPath);

            foreach (string sheetString in sheetArray)
            {
                string[] cellData = sheetString.Split(',');

                SheetData curSheetData = new SheetData();
                curSheetData.SheetNumber = cellData[0];
                curSheetData.SheetName = cellData[1];

                sheetDataList.Add(curSheetData);
            }
            sheetDataList.RemoveAt(0);
            //this.Close();

        }
        
        private static List<dSheets> SheetList()
        {
            string sheetsFilePath = OpenFile();

            List<dSheets> sheets = new List<dSheets>();
            string[] sheetsArray = File.ReadAllLines(sheetsFilePath);
            foreach (var sheetsRowString in sheetsArray)
            {
                string[] sheetsCellString = sheetsRowString.Split(',');
                var sheet = new dSheets
                {
                    Number = sheetsCellString[0],
                    Name = sheetsCellString[1]
                };

                sheets.Add(sheet);
            }

            return sheets;
        }
        private static string OpenFile()
        {
            OpenFileDialog selectFile = new OpenFileDialog();
            selectFile.InitialDirectory = "C:\\";
            selectFile.Filter = "Excel|*.xlsx";
            selectFile.Multiselect = false;

            string fileName = "";
            if ((bool)selectFile.ShowDialog())
            {
                fileName = selectFile.FileName;
            }

            return fileName;
        }
        private View GetViewByName(Document doc, string name)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Views);

            foreach (View currentView in collector)
            {
                if (currentView.Name == name)
                {
                    return currentView;
                }
            }

            return null;
        }
        internal Element GetTitleBlockByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
            foreach (Element currentTblock in collector)
            {
                if (currentTblock.Name == typeName)
                {
                    return currentTblock;
                }
            }

            return null;
        }
        //Place views on Center of Sheet
        private XYZ GetSheetCenterPoint(ViewSheet currentSheet)
        {
            // Get the middle point of the sheet (insertion point)
            BoundingBoxUV outline = currentSheet.Outline;
            double x = (outline.Max.U + outline.Min.U) / 2;
            double y = (outline.Max.V + outline.Min.V) / 2;

            XYZ returnPoint = new XYZ(x, y, 0);

            return returnPoint;
        }
    }

    public struct SheetData
    {
        public string SheetNumber;
        public string SheetName;
    }
    public class SheetNumberNameData
    {
        public static Document doc;

        public string SheetNumber { get; set; }

        public string SheetName { get; set; }

        public bool IsPlaceholder { get; set; }

        //public List<string> TitleblockType
        //{
        //    get => GetAllTitleblocks(doc);
        //}

        public string ViewToPlace { get; set; }

        public SheetNumberNameData(ObservableCollection<string> sheetNumItems, ObservableCollection<string> sheetNameItems)
        {
            SheetNumber = sheetNumItems.ToString();
            SheetName = sheetNameItems.ToString();
        }

        
    }
}
