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
        private ObservableCollection<FamilySymbol> titleblockCollection { get; set; }
        private ObservableCollection<View> allViewsCollection { get; set; }


        public MyForm(List<FamilySymbol> titleblockList,List<View> viewsList)
        {
            InitializeComponent();

            titleblockCollection = new ObservableCollection<FamilySymbol>(titleblockList);
            allViewsCollection = new ObservableCollection<View>(viewsList);

        }


        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
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
            this.Close();
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnLoad_Click(object sender, RoutedEventArgs e)
        {
            //var fileName = OpenFile();
            this.Close();
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
    }
    
    
    public class SheetNumberNameData
    {
        private static Document doc;

        public string SheetNumber { get; set; }
        
        public string SheetName { get;set; }

        public bool IsPlaceholder { get; set; }

        public string TitleblockType { get => TitleblockType;
            set => GetAllTitleblocks(doc);
        }

        public string ViewToPlace { get; set; }

        public SheetNumberNameData(ObservableCollection<string> sheetNumItems, ObservableCollection<string> sheetNameItems)
        {
            SheetNumber = sheetNumItems.ToString();
            SheetName = sheetNameItems.ToString();
        }
        
        public static List<string> GetAllTitleblocks(Document doc)
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
    }
}
