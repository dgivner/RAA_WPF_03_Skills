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
        private static Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
        private static  Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Open("FilePath");
        private static Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Sheets[1];
        private SheetNumberNameData sheetData;
        private Sheets sheets;
        private List<SheetNumberNameData> sheetNumberNameData;

        ObservableCollection<SheetNumberNameData> dataList { get; set; }
        ObservableCollection<string> sheetNumItems { get; set; } 
        ObservableCollection<string> sheetNameItems { get; set; }

        
        public MyForm(Sheets sheets, SheetNumberNameData sheetNumberNameData)
        {
            InitializeComponent();

            List<dSheets> dSheetsList = new List<dSheets>();
            List<SheetNumberNameData> sheetDataList = new List<SheetNumberNameData>();
            foreach (dSheets sheet in dSheetsList)
            {
                SheetNumberNameData sheetData = new SheetNumberNameData(sheetNumItems,sheetNameItems)
                {
                    SheetNumber = sheet.Number,
                    SheetName = sheet.Name
                };
            }
            sheetDataList.Add(sheetData);
            
            dataList = new ObservableCollection<SheetNumberNameData>();
            dataList.Add(new SheetNumberNameData(sheetNumItems,sheetNameItems));

            sheetNumItems = new ObservableCollection<string>();
            sheetNameItems = new ObservableCollection<string>();

            dataGrid.ItemsSource = SheetList();

            //titleBlockItem.ItemsSource = ;
            //viewItem.ItemsSource = ;

        }

        public MyForm(Sheets sheets, List<SheetNumberNameData> sheetNumberNameData)
        {
            this.sheets = sheets;
            this.sheetNumberNameData = sheetNumberNameData;
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
            try
            {
                foreach (SheetNumberNameData curRow in dataList)
                {
                    if (dataGrid.SelectedItem == curRow)
                        dataList.Remove(curRow);
                }
            }
            catch (Exception)
            { }
            //this.Close();
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            dataList.Add(new SheetNumberNameData(sheetNumItems,sheetNameItems));
            //this.Close();
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
        public string SheetNumber { get; set; }
        
        public string SheetName { get;set; }

        public bool IsPlaceholder { get; set; }

        public string TitleblockType { get; set; }

        public string ViewToPlace { get; set; }

        public SheetNumberNameData(ObservableCollection<string> sheetNumItems, ObservableCollection<string> sheetNameItems)
        {
            SheetNumber = sheetNumItems.ToString();
            SheetName = sheetNameItems.ToString();
        }
    }
}
