using System;
using System.Collections.Generic;
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
using NewEventLogDLL;
using DataValidationDLL;
using KeyWordDLL;
using InspectionsDLL;
using VehicleMainDLL;
using Excel =  Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace FindVehicleMileage
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        DataValidationClass TheDataValidationClass = new DataValidationClass();
        KeyWordClass TheKeyWordClass = new KeyWordClass();
        InspectionsClass TheInspectionsClass = new InspectionsClass();
        VehicleMainClass TheVehicleMainClass = new VehicleMainClass();

        FindVehicleMainByVINNumberDataSet TheFindVehicleMainByVINNumberDataSet = new FindVehicleMainByVINNumberDataSet();
        FindDailyVehicleInspectionMaxOdometerDataSet TheFindDailyVehicleInspectionMaxOdometerDataSet = new FindDailyVehicleInspectionMaxOdometerDataSet();
        VehicleMileagleDataSet TheVehicleMileageDataSet = new VehicleMileagleDataSet();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            bool blnFatalError = false;

            try
            {
                blnFatalError = ImportExcel();
            } 
            catch (Exception Ex)
            {
                TheMessagesClass.ErrorMessage("You Have Failed This Import");
            }
        }
        private bool ImportExcel()
        {
            bool blnFatalError = false;
            Excel.Application xlDropOrder;
            Excel.Workbook xlDropBook;
            Excel.Worksheet xlDropSheet;
            Excel.Range range;
            int intColumnRange = 0;
            int intCounter;
            int intNumberOfRecords;
            int intRecordsReturned;
            string strVINNumber;
            int intVehicleID = 0;
            string strVehicleNumber;
            string strManufacturer;
            string strModel;
            string strDescription;
            string strAssignedOffice;
            int intOdometer;
            string strVehicleYear;

            try
            {
                TheVehicleMileageDataSet.vehiclemileage.Rows.Clear();

                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.FileName = "Document"; // Default file name
                dlg.DefaultExt = ".xlsx"; // Default file extension
                dlg.Filter = "Excel (.xlsx)|*.xlsx"; // Filter files by extension

                // Show open file dialog box
                Nullable<bool> result = dlg.ShowDialog();

                // Process open file dialog box results
                if (result == true)
                {
                    // Open document
                    string filename = dlg.FileName;
                }

                PleaseWait PleaseWait = new PleaseWait();
                PleaseWait.Show();

                xlDropOrder = new Excel.Application();
                xlDropBook = xlDropOrder.Workbooks.Open(dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlDropSheet = (Excel.Worksheet)xlDropOrder.Worksheets.get_Item(1);

                range = xlDropSheet.UsedRange;
                intNumberOfRecords = range.Rows.Count;
                intColumnRange = range.Columns.Count;

                for (intCounter = 1; intCounter <= intNumberOfRecords; intCounter++)
                {
                    intVehicleID = intCounter;                    
                    strVehicleNumber = "";
                    strVehicleYear = Convert.ToString((range.Cells[intCounter, 1] as Excel.Range).Value2);
                    strManufacturer = Convert.ToString((range.Cells[intCounter, 2] as Excel.Range).Value2);
                    strModel = Convert.ToString((range.Cells[intCounter, 3] as Excel.Range).Value2);
                    strDescription = Convert.ToString((range.Cells[intCounter, 4] as Excel.Range).Value2);
                    strVINNumber = Convert.ToString((range.Cells[intCounter, 5] as Excel.Range).Value2);
                    strAssignedOffice = "";
                    intOdometer = 0;

                    TheFindVehicleMainByVINNumberDataSet = TheVehicleMainClass.FindVehicleMainByVINNumber(strVINNumber);

                    intRecordsReturned = TheFindVehicleMainByVINNumberDataSet.FindVehicleMainByVINNumber.Rows.Count;

                    if(intRecordsReturned > 0)
                    {
                        intVehicleID = TheFindVehicleMainByVINNumberDataSet.FindVehicleMainByVINNumber[0].VehicleID;
                        strAssignedOffice = TheFindVehicleMainByVINNumberDataSet.FindVehicleMainByVINNumber[0].AssignedOffice;
                        strVehicleNumber = TheFindVehicleMainByVINNumberDataSet.FindVehicleMainByVINNumber[0].VehicleNumber;

                        TheFindDailyVehicleInspectionMaxOdometerDataSet = TheInspectionsClass.FindDailyVehicleInspectionMaxOdometer(intVehicleID);

                        intRecordsReturned = TheFindDailyVehicleInspectionMaxOdometerDataSet.FindDailyVehicleInspectionMaxOdometer.Rows.Count;

                        if(intRecordsReturned > 0)
                        {
                            intOdometer = TheFindDailyVehicleInspectionMaxOdometerDataSet.FindDailyVehicleInspectionMaxOdometer[0].Column1;
                        }
                    }

                    VehicleMileagleDataSet.vehiclemileageRow NewVehicleRow = TheVehicleMileageDataSet.vehiclemileage.NewvehiclemileageRow();

                    NewVehicleRow.AssigedOffice = strAssignedOffice;
                    NewVehicleRow.Description = strDescription;
                    NewVehicleRow.Odometer = intOdometer;
                    NewVehicleRow.VehicleID = intVehicleID;
                    NewVehicleRow.VehicleManufacturer = strManufacturer;
                    NewVehicleRow.VehicleModel = strModel;
                    NewVehicleRow.VehicleNumber = strVehicleNumber;
                    NewVehicleRow.VehicleYear = strVehicleYear;
                    NewVehicleRow.VINNumber = strVINNumber;

                    TheVehicleMileageDataSet.vehiclemileage.Rows.Add(NewVehicleRow);

                }

                PleaseWait.Close();
                dgrResults.ItemsSource = TheVehicleMileageDataSet.vehiclemileage;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Find Vehicle Mileage // Import Excel " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());

                blnFatalError = true;
            }

            return blnFatalError;
        }

        private void btnCopyToExcel_Click(object sender, RoutedEventArgs e)
        {
            int intRowCounter;
            int intRowNumberOfRecords;
            int intColumnCounter;
            int intColumnNumberOfRecords;

            // Creating a Excel object. 
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {


                worksheet = workbook.ActiveSheet;

                worksheet.Name = "OpenOrders";

                int cellRowIndex = 1;
                int cellColumnIndex = 1;
                intRowNumberOfRecords = TheVehicleMileageDataSet.vehiclemileage.Rows.Count;
                intColumnNumberOfRecords = TheVehicleMileageDataSet.vehiclemileage.Columns.Count;

                for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                {
                    worksheet.Cells[cellRowIndex, cellColumnIndex] = TheVehicleMileageDataSet.vehiclemileage.Columns[intColumnCounter].ColumnName;

                    cellColumnIndex++;
                }

                cellRowIndex++;
                cellColumnIndex = 1;

                //Loop through each row and read value from each column. 
                for (intRowCounter = 0; intRowCounter < intRowNumberOfRecords; intRowCounter++)
                {
                    for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                    {
                        worksheet.Cells[cellRowIndex, cellColumnIndex] = TheVehicleMileageDataSet.vehiclemileage.Rows[intRowCounter][intColumnCounter].ToString();

                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                //Getting the location and file name of the excel to save from user. 
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 1;

                saveDialog.ShowDialog();

                workbook.SaveAs(saveDialog.FileName);
                MessageBox.Show("Export Successful");



            }
            catch (System.Exception ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Find Vehicle Mileage // Export to Excel " + ex.Message);

                MessageBox.Show(ex.ToString());
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }
        }
    }
}
