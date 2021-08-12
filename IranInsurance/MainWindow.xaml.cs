using Aspose.Cells;
using Dbf;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Security.Policy;
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

namespace DbfFileTest
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        DataTable insuranceExcelInfo;

        public MainWindow()
        {
            InitializeComponent();

            cmbMonth.Items.Add("فروردين");
            cmbMonth.Items.Add("ارديبهشت");
            cmbMonth.Items.Add("خرداد");
            cmbMonth.Items.Add("تير");
            cmbMonth.Items.Add("مرداد");
            cmbMonth.Items.Add("شهريور");
            cmbMonth.Items.Add("مهر");
            cmbMonth.Items.Add("آبان");
            cmbMonth.Items.Add("آذر");
            cmbMonth.Items.Add("دي");
            cmbMonth.Items.Add("بهمن");
            cmbMonth.Items.Add("اسفند");

            cmbYear.Items.Add("1400");
            cmbYear.Items.Add("1401");
            cmbYear.Items.Add("1402");
            cmbYear.Items.Add("1403");
            cmbYear.Items.Add("1404");
            cmbYear.Items.Add("1405");
            cmbYear.Items.Add("1406");
            cmbYear.Items.Add("1407");
            cmbYear.Items.Add("1408");
            cmbYear.Items.Add("1409");

            cmbSiteCode.Items.Add("عسلویه");// site 1
            cmbSiteCode.Items.Add("تهران");// site 2

            cmbSiteCode.SelectedIndex = 0;

            System.Globalization.PersianCalendar pc = new System.Globalization.PersianCalendar();
            int year = pc.GetYear(DateTime.Today);
            cmbYear.SelectedIndex = year - 1400;


            int currentMonthNo = pc.GetMonth(DateTime.Today);
            int prevMonthNo = currentMonthNo - 1;
            if (prevMonthNo == 0)// current month is Farvardin
            {
                cmbYear.SelectedIndex = cmbYear.SelectedIndex - 1;
                cmbMonth.SelectedIndex = 11;//prevMonthNo = 12;
            }
            else
            {
                cmbMonth.SelectedIndex = prevMonthNo - 1;
            }
        }

        private void btnLoadExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openDialog = new OpenFileDialog();
                openDialog.Filter = "Excel Files|*.xlsx;*.xls";
                if ((bool)openDialog.ShowDialog())
                {
                    string fileName = openDialog.FileName;
                    string path = System.IO.Path.GetDirectoryName(fileName);

                    insuranceExcelInfo = ConvertExcelToDataTable(fileName);

                    insuranceExcelInfo.Columns.Add("HasError");
                    insuranceExcelInfo.Columns.Add("ErrorDesc");

                    bool isValidExcel = ValidateAndFixInsuranceExcelInfo(insuranceExcelInfo);
                    insuranceList.DataContext = insuranceExcelInfo.DefaultView;
                    if (isValidExcel == false)
                    {
                        MessageBox.Show("Execl is not valid, please correct errors and try again.");
                        return;
                    }

                    string selectedMonth = string.Format("{0:D2}", cmbMonth.SelectedIndex + 1);
                    string selectedYear = string.Format("{0:D2}", cmbYear.SelectedIndex);
                    string siteCode = (cmbSiteCode.SelectedIndex == 0) ? "5114100063" : "0256102071";//smaple site no
                    DataTable workDisketDataTable;
                    DataTable karDisketDataTable;
                    GetInsuranceWorkDisket(insuranceExcelInfo, selectedYear, selectedMonth, siteCode, out workDisketDataTable, out karDisketDataTable);

                    string dSKWOR00FileName = path + @"\DSKWOR00.DBF";
                    string dSKKAR00FileName = path + @"\DSKKAR00.DBF";
                    if (File.Exists(dSKWOR00FileName) || File.Exists(dSKKAR00FileName))
                    {
                        if (MessageBox.Show("Files already exists, do you want to replace?", "Warning", MessageBoxButton.YesNoCancel, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                        {
                            return;
                        }
                    }

                    Dbf.DbfFile.Write(dSKWOR00FileName, workDisketDataTable, Encoding.GetEncoding(1256), true, true);
                    Dbf.DbfFile.Write(dSKKAR00FileName, karDisketDataTable, Encoding.GetEncoding(1256), true, true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }
        }

        public bool ValidateAndFixInsuranceExcelInfo(DataTable insuranceExcelInfo)
        {
            bool isValid = true;
            foreach (DataRow excelItem in insuranceExcelInfo.Rows)
            {
                if (string.IsNullOrEmpty(excelItem["EmpNo"].ToString()))//row is empty. ignore it
                {
                    continue;
                }

                string validationMessage = ValidatePersianDate(excelItem["Startdate"], '/', true);
                if (string.IsNullOrEmpty(validationMessage) == false)
                {
                    excelItem["HasError"] = "Yes";
                    excelItem["ErrorDesc"] = "Startdate: " + validationMessage;
                    isValid = false;
                    continue;
                }

                validationMessage = ValidatePersianDate(excelItem["Tarkdate"], '/', true);
                if (string.IsNullOrEmpty(validationMessage) == false)
                {
                    excelItem["HasError"] = "Yes";
                    excelItem["ErrorDesc"] = "Tarkdate: " + validationMessage;
                    isValid = false;
                    continue;
                }

                validationMessage = ValidatePersianDate(excelItem["Sodoordate"], '/', false);
                if (string.IsNullOrEmpty(validationMessage) == false)
                {
                    excelItem["HasError"] = "Yes";
                    excelItem["ErrorDesc"] = "Sodoordate: " + validationMessage;
                    isValid = false;
                    continue;
                }

                validationMessage = ValidatePersianDate(excelItem["Birthdate"], '/', false);
                if (string.IsNullOrEmpty(validationMessage) == false)
                {
                    excelItem["HasError"] = "Yes";
                    excelItem["ErrorDesc"] = "Birthdate: " + validationMessage;
                    isValid = false;
                    continue;
                }

                validationMessage = ValidateNationalId(excelItem["NationalId"].ToString());
                if (string.IsNullOrEmpty(validationMessage) == false)
                {
                    excelItem["HasError"] = "Yes";
                    excelItem["ErrorDesc"] = "NationalId: " + validationMessage;
                    isValid = false;
                    continue;
                }
                else
                {
                    excelItem["NationalId"] = string.Format("{0:D10}", Convert.ToInt64(excelItem["NationalId"].ToString()));
                }


                if (excelItem["Daywork"] == null || string.IsNullOrEmpty(excelItem["Daywork"].ToString()))
                {
                    excelItem["HasError"] = "Yes";
                    excelItem["ErrorDesc"] = "Daywork: " + "can not be empty.";
                    isValid = false;
                    continue;
                }

                if (excelItem["Dastmozderouzane"] == null || string.IsNullOrEmpty(excelItem["Dastmozderouzane"].ToString()))
                {
                    excelItem["HasError"] = "Yes";
                    excelItem["ErrorDesc"] = "Dastmozderouzane: " + "can not be empty.";
                    isValid = false;
                    continue;
                }

                if (excelItem["Dastmozdemahane"] == null || string.IsNullOrEmpty(excelItem["Dastmozdemahane"].ToString()))
                {
                    excelItem["HasError"] = "Yes";
                    excelItem["ErrorDesc"] = "Dastmozdemahane: " + "can not be empty.";
                    isValid = false;
                    continue;
                }

                if (excelItem["Mazayaibime"] == null || string.IsNullOrEmpty(excelItem["Mazayaibime"].ToString()))
                {
                    excelItem["HasError"] = "Yes";
                    excelItem["ErrorDesc"] = "Mazayaibime: " + "can not be empty.";
                    isValid = false;
                    continue;
                }

                if (excelItem["Mashmoulbime"] == null || string.IsNullOrEmpty(excelItem["Mashmoulbime"].ToString()))
                {
                    excelItem["HasError"] = "Yes";
                    excelItem["ErrorDesc"] = "Mashmoulbime: " + "can not be empty.";
                    isValid = false;
                    continue;
                }

                if (excelItem["GROSS"] == null || string.IsNullOrEmpty(excelItem["GROSS"].ToString()))
                {
                    excelItem["HasError"] = "Yes";
                    excelItem["ErrorDesc"] = "GROSS: " + "can not be empty.";
                    isValid = false;
                    continue;
                }

                if (excelItem["Bimekarmand"] == null || string.IsNullOrEmpty(excelItem["Bimekarmand"].ToString()))
                {
                    excelItem["HasError"] = "Yes";
                    excelItem["ErrorDesc"] = "Bimekarmand: " + "can not be empty.";
                    isValid = false;
                    continue;
                }

                
                excelItem["PositionCode"] = excelItem["PositionCode"].ToString().PadLeft(6, '0');
                

                excelItem["HasError"] = "No";
            }

            return isValid;
        }

        public static string ValidatePersianDate(object date, char seperator, bool canBeNull)
        {
            if (canBeNull == false && (date == null || string.IsNullOrEmpty(date.ToString())))
            {
                return "Date can not be null";
            }

            if (canBeNull && (date == null || string.IsNullOrEmpty(date.ToString())))
            {
                return string.Empty;
            }

            string stringDate = date.ToString();
            if (stringDate.Length > 10)
            {
                return "Date should be in YYYY/MM/DD fromat(\"1400/05/22\")";
            }

            try
            {
                string[] splittedDate = date.ToString().Split(seperator);
                if (splittedDate[0].Length != 4)
                {
                    return "Year should be 4 digits";
                }
                if (splittedDate[1].Length > 2)
                {
                    return "Month should be 1 or 2 digits";
                }
                if (splittedDate[2].Length > 2)
                {
                    return "Day should be 1 or 2 digits";
                }


                int year = Convert.ToInt32(splittedDate[0]);
                int month = Convert.ToInt32(splittedDate[1]);
                int day = Convert.ToInt32(splittedDate[2]);
                DateTime dt = new System.Globalization.PersianCalendar().ToDateTime(year, month, day, 0, 0, 0, 0);
                if (dt.Year < 1930)
                {
                    return "wow, Is it Jannati?";
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return string.Empty;
        }

        public static string ValidateNationalId(String melliCode)
        {
            try
            {
                if (string.IsNullOrEmpty(melliCode.Trim()))
                {
                    return "Melli Code is empty";
                }

                else if (melliCode.Length > 10 || melliCode.Length < 8)
                {
                    return "National Id is less than 8 or more than 10 digits.";
                }
                else
                {
                    melliCode = string.Format("{0:D10}", Convert.ToInt64(melliCode));
                    int sum = 0;

                    for (int i = 0; i < 9; i++)
                    {
                        sum += (Convert.ToInt32(melliCode[i].ToString())) * (10 - i);
                    }

                    int lastDigit;
                    int divideRemaining = sum % 11;

                    if (divideRemaining < 2)
                    {
                        lastDigit = divideRemaining;
                    }
                    else
                    {
                        lastDigit = 11 - (divideRemaining);
                    }

                    if (Convert.ToInt32(melliCode[9].ToString()) == lastDigit)
                    {
                        return string.Empty;
                    }
                    else
                    {
                        return "Invalid MelliCode";
                    }
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public static void GetInsuranceWorkDisket(DataTable excelDataTable, string year, string period, string siteCode, out DataTable workDataTable, out DataTable karDataTable)
        {
            workDataTable = new DataTable();

            workDataTable.Columns.Add("DSW_ID").MaxLength = 10;//1
            workDataTable.Columns.Add("DSW_YY").MaxLength = 2;//2
            workDataTable.Columns.Add("DSW_MM").MaxLength = 2;//3
            workDataTable.Columns.Add("DSW_LISTNO").MaxLength = 12;//4
            workDataTable.Columns.Add("DSW_ID1").MaxLength = 10;//5
            workDataTable.Columns.Add("DSW_FNAME").MaxLength = 100;//6
            workDataTable.Columns.Add("DSW_LNAME").MaxLength = 100;//7
            workDataTable.Columns.Add("DSW_DNAME").MaxLength = 100;//8
            workDataTable.Columns.Add("DSW_IDNO").MaxLength = 15;//9
            workDataTable.Columns.Add("DSW_IDPLC").MaxLength = 100;//10
            workDataTable.Columns.Add("DSW_IDATE").MaxLength = 8;//11
            workDataTable.Columns.Add("DSW_BDATE").MaxLength = 8;//12
            workDataTable.Columns.Add("DSW_SEX").MaxLength = 3; //13
            workDataTable.Columns.Add("DSW_NAT").MaxLength = 10;//14
            workDataTable.Columns.Add("DSW_OCP").MaxLength = 100;//15
            workDataTable.Columns.Add("DSW_SDATE").MaxLength = 8;//16
            workDataTable.Columns.Add("DSW_EDATE").MaxLength = 8;//17
            workDataTable.Columns.Add("DSW_DD").MaxLength = 2;//18
            workDataTable.Columns.Add("DSW_ROOZ").MaxLength = 12;//19
            workDataTable.Columns.Add("DSW_MAH").MaxLength = 12;//20
            workDataTable.Columns.Add("DSW_MAZ").MaxLength = 12;//21
            workDataTable.Columns.Add("DSW_MASH").MaxLength = 12;//22
            workDataTable.Columns.Add("DSW_TOTL").MaxLength = 12;//23
            workDataTable.Columns.Add("DSW_BIME").MaxLength = 12;//24
            workDataTable.Columns.Add("DSW_PRATE").MaxLength = 2;//25
            workDataTable.Columns.Add("DSW_JOB").MaxLength = 6;//26
            workDataTable.Columns.Add("PER_NATCOD").MaxLength = 10;//27

            long sumOfWorkingDays = 0;
            long sumOfDailySalary = 0;
            long sumOfMonthlySalary = 0;
            long sumOfMonthlyBenefits = 0;
            long sumOfMonthlyBenefitsAndSalary = 0;
            long sumOfTotalSalary = 0;
            long sumOfEmpInsurance = 0;
            long sumOfCompanyInsurance = 0;
            long sumOfUnEmployementPremium = 0;

            foreach (DataRow excelItem in excelDataTable.Rows)
            {
                if (string.IsNullOrEmpty(excelItem["EmpNo"].ToString()))
                {
                    continue;
                }

                string startDate = excelItem["Startdate"] == null ? "" : excelItem["Startdate"].ToString().Replace("/", "");
                string tarkDate = excelItem["Tarkdate"] == null ? "" : excelItem["Tarkdate"].ToString().Replace("/", "");
                string sodoorDate = excelItem["Sodoordate"] == null ? "" : excelItem["Sodoordate"].ToString().Replace("/", "");
                workDataTable.Rows.Add(
                    siteCode,//1
                    year,//2
                    period,//3
                    1,//4
                    excelItem["INSNO"],//5
                    excelItem["Name"],//6
                    excelItem["Family"],//7
                    excelItem["FatherName"],//8
                    excelItem["IDNo"],//9
                    excelItem["Sodoorplace"],//10
                    sodoorDate,//11
                    excelItem["Birthdate"].ToString().Replace("/", ""),//12
                    excelItem["Sex"],//13
                    "ايرانى",//14
                    excelItem["Position"],//15
                    startDate,//16
                    tarkDate,//17
                    excelItem["Daywork"],//18
                    excelItem["Dastmozderouzane"],//19
                    excelItem["Dastmozdemahane"],//20
                    excelItem["Mazayaibime"],//21
                    excelItem["Mashmoulbime"],//22
                    excelItem["GROSS"],//23
                    excelItem["Bimekarmand"],//24
                    0,//25
                    excelItem["PositionCode"],//26
                    excelItem["NationalId"]//27
                    );

                sumOfWorkingDays += Convert.ToInt64(excelItem["Daywork"]);
                sumOfDailySalary += Convert.ToInt64(excelItem["Dastmozderouzane"]);
                sumOfMonthlySalary += Convert.ToInt64(excelItem["Dastmozdemahane"]);
                sumOfMonthlyBenefits += Convert.ToInt64(excelItem["Mazayaibime"]);
                sumOfMonthlyBenefitsAndSalary += Convert.ToInt64(excelItem["Mashmoulbime"]);
                sumOfTotalSalary += Convert.ToInt64(excelItem["GROSS"]);
                sumOfEmpInsurance += Convert.ToInt64(excelItem["Bimekarmand"]);
                
            }

            sumOfCompanyInsurance = Convert.ToInt64(sumOfMonthlyBenefitsAndSalary * 0.2);// Convert.ToInt32(excelItem["Daywork"]);
            sumOfUnEmployementPremium = Convert.ToInt64(sumOfMonthlyBenefitsAndSalary * 0.03);// Convert.ToInt32(excelItem["Daywork"]);

            karDataTable = new DataTable();

            karDataTable.Columns.Add("DSK_ID").MaxLength = 10;//1
            karDataTable.Columns.Add("DSK_NAME").MaxLength = 100;//2
            karDataTable.Columns.Add("DSK_FARM").MaxLength = 100;//3
            karDataTable.Columns.Add("DSK_ADRS").MaxLength = 100;//4
            karDataTable.Columns.Add("DSK_KIND").MaxLength = 1;//5
            karDataTable.Columns.Add("DSK_YY").MaxLength = 2;//6
            karDataTable.Columns.Add("DSK_MM").MaxLength = 2;//7
            karDataTable.Columns.Add("DSK_LISTNO").MaxLength = 12;//8
            karDataTable.Columns.Add("DSK_DISC").MaxLength = 100;//9
            karDataTable.Columns.Add("DSK_NUM").MaxLength = 5;//10
            karDataTable.Columns.Add("DSK_TDD").MaxLength = 6;//11
            karDataTable.Columns.Add("DSK_TROOZ").MaxLength = 12;//12
            karDataTable.Columns.Add("DSK_TMAH").MaxLength = 12; //13
            karDataTable.Columns.Add("DSK_TMAZ").MaxLength = 12;//14
            karDataTable.Columns.Add("DSK_TMASH").MaxLength = 12;//15
            karDataTable.Columns.Add("DSK_TTOTL").MaxLength = 12;//16
            karDataTable.Columns.Add("DSK_TBIME").MaxLength = 12;//17
            karDataTable.Columns.Add("DSK_TKOSO").MaxLength = 12;//18
            karDataTable.Columns.Add("DSK_BIC").MaxLength = 12;//19
            karDataTable.Columns.Add("DSK_RATE").MaxLength = 5;//20
            karDataTable.Columns.Add("DSK_PRATE").MaxLength = 2;//21
            karDataTable.Columns.Add("DSK_BIMH").MaxLength = 12;//22
            karDataTable.Columns.Add("MON_PYM").MaxLength = 3;//23

            karDataTable.Rows.Add(
                    siteCode,//1
                    "نام شرکت",//2
                    "نام شرکت",//3
                    "آدرس شرکت",//4
                    0,//5
                    year,//6
                    period,//7
                    1,//8
                    0,//9
                    excelDataTable.Rows.Count,//10
                    sumOfWorkingDays,//11
                    sumOfDailySalary,//12
                    sumOfMonthlySalary,//13
                    sumOfMonthlyBenefits,//14
                    sumOfMonthlyBenefitsAndSalary,//15
                    sumOfTotalSalary,//16
                    sumOfEmpInsurance,//17
                    sumOfCompanyInsurance,//18
                    sumOfUnEmployementPremium,//19
                    23,//20
                    0,//21
                    0,//22
                    "001"//23
                    );
        }
        
        public static DataTable ConvertExcelToDataTable(string FileName)
        {
            // Create a file stream containing the Excel file to be opened
            FileStream fstream = new FileStream(FileName, FileMode.Open);

            // Instantiate a Workbook object
            //Opening the Excel file through the file stream
            Workbook workbook = new Workbook(fstream);

            // Access the first worksheet in the Excel file
            Worksheet worksheet = workbook.Worksheets[0];

            // Export the contents of 2 rows and 2 columns starting from 1st cell to DataTable
            DataTable dataTable = worksheet.Cells.ExportDataTableAsString(0, 0, worksheet.Cells.Rows.Count, worksheet.Cells.Columns.Count, true);
            dataTable.Columns.Add("RowNo").SetOrdinal(0);

            bool rowIsEmpty = true;
            DataRow dataRow = dataTable.Rows[dataTable.Rows.Count - 1];
            
            while (rowIsEmpty)
            {
                foreach (object cellValue in dataRow.ItemArray)
                {
                    if (cellValue != null && string.IsNullOrEmpty(cellValue.ToString()) == false)
                    {
                        rowIsEmpty = false;
                        break;
                    }
                }

                if (rowIsEmpty)
                {
                    dataTable.Rows.RemoveAt(dataTable.Rows.Count - 1);
                    dataRow = dataTable.Rows[dataTable.Rows.Count - 1];
                }
            }

            int rowCounter = 1;
            foreach (DataRow row in dataTable.Rows)
            {
                row["RowNo"] = rowCounter;
                rowCounter++;
            }
            // Close the file stream to free all resources
            fstream.Close();
            return dataTable;
        }

        private void btnSaveAsDisket_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (insuranceExcelInfo == null)
                {
                    MessageBox.Show("Please select excel first.");
                    return;
                }

                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.FileName = "";

                if ((bool)saveDialog.ShowDialog())
                {
                    string fileName = saveDialog.FileName;
                    string path = System.IO.Path.GetDirectoryName(fileName);

                    string dSKWOR00FileName = path + @"\DSKWOR00.DBF";
                    string dSKKAR00FileName = path + @"\DSKKAR00.DBF";
                    if (File.Exists(dSKWOR00FileName) || File.Exists(dSKKAR00FileName))
                    {
                        if (MessageBox.Show("Files already exists, do you want to replace?", "Warning", MessageBoxButton.YesNoCancel, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                        {
                            return;
                        }
                    }

                    string selectedMonth = string.Format("{0:D2}", cmbMonth.SelectedIndex + 1);
                    string selectedYear = string.Format("{0:D2}", cmbYear.SelectedIndex);
                    string siteCode = (cmbSiteCode.SelectedIndex == 0) ? "5114100063" : "0256102071";

                    DataTable workDisketDataTable;
                    DataTable karDisketDataTable;
                    GetInsuranceWorkDisket(insuranceExcelInfo, selectedYear, selectedMonth, siteCode, out workDisketDataTable, out karDisketDataTable);
                    Dbf.DbfFile.Write(dSKWOR00FileName, workDisketDataTable, Encoding.GetEncoding(1256), true, true);//.GetEncoding(1252));
                    Dbf.DbfFile.Write(dSKKAR00FileName, karDisketDataTable, Encoding.GetEncoding(1256), true, true);//.GetEncoding(1252));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }
        }
    }
}




