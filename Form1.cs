using System;
using oExcel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.VisualBasic;
using System.Reflection;
using System.Globalization;
using System.Threading;
using System.Runtime.InteropServices;
namespace ImportSEFData
{
    public partial class Form1 : Form
    { 
       public static DatabaseConnection dbc = new DatabaseConnection();
              int i,errorCount, NoOfSheets, GrdRow = 0;
            int CardID = 0, RankNo = 0, BloodGroupNo = 0, OrderType = 0, LicType = 0, SheetCols = 0;
            string WrongDataLogs = "", TableName = "";
            long SheetRows = 0;
            bool hasError = false;
            SqlCommand cmd = new SqlCommand("");
             List<string> OrderList = new List<string>();
            SortedList<string, string> ColoumName = new SortedList<string, string>(); 
        public Form1()
        {
            InitializeComponent();
      
            txt_Servername.Text = System.Environment.MachineName;
            txt_DBname.Text = "SS_cards";
            txtDBPassword.Text = "";

        }
        public void Databaseinit(){
               dbc.ServerName = txt_Servername.Text.ToString();
               dbc.DBName = txt_DBname.Text.ToString();
               dbc.ServerPWD = txtDBPassword.Text.ToString();
               dbc.GlobalConnectionString = "Data Source=" + dbc.ServerName + ";Initial Catalog=" + dbc.DBName + "; user id = sa ; Password =" + dbc.ServerPWD;
              
           }
        private void btnOK_Click(object sender, EventArgs e)
        {
            Databaseinit();
            
             CultureInfo oldCI = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            NoOfSheets = 6;// oxlbook.Worksheets.Count;
            var oxlapp = new oExcel.Application();
            oExcel.Workbook oxlbook = oxlapp.Workbooks.Open(Application.StartupPath + "\\SEF Data.xls");
            oExcel.Worksheet oxlsheet;
            //oExcel.Range range;
            //oxlapp.Visible = true;

            SqlTransaction SQLTrans = null; 
            
            try
            {
                SqlConnection con = new SqlConnection(dbc.GlobalConnectionString);
                con.Open();
                SQLTrans = con.BeginTransaction() ;
               
                Cursor.Current = Cursors.WaitCursor;

                for (i = 1; (i <= NoOfSheets); i++)
                {    // NoOfSheets
                    //if (i == 1 || i == 2 || i == 3) 
                    CardID = i;
                    if (CardID == 1 || CardID == 4) TableName = SEFDataValidation.GetTablename(0, CardID);
                    else if (CardID == 2 || CardID == 5) TableName = SEFDataValidation.GetTablename(0, CardID);
                    else if (CardID == 3) TableName = SEFDataValidation.GetTablename(0, CardID);
                    else if (CardID == 6) TableName = SEFDataValidation.GetTablename(0, CardID);
                   if (CardID==6)
                       WrongDataLogs += SEFDataValidation.GetTablename(CardID,0) + "\r\n" + "رقم الهوية	تاريخ الإصدار	نوع الرخصة	رقم القرار	تاريخ القرار	المشكلة" + "\r\n";
                    else
                       WrongDataLogs += SEFDataValidation.GetTablename(CardID,0) + "\r\n" + "رقم الهوية	الرتبة	الاسم	الرقم العسكري	تاريخ الإصدار	فصيلة الدم	نوع القرار	رقم القرار	تاريخ القرار	المشكلة" + "\r\n";

                    //WrongDataLogs += "\r\n" + CardID+"\r\n";
                    labStatus.Text = ("Applying Sheet ( " + (i + (" ) of " + NoOfSheets)));
                    Application.DoEvents();
                    oxlsheet = (oExcel.Worksheet)oxlbook.Worksheets.get_Item(i);
                    SheetCols = oxlsheet.UsedRange.Columns.Count;
                    SheetRows = oxlsheet.UsedRange.Rows.Count;
                 
                    for (GrdRow = 2; (GrdRow <= SheetRows); GrdRow++)
                    {
                        errorCount = 0;
                        ColoumName.Clear();
                        OrderList.Clear();
                        ColoumName.Add("ID", Convert.ToString(((oExcel.Range)oxlsheet.Cells[GrdRow, 1]).Value2));
                        //SEFDataValidation.ValidateData(CardID, "ID", GrdRow, ColoumName["ID"]);
                      
                        if (CardID == 6)
                        {
                            GetLicData(oxlsheet);
                           bool  Order = SEFDataValidation.CheckOrderData(OrderList);
                            foreach (KeyValuePair<string, string> name in ColoumName)
                            {
                                bool Data = SEFDataValidation.ValidateData(CardID, name.Key, GrdRow, name.Value);
                                if ((!Data || !Order) && errorCount<1)
                                {
                                    WrongDataLogs += ColoumName["ID"] + "\t" + ColoumName["IssueDate"] + "\t" + ColoumName["OrderType"] + "\t" + OrderList[1] + "\t" + OrderList[2] + "\t" + SEFDataValidation.WrongDataLogs + "\r\n";
                                        hasError = true;errorCount++;
                                }
                            }
                        }
                        else
                        {
                            GetCardData(oxlsheet);
                            bool  Order = SEFDataValidation.CheckOrderData(OrderList);
                            foreach (KeyValuePair<string,string> name in ColoumName)
                             {
                                 bool Data = SEFDataValidation.ValidateData(CardID, name.Key, GrdRow, name.Value);

                                 if ((!Data || !Order) && errorCount < 1)
                                 {
                                     WrongDataLogs += ColoumName["ID"] + "\t" + ColoumName["RankText"] + "\t" + ColoumName["PersonName"] + "\t" + ColoumName["MilitaryID"] + "\t" + ColoumName["IssueDate"] + "\t" + ColoumName["BloodGroupText"] + "\t" + ColoumName["OrderType"] + "\t" + OrderList[1] + "\t" + OrderList[2] + "\t" + SEFDataValidation.WrongDataLogs + "\r\n";
                                     hasError = true; errorCount++;
                                  }
                             }
                            
                        }
                        //1111111111111111111111111111111111111111111111111
                     
                        if (CardID == 6)
                            cmd = new SqlCommand("INSERT INTO " + TableName + " (ID, IssueDate, LicType, OrderNo, OrderDate) VALUES (" + ColoumName["ID"] + ", '" + ColoumName["IssueDate"] + "', " + OrderList[0] + ", " + OrderList[1] + ", '" + OrderList[2] + "')", con, SQLTrans);
                        else
                            cmd = new SqlCommand("INSERT INTO " + TableName + " (ID, RankNo ,PersonName , MilitaryID, IssueDate, BloodGroupNo, OrderType, OrderNo, OrderDate) VALUES (" + ColoumName["ID"] + ", " + ColoumName["RankNo"] + ", '" + ColoumName["PersonName"] + "', " + ColoumName["MilitaryID"] + ", '" + ColoumName["IssueDate"] + "', " + ColoumName["BloodGroupNo"] + ", " + OrderList[0] + ", " + OrderList[1] + ", '" + OrderList[2] + "')", con, SQLTrans);
                        
                            
                        //222222222222222222222222222222222222222222222222222
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message.ToString() + "\r\n" + " (Grid Row " + GrdRow + " / Card ID" + CardID + " )");
                goto ExecErr;
            }
           
            if (!hasError &&!string.IsNullOrEmpty( cmd.CommandText)) { 
                ExecuteSQl(cmd);
                SQLTrans.Commit();
                oxlapp.Visible = true;
                MessageBox.Show("Executed Succefully");
                goto ExecSucess;
            }
            else
            {
                SQLTrans.Rollback();
                OpenExcelTxt("WrongDataLogs.txt", WrongDataLogs);
            }
           
        ExecErr:
           
       ExecSucess: 
            oxlapp.Quit(); 
            labStatus.Text = "";
            Application.DoEvents();
            Cursor.Current = Cursors.Default;  
            //oxlapp = null;
            //oxlbook = null;
            oxlsheet = null;
            //Marshal.ReleaseComObject(oxlsheet);
            Marshal.ReleaseComObject(oxlbook);
            Marshal.ReleaseComObject(oxlapp);
            this.Close();
        
        }
        private bool GetLicData(oExcel.Worksheet oxlsheet)
        {
            LicType = SEFDataValidation.GetlicType(Convert.ToString(((oExcel.Range)oxlsheet.Cells[GrdRow, 3]).Text));
            
            //IssueDate
            
            ColoumName.Add("IssueDate", SEFDataValidation.RevDate(Convert.ToString(((oExcel.Range)oxlsheet.Cells[GrdRow, 2]).Text)));
            ColoumName.Add("OrderType", Convert.ToString(((oExcel.Range)oxlsheet.Cells[GrdRow, 3]).Text));
            OrderList.Add(LicType.ToString());
            OrderList.Add(Convert.ToString(((oExcel.Range)oxlsheet.Cells[GrdRow, 4]).Value2));
           //orderDate
            OrderList.Add(SEFDataValidation.RevDate(Convert.ToString(((oExcel.Range)oxlsheet.Cells[GrdRow, 5]).Text)));
            for (int x = 0; x < OrderList.Count; x++)
            {
                if (string.IsNullOrWhiteSpace(OrderList[x]) || string.IsNullOrEmpty(OrderList[x]))
                {
                    OrderList[x] = "0";
                }
            }
            return true;
        }
        private bool GetCardData(oExcel.Worksheet oxlsheet)
        {
            RankNo = SEFDataValidation.GetRankNo(CardID, Convert.ToString(((oExcel.Range)oxlsheet.Cells[GrdRow, 2]).Value2));
            BloodGroupNo = SEFDataValidation.GetBloodGroupNo(Convert.ToString(((oExcel.Range)oxlsheet.Cells[GrdRow, 6]).Text));
            OrderType = SEFDataValidation.GetOrderTypeNo(Convert.ToString(((oExcel.Range)oxlsheet.Cells[GrdRow, 7]).Text));
            ColoumName.Add("RankNo", RankNo.ToString());
            ColoumName.Add("RankText", Convert.ToString(((oExcel.Range)oxlsheet.Cells[GrdRow, 2]).Value2));
            ColoumName.Add("PersonName", Convert.ToString(((oExcel.Range)oxlsheet.Cells[GrdRow, 3]).Text));
            ColoumName.Add("MilitaryID", Convert.ToString(((oExcel.Range)oxlsheet.Cells[GrdRow, 4]).Value2));
            ColoumName.Add("IssueDate", SEFDataValidation.RevDate(Convert.ToString(((oExcel.Range)oxlsheet.Cells[GrdRow, 5]).Text)));//Issuedate
            ColoumName.Add("BloodGroupNo", BloodGroupNo.ToString());
            ColoumName.Add("BloodGroupText", Convert.ToString(((oExcel.Range)oxlsheet.Cells[GrdRow, 6]).Text));
            ColoumName.Add("OrderType", Convert.ToString(((oExcel.Range)oxlsheet.Cells[GrdRow, 7]).Text));
            OrderList.Add(OrderType.ToString());
            OrderList.Add(Convert.ToString(((oExcel.Range)oxlsheet.Cells[GrdRow, 8]).Value2));
            OrderList.Add(SEFDataValidation.RevDate(Convert.ToString(((oExcel.Range)oxlsheet.Cells[GrdRow, 9]).Text)));            //OrderDate
            
            for (int x = 0; x < OrderList.Count; x++)
            {
                if (string.IsNullOrWhiteSpace(OrderList[x]) || string.IsNullOrEmpty(OrderList[x]))
                {
                    OrderList[x] = "0";
                }
            }
            return true;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
            Application.Exit();
        }
        public static bool ExecuteSQl(SqlCommand cmd)
        {
            try
            {
                {
                    SqlConnection cn = cmd.Connection;
                    //cn.Open();
                    int RowsAff = cmd.ExecuteNonQuery();
                    //cn.Close();
                    return true;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message.ToString() + "\r\n" + cmd.CommandText);
                return false;
            }

        }
        bool OpenExcelTxt(string FileName, string Txt)
        {
            try
            {FileName = FileName.Replace(" ", "_");
            FileName = (System.Environment.GetEnvironmentVariable("temp") + ("\\" + FileName));
            System.IO.File.WriteAllText(FileName, Txt,UTF8Encoding.Unicode);
            System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo("Excel", FileName);
            psi.UseShellExecute = true;
            System.Diagnostics.Process.Start(psi);
            return true;

            }
            catch (Exception ex)
            {MessageBox.Show(ex.Message.ToString());
                return false;
            }
            
        }
       
    }
}
