using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.Linq;
namespace ImportSEFData
{   
   public class SEFDataValidation
   {
     
       
      public static string   WrongDataLogs ="";
      public static string LicError = "";
      //public static DatabaseConnection dbc = new DatabaseConnection();

       public static string SqlQuery = "";

        public static bool ValidateData(int CurrentCardID, string ColoumName, int GrdRow,string ColoumValue )
        {
            if ((ColoumName == "ID"))
            {
                
                if (string.IsNullOrEmpty(ColoumValue) || string.IsNullOrWhiteSpace(ColoumValue) || ColoumValue == "0")
                {
                    WrongDataLogs = GetNullMsg(ColoumName, CurrentCardID, GrdRow);
                    return false;
                }
                string ID = ColoumValue.TrimStart(new Char[] { '0' });
                if ((ID.Length != 10))
                {

                    WrongDataLogs = "طول الحقل اقل من 10 ارقام" + "=" + ColoumValue;
                    return false;
                }
                if (ColoumValue.Any(char.IsLetter))
                {
                    WrongDataLogs = "لا يجب ان يحتوى الحقل على حروف  ";
                    return false;
                }
                if (!VerifyReplecation(ColoumName, ID,CurrentCardID))
                {
                    if (CurrentCardID ==6) WrongDataLogs = LicError;
                    else WrongDataLogs = GetConstrMsg(ColoumName, ID);
                    return false;
                }
               

            }
            if (ColoumName == "MilitaryID")
            {

                if (ColoumValue.Any(char.IsLetter))
                {
                    WrongDataLogs = "الحقل لا يجب ان يحتوى على حروف" ;
                    return false;
                }
                if (string.IsNullOrEmpty(ColoumValue) || string.IsNullOrWhiteSpace(ColoumValue) || ColoumValue == "0")
                {
                    WrongDataLogs = GetNullMsg(ColoumName, CurrentCardID, GrdRow);
                    return false;
                }
                
                string MID = ColoumValue.TrimStart(new Char[] { '0' });
                if (!(MID.Length >= 3 && MID.Length <= 6))
                {

                    WrongDataLogs ="طول الرقم العسكرى لايزيد عن 6 ولا يقل عن 3= " + ColoumValue ;
                  
                    return false;
                }
                if ( CurrentCardID != 6)
                {
                    if (VerifyReplecation(ColoumName, MID) == false)
                    {
                        WrongDataLogs = GetConstrMsg(ColoumName, MID);
                        //var confirmResult = DialogResult.None;
                        //confirmResult = MessageBox.Show(GetConstrMsg(ColoumName, MID) + "هل تريد المتابعة", "Database Error", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                        //if (confirmResult == DialogResult.No)
                        //{
                        //    return false;
                        //}
                        return false;
                    }
                }

            }
        
            if (ColoumName == "BloodGroupNo")
            {
                if (string.IsNullOrEmpty(ColoumValue) || string.IsNullOrWhiteSpace(ColoumValue))
                {
                    GetNullMsg(ColoumName, CurrentCardID, GrdRow);
                    return false;
                }
                if (ColoumValue == "0")
                {
                    WrongDataLogs = "خطأ في بيانات حقل فصيلة الدم " + "=" + ColoumValue;
                    return false;
                }

            }
            if (ColoumName == "RankNo")
            {
                if (string.IsNullOrEmpty(ColoumValue) || string.IsNullOrWhiteSpace(ColoumValue) )
                {
                   GetNullMsg(ColoumName, CurrentCardID, GrdRow);
                    return false;
                }
                if ( ColoumValue == "0")
                {
                    WrongDataLogs = "خطأ في بيانات حقل الرتبة " + "=" + ColoumValue;
                    return false;
                }
                
            }
            if (ColoumName == "PersonName")
            {
                if (string.IsNullOrEmpty(ColoumValue) || string.IsNullOrWhiteSpace(ColoumValue) || ColoumValue == "0")
                {
                    WrongDataLogs = GetNullMsg(ColoumName, CurrentCardID, GrdRow);
                    return false;
                }
                if (!IsValidName(ColoumValue))
                {
                    WrongDataLogs = "خطأ في بيانات حقل الاسم " + "=" + ColoumValue ;
                    //MessageBox.Show(msg);
                   
                    return false;
                }
            }
            //taha /11/7/2018
            //to check if enterd date are in true format or not
            if (ColoumName == "IssueDate")
            {
                if (ColoumValue.Any(char.IsLetter))
                {
                    WrongDataLogs = "الحقل تاريخ الاصدار لا يجب ان يحتوى على حروف  ";
                    return false;
                }
                if (string.IsNullOrEmpty(ColoumValue) || string.IsNullOrWhiteSpace(ColoumValue) || ColoumValue == "0")
                {
                  GetNullMsg(ColoumName, CurrentCardID, GrdRow);
                    return false;
                }
                else if(!CheckDate(ColoumValue))
                {
                    WrongDataLogs = "خطأ في صيغةالتاريخ فى الحقل تاريخ الاصدار " + "=" + ColoumValue ;
                    //MessageBox.Show(msg2);
                   
                    return false;
                }
            }

            //if (ColoumName == "OrderType" || ColoumName == "OrderDate" || ColoumName == "LicType" || ColoumName == "OrderNo")
            //{
            
                    //MessageBox.Show("Order Type Not Found!" + "\r\n" + ColoumName + "\r\n" + "Card ( " + CurrentCardID + " / " + GrdRow + " )");
                    //return false;
            //} 

            return true; 
        }
      public static bool CheckOrderData(List <string> OrderData) {

            string type = OrderData[0].ToString();
            string orderNo = OrderData[1].ToString();
            string orderDate = OrderData[2].ToString();

            if ((type == "0" && orderDate == "0" && orderDate == "0") || (type != "0" && orderDate != "0" && orderDate != "0")) {

                if (orderDate != "0")
                {
                    if (orderDate.Any(char.IsLetter))
                    {
                        WrongDataLogs = "لا يجب ان يحتوى على حروف ";
                        return false;
                    }
                    if (!CheckDate(orderDate))
                    {
                        WrongDataLogs = "خطأ في صيغةالتاريخ فى الحقل  " + orderDate;
                        return false;
                    }
                }
                return true; }
            else
                WrongDataLogs = "لابد من أدخال بيانات القرار (الرقم ,التاريخ ,النوع )";
                    return false;
           
        }
      public static bool IsValidName(string name)
        {
            // Check that we were actually passed a name, with a length
            if (name == null || name.Length == 0)
                return false;

            // First character must be a letter
            if (!Char.IsLetter(name, 0))
                return false;

            // Don't allow a leading underscore
            if (name.StartsWith("_"))
                return false;

            // Make sure there's nothing in the proposed name that isn't a letter or digit
            for (int i = 0; i < name.Length; i++)
            {
                if (Char.IsDigit(name, i))
                    return false;
            }

            return true;
        }
      private static Boolean Is_Hijri(string DateText)
        {
            CultureInfo arCI = new CultureInfo("ar-SA");
            UmAlQuraCalendar ul = new UmAlQuraCalendar();
            DateTime convertedDate;
            try
            {
            convertedDate = Convert.ToDateTime(DateText, arCI.DateTimeFormat);
                    string yr, mon, dy, date,Revdate;
                    convertedDate.Kind.ToString();

                    yr = (ul.GetYear(convertedDate)).ToString(); // 1436
                    mon = (ul.GetMonth(convertedDate)).ToString(); // 12
                    dy = (ul.GetDayOfMonth(convertedDate)).ToString(); // 19
                    date = yr + "/" + mon + "/" + dy;
                    Revdate = dy + "/" + mon + "/" + yr;
                    if (date == DateText || Revdate == DateText) return true;
                    else return false;
            }
            catch (Exception)
            {
                return false;
                throw;
            }
        

        }
      public static string RevDate(string DateText)
      { 
          string yr, mon, dy, date, Revdate;
          CultureInfo arCI = new CultureInfo("ar-SA");
          UmAlQuraCalendar ul = new UmAlQuraCalendar();
          DateTime convertedDate;

          if (CheckDate(DateText))
              {
                  convertedDate = Convert.ToDateTime(DateText, arCI.DateTimeFormat);
                  convertedDate.Kind.ToString();
                  yr = (ul.GetYear(convertedDate)).ToString(); // 1436
                  mon = (ul.GetMonth(convertedDate)).ToString(); // 12
                  dy = (ul.GetDayOfMonth(convertedDate)).ToString(); // 19
                  date = yr + "/" + mon + "/" + dy;
                  Revdate = dy + "/" + mon + "/" + yr;
                  if (date == DateText || Revdate == DateText) return date;
                  else return date;
              }
          else return DateText;
         
         
      }
      public static bool CheckDate(string date)
      {
          //CultureInfo arCI = new CultureInfo("ar-SA");
          //DateTime convertedDate = Convert.ToDateTime(date,arCI) ;
          //if (DateTime.TryParseExact(date, "dd/MM/yy", arCI.DateTimeFormat, System.Globalization.DateTimeStyles.None, out convertedDate)) { }
          ////string ReversedDate = DateTime.ParseExact(date, "yyyy-MM-dd", arCI).ToString();
          ////string result = DateTime.ParseExact(date, "dd/mm/yyyy", arCI.DateTimeFormat).ToString("yyyy/MM/dd");
          if ( Is_Hijri(date) && (date.Contains("/")))
                  return true;
              else
                  return false;
          
         
      }

        public static bool VerifyReplecation(string ColoumName, string ColoumValue,int CardID=0)
      {
          if (CardID == 6)
          {
              if (!VerifyLicense(ColoumName,ColoumValue))
              {
                  WrongDataLogs = LicError;
                  return false;
              }
              else
              {
                  return true;
              }
             
          }
            try
            {using (SqlConnection sqlConn = new SqlConnection(Form1.dbc.GlobalConnectionString))
            {                
                  SqlQuery = "select a." + ColoumName + " from (SELECT SEF_Cards_Afrad." + ColoumName + " FROM SEF_Cards_Afrad  union SELECT SEF_Cards_Officers." + ColoumName + " FROM SEF_Cards_Officers  union SELECT SEF_Cards_Civil." + ColoumName + " FROM SEF_Cards_Civil )  a where " + ColoumName + "= " + ColoumValue;
                sqlConn.Open();
                SqlCommand cmd2 = new SqlCommand(SqlQuery, sqlConn);
                using (SqlDataReader sqlReader = cmd2.ExecuteReader())
                {
                    sqlReader.Read();
                    if (sqlReader.HasRows)
                    {
                        return false;
                    }
                    else
                    {

                        return true;
                    }
                }

             }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message.ToString() );
                return false;
            }

        }
        public static bool VerifyLicense(string ColoumName, string ColoumValue)
        {
          
            try
            {
            using (SqlConnection sqlConn = new SqlConnection(Form1.dbc.GlobalConnectionString))
                {
                    SqlQuery = "select a.ID from (SELECT SEF_Cards_Afrad.ID FROM SEF_Cards_Afrad where Not((OrderType = 3) or(OrderType = 4)) And ID Not IN(select SEF_Cards_License.ID from SEF_Cards_License) union SELECT SEF_Cards_Officers.ID FROM SEF_Cards_Officers where Not((OrderType = 3) or(OrderType = 4))and ID Not IN(select SEF_Cards_License.ID from SEF_Cards_License) ) a where  ID= " + ColoumValue;
                    sqlConn.Open();
                    SqlCommand cmd2 = new SqlCommand(SqlQuery, sqlConn);
                    using (SqlDataReader sqlReader = cmd2.ExecuteReader())
                    {
                        sqlReader.Read();
                        if (sqlReader.HasRows)
                        {
                            return true; //person founded and not retierd and not have Lic
                        }
                        else
                        {
                            if (VerifyReplecation(ColoumName, ColoumValue))
                            {
                                LicError = "رقم الهوية غير مسجل";
                                return false;
                            }
                            else
                            {
                                LicError = "رقم الهوية لشخص متقاعد او لديه رخصة بالفعل";
                                return false;
                            }
                        }
                    }
                 }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                return false;
            }
               
        
        }
       public static string GetNullMsg(string colName,int CurrentCardID,int GrdRow )
        {
            //if (MyLocalization.Language == Language.English)
            //    return "Field [" + colName + "] cannot be empty.";
            WrongDataLogs = "الحقل [" + colName + "] لا يمكن أن يترك فارغا. ";
            return WrongDataLogs;
        }

        private static string GetConstrMsg(string colName, string value)
        {
           
            WrongDataLogs= "الحقل [" + colName + "] لا يسمح بوجود تكرارات. القيمة '" + value + "' موجودة بالفعل.";
            return WrongDataLogs;
        }
        public static int GetBloodGroupNo(string BloodGroup)
        {
            switch (BloodGroup)
            {
                case "A": return 1;
                case "A-": return 2;
                case "A+": return 3;
                case "B": return 4;
                case "B-": return 5;
                case "B+": return 6;
                case "AB": return 7;
                case "AB-": return 8;
                case "AB+": return 9;
                case "O": return 10;
                case "O-": return 11;
                case "O+": return 12;
                default: return 0;
            }
        }
        public static int GetlicType(string LicType)
        {
            switch (LicType)
            {
                case "دراجة نارية": return 1;
                case "نقل خفيف": return 2;
                case "نقل متوسط": return 3;
                case "نقل ثقيل": return 4;
                case "نقل فوق ثقيل": return 5;
                default: return 0;
            }
        }

        public static int GetOrderTypeNo(string OrderType)
        {
            switch (OrderType)
            {
                case "تعيين": return 1;
                case "ترقية": return 2;
                case "تقاعد": return 3;
                case "إنهاء خدمة": return 4;
                default: return 0;
            }
        }
        public static string GetTablename(int CardId=0,int Card_Id=0) {
            if (CardId != 0)
            {
                switch (CardId)
                {
                    case 1: return "الضــــباط";
                    case 2: return "الافــــراد";
                    case 3: return "المــدنيين";
                    case 4: return "الضباط المتقاعدين";
                    case 5: return "الافراد المتقاعدين";
                    case 6: return "رخصة السياقة العسكرية";
                    default:
                        return "";

                }
            }
            else
            {
                switch (Card_Id)
                {
                    case 1: return "SEF_Cards_Officers";
                    case 2: return "SEF_Cards_Afrad";
                    case 3: return "SEF_Cards_Civil";
                    case 4: return "SEF_Cards_Officers";
                    case 5: return "SEF_Cards_Afrad";
                    case 6: return "SEF_Cards_License";
                    default:
                        return "";

                }
            }
           
        }
        public static int GetRankNo(int CardID, string Rank)
        {
            //if (Rank.Any(Char.IsDigit))
            //{
            //    return 0;
            //}
            if (CardID == 1 || CardID == 4)           //Officers
            {
                switch (Rank)
                {
                    case "ملازم": return 1;
                    case "ملازم أول": return 2;
                    case "نقيب": return 3;
                    case "رائد": return 4;
                    case "مقدم": return 5;
                    case "عقيد": return 6;
                    case "عميد": return 7;
                    case "لواء": return 8;
                    case "فريق": return 9;
                    case "فريق أول": return 10;
                    default: return 0;

                    //if (Rank == "ملازم") return 1;
                    //else if (Rank == "ملازم أول") return 2;
                    //else if (Rank == "نقيب") return 3;
                    //else if (Rank == "رائد") return 4;
                    //else if (Rank == "مقدم") return 5;
                    //else if (Rank == "عقيد") return 6;
                    //else if (Rank == "عميد") return 7;
                    //else if (Rank == "لواء") return 8;
                    //else if (Rank == "فريق") return 9;
                    //else if (Rank == "فريق أول") return 10;
                    //else return 0;


                }
            }
            else
            {
                if (CardID == 2 || CardID == 5)  //Afrad
                {
                    switch (Rank)
                    {
                        case "جندي": return 1;
                        case "جندي أول": return 2;
                        case "عريف": return 3;
                        case "وكيل رقيب": return 4;
                        case "رقيب": return 5;
                        case "رقيب أول": return 6;
                        case "رئيس رقباء": return 7;
                        default: return 0;
                    }
                }
                else
                    if (CardID == 3)  //Civil
                    {
                        switch (Rank)
                        {
                            case "الأولى": return 1;
                            case "الثانية": return 2;
                            case "الثالثة": return 3;
                            case "الرابعة": return 4;
                            case "الخامسة": return 5;
                            case "السادسة": return 6;
                            case "السابعة": return 7;
                            case "الثامنة": return 8;
                            case "التاسعة": return 9;
                            case "العاشرة": return 10;
                            case "الحادية عشر": return 11;
                            case "الثانية عشر": return 12;
                            case "الثالثة عشر": return 13;
                            case "الرابعة عشر": return 14;
                            case "الخامسة عشر": return 15;
                            default: return 0;
                        }
                    }
                    else
                        return 0;
            }

        }
       //private static void WriteErrorData(string WrongData) {

       //        oExcel.Application xlApp = new oExcel.Application();

       //        if (xlApp == null)
       //        {
       //            MessageBox.Show("Excel is not properly installed!!");
       //            return;
       //        }


       //        oExcel.Workbook xlWorkBook;
       //        oExcel.Worksheet xlWorkSheet;
       //        object misValue = System.Reflection.Missing.Value;

       //        xlWorkBook = xlApp.Workbooks.Add(misValue);
       //        xlWorkSheet = (oExcel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

       //        xlWorkSheet.Cells[1, 1] = "ID";
       //        //xlWorkSheet.Rows[2] = WrongData;




       //        xlWorkBook.SaveAs("d:\\csharp-Excel.xls", oExcel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, oExcel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
       //        xlWorkBook.Close(true, misValue, misValue);
       //        xlApp.Quit();

       //        Marshal.ReleaseComObject(xlWorkSheet);
       //        Marshal.ReleaseComObject(xlWorkBook);
       //        Marshal.ReleaseComObject(xlApp);

       //        MessageBox.Show("Excel file created , you can find the file d:\\csharp-Excel.xls");


       //}

        
   }
}
