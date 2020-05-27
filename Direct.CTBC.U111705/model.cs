using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;

namespace Direct.CTBC.U111705
{
    public class model
    {       
        static int tmp_stroeCaseNoCol = 0;
        static int tmp_stroeNameCol = 0;
        static int tmp_accountDateCol = 0;
        static int tmp_unKnowCardCol = 0;
        static int tmp_cardNumberCol = 0; 
        static int tmp_tradeCaseNoCol = 0;
        static int tmp_moneyCol = 0;
        static int tmp_transationDateCol = 0;     
        static int tmp_approveCol = 0;

        struct customers
        {
            public string storeCaseNo;
            public string storeName;
            public string accountDate;
            public string cardNumber;
            public string tradeCaseNo;
            public string transationDate;
            public string tradeMoney;
            public string approvalCaseNo;
        }

        static List<customers> tempStructureList = new List<customers>();
        static List<string> allLinesText = new List<string>();

        public static string getCardNumner(string excelRoute, string pwd, string unKnowCardNumber, string approvalCode, string money)
        {
            string result_cardNumber = "null";
            bool haveThisCard = false;
            ExcelTool e1 = new ExcelTool();
            e1.openExcel(excelRoute, pwd);
            e1.selectSheet("Sheet1");
            tmp_stroeCaseNoCol = e1.findValueColumn("A1", "S1", "商店代號");
            tmp_stroeNameCol = e1.findValueColumn("A1", "S1", "商店名稱");
            tmp_accountDateCol = e1.findValueColumn("A1", "S1", "帳務日期");
            tmp_unKnowCardCol = e1.findValueColumn("A1", "S1", "卡號前6後4碼");
            tmp_cardNumberCol = e1.findValueColumn("A1", "S1", "卡號");
            tmp_tradeCaseNoCol = e1.findValueColumn("A1", "S1", "交易碼");
            tmp_transationDateCol = e1.findValueColumn("A1", "S1", "交易日期");
            tmp_moneyCol = e1.findValueColumn("A1", "S1", "交易金額");
            tmp_approveCol = e1.findValueColumn("A1", "S1", "授權碼");

            var tempIndex_List = e1.findConditionIndex(unKnowCardNumber, tmp_unKnowCardCol);

            foreach (int i in tempIndex_List)
            {
                string excel_unKonwCardNumber = e1.getCellValue(i, tmp_unKnowCardCol);
                string excel_approvel = e1.getCellValue(i, tmp_approveCol);
                string excel_money = System.Convert.ToString(e1.getDoubleCellValue(i, tmp_moneyCol));
                if (unKnowCardNumber == excel_unKonwCardNumber && approvalCode == excel_approvel && (money == excel_money || "-" + money == excel_money))
                {
                    result_cardNumber = e1.getCellValue(i, tmp_cardNumberCol);
                    haveThisCard = true;
                    break;
                }
            }

            if (haveThisCard)
            {
                foreach (int i in tempIndex_List)
                {
                    string excel_unKonwCardNumber = e1.getCellValue(i, tmp_unKnowCardCol);
                    string excel_approvel = e1.getCellValue(i, tmp_approveCol);
                    string excel_money = System.Convert.ToString(e1.getDoubleCellValue(i, tmp_moneyCol));
                    if (unKnowCardNumber == excel_unKonwCardNumber && (money == excel_money || "-" + money == excel_money))
                    {
                        var c = new customers();
                        c.storeCaseNo = e1.getCellValue(i, tmp_stroeCaseNoCol);
                        c.storeName = e1.getCellValue(i, tmp_stroeNameCol);
                        c.accountDate = e1.getDoubleCellValue(i, tmp_accountDateCol).ToString();
                        c.cardNumber = e1.getCellValue(i, tmp_cardNumberCol);
                        c.tradeCaseNo = e1.getCellValue(i, tmp_tradeCaseNoCol);
                        c.transationDate = e1.getCellValue(i, tmp_transationDateCol);
                        c.tradeMoney = e1.getDoubleCellValue(i, tmp_moneyCol).ToString();
                        c.approvalCaseNo = e1.getCellValue(i, tmp_approveCol);
                        tempStructureList.Add(c);
                    }
                }
            }

            e1.closeExcel();
            return result_cardNumber;
        }

        public static string getCardNumberFromEDC(string excelRoute, string pwd, string storeCaseNumber, string unKnowCardNumber, string approvalCode, string money)
        {
            string result_cardNumberFromEDC = "null";
            ExcelTool e1 = new ExcelTool();
            e1.openExcel(excelRoute, pwd);
            e1.selectSheet(DateTime.Now.ToString("MMdd"));
            tmp_stroeCaseNoCol = 8;  //商店代號
            tmp_cardNumberCol = 7;   //卡號
            tmp_approveCol = 14;     //授權碼
            tmp_moneyCol = 18;       //金額
            int lastRow = e1.getLastRow();
            for (int i = 2; i <= lastRow; i++)
            {
                
                string excel_unKonwCardNumber;
                string excel_storeCaseNumber=e1.getCellValue(i,tmp_stroeCaseNoCol);
                string excel_cardNumber = e1.getCellValue(i, tmp_cardNumberCol);
                string excel_approvel = e1.getCellValue(i, tmp_approveCol);
                string excel_money = e1.getDoubleCellValue(i, tmp_moneyCol).ToString();
                excel_unKonwCardNumber = excel_cardNumber.Substring(0, 6) + "XXXXXX" + excel_cardNumber.Substring(12, 4);
                if ("822" + storeCaseNumber == excel_storeCaseNumber && unKnowCardNumber == excel_unKonwCardNumber && approvalCode == excel_approvel && (money == excel_money || "-" + money == excel_money))
                {
                    result_cardNumberFromEDC = excel_cardNumber;
                    break;
                }
            }
            e1.closeExcel();
            return result_cardNumberFromEDC;
        }
      

        public static void copyExcelSheet(string excelRoute, string pwd, string sheetName)
        {
            ExcelTool e1 = new ExcelTool();
            e1.openExcel(excelRoute,pwd);
            e1.copySheet(sheetName);
            e1.closeExcel();
        }


        // 複製sheet，並將卡號與正負項金額資料撈出
        public static void writeIntoExcel(string excelPwd, bool excelType)  //true-> 開啟全聯異常帳 false -> 全聯異常帳-錢包
        {
            int lastRow = 35;
            string currentDay = DateTime.Now.ToString("MMdd");
            string year = (DateTime.Today.Year - 1911).ToString();
            ExcelTool e1 = new ExcelTool();
            if (excelType)
            {
                e1.openExcel(@"C:\CRPADATA\U111705\downloadFile\全聯異常帳" + year + currentDay.Substring(0, 2) + ".xlsx", excelPwd);
            }
            else
            {
                e1.openExcel(@"C:\CRPADATA\U111705\downloadFile\全聯異常帳" + year + currentDay.Substring(0, 2) + "-錢包.xlsx", excelPwd);
            }      
                    
            if (tempStructureList.Count > 0)
            {
                e1.selectSheet(currentDay);
                foreach (customers c in tempStructureList)
                {
                    e1.setCellValue("'" + c.storeCaseNo, lastRow, 1);
                    e1.setCellValue(c.storeName, lastRow, 2);
                    e1.setCellValue(c.accountDate, lastRow, 3);
                    e1.setCellValue("'" + c.cardNumber, lastRow, 4);
                    e1.setCellValue(c.tradeCaseNo, lastRow, 5);
                    e1.setCellValue(c.transationDate, lastRow, 6);
                    e1.setCellValue(c.tradeMoney, lastRow, 7);
                    e1.setCellValue("'" + c.approvalCaseNo, lastRow, 8);
                    lastRow = lastRow + 1;
                }
            }
            tempStructureList.Clear();
            e1.closeExcel();
        }

         public static void addTextList(string txtRoute_1, string txtRoute_2)
        {

            List<string> D254LinesText = File.ReadAllLines(txtRoute_1, Encoding.Default).ToList();
            List<string> D255LinesText = File.ReadAllLines(txtRoute_2, Encoding.Default).ToList();          
            if (D254LinesText.Count > 0)
            {
                allLinesText.AddRange(D254LinesText);
            }
            if (D255LinesText.Count > 0)
            {
                allLinesText.AddRange(D255LinesText);
            }
        }


         public static bool splitTxt()
         {
             string today = DateTime.Now.ToString("MMdd");
             int lastRow = 2;
             bool flag = false;
             if (allLinesText.Count == 0)
             {
                 Console.WriteLine("TxtFile is empty");
             }
             else
             {
                 ExcelTool e1 = new ExcelTool();
                 e1.openExcel(@"C:\CRPADATA\U111705\downloadFile\全聯EDC異常.xlsx", "12010499");
                 e1.addSheet(today);
                 e1.selectSheet(today);
                 foreach (string line in allLinesText)
                 {
                     e1.setCellValue(line, lastRow, 2);
                     lastRow = lastRow + 1;
                 }
                 e1.dataSplit(allLinesText.Count);
                 for (int i = 2; i <= allLinesText.Count + 1; i++)
                 {
                     e1.setCellValue(DateTime.Now.ToString("MM月dd日"), i, 1);
                 }
                 e1.setExcelTitle();
                 e1.closeExcel();
                 flag = true;
             }
             return flag;
         }

        
    }
}
