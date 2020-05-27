using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Direct.CTBC.U111705
{
    class ExcelTool
    {

        Excel.Application myExcel = null;
        Excel.Workbook myBook = null;
        Excel.Worksheet mySheet = null;
        Excel.Range myRange = null;
        Excel.Range myRange2 = null;
        Excel.Range cellRange = null;

        public void openExcel(string fileRoute, string password)
        {
            myExcel = new Excel.Application();
            myExcel.Visible = true;    //開啟excel觀看
            myExcel.DisplayAlerts = false;
            myBook = myExcel.Workbooks.Open(fileRoute, Password: password);
        }

        public void selectSheet(string sheetName)
        {
            mySheet = myBook.Worksheets[sheetName];
        }

        public void copySheet(string sheetName)
        {
            mySheet = myBook.Worksheets["空白"];
            mySheet.Copy(Type.Missing, myBook.Sheets[myBook.Sheets.Count]);
            myBook.Sheets[myBook.Sheets.Count].Name = sheetName;
        }

        public void addSheet(string sheetName)
        {
            myBook.Sheets.Add(After: myBook.Sheets[myBook.Sheets.Count]);
            myBook.Sheets[myBook.Sheets.Count].Name = sheetName;
        }

        public int getLastRow()  //從最後一欄往上找不為空白列之欄數
        {
            return mySheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
        }

        public int findValueColumn(string x, string y, string target)
        {

            return mySheet.Range[x + ":" + y].Find(target, LookIn: Excel.XlFindLookIn.xlFormulas,
                LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByRows, SearchDirection: Excel.XlSearchDirection.xlNext).Column;
        }

        public List<int> findConditionIndex(string unKnowCardNumber, int unknowCardCol)
        {
            string unknowCardNumberCol = GetExcelColumnName(unknowCardCol);
            List<int> tempIndex = new List<int>();
            Excel.Range currentFind = null;
            Excel.Range firstFind = null;
            Excel.Range unknowCardNumber = myExcel.get_Range(unknowCardNumberCol + ":" + unknowCardNumberCol);
            currentFind = unknowCardNumber.Find(unKnowCardNumber, System.Reflection.Missing.Value,
            Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
            Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
            System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            while (currentFind != null)
            {
                // Keep track of the first range you find. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }

                // If you didn't move to a new range, you are done.
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                      == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                {
                    break;
                }
                tempIndex.Add(currentFind.Row);

                currentFind = unknowCardNumber.FindNext(currentFind);
            }
            return tempIndex;
        }

        public string getCellValue(int x, int y)
        {
            return mySheet.Cells[x, y].value;
        }

        public object getObjectCell(int x, int y)
        {
            return mySheet.Cells[x, y].value();
        }

        public double getDoubleCellValue(int x, int y)
        {
            return mySheet.Cells[x, y].value;
        }

        public void setCellValue(string s, int x, int y)
        {
            mySheet.Cells[x, y] = s;
        }

        public void setCellValue(int s, int x, int y)
        {
            mySheet.Cells[x, y] = s;
        }

        public void setCellValue(double s, int x, int y)
        {
            mySheet.Cells[x, y] = s;
        }

        public void setExcelTitle()
        {
            mySheet.Cells[1, 1] = "處理日";
            mySheet.Cells[1, 3] = "帳務日";
            mySheet.Cells[1, 4] = "店代號";
            mySheet.Cells[1, 7] = "卡號";
            mySheet.Cells[1, 8] = "商店代號";
            mySheet.Cells[1, 9] = "端末機編號";
            mySheet.Cells[1, 12] = "交易日";
            mySheet.Cells[1, 14] = "授權碼";
            mySheet.Cells[1, 17] = "交易別";
            mySheet.Cells[1, 18] = "金額";
            mySheet.Cells[1, 19] = "9";
            mySheet.Cells[1, 20] = "9";
            mySheet.Cells[1, 21] = "7";
            mySheet.Cells[1, 23] = "通知日";
            mySheet.Cells[1, 24] = "結果";
            mySheet.Range["A1:X1"].Font.Bold = true;
            mySheet.Range["A:X"].EntireColumn.AutoFit();
        }

        public string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public void dataSplit(int rowCount)
        {
            myRange = mySheet.Cells[2, 2];
            myRange2 = mySheet.Cells[rowCount + 1, 2];
            cellRange = myExcel.get_Range(myRange, myRange2);
            object[] field_info = { new int[] { 0, 2 }, new int[] { 1, 2 }, new int[] { 10, 2 }, new int[] { 15, 2 }, new int[] { 21, 2 }, new int[] { 27, 2 }, new int[] { 43, 2 }, new int[]{ 55, 2 }, new int[]{ 63, 2 }, new int[]{69, 2 },new int[]{ 70, 2 }, new int[]{ 78 , 2 }, new int[]{ 86 , 2 }
                                      , new int[]{ 92, 2 },new int[]{ 100, 2 }, new int[]{ 112, 2 }, new int[]{ 114, 1 },new int[]{ 121, 2 }, new int[]{ 130 , 2}, new int[]{ 139 , 2}, new int[]{ 146 , 2}  };
            cellRange.TextToColumns(Destination: mySheet.Cells[2, 2], DataType: Excel.XlTextParsingType.xlFixedWidth, FieldInfo: field_info, TrailingMinusNumbers: true);
        }

        public void closeExcel()
        {
            myBook.Save(); //更新數值存檔
            myBook.Close(true); //活頁簿close
            myExcel.Visible = false;
            myExcel.Quit();
            try
            {
                //刪除 Windows工作管理員中的Excel.exe 處理緒.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(this.myExcel);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(this.myBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(this.mySheet);
            }
            catch { }
            this.myExcel = null;
            this.myBook = null;
            this.mySheet = null;
            this.myRange = null;
            this.myRange2 = null;
            GC.Collect();
        }
    }
}
