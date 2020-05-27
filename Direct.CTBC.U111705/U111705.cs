using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Direct.Shared;
using System.IO;
using System.Threading.Tasks;

namespace Direct.CTBC.U111705
{

    [DirectSealed]
    [DirectDom("U111705")]
    [ParameterType(false)]
    public class U111705
    {
        [DirectDom("getCardNumber")]
        [DirectDomMethod("excel絕對路徑{please input excelRoute} excel密碼{input excelPwd} 卡號前6後4碼{input unknowCardNumber} 識別碼{input approveCode} 金額{input money}")]
        [MethodDescriptionAttribute("get cardNumber from excelFile")]
        public static string getCardNumber(string excelRoute, string excelPwd, string unknowCardNumber, string approveCode, string money)
        {
            string cardNumber = "";
            try
            {
                cardNumber = model.getCardNumner(excelRoute, excelPwd, unknowCardNumber, approveCode, money);
            }          
            catch (Exception e)
            {
                using (StreamWriter sr = new StreamWriter(@"C:\CRPADATA\U111705\DLL_LOG.txt", true))
                {
                    sr.WriteLine("["+DateTime.Now.ToString()+"]" + e.ToString());
                }
            }

            return cardNumber;
           
        }


        [DirectDom("getCardNumberFromEDC")]
        [DirectDomMethod("excel絕對路徑{please input excelRoute} excel密碼{input excelPwd} 商店代號{input storeCaseNumber} 卡號前6後4碼{input unknowCardNumber} 識別碼{input approveCode} 金額{input money}")]
        [MethodDescriptionAttribute("get cardNumber from EDCexcelFile")]
        public static string getEDCCardNumber(string excelRoute,string excelPwd, string storeCaseNumber, string unknowCardNumber, string approveCode,string money)
        {
            string cardNumber = "";
            try
            {
                cardNumber = model.getCardNumberFromEDC(excelRoute, excelPwd, storeCaseNumber, unknowCardNumber, approveCode, money);
            }
            catch (Exception e)
            {
                using (StreamWriter sr = new StreamWriter(@"C:\CRPADATA\U111705\DLL_LOG.txt", true))
                {
                    sr.WriteLine("[" + DateTime.Now.ToString() + "]" + e.ToString());
                }
            }
            return cardNumber;
        }
        

        [DirectDom("複製空白Excel分頁")]
        [DirectDomMethod("excel絕對路徑{please input excelRoute} excel密碼{input excelPwd} 新增分頁名稱{input sheetName}")]
        [MethodDescriptionAttribute("在excel中複製指定的空白分頁")]
        public static void copyExcelSheet(string excelRoute, string excelPwd, string sheetName)
        {
            try
            {
                model.copyExcelSheet(excelRoute, excelPwd, sheetName);
            }
            catch (Exception e)
            {
                using (StreamWriter sr = new StreamWriter(@"C:\CRPADATA\U111705\DLL_LOG.txt", true))
                {
                    sr.WriteLine("[" + DateTime.Now.ToString() + "]" + e.ToString());
                }
            }

        }

        [DirectDom("寫入全聯異常帳YYYMM")]
        [DirectDomMethod("excel密碼{input excelPwd} excel檔案類別{input excelType}")]
        [MethodDescriptionAttribute("exceltype = true 將異常卡號交易資料寫入全聯異常帳YYYMM,exceltype = false 將異常卡號交易資料寫入全聯異常帳YYYMM-錢包")]
        public static void writeIntoExcel(string excelPwd, bool excelType)
        {
            try
            {
                model.writeIntoExcel(excelPwd,excelType);
            }           
            catch (Exception e)
            {
                using (StreamWriter sr = new StreamWriter(@"C:\CRPADATA\U111705\DLL_LOG.txt", true))
                {
                    sr.WriteLine("[" + DateTime.Now.ToString() + "]" + e.ToString());
                }
            }
        }


        [DirectDom("將D254&255Txt檔合併")]
        [DirectDomMethod("D254Txt{檔案絕對路徑}D255Txt{檔案絕對路徑}")]
        [MethodDescriptionAttribute("將D254&255Txt")]
        public static void addTextList(string d254TxtRoute, string d255TxtRoute)
        {
            try
            {
                model.addTextList(d254TxtRoute,d255TxtRoute);
            }
            catch (Exception e)
            {
                using (StreamWriter sr = new StreamWriter(@"C:\CRPADATA\U111705\DLL_LOG.txt", true))
                {
                    sr.WriteLine("[" + DateTime.Now.ToString() + "]" + e.ToString());
                }
            }
        }

        [DirectDom("將D254&255Txt檔資料剖析")]
        [DirectDomMethod("D254Txt與D255Txt剖析並寫入excel")]
        [MethodDescriptionAttribute("D254Txt與D255Txt剖析並寫入excel")]
        public static bool splitText()
        {
            bool flag = false;
            try
            {
                flag = model.splitTxt();
            }
            catch (Exception e)
            {
                using (StreamWriter sr = new StreamWriter(@"C:\CRPADATA\U111705\DLL_LOG.txt", true))
                {
                    sr.WriteLine("[" + DateTime.Now.ToString() + "]" + e.ToString());
                }
            }
            return flag;
        }
    }
}
