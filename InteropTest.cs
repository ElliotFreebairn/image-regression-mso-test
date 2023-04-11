using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using word = Microsoft.Office.Interop.Word;
using excel = Microsoft.Office.Interop.Excel;
using powerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Net.Http;

namespace mso_test
{
    internal class InteropTest
    {
        public static word.Application wordApp = new word.Application();
        public static excel.Application excelApp = new excel.Application();
        public static powerPoint.Application powerPointApp = new powerPoint.Application();

        public static HttpClient client = new HttpClient();

        static InteropTest()
        {
            client.BaseAddress = new Uri("https://bugs.documentfoundation.org");
        }

        private static void Main(string[] args)
        {
            Console.WriteLine("Hello World");
            wordApp.Quit();
            excelApp.Quit();
            powerPointApp.Quit();
        }

        private static (bool, string) TestWordDoc(string fileName)
        {
            word.Document doc = null;
            bool testResult = true;
            string errorMessage = "";

            try
            {
                doc = wordApp.Documents.OpenNoRepairDialog(fileName, ReadOnly: true, Visible: false);
            }
            catch (COMException e)
            {
                testResult = false;
                errorMessage = e.Message;
            }
            if (doc != null)
                doc.Close();

            return (testResult, errorMessage);
        }

        private static (bool, string) TestExcelWorkbook(string fileName)
        {
            excel.Workbook wb = null;
            bool testResult = true;
            string errorMessage = "";

            try
            {
                wb = excelApp.Workbooks.Open(fileName, ReadOnly: true);
            }
            catch (COMException e)
            {
                testResult = false;
                errorMessage = e.Message;
            }

            if (wb != null)
                wb.Close();

            return (testResult, errorMessage);
        }

        private static (bool, string) TestPowerPointPresentation(string fileName)
        {
            powerPoint.Presentation presentation = null;
            bool testResult = true;
            string errorMessage = "";

            try
            {
                presentation = powerPointApp.Presentations.Open(fileName, ReadOnly: Microsoft.Office.Core.MsoTriState.msoTrue, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);
            }
            catch (COMException e)
            {
                testResult = false;
                errorMessage = e.Message;
            }

            if (presentation != null)
                presentation.Close();

            return (testResult, errorMessage);
        }

        private static async void downloadBugList()
        {
            // search all the bugs which has MSO or libreoffice file attachments
            string buglistParameters = "/buglist.cgi" +
                "?f1=attachments.mimetype&f10=attachments.mimetype&f11=attachments.mimetype&f2=attachments.mimetype" +
                "&f3=attachments.mimetype&f4=attachments.mimetype&f5=attachments.mimetype&f6=attachments.mimetype" +
                "&f7=attachments.mimetype&f8=attachments.mimetype&f9=attachments.mimetype" +
                "&j_top=OR&o1=equals&o10=equals&o11=equals&o2=equals&o3=equals&o4=equals&o5=equals&o6=equals&o7=equals&o8=equals&o9=equals" +
                "&query_format=advanced&limit=0&ctype=csv&human=1&v1=application%2Fexcel&v10=application%2Fvnd.ms-powerpoint" +
                "&v11=application%2Fvnd.openxmlformats-officedocument.presentationml.presentation&v2=application%2Fmsword" +
                "&v3=application%2Fpowerpoint&v4=application%2Fvnd.oasis.opendocument.text&v5=application%2Fvnd.oasis.opendocument.presentation" +
                "&v6=application%2Fvnd.oasis.opendocument.spreadsheet&v7=application%2Fvnd.openxmlformats-officedocument.wordprocessingml.document" +
                "&v8=application%2Fvnd.ms-excel&v9=application%2Fvnd.openxmlformats-officedocument.spreadsheetml.sheet";

            using (var respose = client.GetAsync(buglistParameters).Result.Content)
            {
                using (var fileStream = new FileStream("buglist.csv", FileMode.Create))
                {
                    await respose.CopyToAsync(fileStream);
                }
            }
        }
    }
}