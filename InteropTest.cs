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
using System.IO;
using Newtonsoft.Json.Linq;
using System.Net;

namespace mso_test
{
    internal class InteropTest
    {
        public static word.Application wordApp;
        public static excel.Application excelApp;
        public static powerPoint.Application powerPointApp;

        public static void startApplication(string application)
        {
            if (application == "word")
            {
                wordApp = new word.Application();
                wordApp.DisplayAlerts = word.WdAlertLevel.wdAlertsNone;
                wordApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            }

            if (application == "excel")
            {
                excelApp = new excel.Application();
                excelApp.DisplayAlerts = false;
                excelApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            }

            if (application == "powerpoint")
            {
                powerPointApp = new powerPoint.Application();
                powerPointApp.DisplayAlerts = powerPoint.PpAlertLevel.ppAlertsNone;
                powerPointApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            }
        }

        public static void quitApplication(string application)
        {
            if (application == "word")
                wordApp.Quit();

            if (application == "excel")
                excelApp.Quit();

            if (application == "powerpoint")
                powerPointApp.Quit();
        }

        public const string bugListFileName = "bugList.csv";
        public static HttpClient bugsClient = new HttpClient() { BaseAddress = new Uri("https://bugs.documentfoundation.org") };
        public static HttpClient coolClient = new HttpClient() { BaseAddress = new Uri("https://staging.eu.collaboraonline.com/cool/convert-to") };

        public static HashSet<string> allowedMimeTypes = new HashSet<string>()
        {
            "application/excel",
            "application/msword",
            "application/powerpoint",
            "application/vnd.ms-excel",
            "application/vnd.ms-powerpoint",
            "application/vnd.oasis.opendocument.presentation",
            "application/vnd.oasis.opendocument.spreadsheet",
            "application/vnd.oasis.opendocument.text",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        };

        public static HashSet<string> docExtention = new HashSet<string>()
        {
            ".odt",
            ".docx",
            ".doc"
        };

        public static HashSet<string> wordExtention = new HashSet<string>()
        {
            ".docx",
            ".doc"
        };

        public static HashSet<string> sheetExtention = new HashSet<string>()
        {
            ".ods",
            ".xls",
            ".xlsx"
        };

        public static HashSet<string> excelExtention = new HashSet<string>()
        {
            ".xls",
            ".xlsx"
        };

        public static HashSet<string> presentationExtention = new HashSet<string>()
        {
            ".odp",
            ".ppt",
            ".pptx"
        };

        public static HashSet<string> powerPointExtention = new HashSet<string>()
        {
            ".ppt",
            ".pptx"
        };

        //args can be word/excel/powerpoint
        //specified appropriate docs will be tested with the specified program
        private static async Task Main(string[] args)
        {
            ServicePointManager.Expect100Continue = false;
            startApplication(args[0]);
            await testDownloadedfiles(args[0]);

            quitApplication(args[0]);
        }

        public static (bool, string) testFile(string application, string fileName)
        {
            if (File.Exists(fileName + ".failed"))
                return (false, "");

            if (application == "word")
            {
                return TestWordDoc(fileName);
            }
            else if (application == "excel")
            {
                return TestExcelWorkbook(fileName);
            }
            else if (application == "powerpoint")
            {
                return TestPowerPointPresentation(fileName);
            }

            return (false, "");
        }

        public static async Task testDirectoriy(string application, DirectoryInfo dir, string convertTo)
        {
            FileInfo[] fileInfos = dir.GetFiles();
            var watch = new System.Diagnostics.Stopwatch();
            foreach (FileInfo file in fileInfos)
            {
                if (file.Extension == ".failed" ||
                    File.Exists(Path.GetFullPath(@"converted\" + convertTo + @"\" + Path.GetFileNameWithoutExtension(file.Name) + "." + convertTo)) ||
                    File.Exists(Path.GetFullPath(@"converted\" + convertTo + @"\" + Path.GetFileNameWithoutExtension(file.Name) + "." + convertTo + ".failed")))
                    continue;

                Console.WriteLine("Starting test for " + file.Name);
                watch.Restart();
                bool DownloadResult = testFile(application, file.FullName).Item1;

                if (DownloadResult)
                {
                    await convertFile(file.FullName, Path.GetFileNameWithoutExtension(file.Name), convertTo).ContinueWith(task =>
                    {
                        if (string.IsNullOrEmpty(task.Result))
                            return;
                        (bool, string) ConvertResult = testFile(application, task.Result);
                        if (!ConvertResult.Item1)
                        {
                            using (FileStream fs = new FileStream("failed_files_" + application + ".txt", FileMode.Append))
                            {
                                byte[] info = new UTF8Encoding(true).GetBytes(file.FullName + ": " + ConvertResult.Item2 + "\n");
                                fs.Write(info, 0, info.Length);
                            }
                        }
                    });
                }
                watch.Stop();
                Console.WriteLine(file.Name + $" testing took {watch.ElapsedMilliseconds} ms\n");
            }
        }

        public static async Task testDownloadedfiles(string application)
        {
            DirectoryInfo downloadedDirInfo = new DirectoryInfo(@"download");
            if (!downloadedDirInfo.Exists)
            {
                return;
            }
            DirectoryInfo[] downloadedDictInfo = downloadedDirInfo.GetDirectories();
            foreach (DirectoryInfo dict in downloadedDictInfo)
            {
                if (application == "word" && (dict.Name == "doc" || dict.Name == "docx" || dict.Name == "odt"))
                {
                    await testDirectoriy(application, dict, "docx");
                }
                else if (application == "excel" && (dict.Name == "xls" || dict.Name == "xlsx" || dict.Name == "ods"))
                {
                    await testDirectoriy(application, dict, "xlsx");
                }
                else if (application == "powerpoint" && (dict.Name == "ppt" || dict.Name == "pptx" || dict.Name == "odp"))
                {
                    await testDirectoriy(application, dict, "pptx");
                }
            }
        }

        public static void testConvertedFile(string application)
        {
            string directoryName = "";
            if (application == "word")
                directoryName = "docx";
            else if (application == "excel")
                directoryName = "xlsx";
            else if (application == "powerpoint")
                directoryName = "pptx";
            DirectoryInfo convertedDirInfo = new DirectoryInfo(@"converted\" + directoryName);
            if (!convertedDirInfo.Exists)
            {
                return;
            }
            FileInfo[] convertedFileInfo = convertedDirInfo.GetFiles();

            using (FileStream fs = new FileStream("failed_files_" + application + ".txt", FileMode.Append))
            {
                foreach (FileInfo file in convertedFileInfo)
                {
                    (bool, string) result = (true, "");
                    if (application == "word" && docExtention.Contains(file.Extension))
                    {
                        if (wordExtention.Contains(file.Extension))
                            result = TestWordDoc(file.FullName);
                    }
                    else if (application == "excel" && sheetExtention.Contains(file.Extension))
                    {
                        if (excelExtention.Contains(file.Extension))
                            result = TestExcelWorkbook(file.FullName);
                    }
                    else if (application == "powerpoint" && presentationExtention.Contains(file.Extension))
                    {
                        if (powerPointExtention.Contains(file.Extension))
                            result = TestPowerPointPresentation(file.FullName);
                    }

                    if (!result.Item1)
                    {
                        byte[] info = new UTF8Encoding(true).GetBytes(file.FullName + ": " + result.Item2 + "\n");
                        fs.Write(info, 0, info.Length);
                    }
                }
            }
        }

        public static async Task<string> convertFile(string fullFileName, string fileName, string convertTo)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), coolClient.BaseAddress + "/" + convertTo))
            {
                var multipartContent = new MultipartFormDataContent
                {
                    { new ByteArrayContent(File.ReadAllBytes(fullFileName)), "data", Path.GetFileName(fullFileName) }
                };
                request.Content = multipartContent;

                try
                {
                    using (var response = await coolClient.SendAsync(request))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            Console.Error.WriteLine("Faild to convert " + fullFileName + ": " + response.StatusCode);
                            return "";
                        }
                        Directory.CreateDirectory(Path.GetDirectoryName(@"converted\" + convertTo + @"\"));

                        string convertedFilePath = @"converted\" + convertTo + @"\" + fileName + "." + convertTo;
                        using (FileStream fs = File.Open(convertedFilePath, FileMode.Create))
                        {
                            return await response.Content.CopyToAsync(fs).ContinueWith(task => { return fs.Name; });
                        }
                    }
                }
                catch (Exception ex) { Console.WriteLine(ex); }
            }
            return "";
        }

        private static (bool, string) TestWordDoc(string fileName)
        {
            word.Document doc = null;
            bool testResult = true;
            string errorMessage = "";

            try
            {
                doc = wordApp.Documents.OpenNoRepairDialog(fileName, ReadOnly: true, Visible: false, PasswordDocument: "'");
            }
            catch (COMException e)
            {
                testResult = false;
                errorMessage = e.Message;
            }
            if (doc != null)
                doc.Close(SaveChanges: false);

            if (!testResult)
            {
                try
                {
                    System.IO.File.Move(fileName, fileName + ".failed");
                }
                catch (System.IO.IOException ex) { Console.WriteLine(ex.Message); }
            }

            return (testResult, errorMessage);
        }

        private static (bool, string) TestExcelWorkbook(string fileName)
        {
            excel.Workbook wb = null;
            bool testResult = true;
            string errorMessage = "";

            try
            {
                wb = excelApp.Workbooks.Open(fileName, ReadOnly: true, Password: "'", UpdateLinks: false);
            }
            catch (COMException e)
            {
                testResult = false;
                errorMessage = e.Message;
            }

            if (wb != null)
                wb.Close(SaveChanges: false);

            if (!testResult)
            {
                try
                {
                    System.IO.File.Move(fileName, fileName + ".failed");
                }
                catch (System.IO.IOException ex) { Console.WriteLine(ex.Message); }
            }

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

            if (!testResult)
            {
                try
                {
                    System.IO.File.Move(fileName, fileName + ".failed");
                }
                catch (System.IO.IOException ex) { Console.WriteLine(ex.Message); }
            }

            return (testResult, errorMessage);
        }

        public static async Task downloadBugList()
        {
            // search all the bugs which has MSO or libreoffice file attachments
            string bugListParameters = "/buglist.cgi" +
                "?f1=attachments.mimetype&f10=attachments.mimetype&f11=attachments.mimetype&f2=attachments.mimetype" +
                "&f3=attachments.mimetype&f4=attachments.mimetype&f5=attachments.mimetype&f6=attachments.mimetype" +
                "&f7=attachments.mimetype&f8=attachments.mimetype&f9=attachments.mimetype" +
                "&j_top=OR&o1=equals&o10=equals&o11=equals&o2=equals&o3=equals&o4=equals&o5=equals&o6=equals&o7=equals&o8=equals&o9=equals" +
                "&query_format=advanced&limit=0&ctype=csv&human=1&v1=application%2Fexcel&v10=application%2Fvnd.ms-powerpoint" +
                "&v11=application%2Fvnd.openxmlformats-officedocument.presentationml.presentation&v2=application%2Fmsword" +
                "&v3=application%2Fpowerpoint&v4=application%2Fvnd.oasis.opendocument.text&v5=application%2Fvnd.oasis.opendocument.presentation" +
                "&v6=application%2Fvnd.oasis.opendocument.spreadsheet&v7=application%2Fvnd.openxmlformats-officedocument.wordprocessingml.document" +
                "&v8=application%2Fvnd.ms-excel&v9=application%2Fvnd.openxmlformats-officedocument.spreadsheetml.sheet";

            using (var response = await bugsClient.GetAsync(bugListParameters))
            {
                using (var fileStream = new FileStream(bugListFileName, FileMode.Create))
                {
                    await response.Content.CopyToAsync(fileStream);
                }
            }
        }

        public static List<string> createBugList()
        {
            List<string> bugList = new List<string>();
            using (var reader = new StreamReader(bugListFileName))
            {
                reader.ReadLine();
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    bugList.Add(values[0]);
                }
            }
            return bugList;
        }

        public static async Task downloadAttachmentsOfBug(string bugId)
        {
            string attachmentQuery = "/rest/bug/" + bugId + "/attachment";

            Directory.CreateDirectory(Path.GetDirectoryName(@"download\"));

            using (var response = await bugsClient.GetAsync(attachmentQuery))
            {
                if (!response.IsSuccessStatusCode)
                {
                    Console.Error.WriteLine("Faild to download attachement of bug " + bugId);
                    return;
                }
                JObject jsonResponse = JObject.Parse(await response.Content.ReadAsStringAsync());
                JArray attachmentList = (JArray)jsonResponse["bugs"][bugId];
                foreach (var attachment in attachmentList)
                {
                    if (allowedMimeTypes.Contains(attachment["content_type"].ToString()))
                        File.WriteAllBytes(@"download\" + attachment["file_name"].ToString(), Convert.FromBase64String(attachment["data"].ToString()));
                }
            }
        }

        public static async Task downloadNBugsAttachment(List<string> bugIds, int numberOfBugs)
        {
            int _numberOfBugs = numberOfBugs == 0 ? bugIds.Count() : numberOfBugs;
            for (int i = 0; i < _numberOfBugs && i < bugIds.Count(); i++)
            {
                await downloadAttachmentsOfBug(bugIds[i]);
            }
        }
    }
}