using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using excel = Microsoft.Office.Interop.Excel;
using powerPoint = Microsoft.Office.Interop.PowerPoint;
using word = Microsoft.Office.Interop.Word;

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

        public static HttpClient coolClient = new HttpClient() { BaseAddress = new Uri("https://staging.eu.collaboraonline.com/cool/convert-to") };

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
    }
}