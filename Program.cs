using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using timesheet.Models;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Net.Mail;

namespace timesheet
{
    public class Program
    {
        public static void Main()
        {
            var appSettings = ConfigurationManager.AppSettings;

            // email settings from config
            var emailSmtpHost = appSettings["EmailSmtpHost"];
            var emailSubject = appSettings["EmailSubject"];
            var emailBody = appSettings["EmailBody"];

            // template filenames from config
            var templateFileName = appSettings["TemplateFileName"];
            var outputFileName = appSettings["ExcelOutputFileName"];

            // variables to hold information from ConsultantInfo.txt
            var consultantName = appSettings["ConsultantName"];
            var consultantPhone = appSettings["ConsultantPhone"];
            var emailSendFrom = appSettings["EmailSendFrom"];
            var emailSendFromDisplay = appSettings["EmailSendFromDisplay"];
            var managerName = appSettings["ManagerName"];
            var emailSendTo = appSettings["EmailSendTo"];

            // is this running from exe outside of bin folder?
            var path = Path.GetFullPath(Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName)) + @"\";

#if DEBUG // if we're debugging, send email to consultant email
            path = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName), @"..\..\"));
            emailSendTo = emailSendFrom;
#endif

            var theDay = DateTime.Today.AddDays(-1);
#if DEBUG
            theDay = new DateTime(2017, 10, 16).AddDays(-1);
#endif

            var curDay = theDay.Day;
            var curMonth = theDay.Month;
            var curYear = theDay.Year;

            var templatePath = $"{path}{templateFileName}";
            object outputPath = $"{path}{outputFileName}";
            var newPath = outputPath.ToString().Replace("Timesheet.xlsx", $"{theDay.AddDays(1):MM.dd.yyyy} Timesheet.xlsx");

            object oMissing = System.Reflection.Missing.Value;

            try
            {
                var tr = new TimeRecords();
                var timesheet = new List<TimeEntry>();

                switch (curDay)
                {
                    case 15:
                        // get days 1 - 15 of current month
                        timesheet = tr.TimeEntries.Where(x => x.EntryDateTime >= new DateTime(curYear, curMonth, 1) &&
                                                              x.EntryDateTime <= new DateTime(curYear, curMonth, 16))
                            .OrderBy(x => x.EntryDateTime).ToList();
                        break;
                    case 28:
                    case 29:
                    case 30:
                    case 31:
                        // get days 16 - end of month from previous month
                        timesheet = tr.TimeEntries.Where(x => x.EntryDateTime >= new DateTime(curYear, curMonth, 16) &&
                                                              x.EntryDateTime <=
                                                              new DateTime(curYear, curMonth,
                                                                  DateTime.DaysInMonth(curYear, curMonth)))
                            .OrderBy(x => x.EntryDateTime).ToList();
                        break;
                }

                if (timesheet.Count > 0)
                {
                    var xlApp = new Application();
                    var xlWorkBook = xlApp.Workbooks.Open(templatePath, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    var xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.Item[1];


                    xlWorkSheet.Cells[4, 8].Value = consultantName;
                    xlWorkSheet.Cells[5, 8].Value = managerName;
                    xlWorkSheet.Cells[6, 8].Value = consultantPhone;
                    xlWorkSheet.Cells[7, 8].Value = emailSendFrom;

                    xlWorkSheet.Range[xlWorkSheet.Cells[12, 4], xlWorkSheet.Cells[28, 4]].ClearContents();
                    xlWorkSheet.Cells[8, 3].Value = timesheet.First().EntryDateTime.Date;

                    var i = 12;
                    foreach (var te in timesheet)
                    {
                        xlWorkSheet.Cells[i, 4].Value = te.Hours;
                        xlWorkSheet.Cells[i, 5].Value = te.Comment;
                        i++;
                    }

                    xlWorkSheet.Cells[31, 6].Value = theDay.AddDays(1);

                    if (File.Exists(newPath))
                    {
                        File.Delete(newPath);
                    }

                    xlWorkBook.SaveAs(newPath, XlFileFormat.xlOpenXMLWorkbook, oMissing, oMissing, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlUserResolution, true, oMissing, oMissing, oMissing);

                    xlWorkBook.Close(false, null, null);
                    xlApp.Quit();
                }

                var space = new string(Uri.EscapeUriString(" ").ToCharArray());

                var client = new SmtpClient(emailSmtpHost);

                using (MailMessage message = new MailMessage())
                {
                    message.IsBodyHtml = true;
                    message.BodyEncoding = System.Text.Encoding.UTF8;
                    message.Subject = emailSubject;
                    message.SubjectEncoding = System.Text.Encoding.UTF8;
                    message.Bcc.Add(new MailAddress(emailSendFrom));
                    message.From = new MailAddress(emailSendFrom, emailSendFromDisplay, System.Text.Encoding.UTF8);
                    message.To.Add(new MailAddress(emailSendTo));
                    message.Attachments.Add(new Attachment(newPath.ToString()));

                    client.Send(message);

                    message.Dispose();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(path);
                Console.WriteLine(ex.ToString());
                Console.Read();
            }
        }
    }
}
