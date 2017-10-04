﻿using System;
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
            try
            {
                var theDay = DateTime.Today.AddDays(-1);
#if DEBUG
                theDay = new DateTime(2017, 10, 1).AddDays(-1);
#endif

                var curDay = theDay.Day;
                var curMonth = theDay.Month;
                var curYear = theDay.Year;

                var appSettings = ConfigurationManager.AppSettings;

                var path = appSettings["TemplatePath"];
                object outputPath = appSettings["ExcelOutputPath"];
                var newPath = outputPath.ToString().Replace("Timesheet.xlsx", $"{theDay.AddDays(1):MM.dd.yyyy} Timesheet.xlsx");
                object oMissing = System.Reflection.Missing.Value;

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
                    var xlWorkBook = xlApp.Workbooks.Open(path, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    var xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.Item[1];


                    xlWorkSheet.Cells[4, 8].Value = appSettings["ConsultantName"];
                    xlWorkSheet.Cells[5, 8].Value = appSettings["ManagerName"];
                    xlWorkSheet.Cells[6, 8].Value = appSettings["ConsultantPhone"];
                    xlWorkSheet.Cells[7, 8].Value = appSettings["ConsultantEmail"];

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

                var emailSendTo = appSettings["EmailTo"];
                var emailSubject = appSettings["EmailSubject"];
                const string emailSmtpHost = "SMTPR.illinois.gov";
                const string emailSendFrom = "alexis.atchison@illinois.gov";
                const string emailSendFromDisplay = "Atchison, Alexis";
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
                    message.Body += $"<span style='font-size:11pt;font-family:Calibri'>{appSettings["EmailBody"]} {newPath.ToString().Replace(" ", space)}</span>";

                    client.Send(message);

                    message.Dispose();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
}