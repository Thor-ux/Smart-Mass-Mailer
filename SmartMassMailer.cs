using System;
using System.Net;
using System.Net.Mail;
using Microsoft.Extensions.Configuration;
using System.IO;
using System.Text;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Threading;
using System.Linq;

class SmartMassMailer
{
    static IConfigurationRoot Configuration { get; set; }
    static SmtpClient SmtpClient { get; set; }

    static SmartMassMailer()
    {
        var builder = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json");
        Configuration = builder.Build();

        SmtpClient = new SmtpClient(Configuration["smtpSettings:host"])
        {
            UseDefaultCredentials = false,
            Credentials = new NetworkCredential(
                Configuration["smtpSettings:username"],
                Configuration["smtpSettings:password"])
        };

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    static void SendEmail(string recipientEmail, string subject, string bodyHtml)
    {
        var msg = new MailMessage
        {
            From = new MailAddress(
                Configuration["fromEmail"],
                Configuration["fromName"]),
            IsBodyHtml = true,
            Body = bodyHtml,
            BodyEncoding = Encoding.UTF8,
            SubjectEncoding = Encoding.UTF8,
            HeadersEncoding = Encoding.UTF8,
            Subject = subject
        };

        msg.To.Add(recipientEmail);
        SmtpClient.Send(msg);
    }

    static void Main()
    {
        int delayBetweenEmails = int.Parse(Configuration["delayBetweenEmailsMilliseconds"]);
        int startRow = int.Parse(Configuration["startRow"]);
        if (startRow < 2)
            startRow = 2;

        var excelFile = new FileInfo(Configuration["recipientsExcelFile"]);
        if (!excelFile.Exists)
            throw new FileNotFoundException(
                $"The file '{excelFile.FullName}' does not exist.",
                excelFile.FullName);

        using (var package = new ExcelPackage(excelFile))
        {
            var workSheet = package.Workbook.Worksheets.FirstOrDefault();
            if (workSheet == null)
                throw new Exception("The Excel file contains no worksheets.");

            if (workSheet.Dimension == null)
                throw new Exception("The worksheet is empty.");

            int mailsLeft = workSheet.Dimension.End.Row - startRow + 1;
            TimeSpan timeLeft = TimeSpan.FromMilliseconds(
                mailsLeft * (delayBetweenEmails + 100));

            Console.WriteLine("Estimated time: {0} hours", timeLeft.TotalHours);

            var columnIndexByName = new Dictionary<string, int>();

            for (int col = 1; col <= workSheet.Dimension.End.Column; col++)
            {
                var cellValue = workSheet.Cells[1, col].Value;
                if (cellValue == null)
                    continue;

                string columnName = cellValue.ToString().ToLower();
                if (!columnIndexByName.ContainsKey(columnName))
                    columnIndexByName[columnName] = col;
            }

            if (!columnIndexByName.ContainsKey("email") ||
                !columnIndexByName.ContainsKey("name"))
            {
                throw new Exception("Excel file must contain 'email' and 'name' columns.");
            }

            for (int row = startRow; row <= workSheet.Dimension.End.Row; row++)
            {
                string email = workSheet.Cells[row, columnIndexByName["email"]]
                    .Value?.ToString();

                string personName = workSheet.Cells[row, columnIndexByName["name"]]
                    .Value?.ToString();

                string subject = Configuration["emailSubject"];
                string bodyHtml = File.ReadAllText(
                    Configuration["emailHtmlTemplate"], Encoding.UTF8);

                bodyHtml = bodyHtml.Replace("[name]", personName);

                Console.Write(
                    $"[{row - 1} of {workSheet.Dimension.End.Row - 1}] " +
                    $"Sending email to: {email} (Name = {personName}) ... ");

                SendEmail(email, subject, bodyHtml);
                Console.WriteLine("Done.");

                Thread.Sleep(delayBetweenEmails);
            }

            Console.WriteLine();
            Console.WriteLine("Done.");
        }
    }
}
