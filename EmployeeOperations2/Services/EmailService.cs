using System;
using System.Configuration;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;

namespace EmployeeOperations2.Services
{
    public class EmailService
    {
        private readonly string _smtpServer;
        private readonly int _smtpPort;
        private readonly string _smtpUsername;
        private readonly string _smtpPassword;
        private readonly string _emailFromAddress;

        public EmailService()
        {
            _smtpServer = ConfigurationManager.AppSettings["SmtpServer"];
            _smtpPort = int.Parse(ConfigurationManager.AppSettings["SmtpPort"]);
            _smtpUsername = ConfigurationManager.AppSettings["SmtpUsername"];
            _smtpPassword = ConfigurationManager.AppSettings["SmtpPassword"];
            _emailFromAddress = ConfigurationManager.AppSettings["EmailFromAddress"];
        }

        public void SendEmail(string emailTo, string emailSubject, string emailBody, string ReceiverName)
        {
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("Sharath", _emailFromAddress ));
            message.To.Add(new MailboxAddress(ReceiverName, emailTo));
            message.Subject = emailSubject;

            var textPart = new TextPart("plain")
            {
                Text = emailBody
            };

            message.Body = textPart;

            using (var client = new SmtpClient())
            {
                client.Connect(_smtpServer, _smtpPort, SecureSocketOptions.StartTls);
                client.Authenticate(_smtpUsername, _smtpPassword);
                client.Send(message);
                client.Disconnect(true);
            }
        }
    }
}




