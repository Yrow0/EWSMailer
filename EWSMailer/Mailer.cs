using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.Security.Authentication;

namespace EWSMailer
{
    public class Mailer
    {
        private ExchangeService exchangeService { get; set; }

        public Mailer(string username, string password, string domain, string? email)
        {
            ExchangeService exchangeService = new ExchangeService(ExchangeVersion.Exchange2010);
            exchangeService.Credentials = new NetworkCredential(username, password, domain);
            if (email != null ) 
            {
                exchangeService.AutodiscoverUrl(email);
            }
        }

        public void SetSender(string email)
        {
            exchangeService.AutodiscoverUrl(email);
        }

        public void SendMail(string subject, string body, List<string> recipients, List<string>? ccRecipients)
        {
            EmailMessage message = new EmailMessage(exchangeService);
            message.Subject = subject;
            message.Body = body;
            foreach (var recipient in recipients)
            {
                message.ToRecipients.Add(recipient);
            }
            if (ccRecipients != null)
            {
                foreach (var recipient in ccRecipients)
                {
                    message.CcRecipients.Add(recipient);
                }
            }
            message.SendAndSaveCopy();
        }
    }
}