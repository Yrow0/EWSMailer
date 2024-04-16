using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWSManager.Models
{
    public  class EWSMail
    {
        public string subject {  get; set; }
        public string body { get; set; }
        public List<string> recipients { get; set; }
        public List<string>? ccRecipients { get; set; }

        public void Send(EWSManager eWSManager)
        {
            EmailMessage message = new EmailMessage(eWSManager.exchangeService);
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
            message.Send();

        }public void SendAndSave(EWSManager eWSManager)
        {
            EmailMessage message = new EmailMessage(eWSManager.exchangeService);
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
