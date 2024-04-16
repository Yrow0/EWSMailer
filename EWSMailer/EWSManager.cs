using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.Security.Authentication;

namespace EWSManager
{
    public class EWSManager
    {
        public ExchangeService exchangeService { get; set; }

        public EWSManager(string username, string password, string domain, string? email)
        {
            ExchangeService exchangeService = new ExchangeService(ExchangeVersion.Exchange2010);
            exchangeService.Credentials = new NetworkCredential(username, password, domain);
            if (email != null ) 
            {
                exchangeService.AutodiscoverUrl(email);
            }
        }

        public void SetMail(string email)
        {
            exchangeService.AutodiscoverUrl(email);
        }
    }
}