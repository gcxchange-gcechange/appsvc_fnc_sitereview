using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SiteReview
{
    public static class Email
    {
        public static async Task<bool> SendWarningEmail(string userEmail, string siteUrl, ILogger log)
        {
            return await SendEmail(
                userEmail,
                "An important message from GCXchange | Un message important de GCÉchange",
                $@"
                (La version française suit)

                Hello,

                TODO: Write warning email (en)

                Regards, 
                The GCX Team
                
                --------------------------------------

                Bonjour,

                TODO: Write warning email (fr)

                Nous vous prions d’agréer l’expression de nos sentiments les meilleurs. 
                Équipe de GCÉchange", 
                log
            );
        }

        public static async Task<bool> SendDeleteEmail(string userEmail, string siteUrl, ILogger log)
        {
            return await SendEmail(
                userEmail,
                "An important message from GCXchange | Un message important de GCÉchange",
                $@"
                (La version française suit)

                Hello, 

                TODO: Write English copy (en)
                 
                Regards, 
                The GCX Team 

                --------------------------------------

                Bonjour, 

                TODO: Write French copy (fr)
                 
                Nous vous prions d’agréer l’expression de nos sentiments les meilleurs. 
                Équipe de GCÉchange", 
                log
            );
        }

        private static async Task<bool> SendEmail(string userEmail, string emailSubject, string emailBody, ILogger log)
        {
            var res = true;
            var auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            try
            {
                var message = new Message
                {
                    Subject = emailSubject,
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Text,
                        Content = emailBody
                    },
                    ToRecipients = new List<Recipient>()
                    {
                        new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = userEmail
                            }
                        }
                    }
                };

                await graphAPIAuth.Users[Globals.emailSenderId]
                .SendMail(message, null)
                .Request()
                .PostAsync();

                log.LogInformation($"Email sent to {userEmail}");
            }
            catch (Exception ex)
            {
                log.LogError($"Error sending email to {userEmail}: {ex.Message}");
                res = false;
            }

            return res;
        }
    }
}
