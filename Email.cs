using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Threading.Tasks;

namespace SiteReview
{
    public static class Email
    {
        public static async Task<List<Tuple<User, bool>>> InformSiteOwners(Site site, GraphServiceClient graphAPIAuth, ILogger log)
        {
            var results = new List<Tuple<User, bool>>();

            var groupQueryOptions = new List<QueryOption>()
            {
                new QueryOption("$search", "\"mailNickname:" + site.Name +"\"")
            };

            var groups = await graphAPIAuth.Groups
            .Request(groupQueryOptions)
            .Header("ConsistencyLevel", "eventual")
            .GetAsync();

            do
            {
                foreach (var group in groups)
                {
                    var owners = await graphAPIAuth.Groups[group.Id].Owners
                    .Request()
                    .GetAsync();

                    do
                    {
                        foreach (var owner in owners)
                        {
                            var user = await graphAPIAuth.Users[owner.Id]
                            .Request()
                            .Select("displayName,mail")
                            .GetAsync();

                            if (user != null)
                            {
                                var result = await SendEmail(user.DisplayName, user.Mail, log);
                                results.Add(new Tuple<User, bool>(user, result));
                            }
                        }
                    }
                    while (owners.NextPageRequest != null && (owners = await owners.NextPageRequest.GetAsync()).Count > 0);
                }
            }
            while (groups.NextPageRequest != null && (groups = await groups.NextPageRequest.GetAsync()).Count > 0);

            return results;
        }

        public static async Task<List<Tuple<User, bool>>> InformTeamOwners(Team team, GraphServiceClient graphAPIAuth, ILogger log)
        {
            var results = new List<Tuple<User, bool>>();

            // TODO: Get team owners and send email

            return results;
        }

        public static async Task<bool> SendEmail(string Username, string UserEmail, ILogger log)
        {
            var res = true;
            var auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            try
            {
                var message = new Message
                {
                    Subject = "English Subject | French Subject",
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Text,
                        Content = @$"
                        (La version française suit)

                        Dear { Username }, 

                        TODO: Write English copy 
                         
                        Regards, 
                        The GCX Team 

                        --------------------------------------

                        Bonjour { Username }, 

                        TODO: Write French copy
                         
                        Nous vous prions d’agréer l’expression de nos sentiments les meilleurs. 
                        Équipe de GCÉchange"
                    },
                    ToRecipients = new List<Recipient>()
                    {
                        new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = UserEmail
                            }
                        }
                    }
                };

                await graphAPIAuth.Users[Globals.emailSenderId]
                .SendMail(message, null)
                .Request()
                .PostAsync();

                log.LogInformation($"Email sent to {UserEmail}");
            }
            catch (Exception ex)
            {
                log.LogError($"Error sending email to {UserEmail}: {ex.Message}");
                res = false;
            }

            return res;
        }
    }
}
