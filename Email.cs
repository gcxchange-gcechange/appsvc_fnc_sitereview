using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SiteReview
{
    public static class Email
    {
        public static async Task<bool> SendReportEmail(string[] userEmails, SiteReport report, GraphServiceClient graphAPIAuth, ILogger log)
        {
            var uniqueListOfSites = report.GetUniqueListSites();
            List<Task> emailTasks = new List<Task>();

            if (uniqueListOfSites.Count == 0)
            {
                foreach (var email in userEmails)
                {
                    emailTasks.Add(SendEmail(
                        email,
                        "Site Review Report",
                        "Greetings,<br><br>We found <b>0</b> sites flagged for review.<br>Woohoo! &#128513;<br><br>" +
                        "Regards,<br>The GCX Team",
                        BodyType.Html,
                        graphAPIAuth,
                        log
                    ));
                }
                return true;
            } 

            foreach (var email in userEmails)
            {
                emailTasks.Add(SendEmail(
                    email,
                    "Site Review Report",
                    $"Greetings,<br><br>We found <b>{uniqueListOfSites.Count + (uniqueListOfSites.Count == 1 ? $" site" : " unique sites")}</b> flagged for review.<br><br>" +
                    "The count for each category violation is as follows...<br>" +
                    "<i style=\"font-size: x-small;\">(Sites can be flagged for multiple category violations and may exceed the unique sites total when tallied.)</i><ul>" +
                    $"<li>{report.WarningSites.Count} sites inactive for {Globals.inactiveDaysWarn} days or more, but less than {Globals.inactiveDaysDelete} days.</li>" +
                    $"<li>{report.DeleteSites.Count} sites inactive for {Globals.inactiveDaysDelete} days or more.</li>" +
                    $"<li>{report.HubAssociationSites.Count} sites not associated with the hub {Globals.hubId}.</li>" +
                    $"<li>{report.NoOwnerSites.Count} sites with less than {Globals.minSiteOwners} owners.</li>" +
                    $"<li>{report.ClassificationSites.Count} sites with no classification label.</li>" +
                    $"<li>{report.PrivacySettingSites.Count} sites without the \"{Globals.expectedPrivacySetting}\" privacy setting.</li>" +
                    $"<li>{report.StorageThresholdSites.Count} sites with {Globals.storageThreshold}% storage used or more.</li></ul>" +
                    "Please note that <b>bolded text</b> indicates a violation of our policy.<br><br><ol>" +
                    string.Join(
                        "<hr>",
                        uniqueListOfSites.Select(item =>
                            "<li>" +
                            $"Site: <a href='{item.SiteUrl}' target='_blank'>{item.SiteDisplayName}</a><br>" +
                            $"In Hub: {(item.InHub == false ? "<b>" + item.InHub + "</b>" : item.InHub)}<br>" +
                            $"Classification: {(item.AssignedLabels.Any() == false ? "<b>None</b>" : string.Join(", ", item.AssignedLabels.Select(m => m.DisplayName)))}<br>" +
                            $"Privacy: {(item.PrivacySetting != "Private" ? "<b>" + (item.PrivacySetting ?? "null") + "</b>" : item.PrivacySetting)}<br>" +
                            $"Owners: {(item.SiteOwners.Count < Globals.minSiteOwners ? "<b>" + item.SiteOwners.Count + "</b>" : item.SiteOwners.Count)}<br>" +
                            $"Inactive: {(item.InactiveDays >= Globals.inactiveDaysWarn ? "<b>" + item.InactiveDays + " days</b>" : item.InactiveDays + " days") }<br>" +
                            $"Storage Used: {(((double)item.StorageUsed / (double)item.StorageAllocated) * (double)100 >= Globals.storageThreshold ? "<b>" + (((double)item.StorageUsed / (double)item.StorageAllocated) * (double)100).ToString("F6") + "%</b>" : (((double)item.StorageUsed / (double)item.StorageAllocated) * (double)100).ToString("F6") + "%")}<br>" +
                            $"Owner Emails: {(item.SiteOwners.Any() ? string.Join(", ", item.SiteOwners.Select(m => m.Mail)) : "<b>None</b>")}" +
                            "</li>"
                        )
                    ) +
                    "</ol><br><br>Regards,<br>The GCX Team",
                    BodyType.Html,
                    graphAPIAuth,
                    log
                ));
            }

            await Task.WhenAll(emailTasks);
            return true;
        }
        public static async Task<bool> SendWarningEmail(string userEmail, string siteUrl, GraphServiceClient graphAPIAuth, ILogger log)
        {
            return await SendEmail(
                userEmail,
                "An important message from GCXchange | Un message important de GCÉchange",
                $@"
                (La version française suit)

                Dear GCXchange community owner,

                Thank you for your interest and participation as a community owner on the GCXchange platform.

                Communities continue to be our most popular feature, and we look forward to seeing new users joining various communities on the platform as activity and engagement grows.

                Due to the popularity of communities, we are continuously ensuring that users can easily browse and find other engaged members that share their interests. Therefore, we permanently remove any community that has been inactive for 120 days. The inactive status changes when you or another registered user views the community page, or once a content update occurs on the site such as editing the text on your splash page or uploading a file to your site.

                Today, we wish to inform you that your community has been inactive for 60 days. You must resume activity within your community by engaging with others, otherwise the site and all materials will be deleted in 60 more days. If you feel that your site is no longer needed or that it has fulfilled its purpose, please let us know and we can supply instructions on how to remove your community presence from the GCXchange platform.

                Please let us know if you have any questions or concerns, and once again thank you for being a valued member of GCXchange.

                Regards, 
                The GCX Team
                
                --------------------------------------

                Cher ou chère responsable de collectivité de GCÉchange,

                Merci de l’intérêt que vous portez et de votre participation à titre de responsable d’une collectivité sur la plateforme GCÉchange.

                Les collectivités demeurent notre fonctionnalité la plus populaire et nous sommes impatients de voir de nouveaux utilisateurs se joindre à différentes collectivités sur la plateforme au fur et à mesure de la croissance de l’activité et de la mobilisation.

                En raison de la popularité des collectivités, nous veillons continuellement à ce que les utilisateurs puissent naviguer facilement et trouver d’autres membres mobilisés qui partagent leurs intérêts. Par conséquent, nous éliminons de manière permanente toute collectivité qui a été inactive pendant 120 jours. L’état inactif change lorsque vous ou un autre utilisateur inscrit visionnez la page de la collectivité ou lorsqu’une mise à jour est apportée au contenu dans le site, par exemple lorsque vous modifiez le texte de votre page d’entrée ou lorsque vous téléversez un fichier dans votre site.

                Aujourd’hui, nous souhaitons vous informer que votre collectivité est inactive depuis 60 jours. Vous devez reprendre l’activité dans votre collectivité en discutant avec d’autres, autrement le site et tous les documents y afférents seront supprimés dans 60 jours. Si vous estimez que votre site n’est plus nécessaire ou qu’il a rempli son objectif, veuillez nous en informer et nous vous fournirons des instructions sur la façon de supprimer la présence de votre collectivité de la plateforme de GCÉchange.

                Veuillez nous indiquer si vous avez des questions ou des préoccupations et, une fois de plus, nous vous remercions d’être un membre important de GCÉchange. 

                Nous vous prions d’agréer l’expression de nos sentiments les meilleurs. 
                Équipe de GCÉchange",
                BodyType.Text,
                graphAPIAuth,
                log
            );
        }

        public static async Task<bool> SendDeleteEmail(string userEmail, string siteUrl, GraphServiceClient graphAPIAuth, ILogger log)
        {
            return await SendEmail(
                userEmail,
                "An important message from GCXchange | Un message important de GCÉchange",
                $@"
                (La version française suit)

                Dear GCXchange community owner, 

                You are receiving this communication today as a follow-up to the 60-day warning for inactive communities e-mail previously sent to you.

                Due to the popularity of communities, we are continuously ensuring that users can easily browse and find other engaged members that share their interests. Therefore, we permanently remove any community that has been inactive for 120 days.

                Today, we wish to inform you that your community has been inactive for 120 days. Therefore, the GCXchange Team has proceeded to remove your community presence from the platform.

                If in the future you would like to create a new community for interdepartmental collaboration, then please return to the Community section of the GCXchange website. We will be happy to welcome you back!

                Please let us know if you have any questions or concerns, and once again thank you for being a valued member of GCXchange.
                 
                Regards, 
                The GCX Team 

                --------------------------------------

                Cher ou chère responsable de la collectivité de GCÉchange, 

                Vous recevez aujourd’hui la présente communication qui fait suite au courriel de mise en garde qui vous a été envoyé précédemment au sujet de l'avis concernant les collectivités inactives pendant 60 jours.

                En raison de la popularité des collectivités, nous veillons continuellement à ce que les utilisateurs puissent naviguer facilement et trouver d’autres membres mobilisés qui partagent leurs intérêts. Par conséquent, nous éliminons de manière permanente toute collectivité qui a été inactive pendant 120 jours.

                Aujourd’hui, nous souhaitons vous informer que votre collectivité a été inactive pendant 120 jours. L’équipe de GCÉchange a donc supprimé votre collectivité de la plateforme. 

                Si, à l’avenir, vous souhaitez créer une nouvelle collectivité pour une collaboration interministérielle, veuillez retourner à la section Collectivités du site GCÉchange. Nous serons heureux de vous accueillir à nouveau. 

                Veuillez nous indiquer si vous avez des questions ou des préoccupations et, une fois de plus, nous vous remercions d’être un membre important de GCÉchange.
                 
                Nous vous prions d’agréer l’expression de nos sentiments les meilleurs. 
                Équipe de GCÉchange",
                BodyType.Text,
                graphAPIAuth,
                log
            );
        }

        private static async Task<bool> SendEmail(string userEmail, string emailSubject, string emailBody, BodyType bodyType, GraphServiceClient graphAPIAuth, ILogger log)
        {
            var res = true;

            try
            {
                var message = new Message
                {
                    Subject = emailSubject,
                    Body = new ItemBody
                    {
                        ContentType = bodyType,
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

                var requestBody = new SendMailPostRequestBody
                {
                    Message = message,
                    SaveToSentItems = true
                };

                await graphAPIAuth.Users[Globals.emailUserName]
                .SendMail
                .PostAsync(requestBody);

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
