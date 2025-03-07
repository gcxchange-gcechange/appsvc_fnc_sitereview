using Microsoft.Azure.WebJobs;
using PnP.Framework.Modernization.Functions;
using System;
using System.Threading.Tasks;

namespace SiteReviewProB
{
	public static class SiteReviewProB 
	{
		[FunctionName("SiteReviewProB")]
		public static async Task RunProB(
            [TimerTrigger("0 0 * * * *")] TimerInfo myTimer, ILogger log, ExecutionContext executionContext)
			//[HttpTrigger(AuthorizationLevel.Anonymous, "get")] HttpRequest req, ILogger log, ExecutionContext executionContext)
		{
            log.LogInformation($"SiteReviewProB timer trigger function executed at: {DateTime.Now}");
        }
		)
	}
}
