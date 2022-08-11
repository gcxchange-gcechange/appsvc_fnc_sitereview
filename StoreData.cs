using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.WindowsAzure.Storage;
using Newtonsoft.Json;

namespace SiteReview
{
    public static class StoreData
    {
        static readonly string FileTitle = "deleteSites.json";

        public static async Task<bool> StoreSitesToDelete(ExecutionContext context, List<string> siteIds, string containerName, ILogger log)
        {
            CreateContainerIfNotExists(context, containerName);

            var storageAccount = GetCloudStorageAccount(context);
            var blobClient = storageAccount.CreateCloudBlobClient();
            var container = blobClient.GetContainerReference(containerName);

            var blob = container.GetBlockBlobReference(FileTitle);
            blob.Properties.ContentType = "application/json";

            var json = JsonConvert.SerializeObject(siteIds.ToArray());

            using (var ms = new MemoryStream())
            {
                LoadStreamWithJson(ms, json);
                await blob.UploadFromStreamAsync(ms);
            }

            log.LogInformation($"Blob {FileTitle} has been uploaded to container {container.Name}");

            await blob.SetPropertiesAsync();

            return true;
        }

        public static async Task<List<string>> GetSitesToDelete(ExecutionContext context, string containerName, ILogger log)
        {
            var storageAccount = GetCloudStorageAccount(context);
            var blobClient = storageAccount.CreateCloudBlobClient();
            var container = blobClient.GetContainerReference(containerName);

            var blob = container.GetBlockBlobReference(FileTitle);
            blob.Properties.ContentType = "application/json";

            var stream = await blob.OpenReadAsync();
            var streamReader = new StreamReader(stream);
            var siteIds = streamReader.ReadToEnd();

            var siteIdList = JsonConvert.DeserializeObject<List<string>>(siteIds);

            streamReader.Close();

            await blob.DeleteAsync();

            log.LogInformation($"Blob {FileTitle} has been deleted from container {container.Name}");

            return siteIdList;
        }

        private static async void CreateContainerIfNotExists(ExecutionContext executionContext, string ContainerName)
        {
            var storageAccount = GetCloudStorageAccount(executionContext);
            var blobClient = storageAccount.CreateCloudBlobClient();
            string[] containers = new string[] { ContainerName };

            foreach (var item in containers)
            {
                var blobContainer = blobClient.GetContainerReference(item);
                await blobContainer.CreateIfNotExistsAsync();
            }
        }

        private static CloudStorageAccount GetCloudStorageAccount(ExecutionContext executionContext)
        {
            var config = new ConfigurationBuilder()
                            .SetBasePath(executionContext.FunctionAppDirectory)
                            .AddJsonFile("local.settings.json", true, true)
                            .AddEnvironmentVariables().Build();
            var storageAccount = CloudStorageAccount.Parse(config["AzureWebJobsStorage"]);
            return storageAccount;
        }

        private static void LoadStreamWithJson(Stream ms, object obj)
        {
            var writer = new StreamWriter(ms);
            writer.Write(obj);
            writer.Flush();
            ms.Position = 0;
        }
    }
}
