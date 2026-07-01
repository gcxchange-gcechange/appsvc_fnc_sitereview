using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.WindowsAzure.Storage;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace SiteReview
{
    public static class StoreData
    {
        static readonly string FileTitle = "deleteSites.json";

        public static async Task<bool> StoreSitesToDelete(List<string> siteIds, string containerName, ILogger log)
        {
            _ = CreateContainerIfNotExists(containerName);

            var storageAccount = GetCloudStorageAccount();
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

        public static async Task<List<string>> GetSitesToDelete(string containerName, ILogger log)
        {
            var storageAccount = GetCloudStorageAccount();
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

        private static async Task CreateContainerIfNotExists(string ContainerName)
        {
            var storageAccount = GetCloudStorageAccount();
            var blobClient = storageAccount.CreateCloudBlobClient();
            string[] containers = new string[] { ContainerName };

            foreach (var item in containers)
            {
                var blobContainer = blobClient.GetContainerReference(item);
                await blobContainer.CreateIfNotExistsAsync();
            }
        }

        private static CloudStorageAccount GetCloudStorageAccount()
        {
            var basePath = Environment.GetEnvironmentVariable("AzureFunctionsJobRoot") ?? Directory.GetCurrentDirectory();

            var config = new ConfigurationBuilder()
                            .SetBasePath(basePath)
                            .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
                            .AddEnvironmentVariables()
                            .Build();

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
