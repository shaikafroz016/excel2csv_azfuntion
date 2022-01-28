using System;
using System.IO;
using System.Threading.Tasks;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using System.Collections.Generic;

namespace ExcelConversionUtility
{
    public class BlobManager
    {
        private string _connectionString;
        private BlobServiceClient _blobServiceClient;

        public BlobManager(string connectionString)
        {
            _connectionString = connectionString;
            _blobServiceClient = new BlobServiceClient(_connectionString);
        }

        public async Task Upload(string containerName, List<BlobInput> inputs)
        {
            try
            {
                // Create service and container client for blob
                BlobContainerClient blobContainerClient = _blobServiceClient.GetBlobContainerClient(containerName);

                foreach (BlobInput item in inputs)
                {
                    // Get a reference to a blob and upload
                    BlobClient blobClient = blobContainerClient.GetBlobClient(item.BlobName.ToString());

                    using (var ms = new MemoryStream(item.BlobContent))
                    {
                        await blobClient.UploadAsync(ms, overwrite: true);
                    }

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public async Task<List<BlobOutput>> Download(string containerName)
        {
            var downloadedData = new List<BlobOutput>();
            try
            {
                // Create service and container client for blob
                BlobContainerClient blobContainerClient = _blobServiceClient.GetBlobContainerClient(containerName);

                // List all blobs in the container
                await foreach (BlobItem item in blobContainerClient.GetBlobsAsync())
                {
                    // Download the blob's contents and save it to a file
                    BlobClient blobClient = blobContainerClient.GetBlobClient(item.Name);
                    BlobDownloadInfo downloadedInfo = await blobClient.DownloadAsync();
                    downloadedData.Add(new BlobOutput { BlobName = item.Name, BlobContent = downloadedInfo.Content });
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return downloadedData;
        }
    }
}
