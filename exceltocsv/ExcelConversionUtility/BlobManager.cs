﻿using System;
using System.IO;
using System.Threading.Tasks;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using System.Collections.Generic;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.WindowsAzure.Storage;

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
        

        public async Task<List<BlobOutput>> Download(string containerName,string name)
        {
            var downloadedData = new List<BlobOutput>();
            try
            {
                
                //---------------------------------------------------------------------------------------
                CloudStorageAccount sourceAccount = CloudStorageAccount.Parse(Constants.ConnectionString);
                CloudBlobClient sourceClient = sourceAccount.CreateCloudBlobClient();
                CloudBlobClient destClient = sourceAccount.CreateCloudBlobClient();
                // To download the contents
                CloudBlobContainer sourceBlobContainer = sourceClient.GetContainerReference(containerName);
                ICloudBlob sourceBlob = await sourceBlobContainer.GetBlobReferenceFromServerAsync(name);
                
                await sourceBlob.DownloadToFileAsync(name, System.IO.FileMode.Create);

                // To Upload the blob contents to destination container
                CloudBlobContainer destBlobContainer = destClient.GetContainerReference(Constants.ExcelCopyContainer);
                string newFilename = $"copy_{name}";
                CloudBlockBlob destBlob = destBlobContainer.GetBlockBlobReference(newFilename);
                await destBlob.UploadFromFileAsync(name);
                //-------------------------------------------------------------------------------------

                // Create service and container client for blob
                BlobContainerClient blobContainerClient = _blobServiceClient.GetBlobContainerClient(Constants.ExcelCopyContainer);
                // Download the blob's contents and save it to a file
                BlobClient blobClient = blobContainerClient.GetBlobClient(newFilename);
                BlobDownloadInfo downloadedInfo = await blobClient.DownloadAsync();

                downloadedData.Add(new BlobOutput { BlobName = name, BlobContent = downloadedInfo.Content });
                
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return downloadedData;
        }
    }
}
