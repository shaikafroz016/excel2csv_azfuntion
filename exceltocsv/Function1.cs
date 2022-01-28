using System;
using System.IO;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;

namespace exceltocsv
{
    [StorageAccount("constr")]
    public static class Function1
    {
        [FunctionName("Function1")]
        public static void Run([BlobTrigger("excelcontainer/{name}")]Stream myBlob, string name, ILogger log)
        {
            log.LogInformation($"C# Blob trigger function Processed blob\n Name:{name} \n Size: {myBlob.Length} Bytes");
            //checking if length is greater then 200kb then it will accept
            if (myBlob.Length >= 200000)
            {
                log.LogInformation("Starting conversion process...");
                ExcelConversionUtility.ExcelConversionUtility excelConversionUtility = new ExcelConversionUtility.ExcelConversionUtility();
                excelConversionUtility.Process().GetAwaiter().GetResult();
                log.LogInformation("Conversion process completed.");
            }
            else
            {
                log.LogInformation("Upload some large file");
            }
            
        }
    }
}
