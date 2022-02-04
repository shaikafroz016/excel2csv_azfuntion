using System;
using System.Threading.Tasks;

namespace ExcelConversionUtility
{
    /// <summary>
    /// This class is responsible for calling the utility methods, required to get Excel files from given container, convert them to CSV format and upload them back to specified blob container. 
    /// </summary>
    public class ExcelConversionUtility
    {
        public async Task Process(string name)
        {
            try
            {

                var blobManager = new BlobManager(Constants.ConnectionString);
                // download the blobs from given blob container 
                var results = await blobManager.Download(Constants.ExcelContainerName,name);
                // convert stremed excel content to csv and send back the content in the form of stream
                var blobs = ExcelToCSVConvertor.Convert(results);

                // upload the converted results back onto supplied container name
                await blobManager.Upload(Constants.CSVContainerName, blobs,name);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
