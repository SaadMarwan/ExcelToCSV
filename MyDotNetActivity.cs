// Comment these lines if using VS 2017
using System.IO;
using System.Linq;
// --------------------

// Comment these lines if using <= VS 2015
using System;
using System.Collections.Generic;
// ---------------------

using Microsoft.Azure.Management.DataFactories.Models;
using Microsoft.Azure.Management.DataFactories.Runtime;

using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;

namespace MyDotNetActivityNS
{
    public class MyDotNetActivity : IDotNetActivity
    {
        public const int FirstRow = 1;
        public const int LastRow = 372;
        public const string SheetName = "Sheet Name";
        /// <summary>
        /// Execute method is the only method of IDotNetActivity interface you must implement.
        /// In this sample, the method invokes the Calculate method to perform the core logic.  
        /// </summary>

        public IDictionary<string, string> Execute(
            IEnumerable<LinkedService> linkedServices,
            IEnumerable<Dataset> datasets,
            Activity activity,
            IActivityLogger logger)
        {
            // get extended properties defined in activity JSON definition
            // (for example: SliceStart)
            DotNetActivity dotNetActivity = (DotNetActivity)activity.TypeProperties;
            string sliceStartString = dotNetActivity.ExtendedProperties["SliceStart"];


            // linked service for input and output data stores
            // in this example, same storage is used for both input/output
            AzureStorageLinkedService inputLinkedService;

            // get the input dataset
            Dataset inputDataset = datasets.Single(dataset => dataset.Name == activity.Inputs.Single().Name);

            // declare variables to hold type properties of input/output datasets
            AzureBlobDataset inputTypeProperties, outputTypeProperties;

            // get type properties from the dataset object
            inputTypeProperties = inputDataset.Properties.TypeProperties as AzureBlobDataset;

            // log linked services passed in linkedServices parameter
            // you will see two linked services of type: AzureStorageLinkedService
            // one for input dataset and the other for output dataset 
            foreach (LinkedService ls in linkedServices)
                logger.Write("linkedService.Name {0}", ls.Name);

            // get the first Azure Storate linked service from linkedServices object
            // using First method instead of Single since we are using the same
            // Azure Storage linked service for input and output.
            inputLinkedService = linkedServices.First(
                linkedService =>
                linkedService.Name ==
                inputDataset.Properties.LinkedServiceName).Properties.TypeProperties
                as AzureStorageLinkedService;

            // get the connection string in the linked service
            string connectionString = inputLinkedService.ConnectionString;

            // get the folder path from the input dataset definition
            string folderPath = GetFolderPath(inputDataset);
            string output = string.Empty; // for use later.

            // create storage client for input. Pass the connection string.
            CloudStorageAccount inputStorageAccount = CloudStorageAccount.Parse(connectionString);
            CloudBlobClient inputClient = inputStorageAccount.CreateCloudBlobClient();

            // initialize the continuation token before using it in the do-while loop.
            BlobContinuationToken continuationToken = null;
            do
            {   // get the list of input blobs from the input storage client object.
                BlobResultSegment blobList = inputClient.ListBlobsSegmented(folderPath,
                                         true,
                                         BlobListingDetails.Metadata,
                                         null,
                                         continuationToken,
                                         null,
                                         null);

                // Calculate method performs the core logic
                output = Calculate(blobList, logger, folderPath, ref continuationToken);

            } while (continuationToken != null);

            // get the output dataset using the name of the dataset matched to a name in the Activity output collection.
            Dataset outputDataset = datasets.Single(dataset => dataset.Name == activity.Outputs.Single().Name);

            // get type properties for the output dataset
            outputTypeProperties = outputDataset.Properties.TypeProperties as AzureBlobDataset;

            // get the folder path from the output dataset definition
            folderPath = GetFolderPath(outputDataset);

            // log the output folder path   
            logger.Write("Writing blob to the folder: {0}", folderPath);

            // create a storage object for the output blob.
            CloudStorageAccount outputStorageAccount = CloudStorageAccount.Parse(connectionString);
            // write the name of the file.
            Uri outputBlobUri = new Uri(outputStorageAccount.BlobEndpoint, folderPath + "/" + GetFileName(outputDataset));

            // log the output file name
            logger.Write("output blob URI: {0}", outputBlobUri.ToString());

            // create a blob and upload the output text.
            CloudBlockBlob outputBlob = new CloudBlockBlob(outputBlobUri, outputStorageAccount.Credentials);
            logger.Write("Writing {0} to the output blob", output);
            outputBlob.UploadText(output);

            // The dictionary can be used to chain custom activities together in the future.
            // This feature is not implemented yet, so just return an empty dictionary.  

            return new Dictionary<string, string>();
        }

        /// <summary>
        /// Gets the folderPath value from the input/output dataset.
        /// </summary>

        private static string GetFolderPath(Dataset dataArtifact)
        {
            if (dataArtifact == null || dataArtifact.Properties == null)
            {
                return null;
            }

            // get type properties of the dataset   
            AzureBlobDataset blobDataset = dataArtifact.Properties.TypeProperties as AzureBlobDataset;
            if (blobDataset == null)
            {
                return null;
            }

            // return the folder path found in the type properties
            return blobDataset.FolderPath;
        }

        /// <summary>
        /// Gets the fileName value from the input/output dataset.   
        /// </summary>

        private static string GetFileName(Dataset dataArtifact)
        {
            if (dataArtifact == null || dataArtifact.Properties == null)
            {
                return null;
            }

            // get type properties of the dataset
            AzureBlobDataset blobDataset = dataArtifact.Properties.TypeProperties as AzureBlobDataset;
            if (blobDataset == null)
            {
                return null;
            }

            // return the blob/file name in the type properties
            return blobDataset.FileName;
        }

        /// <summary>
        /// Iterates through each blob (file) in the folder,
        /// and prepares the output text that is written to the output blob.
        /// </summary>

        public static string Calculate(BlobResultSegment Bresult, IActivityLogger logger, string folderPath, ref BlobContinuationToken token)
        {
            {
                string output = string.Empty;
                logger.Write("number of blobs found: {0}", Bresult.Results.Count<IListBlobItem>());
                foreach (IListBlobItem listBlobItem in Bresult.Results)
                {
                    CloudBlockBlob inputBlob = listBlobItem as CloudBlockBlob;
                    if ((inputBlob != null) && (inputBlob.Name.IndexOf("$$$.$$$") == -1))
                    {
                        var filename = Environment.CurrentDirectory + "\\Book1.xlsx";
                        inputBlob.DownloadToFile(filename, FileMode.Create);


                        List<string> allRows = new List<string>();
                        int j = 0;

                        using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(filename, true))
                        {
                            WorkbookPart workbookPart = myDoc.WorkbookPart;
                            IEnumerable<Sheet> Sheets = myDoc.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == SheetName);
                            if (Sheets.Count() == 0)
                            {
                                throw new ArgumentException("sheetName");
                            }
                            string relationshipId = Sheets.First().Id.Value;
                            WorksheetPart worksheetPart = (WorksheetPart)myDoc.WorkbookPart.GetPartById(relationshipId);
                            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                            int i = 1;
                            string value;
                            string newLine;
                            logger.Write("Sheet is opened. ");
                            foreach (Row r in sheetData.Elements<Row>())
                            {
                                newLine = "";
                                List<string> currentRow = new List<string>();
                                if (i < FirstRow | i > LastRow) { i++; continue; }
                                else
                                {
                                    i++;
                                    foreach (Cell c in r.Elements<Cell>())
                                    {

                                        if (c != null)
                                        {

                                            value = c.InnerText;
                                            if (value == "")
                                            {
                                                value = "Null";
                                                currentRow.Add(value);
                                                j = j + 1;
                                                continue;
                                            }
                                            value = c.CellValue.Text;
                                            if (c.DataType != null)
                                            {
                                                switch (c.DataType.Value)
                                                {
                                                    case CellValues.SharedString:
                                                        var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                                                        if (stringTable != null)
                                                        {
                                                            value = stringTable.SharedStringTable.
                                                            ElementAt(int.Parse(value)).InnerText;

                                                        }

                                                        break;
                                                    case CellValues.Boolean:
                                                        switch (value)
                                                        {
                                                            case "0":
                                                                value = "FALSE";
                                                                break;
                                                            default:
                                                                value = "TRUE";
                                                                break;
                                                        }
                                                        break;


                                                }
                                            }
                                            currentRow.Add(value);
                                            j = j + 1;
                                        }
                                    }
                                }
                                j = 0;
                                newLine = String.Join(", ", currentRow.ToArray());
                                allRows.Add(newLine);
                            }
                            return string.Join(Environment.NewLine, allRows.ToArray());
                        }

                    }
                    return string.Empty;

                }
                return string.Empty;
            }

        }
    }
}
