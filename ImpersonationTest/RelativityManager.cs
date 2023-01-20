using System;
using System.Data;
using System.Linq;
using kCura.Relativity.ImportAPI;
using kCura.Relativity.DataReaderClient;
using System.Configuration;

public static class RelativityManager
{
    public static bool Import(int workspaceArtifactId, DataTable data)
    {
        var iapi = new ImportAPI(
            ConfigurationManager.AppSettings["importUserName"],
            ConfigurationManager.AppSettings["importPassword"],
            ConfigurationManager.AppSettings["importUrl"]
        );

        var importMode = iapi.GetFileUploadMode(workspaceArtifactId).ToString();

        var fields = iapi.GetWorkspaceFields(workspaceArtifactId, artifactTypeID: 10);
        var identityFieldID = fields.First(x => x.FieldCategory == kCura.Relativity.ImportAPI.Enumeration.FieldCategoryEnum.Identifier);
        var idFieldID = identityFieldID.ArtifactID;

        var isSuccess = false;

        var importJob = iapi.NewImageImportJob();

        importJob.Settings.AutoNumberImages = false;
        importJob.Settings.BatesNumberField = "Bates";
        importJob.Settings.CaseArtifactId = workspaceArtifactId;
        importJob.Settings.DocumentIdentifierField = "DocIdentifier";
        importJob.Settings.FileLocationField = "FilePath";
        importJob.Settings.CopyFilesToDocumentRepository = true;
        importJob.Settings.IdentityFieldId = idFieldID;
        importJob.Settings.OverwriteMode = OverwriteModeEnum.Overlay;
        importJob.Settings.DisableImageLocationValidation = true; // we are manually validating the image locations ahead of time
        importJob.SourceData.SourceData = data;
        // Setting the maximum error count to the row count times two because the API seems to either cut this number in half somewhere or double-report errors and then adding one for the 'max errors displayed' error.
        importJob.Settings.MaximumErrorCount = (data.Rows.Count * 2) + 1;

        importJob.OnComplete += report =>
        {
            var errorCount = report.ErrorRowCount;

            if (report.FatalException != null)
            {
                Console.WriteLine($"Fatal Error: {report.FatalException.Message}");
                errorCount++;
            }

            Console.WriteLine($"Import finished. {report.TotalRows} records processed. {errorCount} errors. Total duration (minutes): {(report.EndTime - report.StartTime).TotalMinutes}");

            isSuccess = errorCount == 0;
        };

        importJob.OnError += row =>
        {
            // Log file errors
            string logExpression = "Bates = '" + row["FileID"]?.ToString() + "'";
            var errorLOGLine = data.Select(logExpression).First();
            var error = errorLOGLine.Field<string>(data.Columns.IndexOf("OPTPath")) +
                " [Document Identifier: " + row["DocumentID"]?.ToString() +
                " [File Name: " + row["FileID"]?.ToString() + "] " +
                " [File Path: " + row["FilePath"]?.ToString() + "] " +
                " [Message: " + row["Message"]?.ToString();

            Console.WriteLine("Error: " + error);
        };

        importJob.OnProcessProgress += status =>
        {
            Console.WriteLine($"Progress update - {status.TotalRecordsProcessed} records processed: {DateTime.Now}");
        };

        importJob.Execute();

        iapi = null;

        return isSuccess;
    }
}
