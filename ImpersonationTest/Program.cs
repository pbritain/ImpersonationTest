using System;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Security.Permissions;
using Microsoft.Win32.SafeHandles;
using System.Runtime.ConstrainedExecution;
using System.Security;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Data;
using System.Configuration;

public class ImpersonationTest
{
    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool LogonUser(String lpszUsername, String lpszDomain, String lpszPassword,
        int dwLogonType, int dwLogonProvider, out SafeTokenHandle phToken);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    public extern static bool CloseHandle(IntPtr handle);

    static DataTable AllOPTData { get; set; }

    [PermissionSetAttribute(SecurityAction.Demand, Name = "FullTrust")]
    public static void Main(string[] args)
    {
        Console.WriteLine("Test started. Running application as: " + WindowsIdentity.GetCurrent().Name);

        SafeTokenHandle safeTokenHandle;
        try
        {
            const int LOGON32_PROVIDER_WINNT50 = 3;
            const int LOGON32_LOGON_NEW_CREDENTIALS = 9;

            // Call LogonUser to obtain a handle to an access token.
            bool returnValue = LogonUser(
                ConfigurationManager.AppSettings["optFileAccessAccountUserNameForImpersonation"],
                ConfigurationManager.AppSettings["optFileAccessAccountDomainNameForImpersonation"],
                ConfigurationManager.AppSettings["optFileAccessAccountPasswordForImpersonation"],
                9, 0,
                out safeTokenHandle);

            if (false == returnValue)
            {
                int ret = Marshal.GetLastWin32Error();
                throw new System.ComponentModel.Win32Exception(ret);
            }
            using (safeTokenHandle)
            {
                // Use the token handle returned by LogonUser.
                using (WindowsIdentity newId = new WindowsIdentity(safeTokenHandle.DangerousGetHandle()))
                {
                    using (WindowsImpersonationContext impersonatedUser = newId.Impersonate())
                    {
                        Console.WriteLine("Now impersonating as: " + ConfigurationManager.AppSettings["optFileAccessAccountUserNameForImpersonation"]);

                        // Import logic

                        InitializeOPTDataTable();

                        Console.WriteLine("Getting OPT file paths...");
                        var optFilePaths = GetOPTFilePaths();

                        Console.WriteLine("Validating and gathering data...");
                        foreach (var optFilePath in optFilePaths)
                        {
                            EnsureOPTFileIsValid(optFilePath);
                            GatherAndValidateData(optFilePath);
                        }

                        Console.WriteLine("Reading and importing OPT files...");
                        ReadAndImportOPTFiles();

                        Console.WriteLine("Import completed successfully");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Exception occurred: " + ex.ToString());
        }

        Console.Read();
    }

    private static List<string> GetOPTFilePaths()
    {
        try
        {
            // Run through the Image Load Directory and any directories one level down, find all .opt files, and add them to the optPaths list.
            List<string> optFilePaths = new List<string>();

            string ImageLoadDirectory = ConfigurationManager.AppSettings["ImageLoadDirectory"];

            // Look for OPT files in the root directory
            var newPaths = (Directory.GetFiles(ImageLoadDirectory, "*.opt", SearchOption.TopDirectoryOnly));
            foreach (string path in newPaths)
            {
                optFilePaths.Add(Path.Combine(ImageLoadDirectory, path));
            }

            // Look for OPTs in the first level of directories
            foreach (string directory in Directory.GetDirectories(ImageLoadDirectory, "*", SearchOption.TopDirectoryOnly))
            {
                var newSubPaths = (Directory.GetFiles(directory, "*.opt", SearchOption.TopDirectoryOnly));
                foreach (string path in newSubPaths)
                {
                    optFilePaths.Add(Path.Combine(ImageLoadDirectory, directory, path));
                }
            }

            // If no OPTs are detected, throw an error and explain why.
            if (!optFilePaths.Any())
            {
                throw new IOException("No OPT files have been found in the Image Load Directory.");
            }

            return optFilePaths;
        }
        catch (UnauthorizedAccessException)
        {
            throw new IOException("Access has been denied to the Image Load Directory.");
        }
    }

    private static void EnsureOPTFileIsValid(string optFilePath)
    {
        var firstLine = File.ReadLines(optFilePath).FirstOrDefault();

        // Run through OPT and check the first line.  If the first line of the OPT does not have the "Y" doc break, then the OPT is broken.
        if (string.IsNullOrEmpty(firstLine))
        {
            throw new Exception($"OPT file ({optFilePath}) is either empty or has a empty first line. " +
                "Please add data to OPT file or remove the empty line.");
        }
        else
        {
            var firstLineValues = firstLine.Split(',');
            var docBreak = firstLineValues[3];

            if (!string.Equals(docBreak, "Y", StringComparison.OrdinalIgnoreCase))
            {
                throw new Exception($"OPT file ({optFilePath}) does not start with a document break. " +
                    "Please review and fix the OPT file.");
            }
        }
    }

    private static void GatherAndValidateData(string optFilePath)
    {
        // Load opt file contents into memory so we don't have to keep going back and forth across the network
        var content = File.ReadAllLines(optFilePath);

        var directoryName = Path.GetDirectoryName(optFilePath);
        var docIdentifier = string.Empty;

        foreach (var line in content)
        {
            if (!string.IsNullOrEmpty(line))
            {
                var values = line.Split(',');

                var bates = values[0];
                var volume = values[1];
                var filePath = values[2];
                var docBreak = values[3];
                var pageCount = values[6];

                // Relative Path Handling
                filePath = CheckAndUpdateRelativePaths(filePath, directoryName);

                CheckImageFilePath(filePath);

                if (string.Equals("Y", docBreak, StringComparison.OrdinalIgnoreCase))
                {
                    docIdentifier = bates;
                }

                AllOPTData.Rows.Add(bates, docIdentifier, filePath, volume, pageCount, docBreak, optFilePath);
            }
        }
    }

    private static void CheckImageFilePath(string filePath)
    {
        // Check that each file in the OPT exists and write a useful message if it doesn't or if access is denied.
        try
        {
            if (!File.Exists(filePath))
            {
                throw new Exception($"File does not exist: {filePath}");
            }
        }
        catch (UnauthorizedAccessException)
        {
            throw new Exception($"File cannot be accessed: {filePath}");
        }
        catch (FileNotFoundException)
        {
            throw new Exception($"File does not exist: {filePath}");
        }
        catch (IOException)
        {
            throw new Exception($"File or network path cannot be found: {filePath}");
        }
        catch (Exception ex)
        {
            throw new Exception($"{ex.Message} : {filePath}");
        }
    }

    private static string CheckAndUpdateRelativePaths(string filePath, string directoryName)
    {
        // Relative Path Handling:
        // If the filePath value starts with ".\", then remove the first two characters and add to the ImageLoadDirectory.
        if (filePath.StartsWith(".\\"))
        {
            filePath = Path.Combine(directoryName, filePath.Remove(0, 2));
        }
        // If the filePath value returns a slash as the GetPathRoot value, then remove the first character and add to the ImageLoadDirectory.
        else if (Path.IsPathRooted(filePath) && Path.GetPathRoot(filePath).Equals(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
        {
            filePath = Path.Combine(directoryName, filePath.Remove(0, 1));
        }
        // Lastly, if it doesn't start with a "\" and doesn't contain a ":" for a drive letter, then it's still relative but just needs the ImageLoadDirectory added.
        else if (!filePath.StartsWith("\\") && !filePath.Contains(":"))
        {
            filePath = Path.Combine(directoryName, filePath);
        }

        return filePath;
    }

    private static void InitializeOPTDataTable()
    {
        AllOPTData = new DataTable();

        // Additional columns are added that the ImportAPI doesn't need, these are so we can write the error OPT file and error LOG file.
        AllOPTData.Columns.Add("Bates");
        AllOPTData.Columns.Add("DocIdentifier");
        AllOPTData.Columns.Add("FilePath");
        AllOPTData.Columns.Add("Volume");
        AllOPTData.Columns.Add("PageCount");
        AllOPTData.Columns.Add("DocBreak");
        AllOPTData.Columns.Add("OPTPath");
    }

    private static void ReadAndImportOPTFiles()
    {
        var isSuccess = RelativityManager.Import(int.Parse(ConfigurationManager.AppSettings["WorkspaceArtifactId"]), AllOPTData);
        if (!isSuccess)
        {
            throw new Exception("Error detected during image import");
        }
    }
}

public sealed class SafeTokenHandle : SafeHandleZeroOrMinusOneIsInvalid
{
    private SafeTokenHandle()
        : base(true)
    {
    }

    [DllImport("kernel32.dll")]
    [ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)]
    [SuppressUnmanagedCodeSecurity]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr handle);

    protected override bool ReleaseHandle()
    {
        return CloseHandle(handle);
    }
}