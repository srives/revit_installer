using System;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.Dynamic;
using System.IO;
using System.Net.Security;
using System.Reflection;
using System.Security;
using System.Threading;

namespace DSI
{
    public class Program
    {
        static void Main(string[] args)
        {
            Program run = new Program(args);

            /*
             ** This branch of the installer that should be reached if the
             ** executable is either launched manually or if it is launched
             ** through Revit.
             */
            if (!run.FlagDeployInstall && !run.FlagDeployManifest && !run.FlagDeployAsAdmin.raised)
            {
                run.LockThreadUntilRevitExit();
                run.Install();

                /*
                 ** This branch of the installer should only be reached
                 ** only if the installer is launched manually. The manifest
                 ** will not be written if the executable is launched through
                 ** Revit. This is because if Revit launches the executable,
                 ** then it's assumed that the Revit toolkit is on v2.0 and
                 ** already has a working manifest file.
                 */
                if (!run.FlagAddin)
                {
                    run.WriteManifests();
                }

                Console.WriteLine("Installation complete; please press any key to close this window.");
                Console.ReadKey();
            }
            /*
             ** Used for single installs as admin. This branch should only be reached with
             ** the --deploy-as-admin flag and an appropriate username provided.
             */
            else if (run.FlagDeployAsAdmin.raised)
            {
                run.LockThreadUntilRevitExit();
                run.Install($"\\\\{run.FlagDeployAsAdmin.machine}\\c$\\Users\\{run.FlagDeployAsAdmin.user}\\AppData\\Local");
                run.WriteManifests(run.FlagDeployAsAdmin.machine, run.FlagDeployAsAdmin.user);
            }
            /*
             ** This branch of the installer should only be reached if reached
             ** the installer is launched through the PDQ batch script. This
             ** kicks off step 1 of 2 of the remote install process.
             */
            else if (run.FlagDeployInstall)
            {
                run.LockThreadUntilRevitExit();
                run.Install();

                if (run.FlagUsername.raised && run.FlagPassword.raised && run.FlagDomain.raised)
                {
                    Process exe = new Process();

                    string workingDirectory;
                    if (run.FlagDebug)
                    {
                        workingDirectory = Properties.Resources.DEBUG_EXECUTABLE_PATH;
                    }
                    else
                    {
                        workingDirectory = Properties.Resources.REMOTE_EXECUTABLE_PATH;
                    }

                    exe.StartInfo = new ProcessStartInfo()
                    {
                        WorkingDirectory = workingDirectory,
                        FileName = @"DSIToolkitAddinInstaller.exe",
                        Arguments = $"--deploy-manifest --app-data-directory {Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)} --verbose",
                        UserName = run.FlagUsername.value,
                        Password = run.FlagPassword.value,
                        Domain = run.FlagDomain.value,
                        UseShellExecute = false,
                        RedirectStandardOutput = true
                    };

                    try 
                    {
                        exe.Start();
                        string output = exe.StandardOutput.ReadToEnd();
                        exe.WaitForExit();
                        Console.Write(output);
                    }
                    catch (Exception e)
                    {
                        run.WriteErrorLine(e, $"{e.Message}");
                    }
                }
                else
                {
                    run.WriteErrorLine("error reading either the username, password, or domain provided through the command line; check the arguments and rerun the executable");
                }
            }
            /*
             ** This branch of the installer should only be reached
             ** programatically though the first step of the PDQ branches.
             */
            else if (run.FlagDeployManifest)
            {
                run.LockThreadUntilRevitExit();
                run.WriteManifests(run.FlagAppDataDirectory.value);
            }
        }

        #region **Constructors**

        /// <summary>
        /// Program constructor. Acts mainly as a command line argument parser.
        /// </summary>
        /// <param name="args">The array of arguments provided from the command line.</param>
        public Program(string[] args)
        {
            for (int i = 0; i < args.Length; i++)
            {
                int additionalIncrements = 0;
                if (args[i] == "--verbose")
                {
                    FlagVerbose = true;
                    WriteVerboseLine("verbose flag provided; entering verbose mode");
                }

                if (args[i] == "--launched-from-addin")
                {
                    FlagAddin = true;
                    WriteVerboseLine("addin flag provided");
                }

                if (args[i] == "--deploy-install")
                {
                    FlagDeployInstall = true;
                    WriteVerboseLine("deploy install flag provided");
                }

                if (args[i] == "--deploy-manifest")
                {
                    FlagDeployManifest = true;
                    WriteVerboseLine("deploy manifest flag provided");
                }

                if (args[i] == "--debug")
                {
                    FlagDebug = true;
                    WriteVerboseLine("debug flag provided");
                }


                if (args[i] == "--deploy-as-admin")
                {
                    if (!args[i + 1].StartsWith("-")
                        && !args[i + 2].StartsWith("-"))
                    {
                        FlagDeployAsAdmin = (
                            raised: true,
                            machine: args[i + 1],
                            user: args[i + 2]);

                        WriteVerboseLine($"deploy as admin flag provided");
                        additionalIncrements++;
                    }
                }

                if (args[i] == "--version")
                {
                    if (!args[i + 1].StartsWith("-"))
                    {
                        FlagVersion = (
                            raised: true,
                            value: args[i + 1]);

                        WriteVerboseLine($"version flag and argument provided; only the {FlagVersion.value} version of the toolkit will be installed (if it exists)");
                        additionalIncrements++;
                    }
                }

                if (args[i] == "--app-data-directory")
                {
                    if (!args[i + 1].StartsWith("-"))
                    {
                        FlagAppDataDirectory = (
                            raised: true,
                            value: args[i + 1]);

                        WriteVerboseLine("app data directory flag and argument provided; the manifest file will be written using this argument");
                        additionalIncrements++;
                    }
                }

                if (args[i] == "--username" || args[i] == "-u")
                {
                    if (!args[i + 1].StartsWith("-"))
                    {
                        FlagUsername = (
                            raised: true,
                            value: args[i + 1]);

                        WriteVerboseLine("username flag and argument provided");
                        additionalIncrements++;
                    }
                }

                if (args[i] == "--password" || args[i] == "-p")
                {
                    if (!args[i + 1].StartsWith("-"))
                    {
                        // NOTE: this password is not secure; it can still be accessed though
                        // managed memory attacks since there is an string instance of it
                        string pw = args[i + 1];
                        SecureString spw = new SecureString();

                        foreach (char c in pw) spw.AppendChar(c);
                        spw.MakeReadOnly();

                        FlagPassword = (
                            raised: true,
                            value: spw);

                        WriteVerboseLine("password flag and argument provided");
                        additionalIncrements++;
                    }
                }

                if (args[i] == "--domain" || args[i] == "-d")
                {
                    if (!args[i + 1].StartsWith("-"))
                    {
                        FlagDomain = (
                            raised: true,
                            value: args[i + 1]);

                        WriteVerboseLine("domain flag and argument provided");
                        additionalIncrements++;
                    }
                }

                i += additionalIncrements;
            }

            if (FlagDebug)
            {
                SourcePathPrefix = Properties.Resources.DEBUG_PATH_PREFIX;
                SourcePathPostfix = Properties.Resources.DEBUG_PATH_POSTFIX;
            }
        }

        #endregion

        #region **Properties**

        /// <summary>
        /// The name of the target dll file.
        /// </summary>
        public string DLLName { get; } = Properties.Resources.DLL_NAME;

        public bool FlagAddin { get; }

        public (bool raised, string value) FlagAppDataDirectory { get; } 

        public bool FlagDebug { get; }

        public ( bool raised, string machine, string user) FlagDeployAsAdmin { get; }

        public bool FlagDeployInstall { get; }

        public bool FlagDeployManifest { get; }

        public (bool raised, string value) FlagDomain { get; }

        public (bool raised, SecureString value) FlagPassword { get; }

        public (bool raised, string value) FlagUsername { get; }

        public bool FlagVerbose { get; }

        public (bool raised, string value) FlagVersion { get; }

        public string GUID { get; } = Properties.Resources.GUID;

        public string InstallDirName { get; } = Properties.Resources.INSTALL_DIRECTORY_NAME;

        public DirectoryInfo RevitAddinsDir { get; } = new DirectoryInfo(Properties.Resources.REVIT_ADDINS_PATH);

        public string SourcePathPostfix { get; } = "";

        public string SourcePathPrefix { get; } = Properties.Resources.SOURCE_PATH_PREFIX;

        public string[] SupportedRevitVersions { get; } = new string[4] { "2018", "2019", "2020", "2022" };

        #endregion

        #region **Methods**

        /// <summary>
        /// Copys the remote, up-to-date version of the Revit toolbar to the user's local app data.
        /// </summary>
        public void Install()
        {
            Install(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData));
        }

        /// <summary>
        /// Copys the remote, up-to-date version of the Revit toolbar to the user's local app data.
        /// </summary>
        /// <param name="localAppDataPath">The path to the user's local app data folder.</param>
        public void Install(string localAppDataPath)
        {
            DirectoryInfo toolkitInstallRootDirectory = new DirectoryInfo($"{localAppDataPath}\\{InstallDirName}");

            if (!toolkitInstallRootDirectory.Exists)
            {
                WriteVerboseLine($"{toolkitInstallRootDirectory.FullName} doesn't exist; creating the directory now");
                toolkitInstallRootDirectory.Create();
            }

            foreach (string version in SupportedRevitVersions)
            {
                WriteVerboseLine($"beginning installation process for Revit {version}");

                DirectoryInfo toolkitYearlyVersionDirectory = new DirectoryInfo($"{toolkitInstallRootDirectory.FullName}\\{version}");
                if (!toolkitYearlyVersionDirectory.Exists)
                {
                    WriteVerboseLine($"{toolkitYearlyVersionDirectory.FullName} doesn't exist; creating the directory now");
                    toolkitYearlyVersionDirectory.Create();
                }

                WriteVerboseLine($"now copying {SourcePathPrefix}{version}{SourcePathPostfix} to {toolkitYearlyVersionDirectory.FullName}");
                DirectoryCopy($"{SourcePathPrefix}{version}{SourcePathPostfix}", toolkitYearlyVersionDirectory.FullName);
            }
        }

        /// <summary>
        /// Checks to see if there are any Revit processes running before allowing the installer to continue.
        /// </summary>
        public void LockThreadUntilRevitExit()
        {
            bool isRevitRunning = false;
            int processChecks = 1;

            do
            {
                Process[] processes = Process.GetProcessesByName("Revit");
                if (processes.Length > 0)
                {
                    int sleepInterval = 5000;
                    isRevitRunning = true;
                    Console.WriteLine($"There is an instance of Revit still running; waiting until the instance is closed to continue... ({processChecks})");
                    WriteVerboseLine($"sleeping for {sleepInterval / 1000} seconds; total time spent sleeping is {(processChecks - 1) * 5000 / 1000} seconds");
                    processChecks++;
                    Thread.Sleep(sleepInterval);
                }
            }
            while (isRevitRunning);
        }

        /// <summary>
        /// Writes the new manifest file to the appropriate location.
        /// </summary>
        /// <remarks>User must have admin privileges in order for this to function properly.</remarks>
        public void WriteManifests()
        {
            WriteManifests(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData));
        }

        /// <summary>
        /// Writes the new manifest file to the appropriate location.
        /// </summary>
        /// <param name="localAppDataPath">The path to the user's local app data folder.</param>
        /// <remarks>User must have admin privileges in order for this to function properly.</remarks>
        public void WriteManifests(string localAppDataPath)
        {
            try
            {
                foreach (string version in SupportedRevitVersions)
                {
                    DirectoryInfo revitAddinsVersionDirectory = new DirectoryInfo($"{RevitAddinsDir.FullName}\\{version}");
                    CleanupOldManifests(revitAddinsVersionDirectory.FullName, version);

                    string addinPath = $"{revitAddinsVersionDirectory.FullName}\\{DLLName}{version}.addin";

                    WriteVerboseLine($"checking to see if an addin file already exists...");
                    if (!File.Exists(addinPath))
                    {
                        WriteVerboseLine($"{addinPath} does not exist; creating addin file now");
                        FileStream fs = File.Create(addinPath);
                        fs.Close();
                    }
                    else
                    {
                        WriteVerboseLine($"{addinPath} exists; continuing");
                    }

                    WriteVerboseLine("writing manifest to addin file");
                    File.WriteAllText(addinPath, ConstructManifest(version, $"{localAppDataPath}\\{InstallDirName}\\{version}\\{DLLName}.dll", GUID));
                }
            }
            catch (Exception e)
            {
                WriteErrorLine(e, "unable to write manifests; elevating to admin privilages would probably fix this issue");
            }
        }

        public void WriteManifests(string machine, string user)
        {
            try
            {
                foreach (string version in SupportedRevitVersions)
                {
                    DirectoryInfo revitAddinsVersionDirectory = new DirectoryInfo($"\\\\{machine}\\c$\\ProgramData\\Autodesk\\Revit\\Addins\\{version}");
                    CleanupOldManifests(revitAddinsVersionDirectory.FullName, version);

                    string addinPath = $"{revitAddinsVersionDirectory.FullName}\\{DLLName}{version}.addin";

                    WriteVerboseLine($"checking to see if an addin file already exists...");
                    if (!File.Exists(addinPath))
                    {
                        WriteVerboseLine($"{addinPath} does not exist; creating addin file now");
                        FileStream fs = File.Create(addinPath);
                        fs.Close();
                    }
                    else
                    {
                        WriteVerboseLine($"{addinPath} exists; continuing");
                    }

                    WriteVerboseLine("writing manifest to addin file");
                    File.WriteAllText(addinPath, ConstructManifest(version, $"C:\\Users\\{user}\\AppData\\Local\\{InstallDirName}\\{version}\\{DLLName}.dll", GUID));
                }
            }
            catch (Exception e)
            {
                WriteErrorLine(e, "unable to write manifests; elevating to admin privilages would probably fix this issue");
            }
        }

        /// <summary>
        /// Cleans up the Revit addin folders by deleting the manifest files for older versions of the toolkit.
        /// </summary>
        /// <param name="revitAddinVersionPath">The path to the Revit addin folder for any given year.</param>
        /// <param name="version">The yearly version of Revit that is currently being targeted.</param>
        private void CleanupOldManifests(string revitAddinVersionPath, string version)
        {
            string[] pathsToClean = new string[2]
            {
                            $"{revitAddinVersionPath}\\DSIToolKit - {version}.addin",
                            $"{revitAddinVersionPath}\\DSIToolKit - Debug {version}.addin"
            };

            foreach (string path in pathsToClean)
            {
                WriteVerboseLine($"checking to see if {path} exists...");
                if (File.Exists(path))
                {
                    WriteVerboseLine($"{path} exists; now deleting file");
                    File.Delete(path);
                }
                else
                {
                    WriteVerboseLine($"{path} does not exist, continuing");
                }
            }
        }


        /// <summary>
        /// Constructs the addin manifest from the template and the functions arguments.
        /// </summary>
        /// <param name="revitYear">The Revit version year that the manifest points to.</param>
        /// <param name="dllPath">The path to the DSIRevitToolkit.dll</param>
        /// <param name="clientGuid">A GUID string.</param>
        /// <returns></returns>
        static string ConstructManifest(string revitYear, string dllPath, string clientGuid)
        {
            string addinManifest =
                $"<?xml version=\"1.0\" encoding=\"utf-8\" ?>\n" +
                $"<RevitAddIns>\n" +
                $"  <AddIn Type=\"Application\">\n" +
                $"    <Name>DSI Toolkit</Name>\n" +
                $"    <Description>DSI Toolkit Ribbon for Revit {revitYear}</Description>\n" +
                $"    <Assembly>{dllPath}</Assembly>\n" +
                $"    <FullClassName>DSI.Application</FullClassName>\n" +
                $"    <ClientId>{clientGuid}</ClientId>\n" +
                $"    <VendorId>us.dsi</VendorId>\n" +
                $"    <VendorDescription>Dynamic Systems, Inc.</VendorDescription>\n" +
                $"  </AddIn>\n" +
                $"</RevitAddIns>\n";

            return addinManifest;
        }

        /// <summary>
        /// Recursively copys all files and subdirectories to a target directory.
        /// </summary>
        /// <param name="sourceDirName">The source directory.</param>
        /// <param name="destDirName">The target directory.</param>
        private static void DirectoryCopy(string sourceDirName, string destDirName)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            DirectoryInfo[] dirs = dir.GetDirectories();
            // If the destination directory doesn't exist, create it.
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);
                file.CopyTo(temppath, true);
            }

            foreach (DirectoryInfo subdir in dirs)
            {
                string temppath = Path.Combine(destDirName, subdir.Name);
                DirectoryCopy(subdir.FullName, temppath);
            }
        }

        /// <summary>
        /// Writes an error line to the console if the verbose flag is set to true.
        /// </summary>
        private void WriteErrorLine(string msg)
        {
            Console.WriteLine($"[error] {msg}");
        }

        /// <summary>
        /// Writes an error line to the console if the verbose flag is set to true.
        /// </summary>
        /// <param name="msg">The message to write to the console.</param>
        private void WriteErrorLine(Exception e, string msg)
        {
            Console.WriteLine($"[error] {e.GetType().FullName} - {msg}");
        }

        /// <summary>
        /// Writes a line to the console if the verbose flag is set to true.
        /// </summary>
        /// <param name="msg">The message to write to the console.</param>
        private void WriteVerboseLine(string msg)
        {
            if (FlagVerbose) Console.WriteLine($"[verbose] {msg}");
        }

        #endregion
    }
}
