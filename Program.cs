using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Linq;

// https://github.com/PSAppDeployToolkit/PSAppDeployToolkit/blob/ee612af863d9e7a1ab82f55e597790afb55970b7/Sources/Deploy-Application/Deploy-Application/DeployApplication.cs
// https://github.com/MSEndpointMgr/PSInvoker/tree/master/PSInvoker

namespace Invoker
{
    internal class Program
    {

        static void Main(string[] args)
        {
            try
            {

                int processExitCode = 60010;
                string currentAppPath = new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath;
                string currentAppFolder = Path.GetDirectoryName(currentAppPath);
                //string appName = Application.ExecutablePath;
                string scriptPath = String.Empty;
                string powershellExePath = Path.Combine(Environment.GetEnvironmentVariable("WinDir"), "System32\\WindowsPowerShell\\v1.0\\PowerShell.exe");
                string powershellArgs = "-ExecutionPolicy Bypass -NoProfile -NoLogo -WindowStyle Hidden";
                //List<string> commandLineArgs = new List<string>(Environment.GetCommandLineArgs());
                //List<string> commandLineArgs = new List<string>(Environment.CommandLine.Split(new[] { " -" }, StringSplitOptions.None));
                List<string> commandLineArgs = new List<string>();
                bool isForceX86Mode = false;
                //WriteDebugMessage(String.Join(", ", commandLineArgs.ToArray()));
                bool is64BitOS = false;
                if (Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE").Contains("64"))
                {
                    is64BitOS = true;
                }

                //WriteDebugMessage("RawCommand: " + string.Join(" ", args));
                // Split command line
                commandLineArgs = args.Select(arg =>
                    arg.StartsWith("\"") && arg.EndsWith("\"") && arg.Length >= 2 ? arg.Substring(1, arg.Length - 2) :
                    arg.StartsWith("\"") ? arg.Substring(1) :
                    arg.EndsWith("\"") ? arg.Substring(0, arg.Length - 1) : arg)
                    .Select(arg => arg.Replace("\"", ""))
                    .Select(arg => arg.Replace("'", ""))
                    .Select(arg => arg.StartsWith("-Command") ? arg : ("'" + arg.Trim() + "'"))
                    .ToList();

                // Trim ending & starting empty space from each element in the command-line
                //commandLineArgs = commandLineArgs.ConvertAll(s => s.Trim());
                // Remove first command-line argument as this is always the executable name
                //commandLineArgs.RemoveAt(0);

                WriteDebugMessage(String.Join(",", commandLineArgs.ToArray()));

                // Check if x86 PowerShell mode was specified on command line
                if (commandLineArgs.Exists(x => x == "/32"))
                {
                    isForceX86Mode = true;
                    WriteDebugMessage("'/32' parameter was specified on the command-line. Running in forced x86 PowerShell mode...");
                    // Remove the /32 command line argument so that it is not passed to PowerShell script
                    commandLineArgs.RemoveAll(x => x == "/32");
                }

                // Check for the App Deploy Script file being specified
                string commandLineScriptFileArg = String.Empty;
                //string commandLineAppDeployScriptPath = String.Empty;
                if (commandLineArgs.Exists(x => x.StartsWith("-File ") || x.StartsWith("-File")))
                {
                    throw new Exception("'-File' parameter was specified on the command-line. Please use the '-Command' parameter instead because using the '-File' parameter can return the incorrect exit code in PowerShell 2.0.");
                }
                else if (commandLineArgs.Exists(x => x.StartsWith("-Command ") || x.StartsWith("-Command")))
                {
                    commandLineScriptFileArg = commandLineArgs.Find(x => x.StartsWith("-Command ") || x.StartsWith("-Command"));
                    //scriptPath = commandLineScriptFileArg.Replace("-Command ", String.Empty).Replace("\"", String.Empty);
                    commandLineArgs.RemoveAt(commandLineArgs.FindIndex(x => x.StartsWith("-Command")));
                    WriteDebugMessage("'-Command' parameter specified on command-line. Passing command-line untouched...");
                }
                else if (commandLineArgs.Exists(x => x.EndsWith(".ps1") || x.EndsWith(".ps1\"")))
                {
                    scriptPath = commandLineArgs.Find(x => x.EndsWith(".ps1") || x.EndsWith(".ps1\"")).Replace("\"", String.Empty);
                    commandLineArgs.RemoveAt(commandLineArgs.FindIndex(x => x.EndsWith(".ps1") || x.EndsWith(".ps1\"")));
                    WriteDebugMessage(".ps1 file specified on command-line. Appending '-Command' parameter name...");
                }
                else
                {
                    WriteDebugMessage("No '-Command' parameter specified on command-line. Adding parameter '-Command \"" + scriptPath + "\"'...");
                }

                // Define the command line arguments to pass to PowerShell
                powershellArgs = powershellArgs + " -Command & { & " + scriptPath + "";
                if (commandLineArgs.Count > 0)
                {
                    powershellArgs = powershellArgs + " " + string.Join(" ", commandLineArgs.ToArray());
                    //powershellArgs = powershellArgs + " -" + string.Join(" -", commandLineArgs.ToArray());
                }
                powershellArgs = powershellArgs + "; Exit $LastExitCode }";

                // Switch to x86 PowerShell if requested
                if (is64BitOS & isForceX86Mode)
                {
                    powershellExePath = Path.Combine(Environment.GetEnvironmentVariable("WinDir"), "SysWOW64\\WindowsPowerShell\\v1.0\\PowerShell.exe");
                }

                // Define PowerShell process
                WriteDebugMessage("PowerShell Path: " + powershellExePath);
                WriteDebugMessage("PowerShell Parameters: " + powershellArgs);
                ProcessStartInfo processStartInfo = new ProcessStartInfo();
                processStartInfo.FileName = powershellExePath;
                processStartInfo.Arguments = powershellArgs;
                processStartInfo.WorkingDirectory = Path.GetDirectoryName(powershellExePath);
                processStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                processStartInfo.UseShellExecute = true;

                // Start the PowerShell process and wait for completion
                processExitCode = 60011;
                Process process = new Process();
                try
                {
                    process.StartInfo = processStartInfo;
                    process.Start();
                    process.WaitForExit();
                    processExitCode = process.ExitCode;
                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    if ((process != null))
                    {
                        process.Dispose();
                    }
                }

                // Exit
                WriteDebugMessage("Exit Code: " + processExitCode);
                Environment.Exit(processExitCode);
            }
            catch (Exception ex)
            {
                WriteDebugMessage(ex.Message, true, MessageBoxIcon.Error);
                Environment.Exit(processExitCode);
            }

        }

        public static void WriteDebugMessage(string debugMessage = null, bool IsDisplayError = false, MessageBoxIcon MsgBoxStyle = MessageBoxIcon.Information)
        {
            // Output to the Console
            Console.WriteLine(debugMessage);

            // If we are to display an error message...
            IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
            if (IsDisplayError == true)
            {
                MessageBox.Show(new WindowWrapper(handle), debugMessage, Application.ProductName + " " + Application.ProductVersion, MessageBoxButtons.OK, (MessageBoxIcon)MsgBoxStyle, MessageBoxDefaultButton.Button1);
            }
        }

        public class WindowWrapper : System.Windows.Forms.IWin32Window
        {
            public WindowWrapper(IntPtr handle)
            {
                _hwnd = handle;
            }

            public IntPtr Handle
            {
                get { return _hwnd; }
            }

            private IntPtr _hwnd;
        }

        public static int processExitCode { get; set; }

    }
}