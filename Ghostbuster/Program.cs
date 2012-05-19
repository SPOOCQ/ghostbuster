#region License
/*
Copyright (c) 2009, G.W. van der Vegt
All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided 
that the following conditions are met:

* Redistributions of source code must retain the above copyright notice, this list of conditions and the 
  following disclaimer.

* Redistributions in binary form must reproduce the above copyright notice, this list of conditions and 
  the following disclaimer in the documentation and/or other materials provided with the distribution.

* Neither the name of G.W. van der Vegt nor the names of its contributors may be 
  used to endorse or promote products derived from this software without specific prior written 
  permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY 
EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF 
MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL 
THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, 
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, 
PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS 
INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, 
STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF 
THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. 
*/
#endregion License

using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Text;
using Microsoft.Win32.TaskScheduler;
using System.Threading;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace Ghostbuster
{
    static class Program
    {
        private static readonly String S_CONSOLE = "You need to have Administrative Priviliges to run in Console mode.";
        private static readonly String S_SCHEDULE = "You need to have Administrative Priviliges to register or unregister a Scheduler Task.";
        private static readonly String S_SYNTAX = String.Format("Syntax: {0} [Option [ConfFile]]\r\n" +
                                                               "\r\n" +
                                                               "Options: \r\n" +
                                                               "\r\n" +
                                                               "  /Help\t\t; This Message\r\n" +
                                                               "  /Register ConfFile\t; * Registers a Scheduler Task\r\n" +
                                                               "  /UnRegister\t; * Unregisters a Scheduler Task\r\n" +
                                                               "  /NoGui ConfFile\t; * Removes Ghosted Devices specified in ConfFile\r\n" +
                                                               "\r\n" +
                                                               "Examples: \r\n" +
                                                               "\r\n" +
                                                               "  {0}\r\n" +
                                                               "  {0} /Help\r\n" +
                                                               "  {0} /nogui %appdata%\\{0}\\{0}.ini\r\n" +
                                                               "  {0} /register %appdata%\\{0}\\{0}.ini\r\n" +
                                                               "  {0} /unregister\r\n" +
                                                               "\r\n" +
                                                               "The options marked with a * need an evelvated command prompt.\r\n", Buster.S_TITLE);

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            //! Add Support for CommandLine Options:
            //!
            //! /Remove
            //! /Register (Task)
            //! /Unregister (Task)
            //! IniFile
            //!

            //! Start logging to TaskManager Log, EventLog and Debug Console.

            switch (Environment.GetCommandLineArgs().Length)
            {

                case 3:
                    if (Buster.IsAdmin())
                    {
                        if (Environment.GetCommandLineArgs()[1].ToUpper().Equals("/NOGUI"))
                        {
                            String fn = Environment.ExpandEnvironmentVariables(Environment.GetCommandLineArgs()[2]);

                            if (File.Exists(fn) && fn.ToUpper().EndsWith(".INI"))
                            {
                                using (Buster buster = new Buster(fn))
                                {
                                    Buster.Enumerate();

                                    Buster.WriteToEventLog(String.Format("Found {0} Device(s)", Buster.HwEntries.Count), EventLogEntryType.Information);

                                    Buster.WriteToEventLog(String.Format("Loaded {0} Device Filter(s)", Buster.Devices.Count), EventLogEntryType.Information);
                                    Buster.WriteToEventLog(String.Format("Loaded {0} Class Filter(s)", Buster.Classes.Count), EventLogEntryType.Information);
                                    Buster.WriteToEventLog(String.Format("Loaded {0} Wildcard Filter(s)", Buster.Wildcards.Count), EventLogEntryType.Information);

                                    //! Count Filtered Devices (separate match from deletion in RemoveDevice_IF_Ghosted_AND_Marked).

                                    Int32 ghosted = 0;
                                    Int32 cnt = 0;
                                    foreach (HwEntry he in Buster.HwEntries)
                                    {
                                        //Debug.WriteLine(he.Description);

                                        if (he.Ghosted)
                                        {
                                            ghosted++;
                                        }

                                        if (he.RemoveDevice_IF_Ghosted_AND_Marked())
                                        {
                                            cnt++;
                                        }
                                    }

                                    Buster.WriteToEventLog(String.Format("Removed {0} of {1} Ghosted Device(s).", cnt, ghosted), EventLogEntryType.Information);

                                    Application.Exit();

                                    return;
                                }
                            }
                            else
                            {
                                MessageBox.Show(String.Format("Configuration File not found at \r\n'{0}'.", fn), Buster.S_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else if (Environment.GetCommandLineArgs()[1].ToUpper().Equals("/REGISTER"))
                        {
                            // Get the service on the local machine
                            using (TaskService ts = new TaskService())
                            {
                                String user = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                                String conf = Environment.ExpandEnvironmentVariables(Environment.GetCommandLineArgs()[2]);
                                String appl = Application.ExecutablePath;

                                Debug.WriteLine(appl);
                                Debug.WriteLine(conf);

                                Version ver = ts.HighestSupportedVersion;
                                Boolean newVer = (ver >= new Version(1, 2));

                                TaskDefinition td = ts.NewTask();

                                td.Data = "GhostBuster";
                                td.Principal.UserId = user;
                                td.Principal.LogonType = TaskLogonType.InteractiveToken;
                                td.RegistrationInfo.Author = "Wim van der Vegt";
                                td.RegistrationInfo.Description = String.Format(
                                    "{0} removes selected Ghosted Devices. \r\n" +
                                    "\r\n" +
                                    "Default Settings for this scheduled task are:\r\n" +
                                    " - Running with the required Administrative Privileges,\r\n" +
                                    " - Running 1 minute after a user logs on,\r\n" +
                                    " - Can be started manually with 'schtasks /run /ts {0}'. \r\n" +
                                    "\r\n" +
                                    "See http://ghosutbuster.codeplex.com", Buster.S_TITLE);
                                td.RegistrationInfo.Documentation = "See http://ghostbuster.codeplex.com";
                                td.Settings.DisallowStartIfOnBatteries = true;
                                td.Settings.Hidden = false;

                                if (newVer)
                                {
                                    td.Principal.RunLevel = TaskRunLevel.Highest;
                                    td.RegistrationInfo.Source = Buster.S_TITLE;
                                    td.RegistrationInfo.URI = new Uri("http://ghostbuster.codeplex.com");
                                    td.RegistrationInfo.Version = new Version(0, 9);
                                    td.Settings.AllowDemandStart = true;
                                    td.Settings.AllowHardTerminate = true;
                                    td.Settings.Compatibility = TaskCompatibility.V2;
                                    td.Settings.MultipleInstances = TaskInstancesPolicy.Queue;
                                }

                                LogonTrigger lTrigger = (LogonTrigger)td.Triggers.Add(new LogonTrigger());
                                if (newVer)
                                {
                                    lTrigger.Delay = TimeSpan.FromMinutes(1);
                                    lTrigger.UserId = user;
                                }

                                td.Actions.Add(new ExecAction(
                                    appl,
                                    String.Format("/NOGUI {0}", conf),
                                    Path.GetDirectoryName(appl)));

                                //if (ts.FindTask(Buster.S_TITLE) != null)
                                //{
                                //    ts.RootFolder.DeleteTask(Buster.S_TITLE);
                                //}

                                // Register the task in the root folder
                                Task task = ts.RootFolder.RegisterTaskDefinition(
                                       Buster.S_TITLE, td,
                                       TaskCreation.CreateOrUpdate, null, null,
                                       TaskLogonType.InteractiveToken,
                                       null);

                                task.ShowEditor();
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show(S_CONSOLE, Buster.S_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    break;

                case 2:
                    {
                        if (Buster.IsAdmin())
                        {
                            if (Environment.GetCommandLineArgs()[1].ToUpper().Equals("/UNREGISTER"))
                            {
                                using (TaskService ts = new TaskService())
                                {
                                    if (ts.FindTask(Buster.S_TITLE) != null)
                                    {
                                        ts.RootFolder.DeleteTask(Buster.S_TITLE);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show(S_SYNTAX, Buster.S_TITLE,
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }

                            Application.Exit();

                            return;
                        }
                        else
                        {
                            if (Environment.GetCommandLineArgs()[1].ToUpper().Equals("/REGISTER") ||
                                Environment.GetCommandLineArgs()[1].ToUpper().Equals("/UNREGISTER"))
                            {
                                MessageBox.Show(S_SCHEDULE, Buster.S_TITLE,
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else if (Environment.GetCommandLineArgs()[1].ToUpper().Equals("/NOGUI"))
                            {
                                MessageBox.Show(S_CONSOLE, Buster.S_TITLE,
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show(S_SYNTAX, Buster.S_TITLE,
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    }
                    break;

                default:
                    Application.Run(new Mainform());
                    break;
            }
        }
    }
}
