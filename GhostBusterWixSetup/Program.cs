using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Swiss;
using WixBuilder;

namespace GhostBusterWixSetup
{
    class Program
    {
        static String src = String.Empty; //@"C:\Users\User\Documents\Visual Studio 2008\Projects\Hottack.Net\Hottack.Net\bin\Debug";
        static String exe = String.Empty; //@"Hottack.Net.exe";

        static void Main(string[] args)
        {
            Debug.Print(Directory.GetCurrentDirectory());

            if (CmdLineArgs.Instance.ContainsKey("TargetExe"))
            {
                exe = CmdLineArgs.Instance["TargetExe"];
            }

            if (CmdLineArgs.Instance.ContainsKey("OutputPath"))
            {
                src = Path.GetFullPath(CmdLineArgs.Instance["OutputPath"]);
            }

            if (String.IsNullOrWhiteSpace(src) || String.IsNullOrWhiteSpace(exe) || !File.Exists(Path.Combine(src, exe)))
            {
                Console.WriteLine(CmdLineArgs.Instance.HelpOnArguments(CmdLine.Instance));
            }
            else
            {
                Console.WriteLine("[Commandline Arguments]");
                Console.WriteLine("OutputPath:\r\n   {0}", src);
                Console.WriteLine("TargetExe:\r\n   {0}", exe);
                Console.WriteLine("Combined:\r\n   {0}", Path.GetFullPath(Path.Combine(src, exe)));

                using (WixDefinition wd = new WixDefinition(src, exe))
                {
                    // Make sure CompanyName does not contain strange characters as it's used for directory name etc.
                    // wd.CompanyName = "Vived Management";

                    //! The default set of warnings to suppress.
                    wd.Suppress = new WixIceCodes[] { WixIceCodes.ICE64, WixIceCodes.ICE69, WixIceCodes.ICE91 };

                    //! Use an EULA license.
                    wd.SkipLicence = false;
                    wd.LicenseFile = Path.Combine(src, "License.rtf");

                    //
                    //! Support both install types.
                    wd.WixUISupportPerUser = true;
                    wd.WixUISupportPerMachine = true;

                    ///////////////////////////
                    // Product.wxs.
                    ///////////////////////////

                    wd.BuildProduct();
                    {
                        wd.AddWixBuilderUI();
                    }
                    wd.SaveProduct();

                    ///////////////////////////
                    // Fragment.wxs.
                    ///////////////////////////

                    wd.BuildFragment();
                    {
                        //! 1) Should end up in ProgramFiles\[ProductName].
                        String exeId = wd.AddFile(WixDefinition.InstallLocation, wd.MainExecutable, true, true);

                        //! 2) Should end up in ProgramFiles\[ProductName].
                        wd.AddFiles(WixDefinition.InstallLocation, "*.dll", SearchOption.AllDirectories, true);

                        //! 3) Shortcut Should end up in User|Common Menu\[ProductName].
                        String exeComp1 = wd.AddComponent(WixDefinition.ApplicationProgramsFolder);
                        wd.AddShortCut(
                            WixDefinition.ApplicationProgramsFolder,
                            exeComp1,
                            "[#" + exeId + "]",
                             WixDefinition.InstallLocation,
                            String.Empty,
                            String.Empty,
                            wd.ProductName);

                        //wd.AddFolder(WixDefinition.DesktopFolder, "CID_Desktop");

                        //! 4) Shortcut Should end up on User|Common Desktop.
                        String exeComp2 = wd.AddComponent(WixDefinition.DesktopFolder);
                        wd.AddShortCut(
                            WixDefinition.DesktopFolder,
                            exeComp2,
                            "[#" + exeId + "]",
                             WixDefinition.InstallLocation,
                            String.Empty,
                            String.Empty,
                            wd.ProductName);

                        //! 5) Shortcut Should end up in User|Common Menu\[ProductName].
                        String uninstComp = wd.AddComponent(WixDefinition.ApplicationProgramsFolder);
                        wd.AddShortCut(
                            WixDefinition.ApplicationProgramsFolder,
                            uninstComp,
                            "[System64Folder]msiexec.exe",
                             WixDefinition.InstallLocation,
                            String.Empty,
                            "/x [ProductCode]",
                            String.Format("Uninstall {0}", wd.ProductName));
                    }
                    wd.SaveFragment();

                    ///////////////////////////
                    // Compile and Build msi.
                    ///////////////////////////

                    wd.PrintOmittedFiles();

                    if (!String.IsNullOrEmpty(wd.WiXDir) && Directory.Exists(wd.WiXDir))
                    {
                        if (wd.Compile())
                        {
                            if (wd.Link())
                            {
                                wd.Test();
                                //wd.View();
                            }
                        }
                    }
                }
            }
            Console.WriteLine("");
            Console.Write("Press any key to exit");
            Console.ReadKey();
        }
    }
}
