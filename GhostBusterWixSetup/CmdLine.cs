namespace Swiss
{
    using System;

    /// CmdLine Storage and Definition Class
    /// </summary>
    public class CmdLine : CmdLineBase
    {
        static readonly CmdLine _instance = new CmdLine();

        /// <summary>
        /// Target Executable.
        /// </summary>
        [CmdLinePairAttribute("TargetExe", true, "Target Executable", false, true)]
        public static String TargetExe
        {
            get;
            set;
        }

        /// <summary>
        /// Target Executable (output) Directory.
        /// </summary>
        [CmdLinePairAttribute("OutputPath", true, "Target Executable (output) Directory", false, true)]
        public static String OutputPath
        {
            get;
            set;
        }

        /// <summary>
        /// Explicit static constructor tells # compiler
        /// not to mark type as beforefieldinit.
        /// </summary>
        static CmdLine()
        {
            //
        }

        /// <summary>
        /// Use Default Section 'CmdLineSettings'.
        /// </summary>
        private CmdLine()
            : base("CmdLineSettings")
        {
            // CmdLineBase sets all properties to their DefaulValue Attriibute's Value.
            // LoadSettings();
        }

        /// <summary>
        /// Visible when reflecting.
        /// </summary>
        public static CmdLine Instance
        {
            get
            {
                return _instance;
            }
        }
    }
}
