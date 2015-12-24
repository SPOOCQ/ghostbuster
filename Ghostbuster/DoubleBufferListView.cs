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

namespace Ghostbuster
{
    using System;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    class DoubleBufferListView
    {
        #region Enumerations

        internal enum LVM
        {
            LVM_FIRST = 0x1000,
            LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54),
            LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55),
        }

        internal enum LVS_EX
        {
            LVS_EX_GRIDLINES = 0x00000001,
            LVS_EX_SUBITEMIMAGES = 0x00000002,
            LVS_EX_CHECKBOXES = 0x00000004,
            LVS_EX_TRACKSELECT = 0x00000008,
            LVS_EX_HEADERDRAGDROP = 0x00000010,
            LVS_EX_FULLROWSELECT = 0x00000020,
            LVS_EX_ONECLICKACTIVATE = 0x00000040,
            LVS_EX_TWOCLICKACTIVATE = 0x00000080,
            LVS_EX_FLATSB = 0x00000100,
            LVS_EX_REGIONAL = 0x00000200,
            LVS_EX_INFOTIP = 0x00000400,
            LVS_EX_UNDERLINEHOT = 0x00000800,
            LVS_EX_UNDERLINECOLD = 0x00001000,
            LVS_EX_MULTIWORKAREAS = 0x00002000,
            LVS_EX_LABELTIP = 0x00004000,
            LVS_EX_BORDERSELECT = 0x00008000,
            LVS_EX_DOUBLEBUFFER = 0x00010000,
            LVS_EX_HIDELABELS = 0x00020000,
            LVS_EX_SINGLEROW = 0x00040000,
            LVS_EX_SNAPTOGRID = 0x00080000,
            LVS_EX_SIMPLESELECT = 0x00100000
        }

        #endregion Enumerations

        #region Methods

        /// <summary>
        /// Add DoubleBuffer to ListView
        /// </summary>
        /// <param name="lv">The LisView to adjust</param>
        public static void SetExStyles(ListView lv)
        {
            LVS_EX styles = (LVS_EX)SendMessage(lv.Handle,
                (int)LVM.LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0);

            styles |= LVS_EX.LVS_EX_DOUBLEBUFFER | LVS_EX.LVS_EX_BORDERSELECT;

            SendMessage(lv.Handle,
                (int)LVM.LVM_SETEXTENDEDLISTVIEWSTYLE, 0, (int)styles);
        }

        /// <summary>
        /// Remove DoubleBuffer to ListView
        /// </summary>
        /// <param name="lv">The LisView to adjust</param>
        public static void ResetExStyles(ListView lv)
        {
            LVS_EX styles = (LVS_EX)SendMessage(lv.Handle,
                (int)LVM.LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0);
            
            styles &= ~(LVS_EX.LVS_EX_DOUBLEBUFFER | LVS_EX.LVS_EX_BORDERSELECT);

            SendMessage(lv.Handle,
                (int)LVM.LVM_SETEXTENDEDLISTVIEWSTYLE, 0, (int)styles);
        }
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        internal static extern int SendMessage(IntPtr handle, int messg, int wparam, int lparam);

        #endregion Methods
    }
}