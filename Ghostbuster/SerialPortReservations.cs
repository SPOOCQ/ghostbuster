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
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Windows.Forms;

    /// <summary>
    /// A serial port reservations.
    /// </summary>
    public partial class SerialPortReservations : Form
    {
        #region Fields

        /// <summary>
        /// The comdb maximum ports arbitrated.
        /// </summary>
        private const Int32 COMDB_MAX_PORTS_ARBITRATED = 4096;

        /// <summary>
        /// The comdb minimum ports arbitrated.
        /// </summary>
        private const Int32 COMDB_MIN_PORTS_ARBITRATED = 256;

        /// <summary>
        /// The error access denied.
        /// </summary>
        private const Int32 ERROR_ACCESS_DENIED = 0x0005;

        /// <summary>
        /// Length of the error bad.
        /// </summary>
        private const Int32 ERROR_BAD_LENGTH = 0x0018;

        /// <summary>
        /// The error cantwrite.
        /// </summary>
        private const Int32 ERROR_CANTWRITE = 0x03F5;

        /// <summary>
        /// The error invalid parameter.
        /// </summary>
        private const Int32 ERROR_INVALID_PARAMETER = 0x0057;

        /// <summary>
        /// Information describing the error more.
        /// </summary>
        private const Int32 ERROR_MORE_DATA = 0x00EA;

        /// <summary>
        /// The error not connected.
        /// </summary>
        private const Int32 ERROR_NOT_CONNECTED = 0x8CA;

        /// <summary>
        /// The error sharing violation.
        /// </summary>
        private const Int32 ERROR_SHARING_VIOLATION = 0x0020;

        /// <summary>
        /// The error success.
        /// 
        /// See  See http://msdn.microsoft.com/en-us/library/windows/desktop/ms681382(v=vs.85).aspx.
        /// </summary>
        private const Int32 ERROR_SUCCESS = 0x0000;

        /// <summary>
        /// The hcomdb invalid handle value.
        /// </summary>
        private const Int32 HCOMDB_INVALID_HANDLE_VALUE = INVALID_HANDLE_VALUE;

        /// <summary>
        /// The invalid handle value.
        /// </summary>
        private const Int32 INVALID_HANDLE_VALUE = -1;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the Ghostbuster.SerialPortReservations class.
        /// </summary>
        public SerialPortReservations()
        {
            InitializeComponent();

            DoubleBufferListView.SetExStyles(listView1);
        }

        #endregion Constructors

        #region Enumerations

        /// <summary>
        /// Values that represent CDB_REPORT.
        /// </summary>
        public enum CDB_REPORT
        {
            /// <summary>
            /// An enum constant representing the bits option.
            /// </summary>
            BITS = 0,

            /// <summary>
            /// An enum constant representing the bytes option.
            /// </summary>
            BYTES = 1
        }

        #endregion Enumerations

        #region Methods

        /// <summary>
        /// Gets the execute.
        /// </summary>
        ///
        /// <returns>
        /// A DialogResult.
        /// </returns>
        public DialogResult Execute()
        {
            ProcessReservations();

            ListViewItem lvi = new ListViewItem((99).ToString());
            lvi.SubItems.Add(String.Format("COM{0}", 99));
            lvi.Tag = 99;
            lvi.Checked = false;

            listView1.Items.Add(lvi);

            return ShowDialog();
        }

        /// <summary>
        /// Adds reserved ports.
        /// </summary>
        internal void AddReservedPorts()
        {
            //! Must run Elevated!!
            //
            IntPtr HComDB = IntPtr.Zero;

            Int32 ret = ComDBOpen(out HComDB);

            if ((ret & 0x0000FFFF) != ERROR_ACCESS_DENIED)
            {
                Int32 maxPorts = 0;

                ComDBGetCurrentPortUsage(HComDB, null, 0, CDB_REPORT.BYTES, out maxPorts);

                // MessageBox.Show(String.Format("MaxPorts: {0}", maxPorts));

                Byte[] buffer = new Byte[maxPorts];

                ComDBGetCurrentPortUsage(HComDB, buffer, maxPorts, CDB_REPORT.BYTES, out maxPorts);

                // StringBuilder sb = new StringBuilder();
                for (Int32 i = 0; i < maxPorts; i++)
                {
                    if (buffer[i] != 0)
                    {
                        Boolean found = false;

                        foreach (ListViewItem lvi in listView1.Items)
                        {
                            if (lvi.Tag.Equals(i + 1))
                            {
                                found = true;
                                break;
                            }
                        }

                        if (!found)
                        {
                            ListViewItem lvi = new ListViewItem((i + 1).ToString());
                            lvi.SubItems.Add(String.Format("COM{0}", i + 1));
                            lvi.Tag = i + 1;
                            lvi.Checked = false;

                            listView1.Items.Add(lvi);
                        }
                    }
                }

                // MessageBox.Show(sb.ToString());

                ComDBClose(HComDB);
            }
            else
            {
                Int32 err = Marshal.GetLastWin32Error();
                String errorMessage = new Win32Exception(ret).Message;
                MessageBox.Show(String.Format("ComDBOpen:\r\n\r\n0x{0} - {1}", (ret & 0x0000FFFF).ToString("X4"), errorMessage));
                //MessageBox.Show(String.Format("ComDBOpen: 0x{0}", (ret & 0x0000FFFF).ToString("X4")));
            }
        }

        /// <summary>
        /// Releases the reserved ports.
        /// </summary>
        internal void ReleaseReservedPorts()
        {
            //! Must run Elevated!!
            //
            IntPtr HComDB = IntPtr.Zero;

            Int32 ret = ComDBOpen(out HComDB);

            if ((ret & 0x0000FFFF) != ERROR_ACCESS_DENIED)
            {
                foreach (ListViewItem lvi in listView1.Items)
                {
                    Int32 port = (Int32)lvi.Tag;
                    Boolean check = lvi.Checked;

                    if (
                        lvi.SubItems[0].Text.Equals(String.Format("{0}", lvi.Tag)) &&
                        lvi.SubItems[1].Text.Equals(String.Format("COM{0}", lvi.Tag)) &&
                        !check
                        )
                    {
                        Debug.Print("Releasing COM:{0}", port);

                        ComDBReleasePort(HComDB, port);

                        // break;
                    }
                }

                ComDBClose(HComDB);

                ProcessReservations();
            }
            else
            {
                Int32 err = Marshal.GetLastWin32Error();
                String errorMessage = new Win32Exception(ret).Message;
                MessageBox.Show(String.Format("ComDBOpen:\r\n\r\n0x{0} - {1}", (ret & 0x0000FFFF).ToString("X4"), errorMessage));
            }
        }

        /// <summary>
        /// Releases the reserved port described by port.
        /// </summary>
        ///
        /// <param name="port"> The port. </param>
        internal void ReleaseReservedPort(Int32 port)
        {
            //! Must run Elevated!!
            //
            IntPtr HComDB = IntPtr.Zero;

            Int32 ret = ComDBOpen(out HComDB);

            if ((ret & 0x0000FFFF) != ERROR_ACCESS_DENIED)
            {
                Debug.Print("Releasing COM:{0}", port);

                ComDBReleasePort(HComDB, port);

                ComDBClose(HComDB);

                ProcessReservations();
            }
            else
            {
                Int32 err = Marshal.GetLastWin32Error();
                String errorMessage = new Win32Exception(ret).Message;
                MessageBox.Show(String.Format("ComDBOpen:\r\n\r\n0x{0} - {1}", (ret & 0x0000FFFF).ToString("X4"), errorMessage));
            }
        }

        /// <summary>
        /// Com database claim next free port.
        /// </summary>
        ///
        /// <param name="hComDB">    The com database. </param>
        /// <param name="ComNumber"> The com number. </param>
        ///
        /// <returns>
        /// An Int64.
        /// </returns>
        [DllImport("msports.dll", SetLastError = true)]
        private static extern Int32 ComDBClaimNextFreePort([In] IntPtr hComDB, out Int32 ComNumber);

        /// <summary>
        /// Com database claim port.
        /// </summary>
        ///
        /// <param name="hComDB">     The com database. </param>
        /// <param name="ComNumber">  The com number. </param>
        /// <param name="ForceClaim"> true to force claim. </param>
        /// <param name="Forced">     The forced. </param>
        ///
        /// <returns>
        /// An Int64.
        /// </returns>
        [DllImport("msports.dll", SetLastError = true)]
        private static extern Int32 ComDBClaimPort(IntPtr hComDB, Int32 ComNumber, [MarshalAs(UnmanagedType.Bool)] Boolean ForceClaim, [MarshalAs(UnmanagedType.Bool)] out Boolean Forced);

        /// <summary>
        /// Com database close.
        /// </summary>
        ///
        /// <param name="hComDB"> The com database. </param>
        ///
        /// <returns>
        /// An Int64.
        /// </returns>
        [DllImport("msports.dll", SetLastError = true)]
        private static extern Int32 ComDBClose(IntPtr hComDB);

        /// <summary>
        /// Com database get current port usage.
        /// </summary>
        ///
        /// <param name="HComDB">           The com database. </param>
        /// <param name="buffer">           The buffer. </param>
        /// <param name="bufferSize">       Size of the buffer. </param>
        /// <param name="reportType">       Type of the report. </param>
        /// <param name="maxPortsReported"> The maximum ports reported. </param>
        ///
        /// <returns>
        /// An Int64.
        /// </returns>
        [DllImport("msports.dll", SetLastError = true)]
        private static extern Int32 ComDBGetCurrentPortUsage(
            [In] IntPtr HComDB,
            [In, Out] byte[] buffer,
            [In] Int32 bufferSize,
            [In] CDB_REPORT reportType,
            [Out] out Int32 maxPortsReported);

        /// <summary>
        /// Queries if a given com database open.
        /// </summary>
        ///
        /// <param name="hComDB"> The com database. </param>
        ///
        /// <returns>
        /// An Int64.
        /// </returns>
        [DllImport("msports.dll", SetLastError = true)]
        private static extern Int32 ComDBOpen(out IntPtr hComDB);

        /// <summary>
        /// Com database release port.
        /// </summary>
        ///
        /// <param name="hComDB">    The com database. </param>
        /// <param name="ComNumber"> The com number. </param>
        ///
        /// <returns>
        /// An Int64.
        /// </returns>
        [DllImport("msports.dll", SetLastError = true)]
        private static extern Int32 ComDBReleasePort(IntPtr hComDB, Int32 ComNumber);

        /// <summary>
        /// Gets the last error.
        /// 
        /// Or managed: System.Runtime.InteropServices.Marshal.GetLastWin32Error.
        /// </summary>
        ///
        /// <returns>
        /// The last error.
        /// </returns>
        [DllImport("kernel32.dll")]
        private static extern Int32 GetLastError();

        /// <summary>
        /// Event handler. Called by button1 for click events.
        /// </summary>
        ///
        /// <param name="sender"> Source of the event. </param>
        /// <param name="e">      Event information. </param>
        private void button1_Click(object sender, EventArgs e)
        {
            ReleaseReservedPorts();
        }

        /// <summary>
        /// Event handler. Called by button2 for click events.
        /// </summary>
        ///
        /// <param name="sender"> Source of the event. </param>
        /// <param name="e">      Event information. </param>
        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        /// <summary>
        /// Process the reservations.
        /// </summary>
        private void ProcessReservations()
        {
            listView1.Items.Clear();

            foreach (KeyValuePair<String, Int32> kvp in Buster.SerialPorts)
            {
                ListViewItem lvi = new ListViewItem(kvp.Value.ToString());
                lvi.SubItems.Add(kvp.Key.ToString());
                lvi.Tag = kvp.Value;
                lvi.Checked = true;

                listView1.Items.Add(lvi);
            }

            AddReservedPorts();

            listView1.Columns[1].Width = -1;

            listView1.ListViewItemSorter = new PortComparer();
            listView1.Sort();
        }

        /// <summary>
        /// Event handler. Called by releaseToolStripMenuItem for click events.
        /// </summary>
        ///
        /// <param name="sender"> Source of the event. </param>
        /// <param name="e">      Event information. </param>
        private void releaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 1 && !listView1.SelectedItems[0].Checked)
            {
                Int32 port = (Int32)listView1.SelectedItems[0].Tag;

                ReleaseReservedPort(port);
            }
        }

        /// <summary>
        /// Event handler. Called by contextMenuStrip1 for opening events.
        /// </summary>
        ///
        /// <param name="sender"> Source of the event. </param>
        /// <param name="e">      Cancel event information. </param>
        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            releaseToolStripMenuItem.Enabled =
                listView1.SelectedItems.Count == 1 &&
                listView1.SelectedItems[0].SubItems[0].Text.Equals(String.Format("{0}", listView1.SelectedItems[0].Tag)) &&
                listView1.SelectedItems[0].SubItems[1].Text.Equals(String.Format("COM{0}", listView1.SelectedItems[0].Tag)) &&
                !listView1.SelectedItems[0].Checked;
        }

        /// <summary>
        /// Event handler. Called by listView1 for item check events.
        /// </summary>
        ///
        /// <param name="sender"> Source of the event. </param>
        /// <param name="e">      Item check event information. </param>
        private void listView1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            ListViewItem lvi = listView1.Items[e.Index];

            if (!(
                lvi.SubItems[0].Text.Equals(String.Format("{0}", lvi.Tag)) &&
                lvi.SubItems[1].Text.Equals(String.Format("COM{0}", lvi.Tag))
                ))
            {
                e.NewValue = CheckState.Checked;
            }
        }

        #endregion Methods

        #region Nested Types

        /// <summary>
        /// A port comparer.
        /// </summary>
        public class PortComparer : IComparer
        {
            #region Methods

            /// <summary>
            /// Calls CaseInsensitiveComparer.Compare with the parameters reversed.
            /// </summary>
            ///
            /// <param name="x"> The first object to compare. </param>
            /// <param name="y"> The second object to compare. </param>
            ///
            /// <returns>
            /// Negative if 'x' is less than 'y', 0 if they are equal, or positive if it is greater.
            /// </returns>
            int IComparer.Compare(Object x, Object y)
            {
                ListViewItem lvix = x as ListViewItem;
                ListViewItem lviy = y as ListViewItem;

                return ((Int32)lvix.Tag).CompareTo((Int32)(lviy.Tag));
            }

            #endregion Methods
        }

        #endregion Nested Types
    }
}