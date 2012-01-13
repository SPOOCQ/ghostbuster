namespace Ghostbuster
{
    partial class Mainform
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Mainform));
            this.RemoveBtn = new System.Windows.Forms.Button();
            this.listView1 = new System.Windows.Forms.ListView();
            this.DeviceCol = new System.Windows.Forms.ColumnHeader();
            this.StatusCol = new System.Windows.Forms.ColumnHeader();
            this.MatchTypeCol = new System.Windows.Forms.ColumnHeader();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.AddClassMnu = new System.Windows.Forms.ToolStripMenuItem();
            this.AddDeviceMnu = new System.Windows.Forms.ToolStripMenuItem();
            this.AddWildCardMnu = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.RemoveClassMnu = new System.Windows.Forms.ToolStripMenuItem();
            this.RemoveDeviceMnu = new System.Windows.Forms.ToolStripMenuItem();
            this.removeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cancelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.RefreshBtn = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel3 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.DescriptionCol = new System.Windows.Forms.ColumnHeader();
            this.contextMenuStrip1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // RemoveBtn
            // 
            this.RemoveBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.RemoveBtn.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.RemoveBtn.Location = new System.Drawing.Point(683, 291);
            this.RemoveBtn.Name = "RemoveBtn";
            this.RemoveBtn.Size = new System.Drawing.Size(104, 23);
            this.RemoveBtn.TabIndex = 0;
            this.RemoveBtn.Text = "Remove Ghosts";
            this.RemoveBtn.UseVisualStyleBackColor = true;
            this.RemoveBtn.Click += new System.EventHandler(this.RemoveBtn_Click);
            // 
            // listView1
            // 
            this.listView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.DeviceCol,
            this.StatusCol,
            this.MatchTypeCol,
            this.DescriptionCol});
            this.listView1.ContextMenuStrip = this.contextMenuStrip1;
            this.listView1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.listView1.FullRowSelect = true;
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(13, 12);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(774, 273);
            this.listView1.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.listView1.TabIndex = 1;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            // 
            // DeviceCol
            // 
            this.DeviceCol.Text = "Device";
            this.DeviceCol.Width = 200;
            // 
            // StatusCol
            // 
            this.StatusCol.Text = "Status";
            // 
            // MatchTypeCol
            // 
            this.MatchTypeCol.Text = "Match Type";
            this.MatchTypeCol.Width = 100;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.AddClassMnu,
            this.AddDeviceMnu,
            this.AddWildCardMnu,
            this.toolStripSeparator1,
            this.RemoveClassMnu,
            this.RemoveDeviceMnu,
            this.removeToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(168, 142);
            this.contextMenuStrip1.Opening += new System.ComponentModel.CancelEventHandler(this.contextMenuStrip1_Opening);
            // 
            // AddClassMnu
            // 
            this.AddClassMnu.Name = "AddClassMnu";
            this.AddClassMnu.Size = new System.Drawing.Size(167, 22);
            this.AddClassMnu.Text = "Add Class";
            this.AddClassMnu.Click += new System.EventHandler(this.AddClassMnu_Click);
            // 
            // AddDeviceMnu
            // 
            this.AddDeviceMnu.Name = "AddDeviceMnu";
            this.AddDeviceMnu.Size = new System.Drawing.Size(167, 22);
            this.AddDeviceMnu.Text = "Add Device";
            this.AddDeviceMnu.Click += new System.EventHandler(this.AddDeviceMnu_Click);
            // 
            // AddWildCardMnu
            // 
            this.AddWildCardMnu.Name = "AddWildCardMnu";
            this.AddWildCardMnu.Size = new System.Drawing.Size(167, 22);
            this.AddWildCardMnu.Text = "Add Wildcard";
            this.AddWildCardMnu.Click += new System.EventHandler(this.AddWildCardMnu_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(164, 6);
            // 
            // RemoveClassMnu
            // 
            this.RemoveClassMnu.Name = "RemoveClassMnu";
            this.RemoveClassMnu.Size = new System.Drawing.Size(167, 22);
            this.RemoveClassMnu.Text = "Remove Class";
            this.RemoveClassMnu.Click += new System.EventHandler(this.RemoveClassMnu_Click);
            // 
            // RemoveDeviceMnu
            // 
            this.RemoveDeviceMnu.Name = "RemoveDeviceMnu";
            this.RemoveDeviceMnu.Size = new System.Drawing.Size(167, 22);
            this.RemoveDeviceMnu.Text = "Remove Device";
            this.RemoveDeviceMnu.Click += new System.EventHandler(this.RemoveDeviceMnu_Click);
            // 
            // removeToolStripMenuItem
            // 
            this.removeToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cancelToolStripMenuItem});
            this.removeToolStripMenuItem.Name = "removeToolStripMenuItem";
            this.removeToolStripMenuItem.Size = new System.Drawing.Size(167, 22);
            this.removeToolStripMenuItem.Text = "Remove Wildcard";
            // 
            // cancelToolStripMenuItem
            // 
            this.cancelToolStripMenuItem.Name = "cancelToolStripMenuItem";
            this.cancelToolStripMenuItem.Size = new System.Drawing.Size(110, 22);
            this.cancelToolStripMenuItem.Text = "Cancel";
            this.cancelToolStripMenuItem.Click += new System.EventHandler(this.RemoveWildcardToolStripMenuItem_Click);
            // 
            // RefreshBtn
            // 
            this.RefreshBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.RefreshBtn.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.RefreshBtn.Location = new System.Drawing.Point(573, 291);
            this.RefreshBtn.Name = "RefreshBtn";
            this.RefreshBtn.Size = new System.Drawing.Size(104, 23);
            this.RefreshBtn.TabIndex = 2;
            this.RefreshBtn.Text = "Refresh Devices";
            this.RefreshBtn.UseVisualStyleBackColor = true;
            this.RefreshBtn.Click += new System.EventHandler(this.RefreshBtn_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripStatusLabel2,
            this.toolStripStatusLabel3,
            this.toolStripProgressBar1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 317);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(799, 22);
            this.statusStrip1.TabIndex = 3;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(118, 17);
            this.toolStripStatusLabel1.Text = "toolStripStatusLabel1";
            // 
            // toolStripStatusLabel2
            // 
            this.toolStripStatusLabel2.Name = "toolStripStatusLabel2";
            this.toolStripStatusLabel2.Size = new System.Drawing.Size(118, 17);
            this.toolStripStatusLabel2.Text = "toolStripStatusLabel2";
            // 
            // toolStripStatusLabel3
            // 
            this.toolStripStatusLabel3.Name = "toolStripStatusLabel3";
            this.toolStripStatusLabel3.Size = new System.Drawing.Size(118, 17);
            this.toolStripStatusLabel3.Text = "toolStripStatusLabel3";
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(100, 16);
            // 
            // DescriptionCol
            // 
            this.DescriptionCol.Text = "Description";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(799, 339);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.RefreshBtn);
            this.Controls.Add(this.RemoveBtn);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.Text = "GhostBuster";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.contextMenuStrip1.ResumeLayout(false);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button RemoveBtn;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ColumnHeader DeviceCol;
        private System.Windows.Forms.ColumnHeader StatusCol;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem AddClassMnu;
        private System.Windows.Forms.ToolStripMenuItem AddDeviceMnu;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem RemoveDeviceMnu;
        private System.Windows.Forms.ToolStripMenuItem RemoveClassMnu;
        private System.Windows.Forms.Button RefreshBtn;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel3;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
        private System.Windows.Forms.ColumnHeader MatchTypeCol;
        private System.Windows.Forms.ToolStripMenuItem AddWildCardMnu;
        private System.Windows.Forms.ToolStripMenuItem removeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cancelToolStripMenuItem;
        private System.Windows.Forms.ColumnHeader DescriptionCol;
    }
}

