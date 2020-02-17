namespace WindowsSnap {
    partial class TrayIconForm {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
      this.components = new System.ComponentModel.Container();
      System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TrayIconForm));
      this.trayIcon = new System.Windows.Forms.NotifyIcon(this.components);
      this.trayMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
      this.quitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.snapshotListEndLine = new System.Windows.Forms.ToolStripSeparator();
      this.justNowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.snapshotListStartLine = new System.Windows.Forms.ToolStripSeparator();
      this.clearSnapshotsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.snapshotToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.trayMenu.SuspendLayout();
      this.SuspendLayout();
      // 
      // trayIcon
      // 
      this.trayIcon.ContextMenuStrip = this.trayMenu;
      this.trayIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("trayIcon.Icon")));
      this.trayIcon.Text = "Window Layout Backup";
      this.trayIcon.Visible = true;
      this.trayIcon.MouseClick += new System.Windows.Forms.MouseEventHandler(this.TrayIcon_MouseClick);
      // 
      // trayMenu
      // 
      this.trayMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.quitToolStripMenuItem,
            this.aboutToolStripMenuItem,
            this.snapshotListEndLine,
            this.justNowToolStripMenuItem,
            this.snapshotListStartLine,
            this.clearSnapshotsToolStripMenuItem,
            this.snapshotToolStripMenuItem});
      this.trayMenu.Name = "contextMenuStrip1";
      this.trayMenu.Size = new System.Drawing.Size(159, 126);
      // 
      // quitToolStripMenuItem
      // 
      this.quitToolStripMenuItem.Name = "quitToolStripMenuItem";
      this.quitToolStripMenuItem.Size = new System.Drawing.Size(158, 22);
      this.quitToolStripMenuItem.Text = "Quit";
      this.quitToolStripMenuItem.Click += new System.EventHandler(this.QuitToolStripMenuItem_Click);
      // 
      // aboutToolStripMenuItem
      // 
      this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
      this.aboutToolStripMenuItem.Size = new System.Drawing.Size(158, 22);
      this.aboutToolStripMenuItem.Text = "About";
      this.aboutToolStripMenuItem.Click += new System.EventHandler(this.AboutToolStripMenuItem_Click);
      // 
      // snapshotListEndLine
      // 
      this.snapshotListEndLine.Name = "snapshotListEndLine";
      this.snapshotListEndLine.Size = new System.Drawing.Size(155, 6);
      // 
      // justNowToolStripMenuItem
      // 
      this.justNowToolStripMenuItem.Name = "justNowToolStripMenuItem";
      this.justNowToolStripMenuItem.Size = new System.Drawing.Size(158, 22);
      this.justNowToolStripMenuItem.Text = "(Just now)";
      this.justNowToolStripMenuItem.Click += new System.EventHandler(this.JustNowToolStripMenuItem_Click);
      this.justNowToolStripMenuItem.MouseEnter += new System.EventHandler(this.JustNowToolStripMenuItem_MouseEnter);
      // 
      // snapshotListStartLine
      // 
      this.snapshotListStartLine.Name = "snapshotListStartLine";
      this.snapshotListStartLine.Size = new System.Drawing.Size(155, 6);
      // 
      // clearSnapshotsToolStripMenuItem
      // 
      this.clearSnapshotsToolStripMenuItem.Name = "clearSnapshotsToolStripMenuItem";
      this.clearSnapshotsToolStripMenuItem.Size = new System.Drawing.Size(158, 22);
      this.clearSnapshotsToolStripMenuItem.Text = "Clear Snapshots";
      this.clearSnapshotsToolStripMenuItem.Click += new System.EventHandler(this.ClearSnapshotsToolStripMenuItem_Click);
      // 
      // snapshotToolStripMenuItem
      // 
      this.snapshotToolStripMenuItem.Name = "snapshotToolStripMenuItem";
      this.snapshotToolStripMenuItem.Size = new System.Drawing.Size(158, 22);
      this.snapshotToolStripMenuItem.Text = "Take Snapshot";
      this.snapshotToolStripMenuItem.Click += new System.EventHandler(this.SnapshotToolStripMenuItem_Click);
      // 
      // TrayIconForm
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(181, 82);
      this.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
      this.Name = "TrayIconForm";
      this.Text = "Form1";
      this.VisibleChanged += new System.EventHandler(this.TrayIconForm_VisibleChanged);
      this.trayMenu.ResumeLayout(false);
      this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.NotifyIcon trayIcon;
        private System.Windows.Forms.ContextMenuStrip trayMenu;
        private System.Windows.Forms.ToolStripMenuItem snapshotToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator snapshotListStartLine;
        private System.Windows.Forms.ToolStripMenuItem quitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator snapshotListEndLine;
        private System.Windows.Forms.ToolStripMenuItem clearSnapshotsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem justNowToolStripMenuItem;
    }
}

