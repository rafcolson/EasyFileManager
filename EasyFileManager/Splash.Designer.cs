namespace EasyFileManager
{
    partial class Splash
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
            InfotableLayoutPanel = new TableLayoutPanel();
            VersionLabel = new Label();
            TitleLabel = new Label();
            CopyrightLinkLabel = new LinkLabel();
            InfotableLayoutPanel.SuspendLayout();
            SuspendLayout();
            // 
            // InfotableLayoutPanel
            // 
            InfotableLayoutPanel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            InfotableLayoutPanel.AutoSize = true;
            InfotableLayoutPanel.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            InfotableLayoutPanel.BackColor = Color.Transparent;
            InfotableLayoutPanel.ColumnCount = 1;
            InfotableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            InfotableLayoutPanel.Controls.Add(VersionLabel, 0, 1);
            InfotableLayoutPanel.Controls.Add(TitleLabel, 0, 0);
            InfotableLayoutPanel.Controls.Add(CopyrightLinkLabel, 0, 2);
            InfotableLayoutPanel.Location = new Point(386, 234);
            InfotableLayoutPanel.Name = "InfotableLayoutPanel";
            InfotableLayoutPanel.RowCount = 3;
            InfotableLayoutPanel.RowStyles.Add(new RowStyle());
            InfotableLayoutPanel.RowStyles.Add(new RowStyle());
            InfotableLayoutPanel.RowStyles.Add(new RowStyle());
            InfotableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
            InfotableLayoutPanel.Size = new Size(82, 114);
            InfotableLayoutPanel.TabIndex = 0;
            // 
            // VersionLabel
            // 
            VersionLabel.AutoSize = true;
            VersionLabel.Dock = DockStyle.Fill;
            VersionLabel.Enabled = false;
            VersionLabel.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Regular, GraphicsUnit.Point);
            VersionLabel.Location = new Point(8, 50);
            VersionLabel.Margin = new Padding(8);
            VersionLabel.Name = "VersionLabel";
            VersionLabel.Size = new Size(66, 19);
            VersionLabel.TabIndex = 1;
            VersionLabel.Text = "[Version]";
            VersionLabel.TextAlign = ContentAlignment.MiddleRight;
            VersionLabel.UseCompatibleTextRendering = true;
            // 
            // TitleLabel
            // 
            TitleLabel.AutoSize = true;
            TitleLabel.BackColor = Color.Transparent;
            TitleLabel.Dock = DockStyle.Fill;
            TitleLabel.Enabled = false;
            TitleLabel.Font = new Font("Microsoft Sans Serif", 18F, FontStyle.Bold, GraphicsUnit.Point);
            TitleLabel.Location = new Point(4, 4);
            TitleLabel.Margin = new Padding(4);
            TitleLabel.Name = "TitleLabel";
            TitleLabel.Size = new Size(74, 34);
            TitleLabel.TabIndex = 0;
            TitleLabel.Text = "[Title]";
            TitleLabel.TextAlign = ContentAlignment.MiddleRight;
            TitleLabel.UseCompatibleTextRendering = true;
            // 
            // CopyrightLinkLabel
            // 
            CopyrightLinkLabel.ActiveLinkColor = Color.White;
            CopyrightLinkLabel.AutoSize = true;
            CopyrightLinkLabel.Dock = DockStyle.Fill;
            CopyrightLinkLabel.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            CopyrightLinkLabel.LinkBehavior = LinkBehavior.HoverUnderline;
            CopyrightLinkLabel.LinkColor = Color.AliceBlue;
            CopyrightLinkLabel.Location = new Point(8, 85);
            CopyrightLinkLabel.Margin = new Padding(8);
            CopyrightLinkLabel.Name = "CopyrightLinkLabel";
            CopyrightLinkLabel.Size = new Size(66, 21);
            CopyrightLinkLabel.TabIndex = 4;
            CopyrightLinkLabel.TabStop = true;
            CopyrightLinkLabel.Text = "[Copyright]";
            CopyrightLinkLabel.TextAlign = ContentAlignment.MiddleRight;
            CopyrightLinkLabel.UseCompatibleTextRendering = true;
            // 
            // Splash
            // 
            AutoScaleMode = AutoScaleMode.None;
            BackColor = SystemColors.ControlDark;
            BackgroundImage = Properties.Resources.Splash;
            BackgroundImageLayout = ImageLayout.Center;
            ClientSize = new Size(480, 360);
            ControlBox = false;
            Controls.Add(InfotableLayoutPanel);
            DoubleBuffered = true;
            Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point);
            FormBorderStyle = FormBorderStyle.None;
            Margin = new Padding(4);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Splash";
            ShowIcon = false;
            ShowInTaskbar = false;
            SizeGripStyle = SizeGripStyle.Hide;
            StartPosition = FormStartPosition.CenterScreen;
            InfotableLayoutPanel.ResumeLayout(false);
            InfotableLayoutPanel.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TableLayoutPanel InfotableLayoutPanel;
        private Label TitleLabel;
        private Label VersionLabel;
        private LinkLabel CopyrightLinkLabel;
    }
}