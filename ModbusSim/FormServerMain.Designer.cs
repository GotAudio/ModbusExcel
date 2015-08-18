using System;
using System.Windows.Forms;

namespace ModbusSim
{
    partial class FormServerMain
    {

        class DoubleBufferedPanel : System.Windows.Forms.Panel
        {
            public DoubleBufferedPanel()
            {
                this.SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer, true);
                //SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            }
            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
            }
        }

        class DoubleBufferedTabControl : System.Windows.Forms.TabControl
        {
            public DoubleBufferedTabControl()
            {
//                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
                this.SetStyle(ControlStyles.UserPaint | ControlStyles.DoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
            }

            protected override void OnPaint(PaintEventArgs pevent)
            {
                this.SetStyle(ControlStyles.UserPaint, false);
                base.OnPaint(pevent);
                System.Drawing.Rectangle o = pevent.ClipRectangle;
                System.Drawing.Graphics.FromImage(buffer).Clear(System.Drawing.SystemColors.Control);
                if (o.Width > 0 && o.Height > 0)
                {
                    DrawToBitmap(buffer, new System.Drawing.Rectangle(0, 0, Width, o.Height));
                }
                pevent.Graphics.DrawImageUnscaled(buffer, 0, 0);
                this.SetStyle(ControlStyles.UserPaint, true);
            }

            protected override void OnResize(EventArgs e)
            {
                base.OnResize(e);
                buffer = new System.Drawing.Bitmap(Width, Height);
            }
            private System.Drawing.Bitmap buffer;
        }

        class DoubleBufferedTabPage : System.Windows.Forms.TabPage
        {
            public DoubleBufferedTabPage()
            {
//                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
                this.SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer, true);
            }
        }
       
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormServerMain));
            this.textBoxPort = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxLog = new System.Windows.Forms.TextBox();
            this.buttonStop = new System.Windows.Forms.Button();
            this.buttonStart = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel3 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel4 = new System.Windows.Forms.ToolStripStatusLabel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.checkbox_showevents = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBoxPort
            // 
            this.textBoxPort.Location = new System.Drawing.Point(38, 13);
            this.textBoxPort.Name = "textBoxPort";
            this.textBoxPort.Size = new System.Drawing.Size(43, 20);
            this.textBoxPort.TabIndex = 22;
            this.textBoxPort.Text = "1502";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(4, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(26, 13);
            this.label2.TabIndex = 32;
            this.label2.Text = "Port";
            // 
            // textBoxLog
            // 
            this.textBoxLog.Location = new System.Drawing.Point(0, 667);
            this.textBoxLog.Multiline = true;
            this.textBoxLog.Name = "textBoxLog";
            this.textBoxLog.ReadOnly = true;
            this.textBoxLog.Size = new System.Drawing.Size(789, 59);
            this.textBoxLog.TabIndex = 31;
            // 
            // buttonStop
            // 
            this.buttonStop.Enabled = false;
            this.buttonStop.Location = new System.Drawing.Point(190, 11);
            this.buttonStop.Name = "buttonStop";
            this.buttonStop.Size = new System.Drawing.Size(98, 23);
            this.buttonStop.TabIndex = 24;
            this.buttonStop.Text = "Sto&p Listening";
            this.buttonStop.UseVisualStyleBackColor = true;
            this.buttonStop.Click += new System.EventHandler(this.ButtonStopClick);
            // 
            // buttonStart
            // 
            this.buttonStart.Location = new System.Drawing.Point(89, 11);
            this.buttonStart.Name = "buttonStart";
            this.buttonStart.Size = new System.Drawing.Size(95, 23);
            this.buttonStart.TabIndex = 23;
            this.buttonStart.Text = "Start &Listening";
            this.buttonStart.UseVisualStyleBackColor = true;
            this.buttonStart.Click += new System.EventHandler(this.ButtonStartClick1);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.checkbox_showevents);
            this.groupBox1.Controls.Add(this.pictureBox1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.buttonStart);
            this.groupBox1.Controls.Add(this.buttonStop);
            this.groupBox1.Controls.Add(this.textBoxPort);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(793, 40);
            this.groupBox1.TabIndex = 37;
            this.groupBox1.TabStop = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::ModbusSim.Properties.Resources.cr;
            this.pictureBox1.Location = new System.Drawing.Point(594, 8);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(148, 29);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 33;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "blue_circle.jpg");
            this.imageList1.Images.SetKeyName(1, "green_circle.jpg");
            this.imageList1.Images.SetKeyName(2, "red_circle.jpg");
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel3,
            this.toolStripStatusLabel4});
            this.statusStrip1.Location = new System.Drawing.Point(0, 729);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(793, 22);
            this.statusStrip1.TabIndex = 39;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel3
            // 
            this.toolStripStatusLabel3.Name = "toolStripStatusLabel3";
            this.toolStripStatusLabel3.Size = new System.Drawing.Size(0, 17);
            // 
            // toolStripStatusLabel4
            // 
            this.toolStripStatusLabel4.Name = "toolStripStatusLabel4";
            this.toolStripStatusLabel4.Size = new System.Drawing.Size(0, 17);
            // 
            // panel1
            // 
            this.panel1.AutoSize = true;
            this.panel1.BackgroundImage = global::ModbusSim.Images._930_grid_unchecked;
            this.panel1.Location = new System.Drawing.Point(12, 61);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(767, 585);
            this.panel1.TabIndex = 40;
            // 
            // checkbox_showevents
            // 
            this.checkbox_showevents.AutoSize = true;
            this.checkbox_showevents.Checked = true;
            this.checkbox_showevents.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkbox_showevents.Location = new System.Drawing.Point(294, 17);
            this.checkbox_showevents.Name = "checkbox_showevents";
            this.checkbox_showevents.Size = new System.Drawing.Size(86, 17);
            this.checkbox_showevents.TabIndex = 34;
            this.checkbox_showevents.Text = "ShowEvents";
            this.checkbox_showevents.UseVisualStyleBackColor = true;
            // 
            // FormServerMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(793, 751);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.textBoxLog);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormServerMain";
            this.Text = "ModbusSim";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxPort;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxLog;
        private System.Windows.Forms.Button buttonStop;
        private System.Windows.Forms.Button buttonStart;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ImageList imageList1;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel toolStripStatusLabel3;
        private ToolStripStatusLabel toolStripStatusLabel4;
        private PictureBox pictureBox1;
        private Panel panel1;
        private CheckBox checkbox_showevents;
    }
}

