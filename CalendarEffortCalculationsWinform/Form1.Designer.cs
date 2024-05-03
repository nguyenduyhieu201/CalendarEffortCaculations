namespace CalendarEffortCalculationsWinform
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            button1 = new Button();
            filepath = new TextBox();
            colorDialog1 = new ColorDialog();
            Browser = new Button();
            abNormalBox = new TextBox();
            otBox = new TextBox();
            tmsBox = new TextBox();
            contextMenuStrip1 = new ContextMenuStrip(components);
            label1 = new Label();
            label2 = new Label();
            label3 = new Label();
            label4 = new Label();
            calendarBox = new TextBox();
            label5 = new Label();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(183, 439);
            button1.Name = "button1";
            button1.Size = new Size(140, 38);
            button1.TabIndex = 0;
            button1.Text = "OK";
            button1.UseVisualStyleBackColor = true;
            button1.Click += OK_Click;
            // 
            // filepath
            // 
            filepath.AllowDrop = true;
            filepath.Location = new Point(12, 55);
            filepath.Name = "filepath";
            filepath.Size = new Size(439, 27);
            filepath.TabIndex = 1;
            filepath.TextChanged += filepath_TextChanged;
            filepath.DragDrop += filePath_DragDrop;
            filepath.DragEnter += firePath_DragEnter;
            // 
            // Browser
            // 
            Browser.Location = new Point(469, 53);
            Browser.Name = "Browser";
            Browser.Size = new Size(94, 29);
            Browser.TabIndex = 2;
            Browser.Text = "Browser";
            Browser.UseVisualStyleBackColor = true;
            Browser.Click += BrowserButton_Click;
            // 
            // abNormalBox
            // 
            abNormalBox.Location = new Point(12, 129);
            abNormalBox.Name = "abNormalBox";
            abNormalBox.Size = new Size(439, 27);
            abNormalBox.TabIndex = 3;
            abNormalBox.TextChanged += nameOfAbnormalBox_TextChanged;
            // 
            // otBox
            // 
            otBox.Location = new Point(12, 210);
            otBox.Name = "otBox";
            otBox.Size = new Size(439, 27);
            otBox.TabIndex = 4;
            otBox.TextChanged += nameOfOTBox_TextChanged;
            // 
            // tmsBox
            // 
            tmsBox.Location = new Point(12, 285);
            tmsBox.Name = "tmsBox";
            tmsBox.Size = new Size(439, 27);
            tmsBox.TabIndex = 5;
            tmsBox.TextChanged += nameOfTMSBox_TextChanged;
            // 
            // contextMenuStrip1
            // 
            contextMenuStrip1.ImageScalingSize = new Size(20, 20);
            contextMenuStrip1.Name = "contextMenuStrip1";
            contextMenuStrip1.Size = new Size(61, 4);
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            label1.Location = new Point(12, 96);
            label1.Name = "label1";
            label1.Size = new Size(372, 20);
            label1.TabIndex = 7;
            label1.Text = "Name of abnormal sheet (Default is AbnormalCase)";
            label1.Click += label1_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            label2.Location = new Point(12, 176);
            label2.Name = "label2";
            label2.Size = new Size(241, 20);
            label2.TabIndex = 8;
            label2.Text = "Name of OT sheet (Default is OT)";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            label3.Location = new Point(12, 250);
            label3.Name = "label3";
            label3.Size = new Size(265, 20);
            label3.TabIndex = 9;
            label3.Text = "Name of TMS sheet (Default is TMS)";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            label4.Location = new Point(12, 18);
            label4.Name = "label4";
            label4.Size = new Size(157, 20);
            label4.TabIndex = 10;
            label4.Text = "Directory of the path";
            // 
            // textBox1
            // 
            calendarBox.Location = new Point(12, 367);
            calendarBox.Name = "calendarBox";
            calendarBox.Size = new Size(439, 27);
            calendarBox.TabIndex = 11;
            calendarBox.TextChanged += nameOfcalendarBox_TextChanged;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            label5.Location = new Point(12, 333);
            label5.Name = "label5";
            label5.Size = new Size(325, 20);
            label5.TabIndex = 12;
            label5.Text = "Name of Calendar sheet (Default is Calendar)";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(575, 518);
            Controls.Add(label5);
            Controls.Add(calendarBox);
            Controls.Add(label4);
            Controls.Add(label3);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(tmsBox);
            Controls.Add(otBox);
            Controls.Add(abNormalBox);
            Controls.Add(Browser);
            Controls.Add(filepath);
            Controls.Add(button1);
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            Name = "Form1";
            Text = "Man Month Caculation ";
            Load += Form1_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private TextBox filepath;
        private ColorDialog colorDialog1;
        private Button Browser;
        private TextBox abNormalBox;
        private TextBox otBox;
        private TextBox tmsBox;
        private ContextMenuStrip contextMenuStrip1;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private TextBox calendarBox;
        private Label label5;
    }
}