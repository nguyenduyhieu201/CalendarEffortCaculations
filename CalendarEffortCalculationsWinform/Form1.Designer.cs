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
            button1 = new Button();
            filepath = new TextBox();
            colorDialog1 = new ColorDialog();
            Browser = new Button();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(177, 73);
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
            filepath.Location = new Point(12, 24);
            filepath.Name = "filepath";
            filepath.Size = new Size(439, 27);
            filepath.TabIndex = 1;
            filepath.TextChanged += filepath_TextChanged;
            filepath.DragDrop += filePath_DragDrop;
            filepath.DragEnter += firePath_DragEnter;
            // 
            // Browser
            // 
            Browser.Location = new Point(469, 22);
            Browser.Name = "Browser";
            Browser.Size = new Size(94, 29);
            Browser.TabIndex = 2;
            Browser.Text = "Browser";
            Browser.UseVisualStyleBackColor = true;
            Browser.Click += BrowserButton_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(575, 134);
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
    }
}