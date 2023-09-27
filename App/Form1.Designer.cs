namespace ReadExcelFileApp
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            btnChoose = new Button();
            button1 = new Button();
            button2 = new Button();
            textBox1 = new TextBox();
            button3 = new Button();
            button4 = new Button();
            SuspendLayout();
            // 
            // btnChoose
            // 
            btnChoose.Location = new Point(11, 11);
            btnChoose.Margin = new Padding(2);
            btnChoose.Name = "btnChoose";
            btnChoose.Size = new Size(82, 23);
            btnChoose.TabIndex = 0;
            btnChoose.Text = "Choose File";
            btnChoose.UseVisualStyleBackColor = true;
            btnChoose.Click += btnChoose_Click;
            // 
            // button1
            // 
            button1.Enabled = false;
            button1.Location = new Point(109, 11);
            button1.Name = "button1";
            button1.Size = new Size(97, 23);
            button1.TabIndex = 3;
            button1.Text = "convert to XML";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // button2
            // 
            button2.Enabled = false;
            button2.Location = new Point(221, 12);
            button2.Name = "button2";
            button2.Size = new Size(107, 23);
            button2.TabIndex = 4;
            button2.Text = "convert to SQL";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // textBox1
            // 
            textBox1.Location = new Point(11, 84);
            textBox1.Multiline = true;
            textBox1.Name = "textBox1";
            textBox1.ScrollBars = ScrollBars.Vertical;
            textBox1.Size = new Size(532, 382);
            textBox1.TabIndex = 6;
            // 
            // button3
            // 
            button3.Enabled = false;
            button3.Location = new Point(221, 41);
            button3.Name = "button3";
            button3.Size = new Size(107, 23);
            button3.TabIndex = 7;
            button3.Text = "Report";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // button4
            // 
            button4.Enabled = false;
            button4.Location = new Point(109, 41);
            button4.Name = "button4";
            button4.Size = new Size(97, 23);
            button4.TabIndex = 8;
            button4.Text = "Report";
            button4.UseVisualStyleBackColor = true;
            button4.Click += button4_Click_1;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(6F, 13F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoScroll = true;
            ClientSize = new Size(556, 484);
            Controls.Add(button4);
            Controls.Add(button3);
            Controls.Add(textBox1);
            Controls.Add(button2);
            Controls.Add(button1);
            Controls.Add(btnChoose);
            Margin = new Padding(2);
            Name = "Form1";
            Text = "Read Excel File";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnChoose;
        private Button button1;
        private Button button2;
        private TextBox textBox1;
        private Button button3;
        private Button button4;
    }
}