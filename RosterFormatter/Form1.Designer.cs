namespace RosterFormatter
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
            btn_open = new Button();
            btn_format = new Button();
            txt_input = new TextBox();
            txt_output = new TextBox();
            label1 = new Label();
            btn_output = new Button();
            SuspendLayout();
            // 
            // btn_open
            // 
            btn_open.Location = new Point(689, 129);
            btn_open.Name = "btn_open";
            btn_open.Size = new Size(150, 29);
            btn_open.TabIndex = 0;
            btn_open.Text = "Open CSV File";
            btn_open.UseVisualStyleBackColor = true;
            btn_open.Click += btn_open_Click;
            // 
            // btn_format
            // 
            btn_format.Location = new Point(689, 371);
            btn_format.Name = "btn_format";
            btn_format.Size = new Size(94, 29);
            btn_format.TabIndex = 1;
            btn_format.Text = "Format";
            btn_format.UseVisualStyleBackColor = true;
            btn_format.Click += btn_format_Click;
            // 
            // txt_input
            // 
            txt_input.Enabled = false;
            txt_input.Location = new Point(77, 129);
            txt_input.Name = "txt_input";
            txt_input.Size = new Size(399, 27);
            txt_input.TabIndex = 2;
            // 
            // txt_output
            // 
            txt_output.Enabled = false;
            txt_output.Location = new Point(77, 256);
            txt_output.Name = "txt_output";
            txt_output.Size = new Size(399, 27);
            txt_output.TabIndex = 3;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(12, 50);
            label1.Name = "label1";
            label1.Size = new Size(904, 20);
            label1.TabIndex = 4;
            label1.Text = "Select Open CSV File and find the CSV Roster. Then, change the output location, and hit format to format the Roster to a printable state.";
            // 
            // btn_output
            // 
            btn_output.Location = new Point(689, 254);
            btn_output.Name = "btn_output";
            btn_output.Size = new Size(195, 29);
            btn_output.TabIndex = 5;
            btn_output.Text = "Select Output Location";
            btn_output.UseVisualStyleBackColor = true;
            btn_output.Click += btn_output_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(935, 450);
            Controls.Add(btn_output);
            Controls.Add(label1);
            Controls.Add(txt_output);
            Controls.Add(txt_input);
            Controls.Add(btn_format);
            Controls.Add(btn_open);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btn_open;
        private Button btn_format;
        private TextBox txt_input;
        private TextBox txt_output;
        private Label label1;
        private Button btn_output;
    }
}
