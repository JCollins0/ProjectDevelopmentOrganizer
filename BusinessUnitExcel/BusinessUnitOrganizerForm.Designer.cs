namespace BusinessUnitExcel
{
    partial class BusinessUnitOrganizerForm
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
            this.button_open_file = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.textbox_column_start = new System.Windows.Forms.TextBox();
            this.button_process_data = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.textbox_row_start = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textbox_last_row = new System.Windows.Forms.TextBox();
            this.textbox_log = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button_open_file
            // 
            this.button_open_file.Location = new System.Drawing.Point(131, 37);
            this.button_open_file.Name = "button_open_file";
            this.button_open_file.Size = new System.Drawing.Size(75, 23);
            this.button_open_file.TabIndex = 0;
            this.button_open_file.Text = "Open File";
            this.button_open_file.UseVisualStyleBackColor = true;
            this.button_open_file.Click += new System.EventHandler(this.button_open_file_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.Filter = "Excel Files|*.xlsx|CSV files|*.csv";
            // 
            // textbox_column_start
            // 
            this.textbox_column_start.Location = new System.Drawing.Point(131, 92);
            this.textbox_column_start.MaxLength = 6;
            this.textbox_column_start.Name = "textbox_column_start";
            this.textbox_column_start.Size = new System.Drawing.Size(75, 20);
            this.textbox_column_start.TabIndex = 3;
            this.textbox_column_start.WordWrap = false;
            this.textbox_column_start.TextChanged += new System.EventHandler(this.textbox_column_start_TextChanged);
            // 
            // button_process_data
            // 
            this.button_process_data.Enabled = false;
            this.button_process_data.Location = new System.Drawing.Point(111, 144);
            this.button_process_data.Name = "button_process_data";
            this.button_process_data.Size = new System.Drawing.Size(95, 23);
            this.button_process_data.TabIndex = 4;
            this.button_process_data.Text = "Process Data";
            this.button_process_data.UseVisualStyleBackColor = true;
            this.button_process_data.Click += new System.EventHandler(this.button_process_data_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(25, 95);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Column To Start On";
            // 
            // textbox_row_start
            // 
            this.textbox_row_start.Location = new System.Drawing.Point(133, 66);
            this.textbox_row_start.Name = "textbox_row_start";
            this.textbox_row_start.Size = new System.Drawing.Size(73, 20);
            this.textbox_row_start.TabIndex = 6;
            this.textbox_row_start.TextChanged += new System.EventHandler(this.textbox_row_start_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(47, 69);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Header Row";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(47, 121);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(52, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Last Row";
            // 
            // textbox_last_row
            // 
            this.textbox_last_row.Location = new System.Drawing.Point(132, 118);
            this.textbox_last_row.Name = "textbox_last_row";
            this.textbox_last_row.Size = new System.Drawing.Size(74, 20);
            this.textbox_last_row.TabIndex = 9;
            this.textbox_last_row.TextChanged += new System.EventHandler(this.textbox_last_row_TextChanged);
            // 
            // textbox_log
            // 
            this.textbox_log.Location = new System.Drawing.Point(12, 204);
            this.textbox_log.Multiline = true;
            this.textbox_log.Name = "textbox_log";
            this.textbox_log.ReadOnly = true;
            this.textbox_log.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textbox_log.Size = new System.Drawing.Size(217, 101);
            this.textbox_log.TabIndex = 10;
            this.textbox_log.WordWrap = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(100, 188);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(25, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "Log";
            // 
            // BusinessUnitOrganizerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(241, 317);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textbox_log);
            this.Controls.Add(this.textbox_last_row);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textbox_row_start);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button_process_data);
            this.Controls.Add(this.textbox_column_start);
            this.Controls.Add(this.button_open_file);
            this.Name = "BusinessUnitOrganizerForm";
            this.Text = "PD Report Generator";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.BusinessUnitOrganizerForm_FormClosed);
            this.Load += new System.EventHandler(this.BusinessUnitOrganizerForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button_open_file;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.TextBox textbox_column_start;
        private System.Windows.Forms.Button button_process_data;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textbox_row_start;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textbox_last_row;
        private System.Windows.Forms.TextBox textbox_log;
        private System.Windows.Forms.Label label4;
    }
}

