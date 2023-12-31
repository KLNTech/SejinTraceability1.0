﻿namespace SejinTraceability
{
    partial class TraceabilityForm
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
            this.trace = new System.Windows.Forms.Label();
            this.textBoxtrace = new System.Windows.Forms.TextBox();
            this.rack_qty = new System.Windows.Forms.Label();
            this.textBoxrackqty = new System.Windows.Forms.TextBox();
            this.rack = new System.Windows.Forms.Label();
            this.textBoxrack = new System.Windows.Forms.TextBox();
            this.print_archive = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxtrace2 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxrack2 = new System.Windows.Forms.TextBox();
            this.textBoxPN = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.checkBoxStripping = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.Export = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // trace
            // 
            this.trace.AutoSize = true;
            this.trace.Location = new System.Drawing.Point(12, 18);
            this.trace.Name = "trace";
            this.trace.Size = new System.Drawing.Size(112, 15);
            this.trace.TabIndex = 0;
            this.trace.Text = "Welding traceability";
            // 
            // textBoxtrace
            // 
            this.textBoxtrace.Location = new System.Drawing.Point(12, 45);
            this.textBoxtrace.Name = "textBoxtrace";
            this.textBoxtrace.Size = new System.Drawing.Size(372, 23);
            this.textBoxtrace.TabIndex = 1;
            this.textBoxtrace.TextChanged += new System.EventHandler(this.TextBox_TextChanged);
            // 
            // rack_qty
            // 
            this.rack_qty.AutoSize = true;
            this.rack_qty.Location = new System.Drawing.Point(230, 158);
            this.rack_qty.Name = "rack_qty";
            this.rack_qty.Size = new System.Drawing.Size(52, 15);
            this.rack_qty.TabIndex = 2;
            this.rack_qty.Text = "Rack qty";
            // 
            // textBoxrackqty
            // 
            this.textBoxrackqty.Location = new System.Drawing.Point(230, 187);
            this.textBoxrackqty.Name = "textBoxrackqty";
            this.textBoxrackqty.Size = new System.Drawing.Size(52, 23);
            this.textBoxrackqty.TabIndex = 3;
            this.textBoxrackqty.TextChanged += new System.EventHandler(this.TextBox_TextChanged);
            // 
            // rack
            // 
            this.rack.AutoSize = true;
            this.rack.Location = new System.Drawing.Point(12, 223);
            this.rack.Name = "rack";
            this.rack.Size = new System.Drawing.Size(32, 15);
            this.rack.TabIndex = 4;
            this.rack.Text = "Rack";
            // 
            // textBoxrack
            // 
            this.textBoxrack.Location = new System.Drawing.Point(12, 248);
            this.textBoxrack.Name = "textBoxrack";
            this.textBoxrack.Size = new System.Drawing.Size(67, 23);
            this.textBoxrack.TabIndex = 5;
            this.textBoxrack.TextChanged += new System.EventHandler(this.TextBox_TextChanged);
            // 
            // print_archive
            // 
            this.print_archive.Location = new System.Drawing.Point(230, 221);
            this.print_archive.Name = "print_archive";
            this.print_archive.Size = new System.Drawing.Size(154, 50);
            this.print_archive.TabIndex = 6;
            this.print_archive.Text = "Print label and archive the record";
            this.print_archive.UseVisualStyleBackColor = true;
            this.print_archive.Click += new System.EventHandler(this.OpenAndPrintExcelFileHandler);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 84);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(112, 15);
            this.label1.TabIndex = 7;
            this.label1.Text = "Welding traceability";
            // 
            // textBoxtrace2
            // 
            this.textBoxtrace2.Location = new System.Drawing.Point(12, 113);
            this.textBoxtrace2.Name = "textBoxtrace2";
            this.textBoxtrace2.Size = new System.Drawing.Size(372, 23);
            this.textBoxtrace2.TabIndex = 8;
            this.textBoxtrace2.TextChanged += new System.EventHandler(this.TextBox_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(117, 223);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 15);
            this.label2.TabIndex = 9;
            this.label2.Text = "Rack 2";
            // 
            // textBoxrack2
            // 
            this.textBoxrack2.Location = new System.Drawing.Point(117, 248);
            this.textBoxrack2.Name = "textBoxrack2";
            this.textBoxrack2.Size = new System.Drawing.Size(67, 23);
            this.textBoxrack2.TabIndex = 10;
            this.textBoxrack2.TextChanged += new System.EventHandler(this.TextBox_TextChanged);
            // 
            // textBoxPN
            // 
            this.textBoxPN.Location = new System.Drawing.Point(12, 187);
            this.textBoxPN.Name = "textBoxPN";
            this.textBoxPN.Size = new System.Drawing.Size(183, 23);
            this.textBoxPN.TabIndex = 11;
            this.textBoxPN.TextChanged += new System.EventHandler(this.TextBox_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 158);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(75, 15);
            this.label3.TabIndex = 12;
            this.label3.Text = "Part Number";
            // 
            // checkBoxStripping
            // 
            this.checkBoxStripping.Location = new System.Drawing.Point(332, 186);
            this.checkBoxStripping.Name = "checkBoxStripping";
            this.checkBoxStripping.Size = new System.Drawing.Size(104, 24);
            this.checkBoxStripping.TabIndex = 15;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(313, 158);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(55, 15);
            this.label4.TabIndex = 14;
            this.label4.Text = "Stripping";
            // 
            // Export
            // 
            this.Export.Location = new System.Drawing.Point(413, 48);
            this.Export.Name = "Export";
            this.Export.Size = new System.Drawing.Size(75, 88);
            this.Export.TabIndex = 16;
            this.Export.Text = "Export";
            this.Export.UseVisualStyleBackColor = true;
            this.Export.Click += new System.EventHandler(this.ExportButtonClick);
            // 
            // TraceabilityForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(395, 289);
            this.Controls.Add(this.Export);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.checkBoxStripping);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBoxPN);
            this.Controls.Add(this.textBoxrack2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBoxtrace2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.print_archive);
            this.Controls.Add(this.textBoxrack);
            this.Controls.Add(this.rack);
            this.Controls.Add(this.textBoxrackqty);
            this.Controls.Add(this.rack_qty);
            this.Controls.Add(this.textBoxtrace);
            this.Controls.Add(this.trace);
            this.Name = "TraceabilityForm";
            this.Text = "TraceabilityForm";
            this.Click += new System.EventHandler(this.ExportButtonClick);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Label trace;
        private TextBox textBoxtrace;
        private Label rack_qty;
        private TextBox textBoxrackqty;
        private Label rack;
        private TextBox textBoxrack;
        private Button print_archive;
        private Label label1;
        private TextBox textBoxtrace2;
        private Label label2;
        private TextBox textBoxrack2;
        private TextBox textBoxPN;
        private Label label3;
        private CheckBox checkBoxStripping;
        private Label label4;
        private Button Export;
    }
}