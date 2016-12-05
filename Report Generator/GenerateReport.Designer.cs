/*
    This file is part of Report Generator.

    Report Generator is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    Report Generator is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with Report Generator.  If not, see <http://www.gnu.org/licenses/>.
 */

namespace GeneratorSpace
{
    partial class GenerateReport
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
            this.btnWordBrowse = new System.Windows.Forms.Button();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.txtWordPath = new System.Windows.Forms.TextBox();
            this.txtQlikPath = new System.Windows.Forms.TextBox();
            this.btnQlikBrowse = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.grpStaticSelections = new System.Windows.Forms.GroupBox();
            this.txtStaticSelections = new System.Windows.Forms.TextBox();
            this.grpSetPaths = new System.Windows.Forms.GroupBox();
            this.lstLog = new System.Windows.Forms.ListBox();
            this.lblLog = new System.Windows.Forms.Label();
            this.grpQuickRef = new System.Windows.Forms.GroupBox();
            this.btnRemove = new System.Windows.Forms.Button();
            this.lstQuickRefVars = new System.Windows.Forms.ListBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtRefID = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnSaveRef = new System.Windows.Forms.Button();
            this.txtRefName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.ReloadQVData = new System.Windows.Forms.CheckBox();
            this.grpStaticSelections.SuspendLayout();
            this.grpSetPaths.SuspendLayout();
            this.grpQuickRef.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnWordBrowse
            // 
            this.btnWordBrowse.Location = new System.Drawing.Point(377, 44);
            this.btnWordBrowse.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnWordBrowse.Name = "btnWordBrowse";
            this.btnWordBrowse.Size = new System.Drawing.Size(104, 25);
            this.btnWordBrowse.TabIndex = 1;
            this.btnWordBrowse.Text = "Browse";
            this.btnWordBrowse.UseVisualStyleBackColor = true;
            this.btnWordBrowse.Click += new System.EventHandler(this.btnWordBrowse_Click);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(186, 498);
            this.btnGenerate.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(125, 34);
            this.btnGenerate.TabIndex = 2;
            this.btnGenerate.Text = "Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // txtWordPath
            // 
            this.txtWordPath.Enabled = false;
            this.txtWordPath.Location = new System.Drawing.Point(5, 44);
            this.txtWordPath.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtWordPath.Name = "txtWordPath";
            this.txtWordPath.Size = new System.Drawing.Size(364, 22);
            this.txtWordPath.TabIndex = 3;
            // 
            // txtQlikPath
            // 
            this.txtQlikPath.Enabled = false;
            this.txtQlikPath.Location = new System.Drawing.Point(5, 98);
            this.txtQlikPath.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtQlikPath.Name = "txtQlikPath";
            this.txtQlikPath.Size = new System.Drawing.Size(364, 22);
            this.txtQlikPath.TabIndex = 4;
            // 
            // btnQlikBrowse
            // 
            this.btnQlikBrowse.Location = new System.Drawing.Point(377, 98);
            this.btnQlikBrowse.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnQlikBrowse.Name = "btnQlikBrowse";
            this.btnQlikBrowse.Size = new System.Drawing.Size(104, 25);
            this.btnQlikBrowse.TabIndex = 5;
            this.btnQlikBrowse.Text = "Browse";
            this.btnQlikBrowse.UseVisualStyleBackColor = true;
            this.btnQlikBrowse.Click += new System.EventHandler(this.btnQlikBrowse_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(154, 17);
            this.label1.TabIndex = 6;
            this.label1.Text = "Path to Word Template";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 79);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(178, 17);
            this.label2.TabIndex = 7;
            this.label2.Text = "Path to QlikView Document";
            // 
            // grpStaticSelections
            // 
            this.grpStaticSelections.Controls.Add(this.txtStaticSelections);
            this.grpStaticSelections.Location = new System.Drawing.Point(11, 386);
            this.grpStaticSelections.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.grpStaticSelections.Name = "grpStaticSelections";
            this.grpStaticSelections.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.grpStaticSelections.Size = new System.Drawing.Size(489, 72);
            this.grpStaticSelections.TabIndex = 10;
            this.grpStaticSelections.TabStop = false;
            this.grpStaticSelections.Text = "Static Selections";
            // 
            // txtStaticSelections
            // 
            this.txtStaticSelections.Location = new System.Drawing.Point(8, 27);
            this.txtStaticSelections.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtStaticSelections.Multiline = true;
            this.txtStaticSelections.Name = "txtStaticSelections";
            this.txtStaticSelections.Size = new System.Drawing.Size(476, 30);
            this.txtStaticSelections.TabIndex = 0;
            // 
            // grpSetPaths
            // 
            this.grpSetPaths.Controls.Add(this.btnQlikBrowse);
            this.grpSetPaths.Controls.Add(this.btnWordBrowse);
            this.grpSetPaths.Controls.Add(this.txtWordPath);
            this.grpSetPaths.Controls.Add(this.label2);
            this.grpSetPaths.Controls.Add(this.txtQlikPath);
            this.grpSetPaths.Controls.Add(this.label1);
            this.grpSetPaths.Location = new System.Drawing.Point(11, 20);
            this.grpSetPaths.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.grpSetPaths.Name = "grpSetPaths";
            this.grpSetPaths.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.grpSetPaths.Size = new System.Drawing.Size(489, 139);
            this.grpSetPaths.TabIndex = 11;
            this.grpSetPaths.TabStop = false;
            this.grpSetPaths.Text = "Set Document Paths";
            // 
            // lstLog
            // 
            this.lstLog.FormattingEnabled = true;
            this.lstLog.ItemHeight = 16;
            this.lstLog.Location = new System.Drawing.Point(517, 41);
            this.lstLog.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.lstLog.Name = "lstLog";
            this.lstLog.Size = new System.Drawing.Size(520, 436);
            this.lstLog.TabIndex = 12;
            this.lstLog.SelectedIndexChanged += new System.EventHandler(this.lstLog_SelectedIndexChanged);
            // 
            // lblLog
            // 
            this.lblLog.AutoSize = true;
            this.lblLog.Location = new System.Drawing.Point(517, 20);
            this.lblLog.Name = "lblLog";
            this.lblLog.Size = new System.Drawing.Size(32, 17);
            this.lblLog.TabIndex = 13;
            this.lblLog.Text = "Log";
            // 
            // grpQuickRef
            // 
            this.grpQuickRef.Controls.Add(this.btnRemove);
            this.grpQuickRef.Controls.Add(this.lstQuickRefVars);
            this.grpQuickRef.Controls.Add(this.label4);
            this.grpQuickRef.Controls.Add(this.txtRefID);
            this.grpQuickRef.Controls.Add(this.label3);
            this.grpQuickRef.Controls.Add(this.btnSaveRef);
            this.grpQuickRef.Controls.Add(this.txtRefName);
            this.grpQuickRef.Location = new System.Drawing.Point(11, 171);
            this.grpQuickRef.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.grpQuickRef.Name = "grpQuickRef";
            this.grpQuickRef.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.grpQuickRef.Size = new System.Drawing.Size(489, 208);
            this.grpQuickRef.TabIndex = 14;
            this.grpQuickRef.TabStop = false;
            this.grpQuickRef.Text = "Quick Reference Variables";
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(308, 167);
            this.btnRemove.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(104, 25);
            this.btnRemove.TabIndex = 12;
            this.btnRemove.Text = "Remove";
            this.btnRemove.UseVisualStyleBackColor = true;
            this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
            // 
            // lstQuickRefVars
            // 
            this.lstQuickRefVars.FormattingEnabled = true;
            this.lstQuickRefVars.ItemHeight = 16;
            this.lstQuickRefVars.Location = new System.Drawing.Point(260, 29);
            this.lstQuickRefVars.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.lstQuickRefVars.Name = "lstQuickRefVars";
            this.lstQuickRefVars.Size = new System.Drawing.Size(203, 132);
            this.lstQuickRefVars.TabIndex = 11;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(10, 95);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(117, 17);
            this.label4.TabIndex = 10;
            this.label4.Text = "Text Box Chart ID";
            // 
            // txtRefID
            // 
            this.txtRefID.Location = new System.Drawing.Point(8, 114);
            this.txtRefID.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtRefID.Name = "txtRefID";
            this.txtRefID.Size = new System.Drawing.Size(223, 22);
            this.txtRefID.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(10, 37);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(145, 17);
            this.label3.TabIndex = 8;
            this.label3.Text = "Quick Reference Text";
            // 
            // btnSaveRef
            // 
            this.btnSaveRef.Location = new System.Drawing.Point(56, 167);
            this.btnSaveRef.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnSaveRef.Name = "btnSaveRef";
            this.btnSaveRef.Size = new System.Drawing.Size(104, 25);
            this.btnSaveRef.TabIndex = 8;
            this.btnSaveRef.Text = "Add";
            this.btnSaveRef.UseVisualStyleBackColor = true;
            this.btnSaveRef.Click += new System.EventHandler(this.btnSaveRef_Click);
            // 
            // txtRefName
            // 
            this.txtRefName.Location = new System.Drawing.Point(8, 55);
            this.txtRefName.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtRefName.Name = "txtRefName";
            this.txtRefName.Size = new System.Drawing.Size(223, 22);
            this.txtRefName.TabIndex = 0;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(527, 483);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(509, 17);
            this.label5.TabIndex = 15;
            this.label5.Text = "Copyright (c) 2016 Institute 4 Priority Thinking, LLC under GNU GPL v3 License\r\n";
            // 
            // ReloadQVData
            // 
            this.ReloadQVData.AutoSize = true;
            this.ReloadQVData.Location = new System.Drawing.Point(19, 464);
            this.ReloadQVData.Name = "ReloadQVData";
            this.ReloadQVData.Size = new System.Drawing.Size(334, 21);
            this.ReloadQVData.TabIndex = 16;
            this.ReloadQVData.Text = "Reload QlikView Data Before Generating Report";
            this.ReloadQVData.UseVisualStyleBackColor = true;
            this.ReloadQVData.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // GenerateReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1047, 565);
            this.Controls.Add(this.ReloadQVData);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.grpQuickRef);
            this.Controls.Add(this.lblLog);
            this.Controls.Add(this.lstLog);
            this.Controls.Add(this.grpSetPaths);
            this.Controls.Add(this.grpStaticSelections);
            this.Controls.Add(this.btnGenerate);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "GenerateReport";
            this.Text = "Generate Report";
            this.Load += new System.EventHandler(this.GenerateReport_Load);
            this.grpStaticSelections.ResumeLayout(false);
            this.grpStaticSelections.PerformLayout();
            this.grpSetPaths.ResumeLayout(false);
            this.grpSetPaths.PerformLayout();
            this.grpQuickRef.ResumeLayout(false);
            this.grpQuickRef.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnWordBrowse;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.TextBox txtWordPath;
        private System.Windows.Forms.TextBox txtQlikPath;
        private System.Windows.Forms.Button btnQlikBrowse;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox grpStaticSelections;
        private System.Windows.Forms.GroupBox grpSetPaths;
        private System.Windows.Forms.ListBox lstLog;
        private System.Windows.Forms.Label lblLog;
        private System.Windows.Forms.TextBox txtStaticSelections;
        private System.Windows.Forms.GroupBox grpQuickRef;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtRefID;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnSaveRef;
        private System.Windows.Forms.TextBox txtRefName;
        private System.Windows.Forms.Button btnRemove;
        private System.Windows.Forms.ListBox lstQuickRefVars;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.CheckBox ReloadQVData;
    }
}

