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
            this.cbxOpenWord = new System.Windows.Forms.CheckBox();
            this.cbxQlikReload = new System.Windows.Forms.CheckBox();
            this.grpStaticSelections.SuspendLayout();
            this.grpSetPaths.SuspendLayout();
            this.grpQuickRef.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnWordBrowse
            // 
            this.btnWordBrowse.Location = new System.Drawing.Point(424, 55);
            this.btnWordBrowse.Name = "btnWordBrowse";
            this.btnWordBrowse.Size = new System.Drawing.Size(117, 31);
            this.btnWordBrowse.TabIndex = 1;
            this.btnWordBrowse.Text = "Browse";
            this.btnWordBrowse.UseVisualStyleBackColor = true;
            this.btnWordBrowse.Click += new System.EventHandler(this.btnWordBrowse_Click);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(219, 652);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(141, 43);
            this.btnGenerate.TabIndex = 2;
            this.btnGenerate.Text = "Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // txtWordPath
            // 
            this.txtWordPath.Enabled = false;
            this.txtWordPath.Location = new System.Drawing.Point(6, 55);
            this.txtWordPath.Name = "txtWordPath";
            this.txtWordPath.Size = new System.Drawing.Size(409, 26);
            this.txtWordPath.TabIndex = 3;
            // 
            // txtQlikPath
            // 
            this.txtQlikPath.Enabled = false;
            this.txtQlikPath.Location = new System.Drawing.Point(6, 122);
            this.txtQlikPath.Name = "txtQlikPath";
            this.txtQlikPath.Size = new System.Drawing.Size(409, 26);
            this.txtQlikPath.TabIndex = 4;
            // 
            // btnQlikBrowse
            // 
            this.btnQlikBrowse.Location = new System.Drawing.Point(424, 122);
            this.btnQlikBrowse.Name = "btnQlikBrowse";
            this.btnQlikBrowse.Size = new System.Drawing.Size(117, 31);
            this.btnQlikBrowse.TabIndex = 5;
            this.btnQlikBrowse.Text = "Browse";
            this.btnQlikBrowse.UseVisualStyleBackColor = true;
            this.btnQlikBrowse.Click += new System.EventHandler(this.btnQlikBrowse_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(172, 20);
            this.label1.TabIndex = 6;
            this.label1.Text = "Path to Word Template";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 99);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(202, 20);
            this.label2.TabIndex = 7;
            this.label2.Text = "Path to QlikView Document";
            // 
            // grpStaticSelections
            // 
            this.grpStaticSelections.Controls.Add(this.txtStaticSelections);
            this.grpStaticSelections.Location = new System.Drawing.Point(12, 483);
            this.grpStaticSelections.Name = "grpStaticSelections";
            this.grpStaticSelections.Size = new System.Drawing.Size(550, 90);
            this.grpStaticSelections.TabIndex = 10;
            this.grpStaticSelections.TabStop = false;
            this.grpStaticSelections.Text = "Static Selections";
            // 
            // txtStaticSelections
            // 
            this.txtStaticSelections.Location = new System.Drawing.Point(9, 34);
            this.txtStaticSelections.Multiline = true;
            this.txtStaticSelections.Name = "txtStaticSelections";
            this.txtStaticSelections.Size = new System.Drawing.Size(535, 37);
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
            this.grpSetPaths.Location = new System.Drawing.Point(12, 25);
            this.grpSetPaths.Name = "grpSetPaths";
            this.grpSetPaths.Size = new System.Drawing.Size(550, 174);
            this.grpSetPaths.TabIndex = 11;
            this.grpSetPaths.TabStop = false;
            this.grpSetPaths.Text = "Set Document Paths";
            // 
            // lstLog
            // 
            this.lstLog.FormattingEnabled = true;
            this.lstLog.ItemHeight = 20;
            this.lstLog.Location = new System.Drawing.Point(582, 51);
            this.lstLog.Name = "lstLog";
            this.lstLog.Size = new System.Drawing.Size(584, 584);
            this.lstLog.TabIndex = 12;
            // 
            // lblLog
            // 
            this.lblLog.AutoSize = true;
            this.lblLog.Location = new System.Drawing.Point(582, 25);
            this.lblLog.Name = "lblLog";
            this.lblLog.Size = new System.Drawing.Size(36, 20);
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
            this.grpQuickRef.Location = new System.Drawing.Point(12, 214);
            this.grpQuickRef.Name = "grpQuickRef";
            this.grpQuickRef.Size = new System.Drawing.Size(550, 260);
            this.grpQuickRef.TabIndex = 14;
            this.grpQuickRef.TabStop = false;
            this.grpQuickRef.Text = "Quick Reference Variables";
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(346, 209);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(117, 31);
            this.btnRemove.TabIndex = 12;
            this.btnRemove.Text = "Remove";
            this.btnRemove.UseVisualStyleBackColor = true;
            this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
            // 
            // lstQuickRefVars
            // 
            this.lstQuickRefVars.FormattingEnabled = true;
            this.lstQuickRefVars.ItemHeight = 20;
            this.lstQuickRefVars.Location = new System.Drawing.Point(292, 36);
            this.lstQuickRefVars.Name = "lstQuickRefVars";
            this.lstQuickRefVars.Size = new System.Drawing.Size(228, 164);
            this.lstQuickRefVars.TabIndex = 11;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(11, 119);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(134, 20);
            this.label4.TabIndex = 10;
            this.label4.Text = "Text Box Chart ID";
            // 
            // txtRefID
            // 
            this.txtRefID.Location = new System.Drawing.Point(9, 142);
            this.txtRefID.Name = "txtRefID";
            this.txtRefID.Size = new System.Drawing.Size(250, 26);
            this.txtRefID.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(11, 46);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(162, 20);
            this.label3.TabIndex = 8;
            this.label3.Text = "Quick Reference Text";
            // 
            // btnSaveRef
            // 
            this.btnSaveRef.Location = new System.Drawing.Point(63, 209);
            this.btnSaveRef.Name = "btnSaveRef";
            this.btnSaveRef.Size = new System.Drawing.Size(117, 31);
            this.btnSaveRef.TabIndex = 8;
            this.btnSaveRef.Text = "Add";
            this.btnSaveRef.UseVisualStyleBackColor = true;
            this.btnSaveRef.Click += new System.EventHandler(this.btnSaveRef_Click);
            // 
            // txtRefName
            // 
            this.txtRefName.Location = new System.Drawing.Point(9, 69);
            this.txtRefName.Name = "txtRefName";
            this.txtRefName.Size = new System.Drawing.Size(250, 26);
            this.txtRefName.TabIndex = 0;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(593, 643);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(564, 20);
            this.label5.TabIndex = 15;
            this.label5.Text = "Copyright (c) 2016 Institute 4 Priority Thinking, LLC under GNU GPL v3 License\r\n";
            // 
            // cbxOpenWord
            // 
            this.cbxOpenWord.AutoSize = true;
            this.cbxOpenWord.Location = new System.Drawing.Point(150, 619);
            this.cbxOpenWord.Name = "cbxOpenWord";
            this.cbxOpenWord.Size = new System.Drawing.Size(292, 24);
            this.cbxOpenWord.TabIndex = 16;
            this.cbxOpenWord.Text = "Open Word document when finished";
            this.cbxOpenWord.UseVisualStyleBackColor = true;
            // 
            // cbxQlikReload
            // 
            this.cbxQlikReload.AutoSize = true;
            this.cbxQlikReload.Location = new System.Drawing.Point(150, 589);
            this.cbxQlikReload.Name = "cbxQlikReload";
            this.cbxQlikReload.Size = new System.Drawing.Size(291, 24);
            this.cbxQlikReload.TabIndex = 17;
            this.cbxQlikReload.Text = "Reload data in Qlik before beginning";
            this.cbxQlikReload.UseVisualStyleBackColor = true;
            // 
            // GenerateReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1178, 707);
            this.Controls.Add(this.cbxQlikReload);
            this.Controls.Add(this.cbxOpenWord);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.grpQuickRef);
            this.Controls.Add(this.lblLog);
            this.Controls.Add(this.lstLog);
            this.Controls.Add(this.grpSetPaths);
            this.Controls.Add(this.grpStaticSelections);
            this.Controls.Add(this.btnGenerate);
            this.Name = "GenerateReport";
            this.Text = "Report Generator";
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
        private System.Windows.Forms.CheckBox cbxOpenWord;
        private System.Windows.Forms.CheckBox cbxQlikReload;
    }
}

