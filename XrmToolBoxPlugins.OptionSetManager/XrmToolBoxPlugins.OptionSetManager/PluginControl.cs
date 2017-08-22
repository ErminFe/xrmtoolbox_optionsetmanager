using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XrmToolBox.Extensibility;

namespace smartpoint.XrmToolBoxPlugins.OptionSetManager
{
  public class PluginControl : PluginControlBase
  {
    private System.Windows.Forms.RadioButton rbLocalOptionSet;
    private System.Windows.Forms.RadioButton rbGlobalOptionSet;
    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.ComboBox cmbOptionSet;
    private System.Windows.Forms.Label lblEntity;
    private System.Windows.Forms.ComboBox cmbEntity;
    private System.Windows.Forms.Button btnImport;
    private System.Windows.Forms.TextBox txtImportFile;
    private System.Windows.Forms.Label label2;
    private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    private System.Windows.Forms.TextBox txtResult;
    private System.Windows.Forms.GroupBox groupBox1;

    public PluginControl() {
      InitializeComponent();
    }

    private void InitializeComponent()
    {
      this.groupBox1 = new System.Windows.Forms.GroupBox();
      this.btnImport = new System.Windows.Forms.Button();
      this.txtImportFile = new System.Windows.Forms.TextBox();
      this.label2 = new System.Windows.Forms.Label();
      this.label1 = new System.Windows.Forms.Label();
      this.cmbOptionSet = new System.Windows.Forms.ComboBox();
      this.lblEntity = new System.Windows.Forms.Label();
      this.cmbEntity = new System.Windows.Forms.ComboBox();
      this.rbGlobalOptionSet = new System.Windows.Forms.RadioButton();
      this.rbLocalOptionSet = new System.Windows.Forms.RadioButton();
      this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
      this.txtResult = new System.Windows.Forms.TextBox();
      this.groupBox1.SuspendLayout();
      this.SuspendLayout();
      // 
      // groupBox1
      // 
      this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.groupBox1.Controls.Add(this.btnImport);
      this.groupBox1.Controls.Add(this.txtImportFile);
      this.groupBox1.Controls.Add(this.label2);
      this.groupBox1.Controls.Add(this.label1);
      this.groupBox1.Controls.Add(this.cmbOptionSet);
      this.groupBox1.Controls.Add(this.lblEntity);
      this.groupBox1.Controls.Add(this.cmbEntity);
      this.groupBox1.Controls.Add(this.rbGlobalOptionSet);
      this.groupBox1.Controls.Add(this.rbLocalOptionSet);
      this.groupBox1.Location = new System.Drawing.Point(30, 29);
      this.groupBox1.Name = "groupBox1";
      this.groupBox1.Size = new System.Drawing.Size(630, 143);
      this.groupBox1.TabIndex = 0;
      this.groupBox1.TabStop = false;
      this.groupBox1.Text = "Einstellungen";
      // 
      // btnImport
      // 
      this.btnImport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
      this.btnImport.Location = new System.Drawing.Point(581, 103);
      this.btnImport.Name = "btnImport";
      this.btnImport.Size = new System.Drawing.Size(31, 23);
      this.btnImport.TabIndex = 8;
      this.btnImport.Text = "..";
      this.btnImport.UseVisualStyleBackColor = true;
      // 
      // txtImportFile
      // 
      this.txtImportFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.txtImportFile.Location = new System.Drawing.Point(82, 105);
      this.txtImportFile.Name = "txtImportFile";
      this.txtImportFile.Size = new System.Drawing.Size(493, 20);
      this.txtImportFile.TabIndex = 7;
      // 
      // label2
      // 
      this.label2.AutoSize = true;
      this.label2.Location = new System.Drawing.Point(18, 108);
      this.label2.Name = "label2";
      this.label2.Size = new System.Drawing.Size(58, 13);
      this.label2.TabIndex = 6;
      this.label2.Text = "Import File:";
      // 
      // label1
      // 
      this.label1.AutoSize = true;
      this.label1.Location = new System.Drawing.Point(177, 59);
      this.label1.Name = "label1";
      this.label1.Size = new System.Drawing.Size(57, 13);
      this.label1.TabIndex = 5;
      this.label1.Text = "OptionSet:";
      // 
      // cmbOptionSet
      // 
      this.cmbOptionSet.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.cmbOptionSet.FormattingEnabled = true;
      this.cmbOptionSet.Location = new System.Drawing.Point(243, 56);
      this.cmbOptionSet.Name = "cmbOptionSet";
      this.cmbOptionSet.Size = new System.Drawing.Size(332, 21);
      this.cmbOptionSet.TabIndex = 4;
      // 
      // lblEntity
      // 
      this.lblEntity.AutoSize = true;
      this.lblEntity.Location = new System.Drawing.Point(197, 34);
      this.lblEntity.Name = "lblEntity";
      this.lblEntity.Size = new System.Drawing.Size(36, 13);
      this.lblEntity.TabIndex = 3;
      this.lblEntity.Text = "Entity:";
      // 
      // cmbEntity
      // 
      this.cmbEntity.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.cmbEntity.FormattingEnabled = true;
      this.cmbEntity.Location = new System.Drawing.Point(243, 31);
      this.cmbEntity.Name = "cmbEntity";
      this.cmbEntity.Size = new System.Drawing.Size(332, 21);
      this.cmbEntity.TabIndex = 2;
      // 
      // rbGlobalOptionSet
      // 
      this.rbGlobalOptionSet.AutoSize = true;
      this.rbGlobalOptionSet.Location = new System.Drawing.Point(17, 55);
      this.rbGlobalOptionSet.Name = "rbGlobalOptionSet";
      this.rbGlobalOptionSet.Size = new System.Drawing.Size(105, 17);
      this.rbGlobalOptionSet.TabIndex = 1;
      this.rbGlobalOptionSet.TabStop = true;
      this.rbGlobalOptionSet.Text = "Global OptionSet";
      this.rbGlobalOptionSet.UseVisualStyleBackColor = true;
      // 
      // rbLocalOptionSet
      // 
      this.rbLocalOptionSet.AutoSize = true;
      this.rbLocalOptionSet.Checked = true;
      this.rbLocalOptionSet.Location = new System.Drawing.Point(17, 32);
      this.rbLocalOptionSet.Name = "rbLocalOptionSet";
      this.rbLocalOptionSet.Size = new System.Drawing.Size(101, 17);
      this.rbLocalOptionSet.TabIndex = 0;
      this.rbLocalOptionSet.TabStop = true;
      this.rbLocalOptionSet.Text = "Local OptionSet";
      this.rbLocalOptionSet.UseVisualStyleBackColor = true;
      // 
      // txtResult
      // 
      this.txtResult.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.txtResult.Location = new System.Drawing.Point(30, 195);
      this.txtResult.Multiline = true;
      this.txtResult.Name = "txtResult";
      this.txtResult.Size = new System.Drawing.Size(630, 181);
      this.txtResult.TabIndex = 1;
      // 
      // PluginControl
      // 
      this.Controls.Add(this.txtResult);
      this.Controls.Add(this.groupBox1);
      this.Name = "PluginControl";
      this.Size = new System.Drawing.Size(684, 398);
      this.groupBox1.ResumeLayout(false);
      this.groupBox1.PerformLayout();
      this.ResumeLayout(false);
      this.PerformLayout();

    }
  }
}
