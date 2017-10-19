using GenericParsing;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using smartpoint.XrmToolBoxPlugins.OptionSetManager.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
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
    private System.Windows.Forms.Button btnChooseFile;
    private System.Windows.Forms.TextBox txtImportFile;
    private System.Windows.Forms.Label label2;
    private System.Windows.Forms.TextBox txtResult;
    private OpenFileDialog dlgFileChooser;
    private GroupBox groupBox2;
    private DataGridView grdValues;
    private Button btnLoad;
    private Button btnImport;
    private System.Windows.Forms.GroupBox groupBox1;

    private int languageCodeOfUser = 1033;
    string selectedEntityLogicalName = null;
    string selectedAttributeLogicalName = null;
    private Button btnSample;
    private SaveFileDialog dlgFileSave;
    Model.ImportDataDataTable importData = null;

    public PluginControl()
    {
      InitializeComponent();
    }

    private void InitComboboxes()
    {
      WorkAsync(new WorkAsyncInfo
      {
        Message = "Retrieving the entities...",
        Work = (w, e) =>
        {
          var result = Helper.GetEntyNames(Service).OrderBy(_ => _.Value).ToList();

          cmbEntity.Invoke((MethodInvoker)delegate
          {
            cmbEntity.ValueMember = "Key";
            cmbEntity.DisplayMember = "Value";
            cmbEntity.DataSource = result;
          });

        },
        ProgressChanged = e =>
        {
        },
        PostWorkCallBack = e =>
        {
        },
        AsyncArgument = null,
        IsCancelable = false,
        MessageWidth = 340,
        MessageHeight = 150
      });

    }//InitComboboxes

    private void InitializeComponent()
    {
      this.groupBox1 = new System.Windows.Forms.GroupBox();
      this.btnSample = new System.Windows.Forms.Button();
      this.btnLoad = new System.Windows.Forms.Button();
      this.btnChooseFile = new System.Windows.Forms.Button();
      this.txtImportFile = new System.Windows.Forms.TextBox();
      this.label2 = new System.Windows.Forms.Label();
      this.label1 = new System.Windows.Forms.Label();
      this.cmbOptionSet = new System.Windows.Forms.ComboBox();
      this.lblEntity = new System.Windows.Forms.Label();
      this.cmbEntity = new System.Windows.Forms.ComboBox();
      this.rbGlobalOptionSet = new System.Windows.Forms.RadioButton();
      this.rbLocalOptionSet = new System.Windows.Forms.RadioButton();
      this.txtResult = new System.Windows.Forms.TextBox();
      this.dlgFileChooser = new System.Windows.Forms.OpenFileDialog();
      this.groupBox2 = new System.Windows.Forms.GroupBox();
      this.grdValues = new System.Windows.Forms.DataGridView();
      this.btnImport = new System.Windows.Forms.Button();
      this.dlgFileSave = new System.Windows.Forms.SaveFileDialog();
      this.groupBox1.SuspendLayout();
      this.groupBox2.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize)(this.grdValues)).BeginInit();
      this.SuspendLayout();
      // 
      // groupBox1
      // 
      this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.groupBox1.Controls.Add(this.btnSample);
      this.groupBox1.Controls.Add(this.btnLoad);
      this.groupBox1.Controls.Add(this.btnChooseFile);
      this.groupBox1.Controls.Add(this.txtImportFile);
      this.groupBox1.Controls.Add(this.label2);
      this.groupBox1.Controls.Add(this.label1);
      this.groupBox1.Controls.Add(this.cmbOptionSet);
      this.groupBox1.Controls.Add(this.lblEntity);
      this.groupBox1.Controls.Add(this.cmbEntity);
      this.groupBox1.Controls.Add(this.rbGlobalOptionSet);
      this.groupBox1.Controls.Add(this.rbLocalOptionSet);
      this.groupBox1.Location = new System.Drawing.Point(19, 29);
      this.groupBox1.Name = "groupBox1";
      this.groupBox1.Size = new System.Drawing.Size(641, 166);
      this.groupBox1.TabIndex = 0;
      this.groupBox1.TabStop = false;
      this.groupBox1.Text = "Einstellungen";
      // 
      // btnSample
      // 
      this.btnSample.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
      this.btnSample.Location = new System.Drawing.Point(569, 102);
      this.btnSample.Name = "btnSample";
      this.btnSample.Size = new System.Drawing.Size(54, 23);
      this.btnSample.TabIndex = 10;
      this.btnSample.Text = "Sample";
      this.btnSample.UseVisualStyleBackColor = true;
      this.btnSample.Click += new System.EventHandler(this.btnSample_Click);
      // 
      // btnLoad
      // 
      this.btnLoad.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
      this.btnLoad.Location = new System.Drawing.Point(254, 137);
      this.btnLoad.Name = "btnLoad";
      this.btnLoad.Size = new System.Drawing.Size(75, 23);
      this.btnLoad.TabIndex = 9;
      this.btnLoad.Text = "&Load";
      this.btnLoad.UseVisualStyleBackColor = true;
      this.btnLoad.Click += new System.EventHandler(this.BtnLoad_Click);
      // 
      // btnChooseFile
      // 
      this.btnChooseFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
      this.btnChooseFile.Location = new System.Drawing.Point(532, 102);
      this.btnChooseFile.Name = "btnChooseFile";
      this.btnChooseFile.Size = new System.Drawing.Size(31, 23);
      this.btnChooseFile.TabIndex = 8;
      this.btnChooseFile.Text = "..";
      this.btnChooseFile.UseVisualStyleBackColor = true;
      this.btnChooseFile.Click += new System.EventHandler(this.btnChooseFile_Click);
      // 
      // txtImportFile
      // 
      this.txtImportFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.txtImportFile.Location = new System.Drawing.Point(82, 104);
      this.txtImportFile.Name = "txtImportFile";
      this.txtImportFile.Size = new System.Drawing.Size(444, 20);
      this.txtImportFile.TabIndex = 7;
      this.txtImportFile.Text = "C:\\Projekte\\FuJo\\Github\\XrmToolBox\\OptionSetManager\\XrmToolBoxPlugins.OptionSetMa" +
    "nager\\XrmToolBoxPlugins.OptionSetManager\\bin\\Debug\\Import.csv";
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
      this.cmbOptionSet.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
      this.cmbOptionSet.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
      this.cmbOptionSet.FormattingEnabled = true;
      this.cmbOptionSet.Location = new System.Drawing.Point(243, 56);
      this.cmbOptionSet.Name = "cmbOptionSet";
      this.cmbOptionSet.Size = new System.Drawing.Size(380, 21);
      this.cmbOptionSet.TabIndex = 4;
      // 
      // lblEntity
      // 
      this.lblEntity.AutoSize = true;
      this.lblEntity.Location = new System.Drawing.Point(198, 34);
      this.lblEntity.Name = "lblEntity";
      this.lblEntity.Size = new System.Drawing.Size(36, 13);
      this.lblEntity.TabIndex = 3;
      this.lblEntity.Text = "Entity:";
      // 
      // cmbEntity
      // 
      this.cmbEntity.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.cmbEntity.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
      this.cmbEntity.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
      this.cmbEntity.FormattingEnabled = true;
      this.cmbEntity.Location = new System.Drawing.Point(243, 31);
      this.cmbEntity.Name = "cmbEntity";
      this.cmbEntity.Size = new System.Drawing.Size(380, 21);
      this.cmbEntity.TabIndex = 2;
      this.cmbEntity.SelectedIndexChanged += new System.EventHandler(this.CmbEntity_SelectedIndexChanged);
      // 
      // rbGlobalOptionSet
      // 
      this.rbGlobalOptionSet.AutoSize = true;
      this.rbGlobalOptionSet.Enabled = false;
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
      this.txtResult.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.txtResult.Enabled = false;
      this.txtResult.Location = new System.Drawing.Point(19, 503);
      this.txtResult.Multiline = true;
      this.txtResult.Name = "txtResult";
      this.txtResult.Size = new System.Drawing.Size(641, 105);
      this.txtResult.TabIndex = 1;
      // 
      // dlgFileChooser
      // 
      this.dlgFileChooser.FileName = "openFileDialog1";
      // 
      // groupBox2
      // 
      this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.groupBox2.Controls.Add(this.grdValues);
      this.groupBox2.Location = new System.Drawing.Point(19, 201);
      this.groupBox2.Name = "groupBox2";
      this.groupBox2.Size = new System.Drawing.Size(641, 267);
      this.groupBox2.TabIndex = 2;
      this.groupBox2.TabStop = false;
      this.groupBox2.Text = "Values";
      // 
      // grdValues
      // 
      this.grdValues.AllowUserToAddRows = false;
      this.grdValues.AllowUserToDeleteRows = false;
      this.grdValues.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.grdValues.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
      this.grdValues.Location = new System.Drawing.Point(17, 32);
      this.grdValues.Name = "grdValues";
      this.grdValues.ReadOnly = true;
      this.grdValues.Size = new System.Drawing.Size(606, 217);
      this.grdValues.TabIndex = 0;
      // 
      // btnImport
      // 
      this.btnImport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
      this.btnImport.Location = new System.Drawing.Point(273, 474);
      this.btnImport.Name = "btnImport";
      this.btnImport.Size = new System.Drawing.Size(75, 23);
      this.btnImport.TabIndex = 10;
      this.btnImport.Text = "&Import";
      this.btnImport.UseVisualStyleBackColor = true;
      this.btnImport.Click += new System.EventHandler(this.BtnImport_OnClick);
      // 
      // PluginControl
      // 
      this.Controls.Add(this.btnImport);
      this.Controls.Add(this.groupBox2);
      this.Controls.Add(this.txtResult);
      this.Controls.Add(this.groupBox1);
      this.Name = "PluginControl";
      this.Size = new System.Drawing.Size(684, 630);
      this.Load += new System.EventHandler(this.Control_OnLoad);
      this.groupBox1.ResumeLayout(false);
      this.groupBox1.PerformLayout();
      this.groupBox2.ResumeLayout(false);
      ((System.ComponentModel.ISupportInitialize)(this.grdValues)).EndInit();
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    private void btnChooseFile_Click(object sender, EventArgs e)
    {
      if (DialogResult.OK == dlgFileChooser.ShowDialog(this))
      {
        txtImportFile.Text = dlgFileChooser.FileName;
      }

    }

    private void BtnLoad_Click(object sender, EventArgs e)
    {
      importData = new Model.ImportDataDataTable();
      using (GenericParser parser = new GenericParser())
      {
        parser.SetDataSource(txtImportFile.Text);

        parser.ColumnDelimiter = ';';
        parser.FirstRowHasHeader = true;
        parser.SkipStartingDataRows = 0;
        parser.MaxBufferSize = 4096;
        parser.MaxRows = 500;
        parser.TextQualifier = '\"';

        while (parser.Read())
        {
          Model.ImportDataRow row = importData.NewImportDataRow();

          row.Value = int.Parse(parser["Value"]);
          if (!string.IsNullOrEmpty(parser["Label1031"]))
          {
            row.Label1031 = parser["Label1031"];
            if (!string.IsNullOrEmpty(row.Label1031))
              row.Label1031 = row.Label1031.Replace("\"", string.Empty).Trim();
          }
          if (!string.IsNullOrEmpty(parser["Label1033"]))
          {
            row.Label1033 = parser["Label1033"];
            if (!string.IsNullOrEmpty(row.Label1033))
              row.Label1033 = row.Label1033.Replace("\"", string.Empty).Trim();
          }
          if (!string.IsNullOrEmpty(parser["Label1036"]))
          {
            row.Label1036 = parser["Label1036"];
            if (!string.IsNullOrEmpty(row.Label1036))
              row.Label1036 = row.Label1036.Replace("\"", string.Empty).Trim();
          }
          if (!string.IsNullOrEmpty(parser["Label1040"]))
          {
            row.Label1040 = parser["Label1040"];
            if (!string.IsNullOrEmpty(row.Label1040))
              row.Label1040 = row.Label1040.Replace("\"", string.Empty).Trim();
          }
          if (!string.IsNullOrEmpty(parser["Label3082"]))
          {
            row.Label3082 = parser["Label3082"];
            if (!string.IsNullOrEmpty(row.Label3082))
              row.Label3082 = row.Label3082.Replace("\"", string.Empty).Trim();
          }


          importData.AddImportDataRow(row);
        }//while
      }//using

      grdValues.DataSource = importData;
    }//BtnLoad_Click

    private void CmbEntity_SelectedIndexChanged(object sender, EventArgs args)
    {
      WorkAsync(new WorkAsyncInfo
      {
        Message = "Retrieving the attributes...",
        Work = (w, e) =>
        {

          cmbEntity.Invoke((MethodInvoker)delegate
          {
            selectedEntityLogicalName = (string)cmbEntity.SelectedValue;
          });

          var result = Helper.GetOptionSetsForEntity(Service, selectedEntityLogicalName).OrderBy(_ => _.Value).ToList();

          cmbOptionSet.Invoke((MethodInvoker)delegate
          {
            cmbOptionSet.ValueMember = "Key";
            cmbOptionSet.DisplayMember = "Value";
            cmbOptionSet.DataSource = result;
          });

        },
        ProgressChanged = e =>
        {
        },
        PostWorkCallBack = e =>
        {
        },
        AsyncArgument = null,
        IsCancelable = false,
        MessageWidth = 340,
        MessageHeight = 150
      });
    }

    private void Control_OnLoad(object sender, EventArgs e)
    {
      LoadUserSettings();
      InitComboboxes();
    }

    private void LoadUserSettings()
    {
      WorkAsync(new WorkAsyncInfo
      {
        Message = "Retrieving your user settings...",
        Work = (w, e) =>
        {
          var request = new WhoAmIRequest();
          var response = (WhoAmIResponse)Service.Execute(request);

          Entity userSettings = Service.RetrieveMultiple(new QueryExpression("usersettings")
          {
            ColumnSet = new ColumnSet("uilanguageid"),
            Criteria = new FilterExpression
            {
              Conditions =
              {
                  new ConditionExpression("systemuserid", ConditionOperator.Equal, response.UserId)
              }
            }
          }).Entities.FirstOrDefault();
          //if (userSettings != null && userSettings.Contains("uilanguageid"))
          //{
          //  languageCodeOfUser = userSettings.GetAttributeValue<int>("uilanguageid");
          //}

          languageCodeOfUser = 1033;
        },
        ProgressChanged = e =>
        {
        },
        PostWorkCallBack = e =>
        {
        },
        AsyncArgument = null,
        IsCancelable = true,
        MessageWidth = 340,
        MessageHeight = 150
      });
    }

    private void BtnImport_OnClick(object sender, EventArgs args)
    {
      WorkAsync(new WorkAsyncInfo
      {
        Message = "Importing the option sets...",
        Work = (w, e) =>
        {
          cmbOptionSet.Invoke((MethodInvoker)delegate
          {
            selectedAttributeLogicalName = (string)cmbOptionSet.SelectedValue;
          });

          txtResult.Invoke((MethodInvoker)delegate
          {
            txtResult.Clear();
          });

          OptionMetadataCollection existingOptions = Helper.GetOptionSetEntries(Service, selectedEntityLogicalName, selectedAttributeLogicalName);

          ExecuteMultipleRequest multipleRequest = new ExecuteMultipleRequest();
          multipleRequest.Settings = new ExecuteMultipleSettings();
          multipleRequest.Settings.ContinueOnError = false;
          multipleRequest.Settings.ReturnResponses = true;
          multipleRequest.Requests = new OrganizationRequestCollection();

          foreach (Model.ImportDataRow row in importData)
          {
            if (existingOptions.Any(_ => _.Value == row.Value))
            {
              //merge
              List<KeyValuePair<int, string>> labels = new List<KeyValuePair<int, string>>();
              OptionMetadata metadata = existingOptions.FirstOrDefault(_ => _.Value == row.Value);
              if (metadata.Label != null && metadata.Label.LocalizedLabels != null)
              {
                // initit labels collection with original labels
                foreach (LocalizedLabel item in metadata.Label.LocalizedLabels)
                  labels.Add(new KeyValuePair<int, string>(item.LanguageCode, item.Label));

                // override the existing labels
                if (!row.IsLabel1031Null() && labels.Any(_ => _.Key == 1031))
                {
                  labels.Remove(labels.First(_ => _.Key == 1031));
                  labels.Add(new KeyValuePair<int, string>(1031, row.Label1031));
                }//if

                if (!row.IsLabel1033Null() && labels.Any(_ => _.Key == 1033))
                {
                  labels.Remove(labels.First(_ => _.Key == 1033));
                  labels.Add(new KeyValuePair<int, string>(1033, row.Label1033));
                }//if

                if (!row.IsLabel1036Null() && labels.Any(_ => _.Key == 1036))
                {
                  labels.Remove(labels.First(_ => _.Key == 1036));
                  labels.Add(new KeyValuePair<int, string>(1036, row.Label1036));
                }//if

                if (!row.IsLabel1040Null() && labels.Any(_ => _.Key == 1040))
                {
                  labels.Remove(labels.First(_ => _.Key == 1040));
                  labels.Add(new KeyValuePair<int, string>(1040, row.Label1040));
                }//if

                if (!row.IsLabel3082Null() && labels.Any(_ => _.Key == 3082))
                {
                  labels.Remove(labels.First(_ => _.Key == 3082));
                  labels.Add(new KeyValuePair<int, string>(3082, row.Label3082));
                }//if

                // create insert request
                multipleRequest.Requests.Add(Helper.CreateUpdateOptionValueRequest(Service,
                  languageCodeOfUser,
                  selectedEntityLogicalName,
                  selectedAttributeLogicalName,
                  labels,
                  row.Value));
              }//if
            }
            else
            {
              List<KeyValuePair<int, string>> labels = new List<KeyValuePair<int, string>>();
              if (!row.IsLabel1031Null())
                labels.Add(new KeyValuePair<int, string>(1031, row.Label1031));

              if (!row.IsLabel1033Null())
                labels.Add(new KeyValuePair<int, string>(1033, row.Label1033));

              if (!row.IsLabel1036Null())
                labels.Add(new KeyValuePair<int, string>(1036, row.Label1036));

              if (!row.IsLabel1040Null())
                labels.Add(new KeyValuePair<int, string>(1040, row.Label1040));

              if (!row.IsLabel3082Null())
                labels.Add(new KeyValuePair<int, string>(3082, row.Label3082));

              // create insert request
              multipleRequest.Requests.Add(Helper.CreateInsertOptionValueRequest(Service,
                languageCodeOfUser,
                selectedEntityLogicalName,
                selectedAttributeLogicalName,
                labels,
                row.Value));
            }
          }//foreach

          //multipleRequest.Requests.Add(new PublishAllXmlRequest());

          ExecuteMultipleResponse response = (ExecuteMultipleResponse)Service.Execute(multipleRequest);

          txtResult.Invoke((MethodInvoker)delegate
          {
            if (!response.IsFaulted)
              txtResult.Text = "Successfully performed the import - please publish the entity.";
            else
            {
              foreach (ExecuteMultipleResponseItem responseItem in response.Responses)
              {
                txtResult.Text = responseItem.Fault.Message + Environment.NewLine + txtResult.Text;
              }//foreach
            }
          });


        },
        ProgressChanged = e =>
        {
        },
        PostWorkCallBack = e =>
        {
        },
        AsyncArgument = null,
        IsCancelable = true,
        MessageWidth = 340,
        MessageHeight = 150
      });
    }

    private void btnSample_Click(object sender, EventArgs e)
    {
      dlgFileSave.FileName = "sample.csv";
      dlgFileSave.Filter = "Semicolon Separated File | *.csv";
      if (DialogResult.OK == dlgFileSave.ShowDialog())
      {
        using (StreamWriter writer = new StreamWriter(dlgFileSave.OpenFile()))
        {
          writer.WriteLine(@"""Value"";""Label1031"";""Label1033"";""Label1036"";""Label1040"";""Label3082""");
          writer.WriteLine(@"""123100"";""a"";""b"";""c"";""d"";""e"";");
          writer.WriteLine(@"""123101"";""a"";""b"";""c"";""d"";""e"";");
          writer.WriteLine(@"""123102"";""a"";""b"";""c"";""d"";""e"";");
          writer.WriteLine(@"""123103"";""a"";""b"";""c"";""d"";""e"";");
          writer.WriteLine(@"""123104"";""a"";""b"";""c"";""d"";""e"";");
        }//using
      }

      MessageBox.Show("Sample file written.");
    }
  }
}
