namespace SalvagniniForecastHelper
{
    partial class Form1
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.labelMasterDataFile = new System.Windows.Forms.Label();
            this.textBoxMasterDataFile = new System.Windows.Forms.TextBox();
            this.buttonOpenMasterDataFile = new System.Windows.Forms.Button();
            this.labelForecastFile = new System.Windows.Forms.Label();
            this.textBoxForecastFile = new System.Windows.Forms.TextBox();
            this.buttonForeCastFile = new System.Windows.Forms.Button();
            this.buttonGo = new System.Windows.Forms.Button();
            this.textBoxMasterdata = new System.Windows.Forms.TextBox();
            this.textBoxForecast = new System.Windows.Forms.TextBox();
            this.textBoxResult = new System.Windows.Forms.TextBox();
            this.groupBoxMAsterDataFormat = new System.Windows.Forms.GroupBox();
            this.labelSheetNameMasterData = new System.Windows.Forms.Label();
            this.textBoxSheetNameMasterData = new System.Windows.Forms.TextBox();
            this.textBoxRowToMasterData = new System.Windows.Forms.TextBox();
            this.labelRowToMasterData = new System.Windows.Forms.Label();
            this.textBoxRowFromMasterData = new System.Windows.Forms.TextBox();
            this.labelRowFromMasterData = new System.Windows.Forms.Label();
            this.textBoxThicknesColumn = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxWidthColumn = new System.Windows.Forms.TextBox();
            this.labelWidthColumn = new System.Windows.Forms.Label();
            this.textBoxLengthColumn = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxArticleNumberColumnMasterData = new System.Windows.Forms.TextBox();
            this.labelArticleNumberColumnMasterData = new System.Windows.Forms.Label();
            this.groupBoxForecastFormat = new System.Windows.Forms.GroupBox();
            this.labelRowToForecast = new System.Windows.Forms.Label();
            this.textBoxRowToForecast = new System.Windows.Forms.TextBox();
            this.labelRowFromForecast = new System.Windows.Forms.Label();
            this.textBoxRowFromForecast = new System.Windows.Forms.TextBox();
            this.labelQuantityColumn = new System.Windows.Forms.Label();
            this.textBoxQuantityColumn = new System.Windows.Forms.TextBox();
            this.labelArticleNumberColumnForecast = new System.Windows.Forms.Label();
            this.textBoxArticleNumberColumnForecast = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxSheetNr = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.buttonCopyToClipboardForExcelPaste = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.groupBoxMAsterDataFormat.SuspendLayout();
            this.groupBoxForecastFormat.SuspendLayout();
            this.SuspendLayout();
            // 
            // labelMasterDataFile
            // 
            this.labelMasterDataFile.AutoSize = true;
            this.labelMasterDataFile.Location = new System.Drawing.Point(13, 13);
            this.labelMasterDataFile.Name = "labelMasterDataFile";
            this.labelMasterDataFile.Size = new System.Drawing.Size(84, 13);
            this.labelMasterDataFile.TabIndex = 0;
            this.labelMasterDataFile.Text = "Master Data File";
            // 
            // textBoxMasterDataFile
            // 
            this.textBoxMasterDataFile.AllowDrop = true;
            this.textBoxMasterDataFile.Location = new System.Drawing.Point(103, 10);
            this.textBoxMasterDataFile.Name = "textBoxMasterDataFile";
            this.textBoxMasterDataFile.Size = new System.Drawing.Size(503, 20);
            this.textBoxMasterDataFile.TabIndex = 1;
            this.textBoxMasterDataFile.Text = "G:\\Kunden\\Kalkulation\\salvagnini\\Kalkulation Salvagnini.xlsx";
            this.textBoxMasterDataFile.DragDrop += new System.Windows.Forms.DragEventHandler(this.textBoxMasterDataFile_DragDrop);
            this.textBoxMasterDataFile.DragEnter += new System.Windows.Forms.DragEventHandler(this.textBoxMasterDataFile_DragEnter);
            // 
            // buttonOpenMasterDataFile
            // 
            this.buttonOpenMasterDataFile.Location = new System.Drawing.Point(612, 8);
            this.buttonOpenMasterDataFile.Name = "buttonOpenMasterDataFile";
            this.buttonOpenMasterDataFile.Size = new System.Drawing.Size(75, 23);
            this.buttonOpenMasterDataFile.TabIndex = 2;
            this.buttonOpenMasterDataFile.Text = "Open";
            this.buttonOpenMasterDataFile.UseVisualStyleBackColor = true;
            this.buttonOpenMasterDataFile.Click += new System.EventHandler(this.buttonOpenMasterDataFile_Click);
            // 
            // labelForecastFile
            // 
            this.labelForecastFile.AutoSize = true;
            this.labelForecastFile.Location = new System.Drawing.Point(30, 40);
            this.labelForecastFile.Name = "labelForecastFile";
            this.labelForecastFile.Size = new System.Drawing.Size(67, 13);
            this.labelForecastFile.TabIndex = 3;
            this.labelForecastFile.Text = "Forecast File";
            // 
            // textBoxForecastFile
            // 
            this.textBoxForecastFile.AllowDrop = true;
            this.textBoxForecastFile.Location = new System.Drawing.Point(103, 37);
            this.textBoxForecastFile.Name = "textBoxForecastFile";
            this.textBoxForecastFile.Size = new System.Drawing.Size(503, 20);
            this.textBoxForecastFile.TabIndex = 4;
            this.textBoxForecastFile.DragDrop += new System.Windows.Forms.DragEventHandler(this.textBoxForecastFile_DragDrop);
            this.textBoxForecastFile.DragEnter += new System.Windows.Forms.DragEventHandler(this.textBoxForecastFile_DragEnter);
            // 
            // buttonForeCastFile
            // 
            this.buttonForeCastFile.Location = new System.Drawing.Point(612, 35);
            this.buttonForeCastFile.Name = "buttonForeCastFile";
            this.buttonForeCastFile.Size = new System.Drawing.Size(75, 23);
            this.buttonForeCastFile.TabIndex = 5;
            this.buttonForeCastFile.Text = "Open";
            this.buttonForeCastFile.UseVisualStyleBackColor = true;
            this.buttonForeCastFile.Click += new System.EventHandler(this.buttonForeCastFile_Click);
            // 
            // buttonGo
            // 
            this.buttonGo.Location = new System.Drawing.Point(12, 78);
            this.buttonGo.Name = "buttonGo";
            this.buttonGo.Size = new System.Drawing.Size(675, 105);
            this.buttonGo.TabIndex = 6;
            this.buttonGo.Text = "GO";
            this.buttonGo.UseVisualStyleBackColor = true;
            this.buttonGo.Click += new System.EventHandler(this.buttonGo_Click);
            // 
            // textBoxMasterdata
            // 
            this.textBoxMasterdata.Location = new System.Drawing.Point(12, 229);
            this.textBoxMasterdata.Multiline = true;
            this.textBoxMasterdata.Name = "textBoxMasterdata";
            this.textBoxMasterdata.Size = new System.Drawing.Size(215, 102);
            this.textBoxMasterdata.TabIndex = 7;
            // 
            // textBoxForecast
            // 
            this.textBoxForecast.Location = new System.Drawing.Point(233, 229);
            this.textBoxForecast.Multiline = true;
            this.textBoxForecast.Name = "textBoxForecast";
            this.textBoxForecast.Size = new System.Drawing.Size(239, 102);
            this.textBoxForecast.TabIndex = 7;
            // 
            // textBoxResult
            // 
            this.textBoxResult.Location = new System.Drawing.Point(478, 229);
            this.textBoxResult.Multiline = true;
            this.textBoxResult.Name = "textBoxResult";
            this.textBoxResult.Size = new System.Drawing.Size(209, 102);
            this.textBoxResult.TabIndex = 7;
            // 
            // groupBoxMAsterDataFormat
            // 
            this.groupBoxMAsterDataFormat.Controls.Add(this.labelSheetNameMasterData);
            this.groupBoxMAsterDataFormat.Controls.Add(this.textBoxSheetNameMasterData);
            this.groupBoxMAsterDataFormat.Controls.Add(this.textBoxRowToMasterData);
            this.groupBoxMAsterDataFormat.Controls.Add(this.labelRowToMasterData);
            this.groupBoxMAsterDataFormat.Controls.Add(this.textBoxRowFromMasterData);
            this.groupBoxMAsterDataFormat.Controls.Add(this.labelRowFromMasterData);
            this.groupBoxMAsterDataFormat.Controls.Add(this.textBoxThicknesColumn);
            this.groupBoxMAsterDataFormat.Controls.Add(this.label4);
            this.groupBoxMAsterDataFormat.Controls.Add(this.textBoxWidthColumn);
            this.groupBoxMAsterDataFormat.Controls.Add(this.labelWidthColumn);
            this.groupBoxMAsterDataFormat.Controls.Add(this.textBoxLengthColumn);
            this.groupBoxMAsterDataFormat.Controls.Add(this.label2);
            this.groupBoxMAsterDataFormat.Controls.Add(this.textBoxArticleNumberColumnMasterData);
            this.groupBoxMAsterDataFormat.Controls.Add(this.labelArticleNumberColumnMasterData);
            this.groupBoxMAsterDataFormat.Location = new System.Drawing.Point(710, 8);
            this.groupBoxMAsterDataFormat.Name = "groupBoxMAsterDataFormat";
            this.groupBoxMAsterDataFormat.Size = new System.Drawing.Size(263, 209);
            this.groupBoxMAsterDataFormat.TabIndex = 8;
            this.groupBoxMAsterDataFormat.TabStop = false;
            this.groupBoxMAsterDataFormat.Text = "Master Data Format";
            // 
            // labelSheetNameMasterData
            // 
            this.labelSheetNameMasterData.AutoSize = true;
            this.labelSheetNameMasterData.Location = new System.Drawing.Point(67, 22);
            this.labelSheetNameMasterData.Name = "labelSheetNameMasterData";
            this.labelSheetNameMasterData.Size = new System.Drawing.Size(66, 13);
            this.labelSheetNameMasterData.TabIndex = 13;
            this.labelSheetNameMasterData.Text = "Sheet Name";
            // 
            // textBoxSheetNameMasterData
            // 
            this.textBoxSheetNameMasterData.Location = new System.Drawing.Point(139, 19);
            this.textBoxSheetNameMasterData.Name = "textBoxSheetNameMasterData";
            this.textBoxSheetNameMasterData.Size = new System.Drawing.Size(100, 20);
            this.textBoxSheetNameMasterData.TabIndex = 12;
            this.textBoxSheetNameMasterData.Text = "Tabelle1";
            // 
            // textBoxRowToMasterData
            // 
            this.textBoxRowToMasterData.Location = new System.Drawing.Point(139, 175);
            this.textBoxRowToMasterData.Name = "textBoxRowToMasterData";
            this.textBoxRowToMasterData.Size = new System.Drawing.Size(100, 20);
            this.textBoxRowToMasterData.TabIndex = 11;
            this.textBoxRowToMasterData.Text = "1000";
            // 
            // labelRowToMasterData
            // 
            this.labelRowToMasterData.AutoSize = true;
            this.labelRowToMasterData.Location = new System.Drawing.Point(91, 178);
            this.labelRowToMasterData.Name = "labelRowToMasterData";
            this.labelRowToMasterData.Size = new System.Drawing.Size(42, 13);
            this.labelRowToMasterData.TabIndex = 10;
            this.labelRowToMasterData.Text = "RowTo";
            // 
            // textBoxRowFromMasterData
            // 
            this.textBoxRowFromMasterData.Location = new System.Drawing.Point(139, 149);
            this.textBoxRowFromMasterData.Name = "textBoxRowFromMasterData";
            this.textBoxRowFromMasterData.Size = new System.Drawing.Size(100, 20);
            this.textBoxRowFromMasterData.TabIndex = 9;
            this.textBoxRowFromMasterData.Text = "5";
            // 
            // labelRowFromMasterData
            // 
            this.labelRowFromMasterData.AutoSize = true;
            this.labelRowFromMasterData.Location = new System.Drawing.Point(78, 152);
            this.labelRowFromMasterData.Name = "labelRowFromMasterData";
            this.labelRowFromMasterData.Size = new System.Drawing.Size(55, 13);
            this.labelRowFromMasterData.TabIndex = 8;
            this.labelRowFromMasterData.Text = "Row From";
            // 
            // textBoxThicknesColumn
            // 
            this.textBoxThicknesColumn.Location = new System.Drawing.Point(139, 123);
            this.textBoxThicknesColumn.Name = "textBoxThicknesColumn";
            this.textBoxThicknesColumn.Size = new System.Drawing.Size(100, 20);
            this.textBoxThicknesColumn.TabIndex = 7;
            this.textBoxThicknesColumn.Text = "L";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(39, 126);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(94, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Thickness Column";
            // 
            // textBoxWidthColumn
            // 
            this.textBoxWidthColumn.Location = new System.Drawing.Point(139, 97);
            this.textBoxWidthColumn.Name = "textBoxWidthColumn";
            this.textBoxWidthColumn.Size = new System.Drawing.Size(100, 20);
            this.textBoxWidthColumn.TabIndex = 5;
            this.textBoxWidthColumn.Text = "K";
            // 
            // labelWidthColumn
            // 
            this.labelWidthColumn.AutoSize = true;
            this.labelWidthColumn.Location = new System.Drawing.Point(60, 100);
            this.labelWidthColumn.Name = "labelWidthColumn";
            this.labelWidthColumn.Size = new System.Drawing.Size(73, 13);
            this.labelWidthColumn.TabIndex = 4;
            this.labelWidthColumn.Text = "Width Column";
            // 
            // textBoxLengthColumn
            // 
            this.textBoxLengthColumn.Location = new System.Drawing.Point(139, 71);
            this.textBoxLengthColumn.Name = "textBoxLengthColumn";
            this.textBoxLengthColumn.Size = new System.Drawing.Size(100, 20);
            this.textBoxLengthColumn.TabIndex = 3;
            this.textBoxLengthColumn.Text = "J";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(55, 74);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(78, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Length Column";
            // 
            // textBoxArticleNumberColumnMasterData
            // 
            this.textBoxArticleNumberColumnMasterData.Location = new System.Drawing.Point(139, 45);
            this.textBoxArticleNumberColumnMasterData.Name = "textBoxArticleNumberColumnMasterData";
            this.textBoxArticleNumberColumnMasterData.Size = new System.Drawing.Size(100, 20);
            this.textBoxArticleNumberColumnMasterData.TabIndex = 1;
            this.textBoxArticleNumberColumnMasterData.Text = "C";
            // 
            // labelArticleNumberColumnMasterData
            // 
            this.labelArticleNumberColumnMasterData.AutoSize = true;
            this.labelArticleNumberColumnMasterData.Location = new System.Drawing.Point(19, 48);
            this.labelArticleNumberColumnMasterData.Name = "labelArticleNumberColumnMasterData";
            this.labelArticleNumberColumnMasterData.Size = new System.Drawing.Size(114, 13);
            this.labelArticleNumberColumnMasterData.TabIndex = 0;
            this.labelArticleNumberColumnMasterData.Text = "Article Number Column";
            // 
            // groupBoxForecastFormat
            // 
            this.groupBoxForecastFormat.Controls.Add(this.labelRowToForecast);
            this.groupBoxForecastFormat.Controls.Add(this.textBoxRowToForecast);
            this.groupBoxForecastFormat.Controls.Add(this.labelRowFromForecast);
            this.groupBoxForecastFormat.Controls.Add(this.textBoxRowFromForecast);
            this.groupBoxForecastFormat.Controls.Add(this.labelQuantityColumn);
            this.groupBoxForecastFormat.Controls.Add(this.textBoxQuantityColumn);
            this.groupBoxForecastFormat.Controls.Add(this.labelArticleNumberColumnForecast);
            this.groupBoxForecastFormat.Controls.Add(this.textBoxArticleNumberColumnForecast);
            this.groupBoxForecastFormat.Controls.Add(this.label1);
            this.groupBoxForecastFormat.Controls.Add(this.textBoxSheetNr);
            this.groupBoxForecastFormat.Location = new System.Drawing.Point(710, 224);
            this.groupBoxForecastFormat.Name = "groupBoxForecastFormat";
            this.groupBoxForecastFormat.Size = new System.Drawing.Size(263, 158);
            this.groupBoxForecastFormat.TabIndex = 9;
            this.groupBoxForecastFormat.TabStop = false;
            this.groupBoxForecastFormat.Text = "Forecast Format";
            // 
            // labelRowToForecast
            // 
            this.labelRowToForecast.AutoSize = true;
            this.labelRowToForecast.Location = new System.Drawing.Point(88, 127);
            this.labelRowToForecast.Name = "labelRowToForecast";
            this.labelRowToForecast.Size = new System.Drawing.Size(45, 13);
            this.labelRowToForecast.TabIndex = 9;
            this.labelRowToForecast.Text = "Row To";
            // 
            // textBoxRowToForecast
            // 
            this.textBoxRowToForecast.Location = new System.Drawing.Point(139, 124);
            this.textBoxRowToForecast.Name = "textBoxRowToForecast";
            this.textBoxRowToForecast.Size = new System.Drawing.Size(100, 20);
            this.textBoxRowToForecast.TabIndex = 8;
            this.textBoxRowToForecast.Text = "1000";
            // 
            // labelRowFromForecast
            // 
            this.labelRowFromForecast.AutoSize = true;
            this.labelRowFromForecast.Location = new System.Drawing.Point(78, 101);
            this.labelRowFromForecast.Name = "labelRowFromForecast";
            this.labelRowFromForecast.Size = new System.Drawing.Size(55, 13);
            this.labelRowFromForecast.TabIndex = 7;
            this.labelRowFromForecast.Text = "Row From";
            // 
            // textBoxRowFromForecast
            // 
            this.textBoxRowFromForecast.Location = new System.Drawing.Point(139, 98);
            this.textBoxRowFromForecast.Name = "textBoxRowFromForecast";
            this.textBoxRowFromForecast.Size = new System.Drawing.Size(100, 20);
            this.textBoxRowFromForecast.TabIndex = 6;
            this.textBoxRowFromForecast.Text = "2";
            // 
            // labelQuantityColumn
            // 
            this.labelQuantityColumn.AutoSize = true;
            this.labelQuantityColumn.Location = new System.Drawing.Point(49, 75);
            this.labelQuantityColumn.Name = "labelQuantityColumn";
            this.labelQuantityColumn.Size = new System.Drawing.Size(84, 13);
            this.labelQuantityColumn.TabIndex = 5;
            this.labelQuantityColumn.Text = "Quantity Column";
            // 
            // textBoxQuantityColumn
            // 
            this.textBoxQuantityColumn.Location = new System.Drawing.Point(139, 72);
            this.textBoxQuantityColumn.Name = "textBoxQuantityColumn";
            this.textBoxQuantityColumn.Size = new System.Drawing.Size(100, 20);
            this.textBoxQuantityColumn.TabIndex = 4;
            this.textBoxQuantityColumn.Text = "C";
            // 
            // labelArticleNumberColumnForecast
            // 
            this.labelArticleNumberColumnForecast.AutoSize = true;
            this.labelArticleNumberColumnForecast.Location = new System.Drawing.Point(19, 49);
            this.labelArticleNumberColumnForecast.Name = "labelArticleNumberColumnForecast";
            this.labelArticleNumberColumnForecast.Size = new System.Drawing.Size(114, 13);
            this.labelArticleNumberColumnForecast.TabIndex = 3;
            this.labelArticleNumberColumnForecast.Text = "Article Number Column";
            // 
            // textBoxArticleNumberColumnForecast
            // 
            this.textBoxArticleNumberColumnForecast.Location = new System.Drawing.Point(139, 46);
            this.textBoxArticleNumberColumnForecast.Name = "textBoxArticleNumberColumnForecast";
            this.textBoxArticleNumberColumnForecast.Size = new System.Drawing.Size(100, 20);
            this.textBoxArticleNumberColumnForecast.TabIndex = 2;
            this.textBoxArticleNumberColumnForecast.Text = "1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(88, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Sheet #";
            // 
            // textBoxSheetNr
            // 
            this.textBoxSheetNr.Location = new System.Drawing.Point(139, 20);
            this.textBoxSheetNr.Name = "textBoxSheetNr";
            this.textBoxSheetNr.Size = new System.Drawing.Size(100, 20);
            this.textBoxSheetNr.TabIndex = 0;
            this.textBoxSheetNr.Text = "1";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 190);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(62, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "Debug info:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(13, 213);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(65, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Master Data";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(230, 213);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(48, 13);
            this.label6.TabIndex = 10;
            this.label6.Text = "Forecast";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(478, 213);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(37, 13);
            this.label7.TabIndex = 10;
            this.label7.Text = "Result";
            // 
            // buttonCopyToClipboardForExcelPaste
            // 
            this.buttonCopyToClipboardForExcelPaste.Location = new System.Drawing.Point(13, 338);
            this.buttonCopyToClipboardForExcelPaste.Name = "buttonCopyToClipboardForExcelPaste";
            this.buttonCopyToClipboardForExcelPaste.Size = new System.Drawing.Size(674, 44);
            this.buttonCopyToClipboardForExcelPaste.TabIndex = 11;
            this.buttonCopyToClipboardForExcelPaste.Text = "Copy to Clipboard for Excel Paste";
            this.buttonCopyToClipboardForExcelPaste.UseVisualStyleBackColor = true;
            this.buttonCopyToClipboardForExcelPaste.Click += new System.EventHandler(this.buttonCopyToClipboardForExcelPaste_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(233, 183);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(454, 23);
            this.progressBar.TabIndex = 12;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(997, 399);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.buttonCopyToClipboardForExcelPaste);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.groupBoxForecastFormat);
            this.Controls.Add(this.groupBoxMAsterDataFormat);
            this.Controls.Add(this.textBoxResult);
            this.Controls.Add(this.textBoxForecast);
            this.Controls.Add(this.textBoxMasterdata);
            this.Controls.Add(this.buttonGo);
            this.Controls.Add(this.buttonForeCastFile);
            this.Controls.Add(this.textBoxForecastFile);
            this.Controls.Add(this.labelForecastFile);
            this.Controls.Add(this.buttonOpenMasterDataFile);
            this.Controls.Add(this.textBoxMasterDataFile);
            this.Controls.Add(this.labelMasterDataFile);
            this.Name = "Form1";
            this.Text = "Salvagnini Forecast Helper";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBoxMAsterDataFormat.ResumeLayout(false);
            this.groupBoxMAsterDataFormat.PerformLayout();
            this.groupBoxForecastFormat.ResumeLayout(false);
            this.groupBoxForecastFormat.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelMasterDataFile;
        private System.Windows.Forms.TextBox textBoxMasterDataFile;
        private System.Windows.Forms.Button buttonOpenMasterDataFile;
        private System.Windows.Forms.Label labelForecastFile;
        private System.Windows.Forms.TextBox textBoxForecastFile;
        private System.Windows.Forms.Button buttonForeCastFile;
        private System.Windows.Forms.Button buttonGo;
        private System.Windows.Forms.TextBox textBoxMasterdata;
        private System.Windows.Forms.TextBox textBoxForecast;
        private System.Windows.Forms.TextBox textBoxResult;
        private System.Windows.Forms.GroupBox groupBoxMAsterDataFormat;
        private System.Windows.Forms.TextBox textBoxRowFromMasterData;
        private System.Windows.Forms.Label labelRowFromMasterData;
        private System.Windows.Forms.TextBox textBoxThicknesColumn;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBoxWidthColumn;
        private System.Windows.Forms.Label labelWidthColumn;
        private System.Windows.Forms.TextBox textBoxLengthColumn;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxArticleNumberColumnMasterData;
        private System.Windows.Forms.Label labelArticleNumberColumnMasterData;
        private System.Windows.Forms.TextBox textBoxRowToMasterData;
        private System.Windows.Forms.Label labelRowToMasterData;
        private System.Windows.Forms.Label labelSheetNameMasterData;
        private System.Windows.Forms.TextBox textBoxSheetNameMasterData;
        private System.Windows.Forms.GroupBox groupBoxForecastFormat;
        private System.Windows.Forms.Label labelRowFromForecast;
        private System.Windows.Forms.TextBox textBoxRowFromForecast;
        private System.Windows.Forms.Label labelQuantityColumn;
        private System.Windows.Forms.TextBox textBoxQuantityColumn;
        private System.Windows.Forms.Label labelArticleNumberColumnForecast;
        private System.Windows.Forms.TextBox textBoxArticleNumberColumnForecast;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxSheetNr;
        private System.Windows.Forms.Label labelRowToForecast;
        private System.Windows.Forms.TextBox textBoxRowToForecast;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button buttonCopyToClipboardForExcelPaste;
        private System.Windows.Forms.ProgressBar progressBar;
    }
}

