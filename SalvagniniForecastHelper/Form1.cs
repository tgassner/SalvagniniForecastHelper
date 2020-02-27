using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SalvagniniForecastHelper
{
    public partial class Form1 : Form
    {

        IList<ResultLine> result = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void buttonOpenMasterDataFile_Click(object sender, EventArgs e)
        {
            openFile(this.textBoxMasterDataFile);
        }

        private void openFile(TextBox textBox)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = @"G:\Kunden\Kalkulation\salvagnini";
                openFileDialog.Filter = "excel files (*.xlsx)|*.xlsx"; // "|All files (*.*)|*.*";
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    textBox.Text = filePath;
                }
            }
        }

        private void textBoxMasterDataFile_DragDrop(object sender, DragEventArgs e)
        {
            doDragDrop(this.textBoxMasterDataFile, e);
        }

        private void doDragDrop(TextBox textBox, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Length != 0)
            {
                textBox.Text = files[0];
            }
        }

        private void textBoxMasterDataFile_DragEnter(object sender, DragEventArgs e)
        {
            doDragEnterFiles(e);
        }

        private void doDragEnterFiles(DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length == 1)
                {
                    if (files[0].EndsWith(".xlsx"))
                    {
                        e.Effect = DragDropEffects.Copy;
                    }
                    else
                    {
                        e.Effect = DragDropEffects.None;
                    }
                }
                else
                {
                    e.Effect = DragDropEffects.None;
                }
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void buttonForeCastFile_Click(object sender, EventArgs e)
        {
            openFile(this.textBoxForecastFile);
        }

        private void textBoxForecastFile_DragDrop(object sender, DragEventArgs e)
        {
            doDragDrop(this.textBoxForecastFile, e);
        }

        private void textBoxForecastFile_DragEnter(object sender, DragEventArgs e)
        {
            doDragEnterFiles(e);
        }

        private void buttonGo_Click(object sender, EventArgs e)
        {
            if (!doValidityChecks())
            {
                return;
            }

            doCalculation();
        }

        private void doCalculation()
        {
            progressBar.Style = ProgressBarStyle.Marquee;
            buttonGo.Enabled = false;
            buttonOpenMasterDataFile.Enabled = false;
            buttonForeCastFile.Enabled = false;
            buttonCopyToClipboardForExcelPaste.Enabled = false;
            textBoxMasterDataFile.Enabled = false;
            textBoxForecastFile.Enabled = false;
            textBoxSheetNameMasterData.Enabled = false;
            textBoxArticleNumberColumnMasterData.Enabled = false;
            textBoxArticleNumberColumnMasterData.Enabled = false;
            textBoxLengthColumn.Enabled = false;
            textBoxWidthColumn.Enabled = false;
            textBoxThicknesColumn.Enabled = false;
            textBoxRowFromMasterData.Enabled = false;
            textBoxRowToMasterData.Enabled = false;

            IDictionary<string, MasterDataLine> masterData = null;

            try
            {
                masterData = readMasterData();
            }
            catch (IOException e)
            {
                MessageBox.Show("Error While opening the File!\r\n" + e.Message);
                this.result = null;
                restoreGui();
                return;
            }

            if (masterData == null)
            {
                this.result = null;
                restoreGui();
                return;
            }

            IDictionary<string, ForecastLine> forecast = readForecast();

            IList<ResultLine> result = new List<ResultLine>();

            foreach (ForecastLine forecastLine in forecast.Values)
            {
                Application.DoEvents();
                ResultLine resultLine = new ResultLine();

                resultLine.ArtikelNrForecast = forecastLine.ArtikelNr;
                resultLine.Quantities = forecastLine.Quantities;
                resultLine.Quantity = forecastLine.Quantities.Sum();

                MasterDataLine masterDataLine = null;

                if (masterData.ContainsKey(forecastLine.ArtikelNr.ToUpper()))
                {
                    masterDataLine = masterData[forecastLine.ArtikelNr.ToUpper()];
                }
                else
                {
                    string artNr = forecastLine.ArtikelNr.Substring(1, forecastLine.ArtikelNr.Length - 3).ToUpper();
                    if (masterData.ContainsKey(artNr))
                    {
                        masterDataLine = masterData[artNr];
                    }
                }

                if (masterDataLine != null)
                {
                    resultLine.ArtikelNrMasterData = masterDataLine.ArtikelNr;
                    resultLine.Length = masterDataLine.Length;
                    resultLine.Width = masterDataLine.Width;
                    resultLine.Thickness = masterDataLine.Thickness;
                }

                result.Add(resultLine);
            }

            string s = "";
            foreach (MasterDataLine value in masterData.Values)
            {
                s += value.ToString() + "\r\n";
            }
            textBoxMasterdata.Text = s;

            s = "";
            foreach (ForecastLine value in forecast.Values)
            {
                s += value.ToString() + "\r\n";
            }
            textBoxForecast.Text = s;

            s = "";
            foreach (ResultLine value in result)
            {
                s += value.ToString() + "\r\n";
            }
            textBoxResult.Text = s;

            this.result = result;

            restoreGui();
        }

        private void restoreGui()
        {
            buttonGo.Enabled = true;
            buttonOpenMasterDataFile.Enabled = true;
            buttonForeCastFile.Enabled = true;
            buttonCopyToClipboardForExcelPaste.Enabled = true;
            textBoxMasterDataFile.Enabled = true;
            textBoxForecastFile.Enabled = true;
            textBoxSheetNameMasterData.Enabled = true;
            textBoxArticleNumberColumnMasterData.Enabled = true;
            textBoxArticleNumberColumnMasterData.Enabled = true;
            textBoxLengthColumn.Enabled = true;
            textBoxWidthColumn.Enabled = true;
            textBoxThicknesColumn.Enabled = true;
            textBoxRowFromMasterData.Enabled = true;
            textBoxRowToMasterData.Enabled = true;
            progressBar.Style = ProgressBarStyle.Blocks;
        }

        private IDictionary<string, ForecastLine> readForecast()
        {
            IDictionary<string, ForecastLine> forecast = new SortedDictionary<string, ForecastLine>();

            using (SpreadsheetDocument ssd = SpreadsheetDocument.Open(this.textBoxForecastFile.Text, false))
            {
                WorkbookPart workbookPart = ssd.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();

                WorksheetPart wsPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;

                if (wsPart == null)
                {
                    return forecast;
                }

                Row[] rows = wsPart.Worksheet.Descendants<Row>().ToArray();

                int countLines = 0;

                foreach (Row row in wsPart.Worksheet.Descendants<Row>())
                {
                    Application.DoEvents();
                    List<object> rowData = new List<object>();
                    foreach (Cell c in row.Elements<Cell>())
                    {
                        rowData.Add(GetCellValue(c));
                    }

                    countLines++;

                    if (countLines > 1 && // don't deal with the Header line
                        rowData != null &&
                        rowData.Count >= 3 &&
                        rowData.ElementAt<Object>(0) != null &&
                        !string.IsNullOrWhiteSpace(rowData.ElementAt<object>(0).ToString()))
                    {
                        string artikelNr = "";
                        if (rowData.Count >= 1 && rowData.ElementAt<Object>(0) != null)
                        {
                            artikelNr = rowData.ElementAt<Object>(0).ToString().Trim();
                        }

                        int quantity = 0;
                        if (rowData.Count >= 3 && rowData.ElementAt<Object>(2) != null)
                        {
                            string qs = rowData.ElementAt<Object>(2).ToString().Trim();
                            int.TryParse(qs, out quantity);
                        }


                        if (!forecast.ContainsKey(artikelNr.ToUpper()))
                        {
                            forecast[artikelNr.ToUpper()] = new ForecastLine(artikelNr);
                        }
                        forecast[artikelNr.ToUpper()].Quantities.Add(quantity);
                    }
                }
            }
            return forecast;
        }

        private IDictionary<string, MasterDataLine> readMasterData()
        {
            IDictionary<string, MasterDataLine> masterData = new Dictionary<string, MasterDataLine>();

            using (SpreadsheetDocument ssd = SpreadsheetDocument.Open(this.textBoxMasterDataFile.Text, false))
            {
                WorkbookPart workbookPart = ssd.WorkbookPart;
                // Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();


                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == textBoxSheetNameMasterData.Text).FirstOrDefault();

                // Throw an exception if there is no sheet.
                if (sheet == null)
                {
                    MessageBox.Show("Text Box Sheet Name Master Data is invalid: Sheet " + textBoxSheetNameMasterData.Text + " does not exist!");
                    return null;
                }

                WorksheetPart wsPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;

                if (wsPart == null)
                {
                    MessageBox.Show("No WorksheetPart Found!");
                    return null;
                }


                int countNullValues = 0;

                for (int i = int.Parse(textBoxRowFromMasterData.Text) - 1; i < int.Parse(textBoxRowToMasterData.Text) ; i++)
                {
                    Application.DoEvents();
                    Cell cellArticleNumber = wsPart.Worksheet.Descendants<Cell>().
                        Where(c => c.CellReference == (textBoxArticleNumberColumnMasterData.Text + i)).FirstOrDefault();

                    if (cellArticleNumber == null)
                    {
                        countNullValues++;
                    } 
                    else
                    {
                        countNullValues = 0;
                    }

                    if (countNullValues == 20)
                    {
                        break;
                    }

                    Cell cellLength = wsPart.Worksheet.Descendants<Cell>().
                        Where(c => c.CellReference == (textBoxLengthColumn.Text + i)).FirstOrDefault();

                    Cell cellWidth = wsPart.Worksheet.Descendants<Cell>().
                        Where(c => c.CellReference == (textBoxWidthColumn.Text + i)).FirstOrDefault();

                    Cell cellThickness = wsPart.Worksheet.Descendants<Cell>().
                        Where(c => c.CellReference == (textBoxThicknesColumn.Text + i)).FirstOrDefault();

                    if (cellArticleNumber != null) { 

                    MasterDataLine masterDataLine = new MasterDataLine();

                    string artNr = GetCellValue(cellArticleNumber);

                        if (!string.IsNullOrWhiteSpace(artNr))
                        {

                            if (masterData.ContainsKey(artNr.ToUpper()))
                            {
                                MessageBox.Show("The Article Number " + artNr + " occures more than once!");
                            }

                            masterDataLine.ArtikelNr = artNr;

                            double d = 0;

                            if (cellLength != null)
                            {
                                double.TryParse(GetCellValue(cellLength), NumberStyles.Number, CultureInfo.CreateSpecificCulture("en-US"), out d);
                                masterDataLine.Length = d;
                            }

                            if (cellWidth != null)
                            {
                                double.TryParse(GetCellValue(cellWidth), NumberStyles.Number, CultureInfo.CreateSpecificCulture("en-US"), out d);
                                masterDataLine.Width = d;
                            }

                            if (cellThickness != null)
                            {
                                double.TryParse(GetCellValue(cellThickness), NumberStyles.Number, CultureInfo.CreateSpecificCulture("en-US"), out d);
                                masterDataLine.Thickness = d;
                            }

                            masterData[masterDataLine.ArtikelNr.ToUpper()] = masterDataLine;
                        }
                    }
                }
            }
            return masterData;
        }

        class ResultLine
        {
            private string artikelNrMasterData = "";
            private string artikelNrForecast = "";
            private IList<int> quantities = new List<int>();
            private int quantity = 0;
            private double length = 0.0;
            private double width = 0.0;
            private double thickness = 0.0;

            public string ArtikelNrMasterData { get => artikelNrMasterData; set => artikelNrMasterData = value; }
            public string ArtikelNrForecast { get => artikelNrForecast; set => artikelNrForecast = value; }
            public IList<int> Quantities { get => quantities; set => quantities = value; }
            public int Quantity { get => quantity; set => quantity = value; }
            public double Length { get => length; set => length = value; }
            public double Width { get => width; set => width = value; }
            public double Thickness { get => thickness; set => thickness = value; }

            public override string ToString()
            {
                return "ArtikelNrForecast: " + ArtikelNrForecast +
                     ";  ArtikelNrMasterData: " + ArtikelNrMasterData +
                     ";  Quantities: " + string.Join(", ", Quantities) +
                     "; Quantity: " + Quantity +
                     ";  length: " + Length +
                     ";  width: " + Width +
                     ";  thickness:" + Thickness;
            }
        }

        class ForecastLine
        {
            private string artikelNr = "";
            private IList<int> quantities = new List<int>();

            public ForecastLine (string artikelNr) {
                this.artikelNr = artikelNr;
            }

            public string ArtikelNr { get => artikelNr; set => artikelNr = value; }
            public IList<int> Quantities { get => quantities; set => quantities = value; }

            public override string ToString()
            {
                return ArtikelNr + ": " + string.Join(", ", Quantities);
            }
        }

        class MasterDataLine
        {
            private string artikelNr = "";
            private double length = 0.0;
            private double width = 0.0;
            private double thickness = 0.0;

            public string ArtikelNr { get => artikelNr; set => artikelNr = value; }
            public double Length { get => length; set => length = value; }
            public double Width { get => width; set => width = value; }
            public double Thickness { get => thickness; set => thickness = value; }

            public override string ToString()
            {
                return ArtikelNr + ": " + Length + " * " + Width + " * " + Thickness;
            }
        }

        private string GetCellValue(Cell cell)
        {
            if (cell == null) { 
                return null;
            }

            string value = cell.InnerText;

            if (cell.DataType == null)
            {
                return cell.InnerText;
            }

            //string value = cell.InnerText;
            switch (cell.DataType.Value)
            {
                case CellValues.SharedString:
                    // For shared strings, look up the value in the shared strings table.
                    // Get worksheet from cell
                    OpenXmlElement parent = cell.Parent;
                    while (parent.Parent != null && parent.Parent != parent
                            && string.Compare(parent.LocalName, "worksheet", true) != 0)
                    {
                        parent = parent.Parent;
                    }
                    if (string.Compare(parent.LocalName, "worksheet", true) != 0)
                    {
                        throw new Exception("Unable to find parent worksheet.");
                    }

                    Worksheet ws = parent as Worksheet;
                    SpreadsheetDocument ssDoc = ws.WorksheetPart.OpenXmlPackage as SpreadsheetDocument;
                    SharedStringTablePart sstPart = ssDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                    // lookup value in shared string table
                    if (sstPart != null && sstPart.SharedStringTable != null)
                    {
                        value = sstPart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                    }
                    break;

                case CellValues.Number:
                    value = "Number: " + value;
                    break;
                //this case within a case is copied from msdn. 
                case CellValues.Boolean:
                    switch (value)
                    {
                        case "0":
                            value = "FALSE";
                            break;
                        default:
                            value = "TRUE";
                            break;
                    }
                    break;
            }
            return value;
        }

        private bool doValidityChecks()
        {
            if (string.IsNullOrWhiteSpace(textBoxMasterDataFile.Text))
            {
                MessageBox.Show("Text box Master data file is empty :-(");
                return false;
            }

            if (string.IsNullOrWhiteSpace(textBoxForecastFile.Text))
            {
                MessageBox.Show("Text box foecast file is empty :-(");
                return false;
            }

            if (!File.Exists(textBoxMasterDataFile.Text))
            {
                MessageBox.Show("File in text box Master data file does not exist :-(");
                return false;
            }

            if (!File.Exists(textBoxForecastFile.Text))
            {
                MessageBox.Show("File in text box forecast file does not exist :-(");
                return false;
            }

            if (!textBoxMasterDataFile.Text.EndsWith(".xlsx"))
            {
                MessageBox.Show("File in text box Master data file is not a Excel file (*.xlsx) :-(");
                return false;
            }

            if (!textBoxForecastFile.Text.EndsWith(".xlsx"))
            {
                MessageBox.Show("File in text box forecast file is not a Excel file (*.xlsx) :-(");
                return false;
            }

            if (string.IsNullOrWhiteSpace(textBoxSheetNameMasterData.Text))
            {
                MessageBox.Show("Text Box Master Data Sheet Name must not be empty!");
                return false;
            }

            if (string.IsNullOrWhiteSpace(textBoxArticleNumberColumnMasterData.Text))
            {
                MessageBox.Show("Text Box Article Number Column Master Data must not be empty!");
                return false;
            }

            if (string.IsNullOrWhiteSpace(textBoxLengthColumn.Text))
            {
                MessageBox.Show("Text Box Length must not be empty!");
                return false;
            }

            if (string.IsNullOrWhiteSpace(textBoxWidthColumn.Text))
            {
                MessageBox.Show("Text Box Width must not be empty!");
                return false;
            }

            if (string.IsNullOrWhiteSpace(textBoxThicknesColumn.Text))
            {
                MessageBox.Show("Text Box Thickness must not be empty!");
                return false;
            }

            int from = 0;

            if (!int.TryParse(textBoxRowFromMasterData.Text, out from) || from <= 1)
            {
                MessageBox.Show("Text Box Row From Master Data must be a natural number > 0!");
                return false;
            }

            int to = 0;

            if (!int.TryParse(textBoxRowToMasterData.Text, out to) || from > to)
            {
                MessageBox.Show("Text Box Row To Master Data must be a natural number >= from !");
                return false;
            }

            int sheetNo = 0;
            if (!int.TryParse(textBoxSheetNr.Text, out sheetNo) || sheetNo <= 0)
            {
                MessageBox.Show("Text Box Row To Master Data must be a natural number >= 1 !");
                return false;
            }

            if (string.IsNullOrWhiteSpace(textBoxArticleNumberColumnForecast.Text))
            {
                MessageBox.Show("Text Box Article Number Column Forecast must not be empty!");
                return false;
            }

            if (string.IsNullOrWhiteSpace(textBoxQuantityColumn.Text))
            {
                MessageBox.Show("Text Box Quantity Column Forecast must not be empty!");
                return false;
            }

            if (!int.TryParse(textBoxRowFromForecast.Text, out from) || from <= 1)
            {
                MessageBox.Show("Text Box Row From Forecast must be a natural number > 0!");
                return false;
            }

            if (!int.TryParse(textBoxRowToForecast.Text, out to) || from > to)
            {
                MessageBox.Show("Text Box Row To Forecast must be a natural number >= from !");
                return false;
            }

            return true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            groupBoxForecastFormat.Visible = false;
            this.Text = this.Text + "  Version:" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        }

        private void buttonCopyToClipboardForExcelPaste_Click(object sender, EventArgs e)
        {
            if (this.result == null)
            {
                MessageBox.Show("there is no result! please calculate first");
                return;
            }
            
            string s = "ArtikelNrForecast\tArtikelNrMasterData\tQuantities\tQuantity Sum\tLength\tWidth\tThickness\r\n";

            foreach (ResultLine value in result)
            {
                s += value.ArtikelNrForecast + "\t" +
                    value.ArtikelNrMasterData + "\t" +
                    string.Join(" + ", value.Quantities) + "\t" +
                    value.Quantity + "\t" +
                    value.Length + "\t" +
                    value.Width + "\t" +
                    value.Thickness + "\t" +
                    "\r\n";
            }

            System.Windows.Forms.Clipboard.SetText(s);
        }
    }
}