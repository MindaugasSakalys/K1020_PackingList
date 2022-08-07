using DataLibrary;
using K1020_PackingList.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Squirrel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace K1020_PackingList
{
    public partial class Form_Main : Form
    {

        private bool printDialog = false;
        private List<SerialTmp_model> serialNrTmp = new List<SerialTmp_model>();
        private Code_model selectedCode = new Code_model();

        private string uniqueBoxNumber = "";
        private int boxCount = 0;
        private int productBoxCount = 0;

        private string serialTmpLabel2 = "";
        private int smallBoxTmpLabel2 = 0;

        private async void Form_Main_Load(object sender, EventArgs e)
        {
            StartScreen(false);
            await CheckForUpdates();

            DataAccess.DBmain(2); // 1-Raspberry  2-Server

            if (DataAccess.ConnToMysql())
            {
                StartScreen(true);
                dataGridView1.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(dataGridView1, true, null);
                dataGridView2.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(dataGridView1, true, null);
                CleanFormAll();
            }
            else
            {
                MessageBox.Show("Nėra ryšio su duomenų baze");
                this.Close();
            }
        }

        //******************************************************* Start Form ***************************************************
        public Form_Main()
        {
            InitializeComponent();
        }

        //**************************************************************************************************************************
        //****************************************** AUTO UPDATE  ******************************************************************
        //**************************************************************************************************************************
        private async Task CheckForUpdates()
        {
            try
            {
                ReleaseEntry release = null;

                using (var manager = new UpdateManager(@"http://78.57.2.98/releases/KG1020PackingList/"))
                {
                    UpdateInfo updateInfo = await manager.CheckForUpdate();

                    if (updateInfo.ReleasesToApply.Any()) // Check if we have any update
                    {
                        System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
                        FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                        string msg = "New version available!" +
                                "\n\nCurrent version: " + updateInfo.CurrentlyInstalledVersion.Version +
                                "\nNew version: " + updateInfo.FutureReleaseEntry.Version +
                                "\n\nUpdate application now";
                        DialogResult dialogResult = MessageBox.Show(msg, fvi.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);

                        if (dialogResult == DialogResult.OK)
                        {
                            // Do the update
                            release = await manager.UpdateApp();
                        }
                        else
                        {
                            this.Close();
                        }
                    }

                    // Restart the app
                    if (release != null)
                    {
                        UpdateManager.RestartApp();
                    }
                }
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message, "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            AddVersionNumber();
        }
        private void AddVersionNumber()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            FileVersionInfo versionInfo = FileVersionInfo.GetVersionInfo(assembly.Location);

            VersionLabel.Text = $" v.{versionInfo.FileVersion}";
        }

        private void StartScreen(bool _visible)
        {
            panel1.Visible = _visible;
            tabControl1.Visible = _visible;
        }


        //******************************************************* Button Function **********************************************
        private void button1_Click(object sender, EventArgs e)//Pallet list button
        {
            Form_Pallets formPallets = new Form_Pallets();
            formPallets.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)// Pallet code button
        {
            Form_Code codePallet = new Form_Code();
            codePallet.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)// Clean serial list
        {
            CleanFirstScan();
        }

        //******************************************************* ComboBox Function **********************************************

        private async void comboBox1_DropDown(object sender, EventArgs e)
        {
            comboBox1.DataSource = null;
            comboBox1.DataSource = await LoadPallets();
            comboBox1.DisplayMember = "PalletName";
            comboBox2.DataSource = null;
            comboBox2.Enabled = false;
            checkBox2.Checked = false;
            textBox1.Enabled = false;
            CleanInfoPanel();
        }

        private async void comboBox1_DropDownClosed(object sender, EventArgs e)
        {
            if (comboBox1.Text.Length != 0)
            {
                Pallets_model _pallet = (Pallets_model)comboBox1.SelectedItem;

                if (await IsSerial_palletId(_pallet.Id))
                {
                    comboBox2.Enabled = false;
                    checkBox2.Checked = false;
                    comboBox3.Enabled = true;
                    textBox1.Enabled = true;
                    textBox1.Focus();
                    await ShowAllSerialNbox();
                    await ShowAllSerialNboxInfo();

                    int code_id = int.Parse(this.dataGridView2.CurrentRow.Cells[7].Value.ToString());
                    comboBox2.DataSource = null;
                    comboBox2.DataSource = await GetCodeName(code_id);
                    comboBox2.DisplayMember = "CodeName";
                    selectedCode = (Code_model)comboBox2.SelectedItem;
                    ShowInfoPanel(selectedCode);
                }
                else
                {
                    comboBox2.Enabled = true;
                    checkBox2.Checked = true;
                    comboBox2.Focus();
                    await ShowAllSerialNbox();
                    await ShowAllSerialNboxInfo();
                }
            }
            else
            {
                dataGridView1.DataSource = null;
                dataGridView2.DataSource = null;
                dataGridView3.DataSource = null;
            }
        }

        private void CleanInfoPanel()
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
        }

        private async void comboBox2_DropDown(object sender, EventArgs e)
        {
            comboBox2.DataSource = null;
            comboBox2.DataSource = await LoadCodes();
            comboBox2.DisplayMember = "CodeName";

            serialNrTmp.Clear();
            textBox1.Text = "";
        }

        private void comboBox2_DropDownClosed(object sender, EventArgs e)
        {
            if (comboBox1.Text.Length != 0 && comboBox2.Text.Length != 0)
            {
                comboBox3.Enabled = true;
                textBox1.Enabled = true;
                textBox1.Text = "";
                textBox1.Focus();
                selectedCode = (Code_model)comboBox2.SelectedItem;
                ShowInfoPanel(selectedCode);
            }
        }

        private void ShowInfoPanel(Code_model _code)
        {
            textBox2.Text = _code.CountryCode;
            textBox3.Text = _code.Version;
            textBox5.Text = _code.BatteryCount.ToString();
            textBox4.Text = _code.CountryName;
        }

        //******************************************************* MySQL Function **********************************************
        //--------------------------------Pallet
        private async Task<List<Pallets_model>> LoadPallets()
        {
            List<Pallets_model> output = new List<Pallets_model>();
            string sql = "SELECT * FROM pallet WHERE DonePallet = 'false' ORDER BY PalletName DESC";
            output = await DataAccess.LoadData<Pallets_model, dynamic>(sql, new { });
            return output;
        }

        private async Task<bool> LoadPallets(string _palletName)
        {
            bool output = false;
            List<Pallets_model> allPalletList = new List<Pallets_model>();
            string sql = "SELECT * FROM pallet WHERE PalletName = @PalletName";
            allPalletList = await DataAccess.LoadData<Pallets_model, dynamic>(sql, new { PalletName = @_palletName });
            if (allPalletList.Count > 0) output = true;
            return output;
        }

        //--------------------------------Code
        private async Task<List<Code_model>> LoadCodes()
        {
            List<Code_model> output = new List<Code_model>();
            string sql = "SELECT * FROM code WHERE Disabled = 'false' ORDER BY CodeName";
            output = await DataAccess.LoadData<Code_model, dynamic>(sql, new { });
            return output;
        }

        private async Task<List<Code_model>> GetCodeName(int _id)
        {
            List<Code_model> output = new List<Code_model>();
            string sql = "SELECT * FROM code WHERE Id = @Id";
            output = await DataAccess.LoadData<Code_model, dynamic>(sql, new { Id = @_id });
            return output;
        }

        private async Task<bool> LoadCodes(string _codeName)
        {
            bool output = false;
            List<Code_model> allCodeList = new List<Code_model>();
            string sql = "SELECT * FROM code WHERE CodeName = @CodeName";
            allCodeList = await DataAccess.LoadData<Code_model, dynamic>(sql, new { CodeName = @_codeName });
            if (allCodeList.Count > 0) output = true;
            return output;
        }

        //------------------------- Serial
        private async Task<List<Serials_model>> LoadSerials()
        {
            List<Serials_model> output = new List<Serials_model>();
            string sql = "SELECT * FROM serials";
            output = await DataAccess.LoadData<Serials_model, dynamic>(sql, new { });
            return output;
        }

        private async Task<List<Serials_model>> LoadSerialsB(int _palletId)
        {
            List<Serials_model> output = new List<Serials_model>();
            string sql = "SELECT * FROM serials WHERE PaletId = @PaletId";
            output = await DataAccess.LoadData<Serials_model, dynamic>(sql, new { PaletId = @_palletId });
            return output;
        }

        private async Task<List<Serials_model>> LoadSerials(int _palletId)
        {
            List<Serials_model> output = new List<Serials_model>();
            string sql = "SELECT serials.Id, serials.MainBoxNr, serials.BoxNr, serials.SerialNr, serials.PaletId, " +
                "pallet.PalletName AS PaletName, serials.CodeId, code.CodeName AS CodeName, code.CountryCode AS SimCard, " +
                "code.Version AS Version, code.CountryName AS CountryCode, code.BatteryCount AS BatteryCount, serials.AddDateTime, serials.ModDateTime, serials.UniqNr " +
                "FROM serials " +
                "LEFT JOIN pallet ON serials.PaletId = pallet.Id " +
                "LEFT JOIN code ON serials.CodeId = code.Id " +
                "WHERE PaletId = @PaletId";
            output = await DataAccess.LoadData<Serials_model, dynamic>(sql, new { PaletId = @_palletId });
            return output;
        }

        private async Task<bool> IsSerial_palletId(int _palletId)
        {
            bool output = false;
            List<Serials_model> serialList = new List<Serials_model>();
            string sql = "SELECT * FROM serials WHERE PaletId = @PaletId";
            serialList = await DataAccess.LoadData<Serials_model, dynamic>(sql, new { PaletId = @_palletId });
            if (serialList.Count > 0) output = true;
            return output;
        }

        private async Task<bool> IsSerial_serialNr(string _serialNr)
        {
            bool output = false;
            List<Serials_model> serialList = new List<Serials_model>();
            string sql = "SELECT * FROM serials WHERE SerialNr = @SerialNr";
            serialList = await DataAccess.LoadData<Serials_model, dynamic>(sql, new { SerialNr = @_serialNr });
            if (serialList.Count > 0) output = true;
            return output;
        }

        //******************************************************* CRUD Function ***********************************************************************

        private async Task AddSerial(int _mainBox, int _smallBox, string _serialNr, int _palletId, int _codeId, string _uniqNr)
        {
            //if (await ValidateForm())
            //{
            string sql = "INSERT INTO serials (MainBoxNr, BoxNr, SerialNr, PaletId, CodeId, AddDateTime, ModDateTime, UniqNr) VALUE (@MainBoxNr, @BoxNr, @SerialNr, @PaletId, @CodeId, @AddDateTime, @ModDateTime, @UniqNr);";

            await DataAccess.SaveData(sql, new
            {
                MainBoxNr = _mainBox,
                BoxNr = _smallBox,
                SerialNr = _serialNr,
                PaletId = _palletId,
                CodeId = _codeId,
                AddDateTime = DateTime.Now,
                ModDateTime = DateTime.Now,
                UniqNr = _uniqNr
            });
            //}
        }

        //******************************************************* Show DATA GRID 2 Functions *******************************************
        private async Task ShowAllSerialNbox()
        {
            Pallets_model _pallet = (Pallets_model)comboBox1.SelectedItem;
            dataGridView2.DataSource = null;
            var output = new List<Serials_model>();
            output = await LoadSerials(_pallet.Id);
            dataGridView2.DataSource = output;
            label11.Text = output.Count.ToString();

            if (dataGridView2.Rows.Count != 0)
            {
                dataGridView2.CurrentCell = dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[3];
                dataGridView2.ClearSelection();
            }
            DataGridShowColor(dataGridView2);
        }

        private async Task ShowAllSerialNboxInfo()
        {
            Pallets_model _pallet = (Pallets_model)comboBox1.SelectedItem;
            dataGridView3.DataSource = null;
            var output = new List<Serials_model>();
            output = await LoadSerials(_pallet.Id);
            output = output.Where(x => x.SerialNr.Contains(textBox15.Text)).Where(x => x.UniqNr.Contains(textBox16.Text)).ToList();
            dataGridView3.DataSource = output;
            label13.Text = output.Count.ToString();

            if (dataGridView3.Rows.Count != 0)
            {
                dataGridView3.CurrentCell = dataGridView3.Rows[dataGridView3.Rows.Count - 1].Cells[3];
                dataGridView3.ClearSelection();
            }
            DataGridShowColor(dataGridView3);
        }

        //******************************************************* Clear form Functions ********************************************
        private void CleanFormAll()
        {
            comboBox2.Enabled = false;
            checkBox2.Checked = false;
            comboBox3.Enabled = false;
            comboBox3.SelectedIndex = 0;
            textBox1.Enabled = false;
            textBox1.Text = "";
            comboBox1.Focus();
            checkBox1.Checked = true;

            textBox6.Enabled = false;
            button5.Enabled = true;
            button6.Enabled = false;
        }

        private void CleanFirstScan()
        {
            if (comboBox1.Text.Length != 0 && comboBox2.Text.Length != 0 && comboBox3.Text.Length != 0)
            {
                dataGridView1.DataSource = null;
                serialNrTmp.Clear();
                textBox1.Enabled = true;
                textBox1.Text = "";
                textBox1.Focus();
                comboBox1.Enabled = true;

                if (dataGridView2.Rows.Count != 0 || dataGridView1.Rows.Count != 0)
                {
                    comboBox2.Enabled = false;
                    checkBox2.Checked = false;
                }
                else
                {
                    comboBox2.Enabled = true;
                    checkBox2.Checked = false;
                }

                comboBox3.Enabled = true;

                label3.Text = "0";
            }
        }

        //******************************************************* Scan Functions ********************************************
        private async void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Return))
            {
                if (await ValidateForm())
                {
                    SerialTmp_model srNr = new SerialTmp_model();
                    boxCount = await GetBoxNr();
                    productBoxCount = serialNrTmp.Count + 1;
                    srNr.Id = productBoxCount;
                    srNr.SerialNr = textBox1.Text;

                    serialNrTmp.Add(srNr);

                    if (checkBox1.Checked)
                    {
                        if(printDialog)
                        {
                            printPreviewDialog1.Document = printDocument1;
                            printPreviewDialog1.ShowDialog();
                        }
                        else
                        {
                            printDocument1.Print();
                        }

                    }

                    textBox1.Text = "";
                    textBox1.Focus();

                    dataGridView1.DataSource = null;
                    dataGridView1.DataSource = serialNrTmp;

                    if (dataGridView1.Rows.Count != 0)
                    {
                        dataGridView1.CurrentCell = dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[1];
                    }

                    if (serialNrTmp.Count == int.Parse(comboBox3.Text))
                    {
                        await AddSerialToDB();
                        serialNrTmp.Clear();
                        textBox1.Text = "";
                        textBox1.Focus();
                    }

                    if (serialNrTmp.Count != 0)
                    {
                        comboBox1.Enabled = false;
                        comboBox2.Enabled = false;
                        checkBox2.Checked = false;

                        comboBox3.Enabled = false;
                    }

                    label3.Text = serialNrTmp.Count.ToString();
                }
            }
        }

        private async Task AddSerialToDB()
        {
            Pallets_model _pallet = (Pallets_model)comboBox1.SelectedItem;
            Code_model _code = (Code_model)comboBox2.SelectedItem;

            uniqueBoxNumber = comboBox1.Text + comboBox3.Text + textBox4.Text + boxCount;

            if (checkBox1.Checked)
            {
                var result = MessageBox.Show("Spausdinami gaminių serijinių numerių lipdukai ant dėžės", "Informacija", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            foreach (var item in serialNrTmp)
            {
                serialTmpLabel2 = item.SerialNr;
                smallBoxTmpLabel2 = item.Id;
                await AddSerial(boxCount, item.Id, item.SerialNr, _pallet.Id, _code.Id, uniqueBoxNumber);

                if (checkBox1.Checked)
                {
                    if (printDialog)
                    {
                        printPreviewDialog1.Document = printDocument2;
                        printPreviewDialog1.ShowDialog();
                    }
                    else
                    {
                        printDocument2.Print();
                    }
                }
            }

            if (checkBox1.Checked)
            {
                var result = MessageBox.Show("Spausdinamas dėžės kodo lipdukas", "Informacija", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (printDialog)
                {
                    printPreviewDialog1.Document = printDocument3;
                    printPreviewDialog1.ShowDialog();
                }
                else
                {
                    printDocument3.Print();
                }
            }
            if(boxCount == 18)
            {
                var result = MessageBox.Show("Nepamirškite padaryti paletės nuotrauką ir ikelti ją į šią duomenų bazę !", "Informacija", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            CleanFirstScan();
            comboBox2.Enabled = false;
            checkBox2.Checked = false;
            await ShowAllSerialNbox();
            await ShowAllSerialNboxInfo();
        }

        private async Task<int> GetBoxNr()
        {
            int output = 0;
            Pallets_model _pallet = (Pallets_model)comboBox1.SelectedItem;
            if (dataGridView2.Rows.Count != 0)
            {
                List<Serials_model> listSrBox = await LoadSerials(_pallet.Id);
                output = listSrBox[listSrBox.Count - 1].MainBoxNr;
                output++;
            }
            else
            {
                output = 1;
            }

            return output;
        }

        //******************************************************* Validate form *****************************************************
        private async Task<bool> ValidateForm()
        {
            bool output = true;

            if (String.IsNullOrEmpty(textBox1.Text))
            {
                var result = MessageBox.Show("Neįvestas serijinis kodas", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (String.IsNullOrEmpty(comboBox1.Text.Trim()))
            {
                var result = MessageBox.Show("Nepasirinktas paletės pavadinimas", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (String.IsNullOrEmpty(comboBox2.Text.Trim()))
            {
                var result = MessageBox.Show("Nepasirinktas kodas", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (textBox1.Text.Count() != 8)
            {
                var result = MessageBox.Show("Blogas serijinis numeris", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (textBox1.Text.Substring(0, 3) != "313")
            {
                var result = MessageBox.Show("Blogas serijinis numeris", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (serialNrTmp.Any(x => x.SerialNr.Equals(textBox1.Text.Trim())))
            {
                var result = MessageBox.Show("Toks serijinis nr. jau nuskenuotas!", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (await IsSerial_serialNr(textBox1.Text.Trim()))
            {
                var result = MessageBox.Show("Toks serijinis nr. egzistuoja DB!", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return output;
        }

        //******************************************************* Data Grid Color *****************************************************
        private void DataGridShowColor(DataGridView _dg)
        {
            if (_dg.Rows.Count != 0)
            {
                for (int i = 0; i < _dg.Rows.Count; i++)
                {
                    if ((int.Parse(_dg.Rows[i].Cells[1].Value.ToString()) % 2) == 0)
                    {
                        _dg.Rows[i].DefaultCellStyle.BackColor = Color.Lavender;
                    }
                }
            }
        }

        //******************************************************* PRINT LABELS *****************************************************
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawString("K1020", new Font("Arial", 18, FontStyle.Bold), Brushes.Black, new Point(2, 10));
            e.Graphics.DrawString("" + textBox2.Text, new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(87, 18));
            e.Graphics.DrawString(textBox3.Text + "_" + textBox4.Text, new Font("Arial", 8, FontStyle.Regular), Brushes.Black, new Point(142, 48));
            e.Graphics.DrawString("*" + textBox1.Text.Trim() + "*", new Font("CCode39", 10, FontStyle.Regular), Brushes.Black, new Point(2, 42));
            e.Graphics.DrawString(textBox1.Text.Trim(), new Font("Arial", 8, FontStyle.Regular), Brushes.Black, new Point(29, 76));
            e.Graphics.DrawString("S/N: ", new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(2, 76));
            e.Graphics.DrawString("Box: ", new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(95, 76));
            e.Graphics.DrawString(boxCount + "/" + productBoxCount, new Font("Arial", 8, FontStyle.Regular), Brushes.Black, new Point(125, 76));
            e.Graphics.DrawString("Pallet: ", new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(157, 76));
            e.Graphics.DrawString(comboBox1.Text, new Font("Arial", 8, FontStyle.Regular), Brushes.Black, new Point(197, 76));
        }

        private void printDocument2_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawString("K1020", new Font("Arial", 18, FontStyle.Bold), Brushes.Black, new Point(2, 10));
            e.Graphics.DrawString("" + textBox2.Text, new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(87, 18));
            e.Graphics.DrawString(textBox3.Text + "_" + textBox4.Text, new Font("Arial", 8, FontStyle.Regular), Brushes.Black, new Point(142, 48));
            //e.Graphics.DrawString("SIM: " + textBox2.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, new Point(142, 48));
            //e.Graphics.DrawString(textBox3.Text + "_" + textBox4.Text, new Font("Arial", 12, FontStyle.Regular), Brushes.Black, new Point(87, 18));
            e.Graphics.DrawString("*" + serialTmpLabel2 + "*", new Font("CCode39", 10, FontStyle.Regular), Brushes.Black, new Point(2, 42));
            e.Graphics.DrawString(serialTmpLabel2, new Font("Arial", 8, FontStyle.Regular), Brushes.Black, new Point(29, 76));
            e.Graphics.DrawString("S/N: ", new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(2, 76));
            e.Graphics.DrawString("Box: ", new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(95, 76));
            e.Graphics.DrawString(boxCount + "/" + smallBoxTmpLabel2, new Font("Arial", 8, FontStyle.Regular), Brushes.Black, new Point(125, 76));
            e.Graphics.DrawString("Pallet: ", new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(157, 76));
            e.Graphics.DrawString(comboBox1.Text, new Font("Arial", 8, FontStyle.Regular), Brushes.Black, new Point(197, 76));
        }

        private void printDocument3_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawString(comboBox2.Text, new Font("Arial", 18, FontStyle.Bold), Brushes.Black, new Point(2, 10));
            e.Graphics.DrawString("*" + uniqueBoxNumber + "*", new Font("CCode39", 10, FontStyle.Regular), Brushes.Black, new Point(2, 42));
            e.Graphics.DrawString(uniqueBoxNumber, new Font("Arial", 8, FontStyle.Regular), Brushes.Black, new Point(80, 76));
            e.Graphics.DrawString("Box barcode: ", new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(2, 76));
            e.Graphics.DrawString("Units: ", new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(157, 76));
            e.Graphics.DrawString(comboBox3.Text, new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(190, 76));
        }

        private void button4_Click(object sender, EventArgs e)// Export to EXCEL
        {
            try
            {
                System.IO.Directory.CreateDirectory(@"c:/KGTmp/");
                string spreadsheetPath = "c:/KGTmp/K1020_" + comboBox1.Text + "_Packing_List.xlsx";
                File.Delete(spreadsheetPath);
                FileInfo spreadsheetInfo = new FileInfo(spreadsheetPath);

                ExcelPackage pck = new ExcelPackage(spreadsheetInfo);
                var Worksheet = pck.Workbook.Worksheets.Add("K1020 " + comboBox1.Text + " Packing List");

                Worksheet.PrinterSettings.PaperSize = ePaperSize.A4;
                Worksheet.PrinterSettings.Orientation = eOrientation.Portrait;

                int nr = 1;
                Worksheet.Cells[1, 1].Value = "Nr.";
                Worksheet.Cells[1, 2].Value = "Serial number";
                Worksheet.Cells[1, 3].Value = "MasterBox code";
                Worksheet.Cells[1, 4].Value = "MasterBox/SaleBox QTY";
                Worksheet.Cells[1, 5].Value = "Pallet";
                Worksheet.Cells[1, 6].Value = "Product Name";
                Worksheet.Cells[1, 7].Value = "SIM Card";
                Worksheet.Cells[1, 8].Value = "Scan Date/Time";

                Worksheet.Cells["A1:M1"].Style.Font.Bold = true;
                Worksheet.Cells["A1:M1"].Style.Font.Size = 12;
                Worksheet.Column(1).Width = 8;
                Worksheet.Column(2).Width = 20;
                Worksheet.Column(3).Width = 20;
                Worksheet.Column(4).Width = 30;
                Worksheet.Column(5).Width = 10;
                Worksheet.Column(6).Width = 20;
                Worksheet.Column(7).Width = 25;
                Worksheet.Column(8).Width = 20;

                Worksheet.Cells[1, 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                Worksheet.Cells[1, 2].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                Worksheet.Cells[1, 3].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                Worksheet.Cells[1, 4].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                Worksheet.Cells[1, 5].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                Worksheet.Cells[1, 6].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                Worksheet.Cells[1, 7].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                Worksheet.Cells[1, 8].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                Worksheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                Worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                Worksheet.Cells[1, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                Worksheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                Worksheet.Cells[1, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                Worksheet.Cells[1, 3].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                Worksheet.Cells[1, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                Worksheet.Cells[1, 4].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                Worksheet.Cells[1, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                Worksheet.Cells[1, 5].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                Worksheet.Cells[1, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                Worksheet.Cells[1, 6].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                Worksheet.Cells[1, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                Worksheet.Cells[1, 7].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                Worksheet.Cells[1, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                Worksheet.Cells[1, 8].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    Worksheet.Cells[i + 2, 1].Value = nr++; //Eilės nr
                    Worksheet.Cells[i + 2, 2].Value = dataGridView2.Rows[i].Cells[3].Value.ToString();//Serijinis nr
                    Worksheet.Cells[i + 2, 3].Value = dataGridView2.Rows[i].Cells[4].Value.ToString();// Box code
                    Worksheet.Cells[i + 2, 4].Value = dataGridView2.Rows[i].Cells[1].Value.ToString() + "/" + dataGridView2.Rows[i].Cells[2].Value.ToString();// Boxes
                    Worksheet.Cells[i + 2, 5].Value = dataGridView2.Rows[i].Cells[6].Value.ToString(); // Pallet
                    Worksheet.Cells[i + 2, 6].Value = dataGridView2.Rows[i].Cells[10].Value.ToString() + "_" + dataGridView2.Rows[i].Cells[11].Value.ToString();// Product
                    Worksheet.Cells[i + 2, 7].Value = dataGridView2.Rows[i].Cells[9].Value.ToString(); // Sim
                    Worksheet.Cells[i + 2, 8].Value = dataGridView2.Rows[i].Cells[13].Value.ToString(); // date time

                    Worksheet.Cells[i + 2, 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    Worksheet.Cells[i + 2, 2].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    Worksheet.Cells[i + 2, 3].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    Worksheet.Cells[i + 2, 4].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    Worksheet.Cells[i + 2, 5].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    Worksheet.Cells[i + 2, 6].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    Worksheet.Cells[i + 2, 7].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    Worksheet.Cells[i + 2, 8].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                }
                pck.Save();
                System.Diagnostics.Process.Start(spreadsheetPath);
            }
            catch (Exception)
            {
            }
        }

        //********************************************************************************************************** Info panel

        private Int32 currentId;
        private int editId = 0;
        private bool isEdit = false;

        private void dataGridView3_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (this.dataGridView3.Rows.Count != 0)
            {
                editId = dataGridView3.CurrentRow.Index;// editId
                currentId = (Int32)this.dataGridView3.CurrentRow.Cells[0].Value;// Id
                textBox6.Text = this.dataGridView3.CurrentRow.Cells[3].Value.ToString();
                textBox7.Text = this.dataGridView3.CurrentRow.Cells[4].Value.ToString();
                textBox8.Text = this.dataGridView3.CurrentRow.Cells[9].Value.ToString();
                textBox9.Text = this.dataGridView3.CurrentRow.Cells[10].Value.ToString();
                textBox10.Text = this.dataGridView3.CurrentRow.Cells[11].Value.ToString();
                textBox11.Text = this.dataGridView3.CurrentRow.Cells[1].Value.ToString();
                textBox12.Text = this.dataGridView3.CurrentRow.Cells[2].Value.ToString();
                textBox14.Text = this.dataGridView3.CurrentRow.Cells[6].Value.ToString();
                textBox13.Text = this.dataGridView3.CurrentRow.Cells[8].Value.ToString();

                isEdit = true;

                textBox6.Enabled = false;
                button5.Enabled = true;
                button6.Enabled = false;
                //button7.Enabled = true;
                button11.Enabled = true;
                button10.Enabled = true;
                button8.Enabled = true;
            }
        }

        private void tabControl1_Click(object sender, EventArgs e)
        {
            textBox6.Enabled = false;
            textBox7.Enabled = false;
            button5.Enabled = true;
            button6.Enabled = false;
            button7.Enabled = false;
            button11.Enabled = false;
            button10.Enabled = false;
            button9.Enabled = false;
            button8.Enabled = false;
            button5.Enabled = false;
        }

        //**************************************** info Button function ***********************************
        private void button5_Click(object sender, EventArgs e)// serial Edit
        {
            if (textBox6.Text.Length != 0)
            {
                textBox6.Enabled = true;
                button5.Enabled = false;
                button6.Enabled = true;
            }
        }

        private async void button6_Click(object sender, EventArgs e) // Save
        {
            await EditSerial();
            textBox6.Enabled = false;
            button5.Enabled = true;
            button6.Enabled = false;
            isEdit = true;
        }

        private void button8_Click(object sender, EventArgs e)// box edit
        {
            if (textBox7.Text.Length != 0)
            {
                //textBox7.Enabled = true;
                button8.Enabled = false;
                button9.Enabled = true;
            }
        }
        private async void button9_Click(object sender, EventArgs e)// delete button
        {
            await DeleteBoxSerial();
            //textBox9.Enabled = false;
            button8.Enabled = true;
            button9.Enabled = false;
            isEdit = true;
        }
        //*****************************************************************************************************
        private async void Edit_serial(int _id)
        {
            if (await ValidateFormSerial())
            {
                string sql = "UPDATE serials SET SerialNr = @SerialNr, ModDateTime = @ModDateTime where Id = @Id";
                await DataAccess.SaveData(sql, new
                {
                    SerialNr = textBox6.Text.Trim(),
                    ModDateTime = DateTime.Now,
                    id = _id
                });
            }
        }

        private async Task EditSerial()
        {
            if (isEdit)
            {
                currentId = (Int32)this.dataGridView3.CurrentRow.Cells[0].Value;

                const string message = "Išsaugoti pakeitimus?";
                const string caption = "Įrašo koregavimas";

                var result = MessageBox.Show(message, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.No)
                {
                }
                else
                {
                    try
                    {
                        Edit_serial(currentId);
                        await ShowAllSerialNboxInfo();
                        CleanInfoTabPanel();

                        //CleanFormAll();
                        dataGridView3.ClearSelection();
                        try
                        {
                            if (dataGridView3.Rows.Count != 0)
                            {
                                if (dataGridView3.Rows[editId] != null)
                                {
                                    dataGridView3.Rows[editId].Selected = true;
                                    dataGridView3.Rows[editId].Cells[1].Selected = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }
        //---------------------------------------- Delete box -------------------------------------------------------------
        public async Task DeleteBox(string _uniqNr)
        {
            string sql = "DELETE FROM serials WHERE UniqNr = @UniqNr";

            await DataAccess.SaveData(sql, new { UniqNr = _uniqNr });
        }
        private async Task DeleteBoxSerial()
        {
            if (isEdit)
            {
                currentId = (Int32)this.dataGridView3.CurrentRow.Cells[0].Value;

                const string message = "Ar tikrai norite ištrinti serijinius numerius?";
                const string caption = "Įrašų šalinimas";

                var result = MessageBox.Show(message, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.No)
                {
                }
                else
                {
                    try
                    {
                        await DeleteBox(textBox7.Text);
                        await ShowAllSerialNboxInfo();
                        CleanInfoTabPanel();

                        dataGridView3.ClearSelection();
                        try
                        {
                            if (dataGridView3.Rows.Count != 0)
                            {
                                if (dataGridView3.Rows[editId] != null)
                                {
                                    dataGridView3.Rows[editId].Selected = true;
                                    dataGridView3.Rows[editId].Cells[1].Selected = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }


        //*********************************************************************************************************************
        private void CleanInfoTabPanel()
        {
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
            textBox13.Text = "";
            textBox14.Text = "";
        }

        private async Task<bool> ValidateFormSerial()
        {
            bool output = true;

            if (String.IsNullOrEmpty(textBox6.Text.Trim()))
            {
                var result = MessageBox.Show("Neįvestas serijinis kodas", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (textBox6.Text.Count() != 8)
            {
                var result = MessageBox.Show("Blogas serijinis numeris", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (textBox6.Text.Substring(0, 3) != "313")
            {
                var result = MessageBox.Show("Blogas serijinis numeris", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (await IsSerial_serialNr(textBox6.Text.Trim()))
            {
                var result = MessageBox.Show("Toks serijinis nr. egzistuoja DB!", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return output;
        }

        private void printDocument4_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawString("K1020", new Font("Arial", 18, FontStyle.Bold), Brushes.Black, new Point(2, 10));
            e.Graphics.DrawString("" + textBox8.Text, new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(87, 18));// sim card
            e.Graphics.DrawString(textBox9.Text + "_" + textBox10.Text, new Font("Arial", 8, FontStyle.Regular), Brushes.Black, new Point(142, 48)); // Gaminys
            //e.Graphics.DrawString("" + textBox8.Text, new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(142, 48));// sim card
            //e.Graphics.DrawString(textBox9.Text + "_" + textBox10.Text, new Font("Arial", 12, FontStyle.Regular), Brushes.Black, new Point(87, 18)); // Gaminys
            e.Graphics.DrawString("*" + textBox6.Text.Trim() + "*", new Font("CCode39", 10, FontStyle.Regular), Brushes.Black, new Point(2, 42));
            e.Graphics.DrawString(textBox6.Text.Trim(), new Font("Arial", 8, FontStyle.Regular), Brushes.Black, new Point(29, 76));
            e.Graphics.DrawString("S/N: ", new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(2, 76));
            e.Graphics.DrawString("Box: ", new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(95, 76));
            e.Graphics.DrawString(textBox11.Text + "/" + textBox12.Text, new Font("Arial", 8, FontStyle.Regular), Brushes.Black, new Point(125, 76));
            e.Graphics.DrawString("Pallet: ", new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(157, 76));
            e.Graphics.DrawString(textBox14.Text, new Font("Arial", 8, FontStyle.Regular), Brushes.Black, new Point(197, 76));
        }

        private void button11_Click(object sender, EventArgs e)// Print label serial
        {
            const string message = "Spausdinti lipduką ant dėžutės?";
            const string caption = "Lipduko ant dėžutės spausdinimas";

            var result = MessageBox.Show(message, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.No)
            {
            }
            else
            {

                if (printDialog)
                {
                    printPreviewDialog1.Document = printDocument4;
                    printPreviewDialog1.ShowDialog();
                }
                else
                {
                    printDocument4.Print();
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)// Print labels serial 10 vnt
        {
        }

        private void button10_Click(object sender, EventArgs e)// print box
        {
            const string message = "Spausdinti lipduką ant dėžės?";
            const string caption = "Lipduko ant dėžės spausdinimas";

            var result = MessageBox.Show(message, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.No)
            {
            }
            else
            {
                if (printDialog)
                {
                    printPreviewDialog1.Document = printDocument5;
                    printPreviewDialog1.ShowDialog();
                }
                else
                {
                    printDocument5.Print();
                }
            }
        }

        private void printDocument5_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)// Print box
        {
            e.Graphics.DrawString(textBox13.Text, new Font("Arial", 18, FontStyle.Bold), Brushes.Black, new Point(2, 10));
            e.Graphics.DrawString("*" + textBox7.Text + "*", new Font("CCode39", 10, FontStyle.Regular), Brushes.Black, new Point(2, 42));
            e.Graphics.DrawString(textBox7.Text, new Font("Arial", 8, FontStyle.Regular), Brushes.Black, new Point(80, 76));
            e.Graphics.DrawString("Box barcode: ", new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(2, 76));
            e.Graphics.DrawString("Units: ", new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(157, 76));
            e.Graphics.DrawString(comboBox3.Text, new Font("Arial", 8, FontStyle.Bold), Brushes.Black, new Point(190, 76));
        }

        private async void textBox15_TextChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text.Length != 0) await ShowAllSerialNboxInfo();
        }

        private async void textBox16_TextChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text.Length != 0) await ShowAllSerialNboxInfo();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked) comboBox2.Enabled = true;
            else comboBox2.Enabled = false;
        }
    }
}