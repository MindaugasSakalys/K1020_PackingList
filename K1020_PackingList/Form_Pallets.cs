using K1020_PackingList.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DataLibrary;
using FluentFTP;
using System.IO;
using System.Threading;
using System.Diagnostics;
using System.Reflection;

namespace K1020_PackingList
{
    public partial class Form_Pallets : Form
    {
        Int32 currentId;
        int editId = 0;
        bool isEdit = false;

        public Form_Pallets()
        {
            InitializeComponent();
        }
        //******************************************************* Start Form ***************************************************
        private async void Form_Pallets_Load(object sender, EventArgs e)
        {
            dataGridView1.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(dataGridView1, true, null);
            CleanFormAll();
            await ShowPalletListOnGrid();
            dataGridView1.ClearSelection();
        }
        //******************************************************* Button Function **********************************************
        private async void button1_Click(object sender, EventArgs e) //Add Button
        {
            if (!isEdit)
            {
                await AddPallet();
                dataGridView1.ClearSelection();
            }
            else
            {
                await EditPallet();
                dataGridView1.ClearSelection();
            }
        }

        private void button2_Click(object sender, EventArgs e) //Undo Button
        {
            CleanFormAll();
            dataGridView1.ClearSelection();
        }

        private async void button3_Click(object sender, EventArgs e) //Delete Button
        {
            await Delete_Pallet();
            dataGridView1.ClearSelection();
        }
        //******************************************************* MySQL Function **********************************************

        private async Task<List<Pallets_model>> LoadPallets()
        {
            List<Pallets_model> output = new List<Pallets_model>();
            string sql = "(SELECT Id, PalletName, DonePallet, AddDate, PhotoPath, 'True' AS Foto " +
                "FROM pallet " +
                "WHERE PhotoPath LIKE '/%') " +
                "UNION " +
                "(SELECT Id, PalletName, DonePallet, AddDate, PhotoPath, 'False' AS Foto " +
                "FROM pallet " +
                "WHERE PhotoPath = '') " +
                "ORDER BY PalletName";
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
        private async Task<bool> IsSerial_palletId(string _palletName)
        {
            bool output = false;
            List<Serials_model> serialList = new List<Serials_model>();
            string sql = "SELECT * FROM serials " +
                "LEFT JOIN pallet ON serials.PaletId = pallet.Id " +
                "LEFT JOIN code ON serials.CodeId = code.Id " +
                "WHERE pallet.PalletName = @PaletName";
            serialList = await DataAccess.LoadData<Serials_model, dynamic>(sql, new { PaletName = @_palletName });
            if (serialList.Count > 0) output = true;
            return output;
        }
        //******************************************************* CRUD Function ************************************************
        private async Task AddPallet()
        {
            if (await ValidateForm())
            {
                string sql = "INSERT INTO pallet (PalletName, DonePallet, AddDate, PhotoPath) VALUE (@PalletName, @DonePallet, @AddDate, @PhotoPath);";


                await DataAccess.SaveData(sql, new
                {
                    PalletName = textBox1.Text.Trim(),
                    DonePallet = checkBox1.Checked,
                    AddDate = DateTime.Now,
                    PhotoPath = ""
                });

                await ShowPalletListOnGrid();
                CleanFormAll();
            }
        }
        //------------------------------------------------- Edit -----------------------------------------------------------------------
        private async void Edit_Pallet(int _id)
        {
            if (await ValidateForm())
            {
                string sql = "UPDATE pallet SET DonePallet = @DonePallet where Id = @Id";
                await DataAccess.SaveData(sql, new
                {
                    DonePallet = checkBox1.Checked,
                    id = _id
                });
            }
        }
        private async Task EditPallet()
        {
            if (isEdit)
            {
                currentId = (Int32)this.dataGridView1.CurrentRow.Cells[0].Value;

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
                        Edit_Pallet(currentId);
                        await ShowPalletListOnGrid();
                        CleanFormAll();
                        dataGridView1.ClearSelection();
                        try
                        {
                            if (dataGridView1.Rows.Count != 0)
                            {
                                if (dataGridView1.Rows[editId] != null)
                                {
                                    dataGridView1.Rows[editId].Selected = true;
                                    dataGridView1.Rows[editId].Cells[1].Selected = true;
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
        //------------------------------------------------- DELETE-----------------------------------------------------------------------
        public async Task DeletePallet(int _id)
        {
            string sql = "DELETE FROM pallet WHERE Id = @Id";

            await DataAccess.SaveData(sql, new { Id = _id });
        }
        //--------------------------------------------------------------
        private async Task Delete_Pallet()
        {
            if (isEdit)
            {
                if (await ValidateFormDelete())
                {
                    currentId = (Int32)this.dataGridView1.CurrentRow.Cells[0].Value;

                    const string message = "Ar tikrai norite pašalinti pasirinktą įrašą?";
                    const string caption = "Įrašo šalinimas";

                    var result = MessageBox.Show(message, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.No)
                    {

                    }
                    else
                    {
                        try
                        {
                            await DeletePallet(currentId);


                            var _token = new CancellationToken();

                            using (FtpClient ftp = DataAccess.CreateFtpClient())
                            {
                                await ftp.ConnectAsync(_token);

                                if (await ftp.FileExistsAsync(ftp_pathName + palletName + ".jpg", _token))
                                {
                                    await ftp.DeleteFileAsync(ftp_pathName + palletName + ".jpg", _token);
                                }
                                CleanImg(pictureBox1);
                            }

                            await ShowPalletListOnGrid();
                            CleanFormAll();
                            dataGridView1.ClearSelection();
                            try
                            {
                                if (dataGridView1.Rows.Count != 0)
                                {
                                    if (dataGridView1.Rows[editId - 1] != null)
                                    {
                                        dataGridView1.Rows[editId - 1].Selected = true;
                                        dataGridView1.Rows[editId - 1].Cells[1].Selected = true;
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

        }

        //******************************************************* Validate Function *******************************************************************
        private async Task<bool> ValidateFormDelete()
        {
            bool output = true;

            if (await IsSerial_palletId(textBox1.Text))
            {
                var result = MessageBox.Show("Šio įrašo pašalinti negalima", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return output;
        }
        private async Task<bool> ValidateForm()
        {
            bool output = true;

            if (!isEdit)
            {
                if (await LoadPallets(textBox1.Text.Trim()))
                {
                    var result = MessageBox.Show("Toks paletės pavadinimas jau egzistuoja!", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            if (String.IsNullOrEmpty(textBox1.Text))
            {
                var result = MessageBox.Show("Neįvestas paletės pavadinimas", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return output;
        }

        //******************************************************* DB Functions **********************************************************************

        private async Task ShowPalletListOnGrid()
        {
            dataGridView1.DataSource = null;
            List<Pallets_model> output = new List<Pallets_model>();
            output = await LoadPallets();
            dataGridView1.DataSource = output;
            dataGridView1.Columns[2].DefaultCellStyle.Format = "yyyy-MM-dd";
            label4.Text = output.Count.ToString();

            for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
            {
                if (Convert.ToBoolean(this.dataGridView1.Rows[i].Cells[3].Value) == true)
                {
                    this.dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                }
                if (Convert.ToBoolean(this.dataGridView1.Rows[i].Cells[5].Value) == true)
                {
                    this.dataGridView1.Rows[i].Cells[5].Style.BackColor = Color.LightBlue;
                }
            }
            if (dataGridView1.Rows.Count != 0)
            {
                dataGridView1.CurrentCell = dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[1];
            }
        }
        //******************************************************* Clear form Functions **************************************************************
        private void CleanFormAll()
        {

            textBox1.Text = "";
            checkBox1.Checked = false;
            textBox1.Focus();
            textBox1.Enabled = true;
            button3.Enabled = false;
            isEdit = false;
            button4.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;
            pictureBox1.Image = new Bitmap(K1020_PackingList.Properties.Resources.PHimg);

        }
        //******************************************************* Data Grids *************************************************************************
        private async void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            CleanFormAll();

            if (this.dataGridView1.Rows.Count != 0)
            {

                editId = dataGridView1.CurrentRow.Index;// editId
                currentId = (Int32)this.dataGridView1.CurrentRow.Cells[0].Value;// Id
                textBox1.Text = this.dataGridView1.CurrentRow.Cells[1].Value.ToString();
                checkBox1.Checked = bool.Parse(this.dataGridView1.CurrentRow.Cells[3].Value.ToString());

                ftp_pathNameDel = "";
                CleanImg(pictureBox1);
                palletName = this.dataGridView1.CurrentRow.Cells[1].Value.ToString();
                ftp_pathNameDel = ftp_pathName;

                string _pathTmp = "";

                if (this.dataGridView1.CurrentRow.Cells[4].Value.ToString() == "")
                {
                    pictureBox1.Image = new Bitmap(K1020_PackingList.Properties.Resources.PHimg);
                    button4.Enabled = true;
                    button5.Enabled = false;
                    button6.Enabled = false;
                }
                else
                {
                    _pathTmp = this.dataGridView1.CurrentRow.Cells[4].Value.ToString();
                    await ShowImgAsync(pictureBox1, _pathTmp);
                    button4.Enabled = true;
                    button5.Enabled = true;
                    button6.Enabled = true;
                }

                isEdit = true;
                textBox1.Enabled = false;
                button3.Enabled = true;
            }
        }

        //*******************************************************FOTO Functions *********************************************************************

        private string palletName = "";
        private string ftp_pathName = "/files/TeltonikaDB/K1020_Packing_List/";
        private string ftp_pathNameDel = "";
        private async void button4_Click(object sender, EventArgs e) //Upload foto
        {
            await UploadImg(pictureBox1);
        }
        private async void button5_Click(object sender, EventArgs e)// Download
        {
            var _token = new CancellationToken();
            try
            {
                using (FtpClient conn = DataAccess.CreateFtpClient())
                {
                    await conn.ConnectAsync(_token);

                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "JPEG Image|*.jpg";
                    save.Title = "Save an Image File";
                    save.FileName = palletName;
                    save.ShowDialog();
                    await conn.DownloadFileAsync(Path.GetFullPath(save.FileName).ToString(), ftp_pathName + palletName + ".jpg", FtpLocalExists.Overwrite, FtpVerify.Retry, null, _token);
                }
            }
            catch (OperationCanceledException)
            {

            }
            finally
            {
            }
        }
        private async void button6_Click(object sender, EventArgs e)// Delete
        {
            currentId = (Int32)this.dataGridView1.CurrentRow.Cells[0].Value;

            const string message = "Ar tikrai norite ištrinti paletės nuotrauką?";
            const string caption = "Paletės nuotraukos šalinimas";

            var result = MessageBox.Show(message, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.No)
            {

            }
            else
            {
                if (ValidateFormPhoto())
                {
                    await DeleteImg(pictureBox1, currentId);
                }
            }
        }

        private async Task DeleteImg(PictureBox _pic, int _id)
        {
            var _token = new CancellationToken();

            using (FtpClient ftp = DataAccess.CreateFtpClient())
            {
                await ftp.ConnectAsync(_token);

                if (await ftp.FileExistsAsync(ftp_pathName + palletName + ".jpg", _token))
                {
                    await ftp.DeleteFileAsync(ftp_pathName + palletName + ".jpg", _token);

                    string sql = "UPDATE pallet SET PhotoPath = @PhotoPath WHERE Id = @Id";
                    await DataAccess.SaveData(sql, new
                    {
                        PhotoPath = "",
                        Id = _id
                    });

                    await ShowPalletListOnGrid();
                    dataGridView1.ClearSelection();
                    try
                    {
                        if (dataGridView1.Rows.Count != 0)
                        {
                            if (dataGridView1.Rows[editId] != null)
                            {
                                dataGridView1.Rows[editId].Selected = true;
                                dataGridView1.Rows[editId].Cells[1].Selected = true;
                            }
                        }
                    }
                    catch (Exception)
                    {

                    }
                }
                CleanImg(_pic);
            }
        }

        private async Task ShowImgAsync(PictureBox _pic, string _path)
        {
            var _token = new CancellationToken();
            try
            {
                using (FtpClient conn = DataAccess.CreateFtpClient())
                {
                    await conn.ConnectAsync(_token);

                    if (await conn.FileExistsAsync(_path))
                    {
                        using (var istream = await conn.OpenReadAsync(_path, _token))
                        {
                            try
                            {
                                _pic.Image = new Bitmap(istream);
                                _pic.Enabled = true;
                            }
                            finally
                            {
                                Console.WriteLine();
                                istream.Close();
                            }
                        }
                    }
                    else
                    {
                        var bmp = new Bitmap(K1020_PackingList.Properties.Resources.PHimg);
                        _pic.Image = bmp;
                        _pic.Enabled = false;
                    }
                }
            }
            catch (OperationCanceledException)
            {

            }
            finally
            {
                //cts.Dispose();
            }

        }
        private async Task UploadImg(PictureBox _pic)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Title = "Pasirinkite paletės nuotrauką";
            openFile.InitialDirectory = "";
            openFile.FileName = "";
            openFile.Filter = "Image Files(*.jpg;)|*.jpg;";
            openFile.RestoreDirectory = true;

            Stream _stream;
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFile.FileName;

                try
                {
                    using (FtpClient ftp = DataAccess.CreateFtpClient())
                    {

                        await ftp.CreateDirectoryAsync(ftp_pathName);
                        await ftp.UploadFileAsync(filePath, ftp_pathName + palletName + ".jpg");


                        if (await ftp.FileExistsAsync(ftp_pathName + palletName + ".jpg"))
                        {
                            _stream = await ftp.OpenReadAsync(ftp_pathName + palletName + ".jpg");
                            _pic.Image = new Bitmap(_stream);
                            _pic.Enabled = true;

                            string tempPath = ftp_pathName + palletName + ".jpg";
                            await UpdatePhoto(currentId.ToString(), tempPath);

                            await ShowPalletListOnGrid();
                            dataGridView1.ClearSelection();
                            try
                            {
                                if (dataGridView1.Rows.Count != 0)
                                {
                                    if (dataGridView1.Rows[editId] != null)
                                    {
                                        dataGridView1.Rows[editId].Selected = true;
                                        dataGridView1.Rows[editId].Cells[1].Selected = true;
                                    }
                                }
                                var _pathTmp = this.dataGridView1.Rows[editId].Cells[4].Value.ToString();
                                await ShowImgAsync(pictureBox1, _pathTmp);
                            }
                            catch (Exception)
                            {

                            }
                        }
                        else
                        {
                            var bmp = new Bitmap(K1020_PackingList.Properties.Resources.PHimg);
                            _pic.Image = bmp;
                            _pic.Enabled = false;
                        }

                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Iškilo problema įkeliant nuotrauką");
                }
            }
        }

        private async Task UpdatePhoto(string _id, string _path)
        {
            if (ValidateFormPhoto())
            {
                string sql = "UPDATE pallet SET PhotoPath = @PhotoPath WHERE Id = @Id";
                await DataAccess.SaveData(sql, new
                {
                    PhotoPath = _path,
                    Id = _id
                });
            }
        }

        private bool ValidateFormPhoto()
        {
            bool output = true;

            if (palletName == "")
            {
                return false;
            }
            return output;
        }

        private void CleanImg(PictureBox _pic)
        {
            var bmp = new Bitmap(K1020_PackingList.Properties.Resources.PHimg);
            _pic.Image = bmp;
            _pic.Enabled = false;
        }

        private async void pictureBox1_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                var _token = new CancellationToken();
                using (FtpClient conn = DataAccess.CreateFtpClient())
                {
                    await conn.ConnectAsync(_token);
                    await conn.DownloadFileAsync(@".\Temp\temp.jpg", ftp_pathName + palletName + ".jpg", FtpLocalExists.Overwrite, FtpVerify.Retry, null, _token);
                    Process.Start(@".\Temp\temp.jpg");
                }
            }
            catch (Exception)
            {
            }
        }
    }

}
