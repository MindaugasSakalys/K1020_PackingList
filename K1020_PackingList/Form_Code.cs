using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DataLibrary;
using K1020_PackingList.Models;

namespace K1020_PackingList
{
    public partial class Form_Code : Form
    {
        Int32 currentId;
        int editId = 0;
        bool isEdit = false;
        public Form_Code()
        {
            InitializeComponent();
        }
        //******************************************************* Start Form *************************************************************************
        private async void Form_Code_Load(object sender, EventArgs e)
        {
            dataGridView1.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(dataGridView1, true, null);
            await ShowCodeListOnGrid();
            CleanFormAll();
            dataGridView1.ClearSelection();
        }
        //******************************************************* Button Function ********************************************************************
        private async void button1_Click(object sender, EventArgs e)//Add Button
        {
            if (!isEdit)
            {
                await AddCode();
                dataGridView1.ClearSelection();
            }
            else
            {
                await EditCode();
                dataGridView1.ClearSelection();
            }
        }

        private void button2_Click(object sender, EventArgs e)//Undo Button
        {
            CleanFormAll();
            dataGridView1.ClearSelection();
        }

        private async void button3_Click(object sender, EventArgs e)//Delete Button
        {
            await Delete_Code();
            dataGridView1.ClearSelection();
        }
        //******************************************************* MySQL Function **********************************************************************
        private async Task<List<Code_model>> LoadCodes()
        {
            List<Code_model> output = new List<Code_model>();
            string sql = "SELECT * FROM code ORDER BY CodeName";
            output = await DataAccess.LoadData<Code_model, dynamic>(sql, new { });
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

        private async Task<bool> IsSerial_CodeId(string _codeName)
        {
            bool output = false;
            List<Serials_model> serialList = new List<Serials_model>();
            string sql = "SELECT * FROM serials " +
                "LEFT JOIN pallet ON serials.PaletId = pallet.Id " +
                "LEFT JOIN code ON serials.CodeId = code.Id " +
                "WHERE code.CodeName = @CodeName";
            serialList = await DataAccess.LoadData<Serials_model, dynamic>(sql, new { CodeName = @_codeName });
            if (serialList.Count > 0) output = true;
            return output;
        }

        //******************************************************* CRUD Function ***********************************************************************

        private async Task AddCode()
        {
            if (await ValidateForm())
            {
                string sql = "INSERT INTO code (CodeName, CountryCode, CountryName, Version, BatteryCount, Disabled) VALUE (@CodeName, @CountryCode, @CountryName, @Version, @BatteryCount, @Disabled);";


                await DataAccess.SaveData(sql, new
                {
                    CodeName = textBox1.Text.Trim(),
                    CountryCode = comboBox1.Text,
                    CountryName = comboBox2.Text,
                    Version = comboBox4.Text,
                    BatteryCount = comboBox3.Text,
                    Disabled = checkBox1.Checked,
                });

                await ShowCodeListOnGrid();
                CleanFormAll();
            }
        }
        //------------------------------------------------- Edit -----------------------------------------------------------------------
        private async void Edit_Code(int _id)
        {
            if (await ValidateForm())
            {
                string sql = "UPDATE code SET CountryCode = @CountryCode, CountryName = @CountryName, Version = @Version, BatteryCount = @BatteryCount, Disabled = @Disabled WHERE Id = @Id";
                await DataAccess.SaveData(sql, new
                {
                    CountryCode = comboBox1.Text,
                    CountryName = comboBox2.Text,
                    Version = comboBox4.Text,
                    BatteryCount = comboBox3.Text,
                    Disabled = checkBox1.Checked,
                    id = _id
                });
            }
        }
        private async Task EditCode()
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
                        Edit_Code(currentId);
                        await ShowCodeListOnGrid();
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
        public async Task DeleteCode(int _id)
        {
            string sql = "DELETE FROM code WHERE Id = @Id";

            await DataAccess.SaveData(sql, new { Id = _id });
        }
        //--------------------------------------------------------------
        private async Task Delete_Code()
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
                            await DeleteCode(currentId);

                            await ShowCodeListOnGrid();
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

        private async Task <bool> ValidateFormDelete()
        {
            bool output = true;

            if (await IsSerial_CodeId(textBox1.Text))
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
                if (await LoadCodes(textBox1.Text.Trim()))
                {
                    var result = MessageBox.Show("Toks kodo pavadinimas jau egzistuoja!", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            if (String.IsNullOrEmpty(textBox1.Text))
            {
                var result = MessageBox.Show("Neįvestas kodo pavadinimas", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            //if (String.IsNullOrEmpty(comboBox1.Text))
            //{
            //    var result = MessageBox.Show("Laukas 'SIM kortelė' negali būti tuščias", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return false;
            //}
            //if (String.IsNullOrEmpty(comboBox2.Text))
            //{
            //    var result = MessageBox.Show("Laukas 'Baterijų kiekis' negali būti tuščias", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return false;
            //}
            //if (String.IsNullOrEmpty(comboBox3.Text))
            //{
            //    var result = MessageBox.Show("Laukas 'Šalis' negali būti tuščias", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return false;
            //}
            if (textBox1.Text.Count() != 12)
            {
                var result = MessageBox.Show("Blogas kodo pavadinimas", "Klaida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }


            return output;
        }
        //******************************************************* DB Functions **********************************************************************

        private async Task ShowCodeListOnGrid()
        {
            dataGridView1.DataSource = null;
            List<Code_model> output = new List<Code_model>();
            output = await LoadCodes();
            dataGridView1.DataSource = output;
            label4.Text = output.Count.ToString();

            for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
            {
                if (Convert.ToBoolean(this.dataGridView1.Rows[i].Cells[6].Value) == true)
                {
                    this.dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGray;
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
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;
            textBox1.Enabled = true;
            checkBox1.Checked = false;
            button3.Enabled = false;
            isEdit = false;
        }
        //******************************************************* Data Grids *************************************************************************
        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            CleanFormAll();

            if (this.dataGridView1.Rows.Count != 0)
            {

                editId = dataGridView1.CurrentRow.Index;// editId
                currentId = (Int32)this.dataGridView1.CurrentRow.Cells[0].Value;// Id
                textBox1.Text = this.dataGridView1.CurrentRow.Cells[1].Value.ToString();
                comboBox1.Text = this.dataGridView1.CurrentRow.Cells[2].Value.ToString();
                comboBox2.Text = this.dataGridView1.CurrentRow.Cells[3].Value.ToString();
                comboBox3.Text = this.dataGridView1.CurrentRow.Cells[5].Value.ToString();
                comboBox4.Text = this.dataGridView1.CurrentRow.Cells[4].Value.ToString();
                checkBox1.Checked = bool.Parse(this.dataGridView1.CurrentRow.Cells[6].Value.ToString());

                isEdit = true;
                textBox1.Enabled = false;
                button3.Enabled = true;
            }
        }
    }
}
