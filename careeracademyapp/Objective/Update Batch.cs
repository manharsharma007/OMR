using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace omrapplication
{
    public partial class Update_Batch : Form
    {
        data_class data;
        DataGridView new_temp;
        DataSet ds;
        SqlDataAdapter da;

        public Update_Batch()
        {
            InitializeComponent();
            data = new data_class();
            data.open_connection();
        }

        private void Update_Batch_Load(object sender, EventArgs e)
        {
            DataTable ds = data.select_batch();
            batch_combo_box.ValueMember = "batch_id";
            batch_combo_box.DisplayMember = "batch_name";
            batch_combo_box.DataSource = ds;

        }


        private void UpdateAnswers()
        {
            try
            {
                SqlCommandBuilder cmd = new SqlCommandBuilder(da);
                da.Update(ds, "options");
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message.ToString());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                if (Convert.ToInt32(batch_combo_box.SelectedValue) <= 0 || da == null)
                    throw new Exception("Batch is not selected");
                
                UpdateAnswers();



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error",
MessageBoxButtons.OK,
MessageBoxIcon.Error,
MessageBoxDefaultButton.Button1);
            }
        }

        private void add_questions_Click(object sender, EventArgs e)
        {
            try
            {

                if (Convert.ToInt32(batch_combo_box.SelectedValue) <= 0)
                    throw new Exception("Batch is not selected");
                

                ds = data.select_batch_options(Convert.ToInt32(batch_combo_box.SelectedValue));
                da = data.da;

                dataGridView1.DataSource = ds.Tables[0];

                dataGridView1.Columns[0].HeaderText = "Unique ID";
                dataGridView1.Columns[1].HeaderText = "Question";
                dataGridView1.Columns[2].HeaderText = "Answer";
                dataGridView1.Columns[3].HeaderText = "Marks";
                dataGridView1.Columns[4].HeaderText = "Negative";
                dataGridView1.Columns[5].HeaderText = "Topic";
                dataGridView1.Columns[6].HeaderText = "Sub Topic";
                dataGridView1.Refresh();

                nothing_lbl.Visible = false;

               


             
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error",
MessageBoxButtons.OK,
MessageBoxIcon.Error,
MessageBoxDefaultButton.Button1);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        { 
            saveFileDialog1 = new SaveFileDialog();           
            saveFileDialog1.Filter = "Excel Documents (*.xlsx)|*.xls; *xlsx";

            Cursor = Cursors.WaitCursor;
            try
            {


                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {


                    new_temp = CopyDataGridView(dataGridView1);


                    foreach (DataGridViewRow row in new_temp.Rows)
                    {
                        //get key
                        int topicId = Convert.ToInt32(row.Cells[5].Value);
                        int subTopicId = Convert.ToInt32(row.Cells[6].Value);
                        int number = Convert.ToInt32(row.Cells[1].Value);

                        //avoid updating the last empty row in datagrid

                        row.Cells[5].Value = data.get_topic(topicId, number).ToString();
                        row.Cells[6].Value = data.get_sub_topic(subTopicId, number).ToString();
                    }
                    string name = saveFileDialog1.FileName;
                    data.ExportToExcel(new_temp, name);
                }
            }
            catch(Exception ex)
            {

                MessageBox.Show(ex.Message, "Error",
MessageBoxButtons.OK,
MessageBoxIcon.Error,
MessageBoxDefaultButton.Button1);
            }
            
            Cursor = Cursors.Arrow;
        }
        private DataGridView CopyDataGridView(DataGridView dgv_org)
        {
            DataGridView dgv_copy = new DataGridView();
            try
            {
                if (dgv_copy.Columns.Count == 0)
                {
                    foreach (DataGridViewColumn dgvc in dgv_org.Columns)
                    {
                        dgv_copy.Columns.Add(dgvc.Clone() as DataGridViewColumn);
                    }
                }

                DataGridViewRow row = new DataGridViewRow();

                for (int i = 0; i < dgv_org.Rows.Count; i++)
                {
                    row = (DataGridViewRow)dgv_org.Rows[i].Clone();
                    int intColIndex = 0;
                    foreach (DataGridViewCell cell in dgv_org.Rows[i].Cells)
                    {
                        row.Cells[intColIndex].Value = cell.Value;
                        intColIndex++;
                    }
                    dgv_copy.Rows.Add(row);
                }
                dgv_copy.AllowUserToAddRows = false;
                dgv_copy.Refresh();

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return dgv_copy;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                try
                {
                    int marks = 4;
                    int neg = -1;

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        row.Cells[3].Value = marks;
                        row.Cells[4].Value = neg;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error",
    MessageBoxButtons.OK,
    MessageBoxIcon.Error,
    MessageBoxDefaultButton.Button1);
                }
            }
        }
    }
}
