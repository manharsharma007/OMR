using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace omrapplication
{
    public partial class Delete_Batch : Form
    {
        data_class data;
        public Delete_Batch()
        {
            InitializeComponent();
            data = new data_class();
            data.open_connection();
        }

        private void Delete_Batch_Load(object sender, EventArgs e)
        {
            
            DataTable dt = data.select_batch();
            dataGridView1.DataSource = dt;
        }

        private void delete_btn_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                //get key
                int rowId = Convert.ToInt32(row.Cells[0].Value);

                //avoid updating the last empty row in datagrid
                if (rowId > 0)
                {
                    //delete 
                    data.delete_batch(rowId);

                    //refresh datagrid
                    dataGridView1.Rows.RemoveAt(row.Index);
                }
            }
        }

    }
}
