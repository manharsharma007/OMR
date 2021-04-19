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
    public partial class Add_Subject : Form
    {
        DataTable dt;
        data_class data;
        public Add_Subject()
        {
            InitializeComponent();
        }

        private void Add_Subject_Load(object sender, EventArgs e)
        {
            data = new data_class();
            data.open_connection();
            dt = data.select_subjects();
            sub_grid_view.DataSource = dt;
        }

        private void add_btn_Click(object sender, EventArgs e)
        {
            try
            {
                if (sub_txt.Text == "")
                    throw new Exception("Subject Name is empty!");

                data.insert_subject(sub_txt.Text);
                dt = data.select_subjects();
                sub_grid_view.DataSource = dt;
                sub_grid_view.Refresh();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void delete_btn_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in sub_grid_view.SelectedRows)
            {
                //get key
                int rowId = Convert.ToInt32(row.Cells[0].Value);

                //avoid updating the last empty row in datagrid
                if (rowId > 0)
                {
                    //delete 
                    data.delete_subject(rowId);

                    //refresh datagrid
                    sub_grid_view.Rows.RemoveAt(row.Index);
                }
            }
        }

    }
}
