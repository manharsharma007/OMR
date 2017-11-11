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
    public partial class Add_Subtopic : Form
    {

        DataTable dt;
        data_class data;
        public Add_Subtopic()
        {
            InitializeComponent();
        }

        private void Add_Subtopic_Load(object sender, EventArgs e)
        {
            
            data = new data_class();
            data.open_connection();
            dt = data.select_sub_topics_all();
            sub_grid_view.DataSource = dt;

            DataTable sub = data.select_topics_all();
            comboBox1.ValueMember = "topic_id";
            comboBox1.DisplayMember = "topic_name";
            comboBox1.DataSource = sub;
        }

        private void add_btn_Click(object sender, EventArgs e)
        {
            try
            {
                if (topic_txt.Text == "")
                    throw new Exception("Sub Topic Name is empty!");
                if (Convert.ToInt32(comboBox1.SelectedValue) <= 0)
                    throw new Exception("Topic is empty!");

                data.insert_sub_topic(Convert.ToInt32(comboBox1.SelectedValue), topic_txt.Text);
                dt = data.select_sub_topics_all();
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
                    data.delete_sub_topic(rowId);

                    //refresh datagrid
                    sub_grid_view.Rows.RemoveAt(row.Index);
                }
            }
        }
        
    }
}
