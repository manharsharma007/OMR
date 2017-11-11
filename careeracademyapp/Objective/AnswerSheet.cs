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
    public partial class AnswerSheet : Form
    {
        int template_id, _maximum_question = 0, _count = 0, batch_id = 0;
        DataTable _ds_questions, _ds_subject, _ds_topics, _ds_sub_topics;
        data_class data;
        DataSet ds;
        SqlDataAdapter da;
        bool new_instance = true;
        public AnswerSheet()
        {
            InitializeComponent();
            data = new data_class();
            data.open_connection();

        }

        private void AnswerSheet_Load(object sender, EventArgs e)
        {
            DataTable ds = data.select_template();
            template_combo_box.ValueMember = "template_id";
            template_combo_box.DisplayMember = "template_name";
            template_combo_box.DataSource = ds;

            _ds_subject = data.select_subjects();
            subject_combo_box.ValueMember = "subject_id";
            subject_combo_box.DisplayMember = "subject_name";
            subject_combo_box.DataSource = _ds_subject;

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Convert.ToInt32(template_combo_box.SelectedValue) > 0)
            {
                template_id = Convert.ToInt32(template_combo_box.SelectedValue);
                calculate_maximum(template_id);
                
            }
            /*
            if (form_loaded)
            {
                template_combo_box.Visible = false;
                label1.Visible = false;
                ds = data.select_options(Convert.ToInt32(template_combo_box.SelectedValue));
                da = data.da;
                
                dataGridView1.DataSource = ds.Tables[0];
            }
             * */
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
                MessageBox.Show(e.Message, "Error",
MessageBoxButtons.OK,
MessageBoxIcon.Error,
MessageBoxDefaultButton.Button1);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            UpdateAnswers();
        }


        private void calculate_maximum(int id)
        {
            _ds_questions = data.select_template_options(id, "Question");
            foreach (DataRow row in _ds_questions.Rows)
            {
                _maximum_question += Convert.ToInt32(row["rows"]);
            }
        }

        private void subject_combo_box_SelectedIndexChanged(object sender, EventArgs e)
        {
            _ds_topics = data.select_topics(Convert.ToInt32(subject_combo_box.SelectedValue));
            topic_combo_box.ValueMember = "topic_id";
            topic_combo_box.DisplayMember = "topic_name";
            topic_combo_box.DataSource = _ds_topics;
        }

        private void topic_combo_box_SelectedIndexChanged(object sender, EventArgs e)
        {
            _ds_sub_topics = data.select_sub_topics(Convert.ToInt32(topic_combo_box.SelectedValue));
            sub_topic_combo_box.ValueMember = "sub_id";
            sub_topic_combo_box.DisplayMember = "sub_name";
            sub_topic_combo_box.DataSource = _ds_sub_topics;
        }

        private void add_questions_Click(object sender, EventArgs e)
        {
            try
            {
                if (name_txt.Text == "")
                    throw new Exception("Name is required for new test batch");
                if (Convert.ToInt32(template_combo_box.SelectedValue) <= 0)
                    throw new Exception("Template is not selected");
                if (Convert.ToInt32(subject_combo_box.SelectedValue) <= 0)
                    throw new Exception("Subject is not selected");

                batch_id = data.check_batch(template_id, name_txt.Text);
                if (batch_id <= 0)
                    batch_id = data.insert_batch(template_id, Convert.ToInt32(subject_combo_box.SelectedValue), name_txt.Text);
                
                if(new_instance)
                {
                    _count = data.count_batch_options(batch_id);
                    _maximum_question -= _count;
                }

                if (_maximum_question <= 0 || Convert.ToInt32(questions_count_combo_box.SelectedItem) > _maximum_question)
                    throw new Exception("Maximum questions reached");

                for (int i = 0; i < Convert.ToInt32(questions_count_combo_box.SelectedItem); i++)
                {
                    data.insert_batch_option(batch_id, Convert.ToInt32(template_combo_box.SelectedValue), _count + 1, 0, 0, 0, Convert.ToInt32(topic_combo_box.SelectedValue), Convert.ToInt32(sub_topic_combo_box.SelectedValue));
                    _maximum_question--;
                    _count++;
                }

                ds = data.select_batch_options(batch_id);
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

                new_instance = false;
             
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error",
MessageBoxButtons.OK,
MessageBoxIcon.Error,
MessageBoxDefaultButton.Button1);
            }
        }

        private void refresh_btn_Click(object sender, EventArgs e)
        {
            dataGridView1.Refresh();
        }

        private void button2_Click(object sender, EventArgs e)
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
