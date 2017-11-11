using omrapplication.classes;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace omrapplication.Subjective
{
    public partial class SubReports : Form
    {
        int test_id, batch_id, subject_id;
        data_class data;
        calculation_class calculation;
        DataTable dt, ds, topics;
        Hashtable topic_list, sub_list;
        SortedDictionary<int, int> question_list;
        public SubReports()
        {
            InitializeComponent();
        }

        private void select_test_btn_Click(object sender, EventArgs e)
        {
            select_box_panel.Visible = true;
            select_box_panel.BringToFront();
        }

        private void cancel_btn_Click(object sender, EventArgs e)
        {
            select_box_panel.Visible = false;
        }

        private void select_btn_Click(object sender, EventArgs e)
        {
            try
            {
                if (select_box.Items.Count == 0)
                    throw new Exception("Test selected : NONE");

                test_id = (int)select_box.SelectedValue;
                dt = data.select_test_custom_details(test_id);
                test_details_grid.DataSource = dt;
                test_details_grid.Visible = true;
                test_details_grid.Columns[0].HeaderText = "Name";
                test_details_grid.Columns[1].HeaderText = "Roll";
                test_details_grid.Columns[2].HeaderText = "Total";
                test_details_grid.Columns[3].HeaderText = "Marks/Question";
                test_details_grid.Columns[4].HeaderText = "Total Marks";
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error",
    MessageBoxButtons.OK,
    MessageBoxIcon.Error,
    MessageBoxDefaultButton.Button1);
            }
        }

        private void SubReports_Load(object sender, EventArgs e)
        {
            data = new data_class();
            calculation = new calculation_class();
            dt = new DataTable();
            data.open_connection();
            ds = data.select_custom_test();
            select_box.ValueMember = "test_id";
            select_box.DisplayMember = "test_name";
            select_box.DataSource = ds;
        }

        private void generate_report_btn_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                if (dt.Rows.Count <= 0 || dt == null)
                    throw new Exception("Test not loaded correctly");

                foreach (DataRow row in ds.Rows)
                {

                    if (row["test_id"].ToString() == select_box.SelectedValue.ToString())
                    {
                        batch_id = Convert.ToInt32(row["batch_id"]);
                        subject_id = data.select_subject_id_from_batch(batch_id);
                        topics = data.select_topics(subject_id);
                        question_list = new SortedDictionary<int, int>();
                        topic_list = new Hashtable();
                        sub_list = new Hashtable();

                    }
                }


                DialogResult result = folderBrowserDialog1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    //
                    // The user selected a folder and pressed the OK button.
                    // We print the number of files found.
                    //
                    string path = folderBrowserDialog1.SelectedPath;

                    Bitmap img = new Bitmap(1066, 1507);
                    Graphics g = Graphics.FromImage(img);


                    int x = 80, y = 20;
                    int highest_marks = data.select_highest_marks(test_id);
                    int average_marks = data.select_average_marks(test_id);

                    string subject = data.subject_name(subject_id);

                    #region
                    foreach (DataRow row in dt.Rows)
                    {
                        g.Clear(Color.White);
                        g.SmoothingMode =
           SmoothingMode.AntiAlias;
                        g.TextRenderingHint =
                             TextRenderingHint.AntiAlias;
                        g.InterpolationMode =
                            InterpolationMode.HighQualityBicubic;

                        g.CompositingMode = CompositingMode.SourceOver;


                        int marks = Convert.ToInt32(row["t_total"]);
                        int total = Convert.ToInt32(row["total_marks"]);

                        double percent = calculation.percentageIncrease(total, 450);
                        double height = 0;

                        string line_break = "----------------------------------------------------------------------------------------------------------------------------------------------------------------------";
                        string name = row["name"].ToString();
                        string roll = row["roll"].ToString();
                        string questions = row["questions"].ToString();
                        if (questions == "")
                            questions = "NONE";
                        g.DrawString("REPORT CARD (CAREER ACADEMY)", new Font("Arial", 17.0f, FontStyle.Bold), new SolidBrush(Color.FromArgb(204, 17, 34)), new Point((img.Width / 2) - 200, y));
                        g.DrawString(line_break, new Font("Arial", 10.0f, FontStyle.Bold), Brushes.Gray, new Point(x, y + 50));
                        g.DrawString("STUDENT NAME : " + name, new Font("Arial", 10.0f, FontStyle.Bold), Brushes.Gray, new Point(x, y + 80));
                        g.DrawString("ROLL NUMBER : " + roll, new Font("Arial", 10.0f, FontStyle.Bold), Brushes.Gray, new Point(x, y + 100));
                        g.DrawString("SUBJECT : " + subject, new Font("Arial", 10.0f, FontStyle.Bold), Brushes.Gray, new Point(x, y + 120));
                        g.DrawString(line_break, new Font("Arial", 10.0f, FontStyle.Bold), Brushes.Gray, new Point(x, y + 150));

                        g.DrawString("-- PERFORMANCE GRAPH --", new Font("Arial", 14.0f, FontStyle.Bold), Brushes.Gray, new Point(x, y + 180));

                        int l_x = x, l_y = y + 220, space = 20, width = 80;
                        g.DrawLine(new Pen(Brushes.Gray, 2f), new Point(l_x + 30, l_y), new Point(x + 30, l_y + 500));
                        l_y += 500;
                        g.DrawLine(new Pen(Brushes.Gray, 2f), new Point(x + 30, l_y), new Point(l_x + 30 + 600, l_y));

                        l_x += 30;
                        DrawRotatedTextAt(g, -90, "Marks ->",
            l_x - 20, l_y - 100, new Font("Arial", 13.0f), Brushes.DarkGray);


                        l_x += space;
                        height = marks + (marks * (percent / 100));
                        g.FillRectangle(
                        new SolidBrush(Color.DarkGreen), l_x, l_y - (int)height, width, (int)height);
                        g.DrawString(marks + " Marks", new Font("Arial", 10.0f, FontStyle.Regular), Brushes.Gray, new Point(l_x, l_y + 8));
                        g.DrawString("(Obtained)", new Font("Arial", 10.0f, FontStyle.Regular), Brushes.CadetBlue, new Point(l_x, l_y + 22));


                        l_x += space;
                        l_x += width;
                        height = (total) + ((total) * (percent / 100));
                        g.FillRectangle(
                        new SolidBrush(Color.DeepSkyBlue), l_x, l_y - (int)height, width, (int)height);
                        g.DrawString(total + " Marks", new Font("Arial", 10.0f, FontStyle.Regular), Brushes.Gray, new Point(l_x, l_y + 8));
                        g.DrawString("(Total)", new Font("Arial", 10.0f, FontStyle.Regular), Brushes.CadetBlue, new Point(l_x, l_y + 22));


                        l_x += space;
                        l_x += width;
                        height = (average_marks) + ((average_marks) * (percent / 100));
                        g.FillRectangle(
                        new SolidBrush(Color.Orange), l_x, l_y - (int)height, width, (int)height);
                        g.DrawString(average_marks + " Marks", new Font("Arial", 10.0f, FontStyle.Regular), Brushes.Gray, new Point(l_x, l_y + 8));
                        g.DrawString("(Average)", new Font("Arial", 10.0f, FontStyle.Regular), Brushes.CadetBlue, new Point(l_x, l_y + 22));


                        l_x += space;
                        l_x += width;
                        height = (highest_marks) + ((highest_marks) * (percent / 100));
                        g.FillRectangle(
                        new SolidBrush(Color.OliveDrab), l_x, l_y - (int)height, width, (int)height);
                        g.DrawString(highest_marks + " Marks", new Font("Arial", 10.0f, FontStyle.Regular), Brushes.Gray, new Point(l_x, l_y + 8));
                        g.DrawString("(Highest)", new Font("Arial", 10.0f, FontStyle.Regular), Brushes.CadetBlue, new Point(l_x, l_y + 22));


                        l_y += 90;

                        g.DrawString(line_break, new Font("Arial", 10.0f, FontStyle.Bold), Brushes.Gray, new Point(x, l_y), StringFormat.GenericTypographic);

                        l_y += 30;

                        g.DrawString("MARKS/QUESTIONS : ", new Font("Arial", 12.0f, FontStyle.Regular), Brushes.DarkBlue, new Rectangle(x, l_y, x + 750, y + 100), StringFormat.GenericTypographic);

                        l_y += 40;
                        g.DrawString(questions, new Font("Arial", 12.0f, FontStyle.Regular), Brushes.Black, new Rectangle(x, l_y, x + 750, y + 100), StringFormat.GenericTypographic);

                        l_y += 70;
                        g.DrawString(line_break, new Font("Arial", 10.0f, FontStyle.Bold), Brushes.Gray, new Point(x, l_y), StringFormat.GenericTypographic);


                        img.Save(path + "/" + name + "_" + roll + ".jpg", ImageFormat.Jpeg);

                    }

                    data.ExportToExcel(test_details_grid, path + "/" + subject + ".xlsx");

                    #endregion

                    MessageBox.Show("Reports generated", "Note",
        MessageBoxButtons.OK,
        MessageBoxIcon.Information,
        MessageBoxDefaultButton.Button1);
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error",
    MessageBoxButtons.OK,
    MessageBoxIcon.Error,
    MessageBoxDefaultButton.Button1);
            }

            Cursor = Cursors.Arrow;
        }
        private void DrawRotatedTextAt(Graphics gr, float angle,
            string txt, int x, int y, Font the_font, Brush the_brush)
        {
            // Save the graphics state.
            GraphicsState state = gr.Save();
            gr.ResetTransform();

            // Rotate.
            gr.RotateTransform(angle);

            // Translate to desired position. Be sure to append
            // the rotation so it occurs after the rotation.
            gr.TranslateTransform(x, y, MatrixOrder.Append);

            // Draw the text at the origin.
            gr.DrawString(txt, the_font, the_brush, 0, 0);

            // Restore the graphics state.
            gr.Restore(state);
        }

        private void delete_btn_Click(object sender, EventArgs e)
        {
            try
            {
                if (test_id <= 0)
                    throw new Exception("Select test to delete");

                data.delete_test(test_id);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error",
    MessageBoxButtons.OK,
    MessageBoxIcon.Error,
    MessageBoxDefaultButton.Button1);
            }
        }


    }
}
