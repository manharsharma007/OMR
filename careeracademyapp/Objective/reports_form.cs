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
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace omrapplication
{
    public partial class reports_form : Form
    {
        int test_id, batch_id, subject_id;
        data_class data;
        calculation_class calculation;
        DataTable dt, ds, topics;
        Hashtable topic_list, sub_list;
        SortedDictionary<int, int> question_list;
        public reports_form()
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
                dt = data.select_test_details(test_id);
                test_details_grid.DataSource = dt;
                test_details_grid.Visible = true;
                test_details_grid.Columns[0].HeaderText = "Name";
                test_details_grid.Columns[1].HeaderText = "Roll";
                test_details_grid.Columns[2].HeaderText = "Right";
                test_details_grid.Columns[3].HeaderText = "Wrong";
                test_details_grid.Columns[4].HeaderText = "Total";
                test_details_grid.Columns[5].HeaderText = "Questions";
                test_details_grid.Columns[6].HeaderText = "Total Marks";
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error",
    MessageBoxButtons.OK,
    MessageBoxIcon.Error,
    MessageBoxDefaultButton.Button1);
            }
        }

        private void reports_form_Load(object sender, EventArgs e)
        {
            data = new data_class();
            calculation = new calculation_class();
            dt = new DataTable();
            data.open_connection();
            ds = data.select_test();
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
                        int neg = Convert.ToInt32(row["t_wrong"]);
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
                        height = (neg) + ((neg) * (percent / 100));
                        g.FillRectangle(
                        new SolidBrush(Color.OrangeRed), l_x, l_y - (int)height, width, (int)height);
                        g.DrawString(neg + " Marks", new Font("Arial", 10.0f, FontStyle.Regular), Brushes.Gray, new Point(l_x, l_y + 8));
                        g.DrawString("(Negative)", new Font("Arial", 10.0f, FontStyle.Regular), Brushes.CadetBlue, new Point(l_x, l_y + 22));

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

                        g.DrawString(line_break, new Font("Arial", 10.0f, FontStyle.Bold), Brushes.Gray, new Point(x, l_y),StringFormat.GenericTypographic);

                        l_y += 30;

                        g.DrawString("Wrong Questions : ", new Font("Arial", 12.0f, FontStyle.Regular), Brushes.DarkBlue, new Rectangle(x, l_y, x + 750, y + 100), StringFormat.GenericTypographic);

                        l_y += 40;
                        g.DrawString(questions, new Font("Arial", 12.0f, FontStyle.Regular), Brushes.Black, new Rectangle(x, l_y, x + 750, y + 100), StringFormat.GenericTypographic);

                        l_y += 70;
                        g.DrawString(line_break, new Font("Arial", 10.0f, FontStyle.Bold), Brushes.Gray, new Point(x, l_y), StringFormat.GenericTypographic);

                        
                        img.Save(path + "/" + name + "_" + roll + ".jpg", ImageFormat.Jpeg);

                    }

                    data.ExportToExcel(test_details_grid, path + "/" + subject + ".xlsx");

                    #endregion

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string str = dt.Rows[i]["questions"].ToString();
                        if (str == "")
                            continue;
                        string[] words = str.Split(',');
                        string topic = "";
                        string sub_topic = "";
                        foreach (string word in words)
                        {
                            if (!question_list.ContainsKey(Convert.ToInt32(word)))
                            {
                                question_list.Add(Convert.ToInt32(word), 1);
                                topic = data.get_topic(batch_id, Convert.ToInt32(word));
                                sub_topic = data.get_sub_topic(batch_id, Convert.ToInt32(word));
                                topic_list.Add(word, topic);
                                sub_list.Add(word, sub_topic);
                            }
                            else
                                question_list[Convert.ToInt32(word)] = Convert.ToInt32(question_list[Convert.ToInt32(word)]) + 1;
                        }

                    }
                    int count = 1, _count_jpeg = 0, _t_students = dt.Rows.Count, _t_width = 100, _t_height = 18;

                    double percentage = 0, percent_width = 0; Brush brush = new SolidBrush(Color.DarkGreen);
                    g.Clear(Color.White);

                    x = 30;
                    y = 70;


                    foreach (KeyValuePair<int, int> de in question_list)
                    {
                        if (count == 1)
                        {
                            g.Clear(Color.White);
                            g.SmoothingMode =
               SmoothingMode.AntiAlias;
                            g.TextRenderingHint =
                                TextRenderingHint.AntiAlias;
                            g.InterpolationMode =
                                InterpolationMode.HighQualityBicubic;

                            g.CompositingMode = CompositingMode.SourceOver;
                            g.DrawString("ANALYSIS REPORT", new Font("Arial", 14.0f, FontStyle.Bold), new SolidBrush(Color.FromArgb(204, 17, 34)), new Point((img.Width / 2) - 200, y - 40));
                        }
                        else if (((count - 1) % 35) == 0 && (count) % 2 == 0)
                        {
                            x = img.Width / 2;
                            y = 70;
                        }
                        else if (((count - 1) % 35) == 0 && (count) % 2 != 0)
                        {
                            img.Save(path + "/Teachers_analysis_report_" + _count_jpeg + ".jpg", ImageFormat.Jpeg);
                            g.Clear(Color.White);
                            g.SmoothingMode =
               SmoothingMode.AntiAlias;
                            g.TextRenderingHint =
                                TextRenderingHint.AntiAlias;
                            g.InterpolationMode =
                                InterpolationMode.Bilinear;

                            g.CompositingMode = CompositingMode.SourceOver;

                            x = 30; y = 70;
                            g.DrawString("ANALYSIS REPORT", new Font("Arial", 14.0f, FontStyle.Bold), new SolidBrush(Color.FromArgb(204, 17, 34)), new Point((img.Width / 2) - 200, y - 40));

                            _count_jpeg++;
                        }

                        percentage = Convert.ToDouble(de.Value) / _t_students * 100;
                        percent_width = (double)_t_width * (percentage / 100);

                        if (percentage > 75.00)
                            brush = new SolidBrush(Color.Red);

                        else if (percentage > 55.00)
                            brush = new SolidBrush(Color.OrangeRed);

                        else if (percentage > 45.00)
                            brush = new SolidBrush(Color.Orange);

                        else if (percentage > 30.00)
                            brush = new SolidBrush(Color.GreenYellow);
                        else
                            brush = new SolidBrush(Color.LightGreen);


                        string measureString = string.Format("{0}. Question {1} : {2} (Topic : {3} | Sub Topic : {4} )", count, de.Key, de.Value, topic_list[de.Key.ToString()], sub_list[de.Key.ToString()]);
                        Font stringFont = new Font("Arial", 9.0f, FontStyle.Italic);

                        // Measure string.
                        SizeF stringSize = new SizeF();
                        stringSize = g.MeasureString(measureString, stringFont);


                        g.DrawString(measureString, new Font("Arial", 9.0f, FontStyle.Bold), Brushes.Black, new Point(x, y));

                        g.DrawRectangle(new Pen(Color.Olive, 1.7f), x + stringSize.Width + 20, y, _t_width, _t_height);
                        g.FillRectangle(
                        brush, x + stringSize.Width + 20, y, (int)percent_width, _t_height);

                        // Draw string to screen.
                        string pp = string.Format(" {0} %", percentage);

                        g.DrawString(pp, stringFont, Brushes.Black, new Point(x + (int)stringSize.Width + 21, y));

                        y += 40;
                        count++;

                    }

                    img.Save(path + "/Teachers_analysis_report_" + _count_jpeg + ".jpg", ImageFormat.Jpeg);


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
