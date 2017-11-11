using omrapplication.classes;
using omrapplication.Main;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace omrapplication
{
    public partial class Form1 : Form
    {
        int template_id, batch_id;
        data_class data;
        calculation_class calculation;

        string file = "", folder_path = "";
        string name = "", roll = "";
        string[] files;
        bool template_loaded = false, files_loaded = false, batch_loaded = false;
        Bitmap img;
        int total = 0, right = 0, wrong = 0, total_marks = 0;
        DataTable dt, name_table, roll_table, question_table, options;
        int lc_x, lc_y, rc_x, bc_y, width, height, x_lc, y_lc, x_rc, y_rc, x_bc, y_bc;


        char[] characters = { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
        int[] numbers = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };
        string questions = "";
        Thread th = null;



        public Form1()
        {
            InitializeComponent();
            data = new data_class();
            data.open_connection();
        }

        private void sheet_designer_btn_Click(object sender, EventArgs e)
        {
            sheet_designer sheet = new sheet_designer();
            sheet.ShowDialog();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DataTable ds = data.select_template();
            temp_comboBox.ValueMember = "template_id";
            temp_comboBox.DisplayMember = "template_name";
            temp_comboBox.DataSource = ds;
            dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Roll");
            dt.Columns.Add("Right");
            dt.Columns.Add("Wrong");
            dt.Columns.Add("Marks");
            dt.Columns.Add("Wrong Questions");
            dt.Columns.Add("Total Marks");


            if (ds.Rows.Count <= 0)
            {
                batch_panel.Visible = false;
                nothing_lbl.Visible = true;
            }
            status_lbl.Text = "Ready";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            AnswerSheet answersheet = new AnswerSheet();
            answersheet.ShowDialog();
        }

        private void delete_btn_Click(object sender, EventArgs e)
        {
            delete_template delete = new delete_template();
            delete.ShowDialog();
        }


        private void Form1_Activated(object sender, EventArgs e)
        {

            DataTable ds = data.select_template();
            temp_comboBox.ValueMember = "template_id";
            temp_comboBox.DisplayMember = "template_name";
            temp_comboBox.DataSource = ds;
            if (ds.Rows.Count <= 0)
            {
                batch_panel.Visible = false;
                nothing_lbl.Visible = true;
            }
            else if (!template_loaded || !batch_loaded)
            {
                batch_panel.Visible = true;
                nothing_lbl.Visible = false;

            }
        }
        /// <summary>
        /// OMR Reader Engine started
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void select_temp_btn_Click(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToInt32(temp_comboBox.SelectedValue) <= 0)
                    throw new Exception("Template is not selected");
                template_id = (int)temp_comboBox.SelectedValue;
                template_loaded = true;

                DataTable dt = data.select_batch(template_id);
                batch_comboBox.ValueMember = "batch_id";
                batch_comboBox.DisplayMember = "batch_name";
                batch_comboBox.DataSource = dt;
                status_lbl.Text = "Templated loaded";
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error",
MessageBoxButtons.OK,
MessageBoxIcon.Error,
MessageBoxDefaultButton.Button1);
            }

        }


        private void select_batch_btn_Click(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToInt32(batch_comboBox.SelectedValue) <= 0)
                    throw new Exception("Batch is not selected");
                batch_id = (int)batch_comboBox.SelectedValue;
                status_lbl.Text = "Batch loaded";
                batch_loaded = true;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error",
MessageBoxButtons.OK,
MessageBoxIcon.Error,
MessageBoxDefaultButton.Button1);
            }

        }
        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                //
                // The user selected a folder and pressed the OK button.
                // We print the number of files found.
                //
                files_loaded = true;
                folder_path = folderBrowserDialog1.SelectedPath;
                files = Directory.GetFiles(folderBrowserDialog1.SelectedPath);
                if (files.Length <= 0)
                {
                    files_loaded = false; 
                    MessageBox.Show("No Files Loaded", "Error",
 MessageBoxButtons.OK,
 MessageBoxIcon.Error,
 MessageBoxDefaultButton.Button1);
                }
            }
        }

        private void process_btn_Click(object sender, EventArgs e)
        {
            if (!files_loaded || !template_loaded || !batch_loaded)
            {
                status_lbl.Text = "Template or files not loaded\n";
                return;
            }

            nothing_lbl.Visible = false;
            batch_panel.Visible = false;
            sheet_panel.Visible = true;


            try
            {
                width = Convert.ToInt32(data.select_template_data(template_id, "width"));
                height = Convert.ToInt32(data.select_template_data(template_id, "height"));


                x_lc = Convert.ToInt32(data.select_template_option_data(template_id, "Markers", "LC", "x_cor"));
                y_lc = Convert.ToInt32(data.select_template_option_data(template_id, "Markers", "LC", "y_cor"));

                x_rc = Convert.ToInt32(data.select_template_option_data(template_id, "Markers", "RC", "x_cor"));
                y_rc = Convert.ToInt32(data.select_template_option_data(template_id, "Markers", "RC", "y_cor"));

                x_bc = Convert.ToInt32(data.select_template_option_data(template_id, "Markers", "BC", "x_cor"));
                y_bc = Convert.ToInt32(data.select_template_option_data(template_id, "Markers", "BC", "y_cor"));


                name_table = data.select_template_options(template_id, "Name");
                roll_table = data.select_template_options(template_id, "Roll");
                question_table = data.select_options_all(template_id);
                options = data.select_batch_options_all(batch_id);

                findtotal();

                th = new Thread(start_processing);
                th.Start();

                status_lbl2.Text = "Processing. Please Wait";
                status_lbl2.Visible = true;
                status_lbl2.BringToFront();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1);
            }


        }
        delegate void SetTextCallback(string text);

        private void SetText(string text)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.status_lbl.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.status_lbl.Text = text;
            }
        }
        private void SetBitmap(string imgage)
        {
            if (this.status_lbl.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetBitmap);
                this.Invoke(d, new object[] { imgage });
            }
            else
            {
                Bitmap img_local = new Bitmap(imgage);
                sheet_panel.BackgroundImage = img_local;
            }
        }
        private void start_processing()
        {
            int count = files.Length;
            int i = 1;
            foreach (string file_load in files)
            {
                if (Path.GetExtension(file_load) == ".jpeg" || Path.GetExtension(file_load) == ".jpg" || Path.GetExtension(file_load) == ".JPG")
                {

                    SetText("Processing file " + file_load + "\n------------------------------------------\n");
                    if (status_lbl2.InvokeRequired)
                        status_lbl2.Invoke(new Action(() =>
                        {
                            status_lbl2.Text = "Processing File : " + i + " of " + count;
                        }));
                    else
                    {
                        status_lbl2.Text = "Processing File : " + i + " of " + count;
                    }
                    SetBitmap(file_load);

                    img = new Bitmap(file_load);



                    calculation = new calculation_class();
                    double percent = 0.0;
                    if (img.Width >= width)
                        percent = calculation.percentageDecrease(img.Width, width);
                    else
                        percent = calculation.percentageIncrease(img.Width, width);

                    Size sz = calculation.increaseScale(img.Width, img.Height, percent);
                    img = resizeImage((Image)img, sz.Width, sz.Height);
                    img = new Bitmap(img, new Size(width, height));

                    img = greyscale(img);


                    if (!fix_orientiation())
                    {
                        MessageBox.Show("Not Properly Aligned. Image Path : " + file_load);
                        continue;
                    }

                    file = file_load;

                    //img.Save(file_load + "_processed"); // FOR DEBUGGING
                    findname();
                    findroll();
                    findresult();
                    name = name.Trim();
                    questions = questions.Trim(new char[] { ',', ' ' });
                    if (roll == "")
                        roll = "00";
                    dt.Rows.Add(name,
                                roll, right, wrong, total, questions, total_marks
                               );

                    name = "";
                    roll = "";
                    right = 0;
                    wrong = 0;
                    total = 0;
                    questions = "";

                    img.Dispose();
                    i++;
                }

            }
            if (result_gridview.InvokeRequired)
                result_gridview.Invoke(new Action(() =>
                {
                    result_gridview.DataSource = dt;
                    status_lbl2.Visible = false;
                    result_gridview.Visible = true;
                    result_gridview.Refresh();
                }));
            else
            {
                result_gridview.DataSource = dt;
                status_lbl2.Visible = false;
                result_gridview.Visible = true;
                result_gridview.Refresh();
            }
            if (save_result_txt.InvokeRequired)
                save_result_txt.Invoke(new Action(() =>
                {
                    save_result_txt.Visible = true;
                    save_result_btn.Visible = true;
                    csv_btn.Visible = true;
                    reset_btn.Visible = true;
                    status_lbl.Text = "Processing Completed";
                }));
            else
            {

                save_result_txt.Visible = true;
                save_result_btn.Visible = true;
                csv_btn.Visible = true;
                reset_btn.Visible = true;
                status_lbl.Text = "Processing Completed";
            }
            //result_gridview.DataSource = dt;
        }

        private bool fix_orientiation()
        {
            bool flag = true;

            deskew();

            flag = search_left_marker(0, 0, 20);
            if (flag)
                flag = search_right_marker(img.Width / 2, 0, 20);
            if (flag)
                flag = search_bottom_marker(0, img.Size.Height / 2, 20);

            /*
            int percent = findcircle(x_lc, y_lc, width/2);
            if (percent <= 5)
            {
                flag = false;
            }*/
            int diff_x = 0, diff_y = 0, n_width = width, n_height = height;
            diff_x = lc_x - 10;
            diff_y = lc_y - 10;


            try
            {
                img = cropAtRect(img, new Rectangle(lc_x, lc_y, rc_x + 20 - diff_x, bc_y + 20 - diff_y));
            }

            catch (Exception ex)
            {
                MessageBox.Show("Aborting : " + ex.Message, "ERROR",
   MessageBoxButtons.OK,
   MessageBoxIcon.Error,
   MessageBoxDefaultButton.Button1);
            }
            double percent = 0.0;
            if (img.Width >= width)
                percent = calculation.percentageDecrease(img.Width, width);
            else
                percent = calculation.percentageIncrease(img.Width, width);

            Size sz = calculation.increaseScale(img.Width, img.Height, percent);
            img = new Bitmap(img, new Size(width, height));
            return flag;
        }
        private bool search_left_marker(int x, int y, int width, bool decrement = false)
        {
            int col = img.Size.Width / 2;
            int height = col;
            if (col > img.Size.Height)
                height = img.Size.Height;

            bool flag = false;

            try
            {

                for (int i = x; i < col; i++)
                {
                    for (int j = y; j < height; j++)
                    {
                        Color pixel = img.GetPixel(i, j);
                        int average = (pixel.R + pixel.G + pixel.B) / 3;
                        if (average < 120)
                        {
                            lc_x = i;
                            lc_y = j;

                            flag = check_square(lc_x, lc_y, width);
                            if (flag)
                                break;
                        }
                        if (flag)
                            break;
                    }
                    if (flag)
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Aborting : " + ex.Message, "ERROR",
    MessageBoxButtons.OK,
    MessageBoxIcon.Error,
    MessageBoxDefaultButton.Button1);
            }
            return flag;
        }
        private bool search_right_marker(int x, int y, int width, bool decrement = false)
        {
            int col = this.width;

            int height = col;
            if (col > this.height / 2)
                height = this.height / 2;

            bool flag = false;
            try
            {
                for (int j = 0; j < height; j++)
                {
                    for (int i = col / 2; i < col; i++)
                    {
                        Color pixel = img.GetPixel(i, j);
                        int average = (pixel.R + pixel.G + pixel.B) / 3;
                        if (average < 120)
                        {
                            rc_x = i;

                            flag = check_square(rc_x, j, width);
                            if (flag)
                                break;
                        }
                        if (flag)
                            break;
                    }
                    if (flag)
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Aborting : " + ex.Message, "ERROR",
    MessageBoxButtons.OK,
    MessageBoxIcon.Error,
    MessageBoxDefaultButton.Button1);
            }
            return flag;
        }

        private bool search_bottom_marker(int x, int y, int width, bool decrement = false)
        {
            int col = img.Size.Height;
            bool flag = false;
            try
            {

                for (int i = x; i < img.Size.Width / 2; i++)
                {
                    for (int j = y; j < col; j++)
                    {
                        Color pixel = img.GetPixel(i, j);
                        int average = (pixel.R + pixel.G + pixel.B) / 3;
                        if (average < 120)
                        {
                            bc_y = j;

                            flag = check_square(i, bc_y, width);
                            if (flag)
                                break;
                        }
                        if (flag)
                            break;
                    }
                    if (flag)
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Aborting : " + ex.Message, "ERROR",
    MessageBoxButtons.OK,
    MessageBoxIcon.Error,
    MessageBoxDefaultButton.Button1);
            }
            return flag;
        }
        private void deskew()
        {
            Bitmap bm = new Bitmap(img);
            gmseDeskew sk = new gmseDeskew(bm);
            sk.Binary(bm, true);
            sk.NoiseRemoval(bm);
            double skewangle = sk.GetSkewAngle();
            img = RotateImage(img, (float)-skewangle);

        }
        private static Bitmap RotateImage(Bitmap bmp, float angle)
        {
            Graphics g = default(Graphics);
            Bitmap tmp = new Bitmap(bmp.Width, bmp.Height, PixelFormat.Format32bppRgb);

            tmp.SetResolution(bmp.HorizontalResolution, bmp.VerticalResolution);
            g = Graphics.FromImage(tmp);
            try
            {
                g.FillRectangle(Brushes.White, 0, 0, bmp.Width, bmp.Height);
                g.RotateTransform(angle);
                g.DrawImage(bmp, 0, 0);
            }
            finally
            {
                g.Dispose();
            }
            return tmp;
        }
        public Bitmap cropAtRect(Bitmap b, Rectangle r)
        {
            Bitmap nb = new Bitmap(r.Width + 10, r.Height + 10);
            Graphics g = Graphics.FromImage(nb);
            g.Clear(Color.White);
            Rectangle dest_rect = new Rectangle(10, 10, r.Width, r.Height);
            g.DrawImage(b, dest_rect, r, GraphicsUnit.Pixel);
            //g.DrawImage(b, -r.X, -r.Y);
            //nb.Save("hghvh.jpg");
            return nb;
        }

        private bool check_square(int x, int y, int width)
        {
            bool flag = false;
            int width_x = width + x;
            if (width_x > this.width)
                width_x = this.width - 1;
            int width_y = width + y;
            int count = 0;
            double percent = 0.00;

            for (int i = x; i <= width_x; i++)
            {
                Color pixel = img.GetPixel(i, y);
                int average = (pixel.R + pixel.G + pixel.B) / 3;
                if (i == width_x)
                {
                    if (average > 120)
                    {
                        count++;
                    }
                    percent = ((double)count / (width)) * 100;
                    count = 0;
                    if ((int)percent > 90)
                        flag = true;

                }
                if (average < 120)
                {
                    count++;
                }
            }
            if (flag)
            {

                for (int j = y; j <= width_y; j++)
                {
                    Color pixel = img.GetPixel(x, j);
                    int average = (pixel.R + pixel.G + pixel.B) / 3;
                    if (j == width_y)
                    {
                        if (average > 120)
                        {
                            count++;
                        }
                        percent = ((double)count / (width)) * 100;
                        if ((int)percent > 80)
                            flag = true;

                    }
                    if (average < 120)
                    {
                        count++;
                    }
                }
            }
            return flag;
        }
        private void findname()
        {
            if (name_table.Rows.Count > 0)
            {
                DataRow row_tb = name_table.Rows[0];
                int x = Convert.ToInt32(row_tb["x_cor"]);
                int y = Convert.ToInt32(row_tb["y_cor"]);
                int radius = Convert.ToInt32(row_tb["radius"]);
                int x_space = Convert.ToInt32(row_tb["x_space"]);
                int y_space = Convert.ToInt32(row_tb["y_space"]);


                int col = 20;
                int row = 26;

                int x_new = x, y_new = y;
                for (int i = 0; i < col; i++)
                {
                    for (int j = 0; j < row; j++)
                    {

                        int percent = findcircle(x, y, radius);
                        if (percent >= 40)
                        {
                            name += characters[j];
                            y = y_new;
                            break;
                        }
                        if (j == row - 1)
                        {
                            y = y_new;
                            name += " ";
                        }
                        else
                        {
                            y += radius * 2;
                            y += y_space;
                        }
                    }

                    if (i == col - 1)
                    {
                        x = x_new;
                    }
                    else
                    {
                        x += radius * 2;
                        x += x_space;
                    }
                }
            }
        }

        private void findroll()
        {


            if (roll_table.Rows.Count > 0)
            {
                DataRow row_tb = roll_table.Rows[0];
                int x = Convert.ToInt32(row_tb["x_cor"]);
                int y = Convert.ToInt32(row_tb["y_cor"]);
                int radius = Convert.ToInt32(row_tb["radius"]);
                int x_space = Convert.ToInt32(row_tb["x_space"]);
                int y_space = Convert.ToInt32(row_tb["y_space"]);

                int col = 4;
                int row = 10;

                int x_new = x, y_new = y;
                for (int i = 0; i < col; i++)
                {
                    for (int j = 0; j < row; j++)
                    {

                        int percent = findcircle(x, y, radius);
                        if (percent >= 40)
                        {
                            roll += numbers[j].ToString();
                            y = y_new;
                            break;
                        }
                        if (j == row - 1)
                        {
                            y = y_new;
                        }
                        else
                        {
                            y += radius * 2;
                            y += y_space;
                        }
                    }

                    if (i == col - 1)
                    {
                        x = x_new;
                    }
                    else
                    {
                        x += radius * 2;
                        x += x_space;
                    }
                }
            }
        }

        private void findresult()
        {
            int q = 1;
            for (int index = 0; index < options.Rows.Count; index++)
            {
                int x = Convert.ToInt32(question_table.Rows[index]["option_x"]);
                int y = Convert.ToInt32(question_table.Rows[index]["option_y"]);
                int radius = Convert.ToInt32(question_table.Rows[index]["radius"]) / 2;
                int x_space = Convert.ToInt32(question_table.Rows[index]["x_space"]);
                int y_space = Convert.ToInt32(question_table.Rows[index]["y_space"]);
                int column = Convert.ToInt32(question_table.Rows[index]["columns"]);
                int marks = Convert.ToInt32(options.Rows[index]["option_marks"]);
                int neg = Convert.ToInt32(options.Rows[index]["option_neg"]);
                int answer = Convert.ToInt32(options.Rows[index]["option_answer"]);

                int x_new = x, y_new = y;
                int selected_index = 0;

                for (int i = 0; i < column; i++)
                {
                    int percent = findcircle(x, y, radius);
                    if (percent >= 35)
                    {
                        if (selected_index == 0 && i < column - 1)
                        {
                            selected_index = i + 1;
                        }
                        else if (selected_index == 0 && i == column - 1)
                        {
                            if (answer == i + 1)
                            {
                                right += 1;
                                total += marks;


                                x += radius * 2;
                                x += x_space;
                                break;
                            }
                            else
                            {
                                wrong += 1;
                                total += neg;
                                if (q == 1)
                                {
                                    questions += (index+1).ToString();

                                    q++;
                                }
                                else
                                {

                                    questions += "," + (index + 1).ToString();

                                    q++;
                                }


                                x += radius * 2;
                                x += x_space;
                                break;
                            }
                        }

                        else if (selected_index > 0 && (i < column - 1 || i == column - 1))
                        {
                            wrong += 1;
                            total += neg;
                            if (q == 1)
                            {
                                questions += (index + 1).ToString();

                                q++;
                            }
                            else
                            {

                                questions += "," + (index + 1).ToString();

                                q++;
                            }


                            x += radius * 2;
                            x += x_space;
                            break;
                        }
                    }
                    else if (i == column - 1)
                    {
                        if (selected_index > 0)
                        {
                            if (answer == selected_index)
                            {
                                right += 1;
                                total += marks;


                                x += radius * 2;
                                x += x_space;
                                break;
                            }
                            else
                            {
                                wrong += 1;
                                total += neg;
                                if (q == 1)
                                {
                                    questions += (index + 1).ToString();

                                    q++;
                                }
                                else
                                {

                                    questions += "," + (index + 1).ToString();

                                    q++;
                                }
                            }
                        }
                        else
                        {
                            wrong += 1;
                            if (q == 1)
                            {
                                questions += (index + 1).ToString();

                                q++;
                            }
                            else
                            {

                                questions += "," + (index + 1).ToString();

                                q++;
                            }
                        }
                    }
                    x += radius * 2;
                    x += x_space;
                }
            }
        }

        private void findtotal()
        {
            for (int index = 0; index < options.Rows.Count; index++)
            {
                total_marks += Convert.ToInt32(options.Rows[index]["option_marks"]);
            }
        }
        private int findcircle(int x, int y, int radius)
        {
            double percent = 0; int count = 0;
            int total = (radius * 2) * (radius * 2);
            int width = radius * 2;
            int height = radius * 2;

            int x_new = x, y_new = y;
            for (int i = 0; i < height; i++)
            {
                for (int j = 0; j < width; j++)
                {
                    Color pixel = img.GetPixel(x, y);
                    int average = (pixel.R + pixel.G + pixel.B) / 3;

                    if (average < 120)
                    {
                        img.SetPixel(x, y, Color.Blue);
                        count += 1;
                    }
                    else if (x > 0 && y > 0)
                    {

                        Color top = img.GetPixel(x - 1, y);
                        int average_top = (top.R + top.G + top.B) / 3;


                        Color bottom = img.GetPixel(x + 1, y);
                        int average_bottom = (bottom.R + bottom.G + bottom.B) / 3;

                        Color left = img.GetPixel(x, y - 1);
                        int average_left = (left.R + left.G + left.B) / 3;

                        Color right = img.GetPixel(x, y + 1);
                        int average_right = (right.R + right.G + right.B) / 3;

                        if (average_top < 120 && average_bottom < 120 && average_left < 120 && average_right < 120)
                        {
                            img.SetPixel(x, y, Color.Red);
                            count += 1;
                        }

                    }

                    x++;
                }
                x = x_new;
                y++;
            }

            percent = ((double)count / total) * 100;
            return (int)percent;
        }


        private Bitmap greyscale(Bitmap img)
        {
            var rect = new Rectangle(0, 0, img.Width, img.Height);
            var data = img.LockBits(rect, ImageLockMode.ReadWrite, img.PixelFormat);
            var depth = Bitmap.GetPixelFormatSize(data.PixelFormat) / 8; //bytes per pixel

            var buffer = new byte[data.Width * data.Height * depth];

            //copy pixels to buffer
            Marshal.Copy(data.Scan0, buffer, 0, buffer.Length);

            for (int i = 0; i < data.Width; i++)
            {
                for (int j = 0; j < data.Height; j++)
                {
                    var offset = ((j * data.Width) + i) * depth;
                    // Dummy work    
                    // To grayscale (0.2126 R + 0.7152 G + 0.0722 B)
                    var b = 0.2126 * buffer[offset + 0] + 0.7152 * buffer[offset + 1] + 0.0722 * buffer[offset + 2];
                    buffer[offset + 0] = buffer[offset + 1] = buffer[offset + 2] = (byte)b;
                }
            }

            //Copy the buffer back to image
            Marshal.Copy(buffer, 0, data.Scan0, buffer.Length);

            img.UnlockBits(data);
            return img;
        }
        /*
        private void ToCsV(DataGridView dGV, string filename)
        {
            string stOutput = "";
            // Export titles:
            string sHeaders = "";

            for (int j = 0; j < dGV.Columns.Count; j++)
                sHeaders = sHeaders.ToString() + Convert.ToString(dGV.Columns[j].HeaderText) + "\t";
            stOutput += sHeaders + "\r\n";
            // Export data.
            for (int i = 0; i < dGV.RowCount; i++)
            {
                string stLine = "";
                for (int j = 0; j < dGV.Rows[i].Cells.Count; j++)
                    stLine = stLine.ToString() + Convert.ToString(dGV.Rows[i].Cells[j].Value) + "\t";
                stOutput += stLine + "\r\n";
            }
            Encoding utf16 = Encoding.GetEncoding(1254);
            byte[] output = utf16.GetBytes(stOutput);
            FileStream fs = new FileStream(filename, FileMode.Create);
            BinaryWriter bw = new BinaryWriter(fs);
            bw.Write(output, 0, output.Length); //write the encoded file
            bw.Flush();
            bw.Close();
            fs.Close();
        }
        */
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (th != null)
                if (th.ThreadState == ThreadState.Running)
                {
                    th.Abort();
                }
        }

        private void save_result_btn_Click(object sender, EventArgs e)
        {
            try
            {
                if (save_result_txt.Text == "")
                {
                    throw new Exception("Test name not provided.");
                }
                else if (dt == null)
                    throw new Exception("OMR data not initialized.");
                data.save_test(batch_id, save_result_txt.Text, dt);
                dt.Clear();
                sheet_panel.BackgroundImage = null;
                folder_path = "";
                name = "";
                roll = "";
                total = 0;
                right = 0;
                wrong = 0;
                total_marks = 0;
                questions = "";
                save_result_txt.Text = "";
                result_gridview.DataSource = null;
                MessageBox.Show("Test Details Saved", "Note",
    MessageBoxButtons.OK,
    MessageBoxIcon.Information,
    MessageBoxDefaultButton.Button1);
                save_result_btn.Visible = false;
                save_result_txt.Visible = false;
                csv_btn.Visible = false;
                reset_btn.Visible = false;
                template_loaded = false;
                template_id = 0;
                batch_loaded = false;
                batch_id = 0;
                Array.Clear(files, 0, files.Length);

                status_lbl.Text = "Test Details Saved";
                batch_panel.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Note",
    MessageBoxButtons.OK,
    MessageBoxIcon.Information,
    MessageBoxDefaultButton.Button1);
            }
        }

        private void reports_btn_Click(object sender, EventArgs e)
        {
            reports_form reports = new reports_form();
            reports.Show();
        }

        private void csv_btn_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel Documents (*.xls)|*.xls; *xlsx";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string excel_name = saveFileDialog1.FileName;
                data.ExportToExcel(result_gridview, excel_name);

                MessageBox.Show("File Saved", "Note",
    MessageBoxButtons.OK,
    MessageBoxIcon.Information,
    MessageBoxDefaultButton.Button1);
            }
            saveFileDialog1.Dispose();
        }

        private void addSubjectToolStripMenuItem_Click(object sender, EventArgs e)
        {

            Add_Subject obj = new Add_Subject();
            obj.ShowDialog();
        }

        private void addTopicToolStripMenuItem_Click(object sender, EventArgs e)
        {

            Add_Topic obj = new Add_Topic();
            obj.ShowDialog();
        }

        private void addSubTopicToolStripMenuItem_Click(object sender, EventArgs e)
        {

            Add_Subtopic obj = new Add_Subtopic();
            obj.ShowDialog();
        }

        private void deleteBatchToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Delete_Batch obj = new Delete_Batch();
            obj.ShowDialog();
        }

        private void deleteTestDetailsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            delete_test obj = new delete_test();
            obj.ShowDialog();
        }

        private void updateBatchToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Update_Batch obj = new Update_Batch();
            obj.ShowDialog();
        }
        public Bitmap resizeImage(Image imgPhoto, int new_width, int new_height)
        {
            Bitmap new_image = new Bitmap(new_width, new_height);
            Graphics g = Graphics.FromImage((Image)new_image);
            g.InterpolationMode = InterpolationMode.Bilinear;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.DrawImage(imgPhoto, 0, 0, new_width, new_height);
            return new_image;
        }

        private void reset_btn_Click(object sender, EventArgs e)
        {
            dt.Clear();
            sheet_panel.BackgroundImage = null;
            folder_path = "";
            name = "";
            roll = "";
            total = 0;
            right = 0;
            wrong = 0;
            total_marks = 0;
            questions = "";
            save_result_txt.Text = "";
            result_gridview.DataSource = null;
            MessageBox.Show("Test Details Saved", "Note",
MessageBoxButtons.OK,
MessageBoxIcon.Information,
MessageBoxDefaultButton.Button1);
            save_result_btn.Visible = false;
            save_result_txt.Visible = false;
            csv_btn.Visible = false;
            reset_btn.Visible = false;
            template_loaded = false;
            template_id = 0;
            batch_loaded = false;
            batch_id = 0;
            Array.Clear(files, 0, files.Length);

            status_lbl.Text = "Test Details Saved";
            batch_panel.Visible = true;
        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            help obj = new help();
            obj.ShowDialog();
        }
    }
}
