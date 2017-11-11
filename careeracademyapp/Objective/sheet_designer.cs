using omrapplication.classes;
using System;
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
    public partial class sheet_designer : Form
    {
        bool drag = false, editing = false;
        Bitmap img;
        string file = "";
        int X_COR, Y_COR, selectedIndex, IMAGE_X, IMAGE_Y, template_id, m_width = 20, m_height = 20, m_x = 10, m_y = 10, index = 1, indBlockWid = 0;

        enum option_type : int { NONE, QUESTION, NAME, ROLL };
        int option = (int)option_type.NONE;
        int bubbleWidth, bubbleHeight, blockRow, blockColumn, X_bubbleSpace, Y_bubbleSpace;


        bool mouseClicked = false;

        List<Panel> blocks = new List<Panel>();
        List<int> type = new List<int>();
        Dictionary<String, int> answers_option = new Dictionary<string, int>();
        Dictionary<String, int> custom_option = new Dictionary<string, int>();

        int drawingScale = 1;

        calculation_class calculate;
        data_class data;

        public sheet_designer()
        {
            InitializeComponent();
            
        }
        private void sheet_designer_Load(object sender, EventArgs e)
        {
            option = (int)option_type.NONE;
            bubbleWidth = 30;
            bubbleHeight = 30;
            X_bubbleSpace = x_space_bar.Value;
            Y_bubbleSpace = y_space_bar.Value;
            radius_bar.Value = bubbleWidth / 2;

            calculate = new calculation_class();
            data = new data_class();
            data.open_connection();
            status_lbl.Text = "Ready";
        }
        private void start_btn_Click(object sender, EventArgs e)
        {
            try
            {
                if (tpl_txt.Text == "")
                {
                    throw new Exception("Template name not provided.");
                }
                template_id = data.create_template(tpl_txt.Text, 'O');
                sheet_designer_btn.Enabled = true;
                delete_btn.Enabled = true;
                name_btn.Enabled = true;
                numeric_btn.Enabled = true;
                questions_btn.Enabled = true;
                x_space_bar.Enabled = true;
                y_space_bar.Enabled = true;
                radius_bar.Enabled = true;
                row_bar.Enabled = true;
                start_btn.Visible = false;
                tpl_txt.Visible = false;

                status_lbl.Text = "Template Created";
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error",
    MessageBoxButtons.OK,
    MessageBoxIcon.Stop,
    MessageBoxDefaultButton.Button1);
            }
        }

        private void sheet_designer_btn_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Image files | *.jpg; *.jpeg; *.jpe; *.jfif; *.png";
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                file = openFileDialog1.FileName;
                imagePanel.Visible = true;
                nothing_lbl.Visible = false;
                save_btn.Visible = true;

                Stream BitmapStream = System.IO.File.Open(file, System.IO.FileMode.Open);
                Image imgage = Image.FromStream(BitmapStream);

                img = new Bitmap(imgage);


                imagePanel.BackgroundImage = img;
                imagePanel.BackgroundImageLayout = ImageLayout.Stretch;
                imagePanel.Width = imagePanel.BackgroundImage.Width;
                imagePanel.Height = imagePanel.BackgroundImage.Height;

                double decrease = calculate.percentageDecrease(imagePanel.BackgroundImage.Width, sheet.Width);
                Size size = calculate.decreaseScale(imagePanel.Width, imagePanel.Height, decrease);

                imagePanel.Size = size;
                imagePanel.Location = new Point(0, 0);

                data.update_template(template_id, size.Width, size.Height);

                Panel panel = new Panel();
                Panel panel1 = new Panel();
                Panel panel2 = new Panel();

                imagePanel.Controls.Add(panel);
                panel.BackgroundImage = drawmarkers(m_width, m_height);
                panel.BackgroundImageLayout = ImageLayout.Zoom;
                panel.Size = new Size(m_width, m_height);
                panel.Location = new Point(m_x, m_y);

                imagePanel.Controls.Add(panel1);
                panel1.BackgroundImage = drawmarkers(m_width, m_height);
                panel1.BackgroundImageLayout = ImageLayout.Zoom;
                panel1.Size = new Size(m_width, m_height);
                panel1.Location = new Point(imagePanel.Width - m_x - m_width, m_y);

                imagePanel.Controls.Add(panel2);
                panel2.BackgroundImage = drawmarkers(m_width, m_height);
                panel2.BackgroundImageLayout = ImageLayout.Zoom;
                panel2.Size = new Size(m_width, m_height);
                panel2.Location = new Point(m_x, imagePanel.Height - m_y - m_height);


                data.insert_new_option(template_id, "Markers", "LC");
                data.update_option(template_id, "Markers", "LC", "x_cor", m_x.ToString());
                data.update_option(template_id, "Markers", "LC", "y_cor", m_y.ToString());

                data.insert_new_option(template_id, "Markers", "RC");
                data.update_option(template_id, "Markers", "RC", "x_cor", (imagePanel.Width - m_x - m_width).ToString());
                data.update_option(template_id, "Markers", "RC", "y_cor", m_y.ToString());

                data.insert_new_option(template_id, "Markers", "BC");
                data.update_option(template_id, "Markers", "BC", "x_cor", m_x.ToString());
                data.update_option(template_id, "Markers", "BC", "y_cor", (imagePanel.Height - m_y - m_height).ToString());
                
                BitmapStream.Dispose();
                editing = true;

                status_lbl.Text = "File Created";
            }
        }
        public Bitmap drawmarkers(int width, int height)
        {
            Bitmap bmp = new Bitmap(width, height);

            Graphics g2 = Graphics.FromImage(bmp);

            g2.FillRectangle(Brushes.Black, new Rectangle(0, 0, width, height));

            g2.Dispose();

            return bmp;


        }
        /*
        private void sheet_designer_Resize(object sender, EventArgs e)
        {
            if (!editing)
                return;
            imagePanel.Width = imagePanel.BackgroundImage.Width;
            imagePanel.Height = imagePanel.BackgroundImage.Height;

            double decrease = calculate.percentageDecrease(imagePanel.BackgroundImage.Width, sheet.Width);
            Size size = calculate.decreaseScale(imagePanel.Width, imagePanel.Height, decrease);

            data.update_template(template_id, size.Width, size.Height);

            imagePanel.Size = size;
        }
        */
        private void name_btn_Click(object sender, EventArgs e)
        {
            option = (int)option_type.NAME;
            blockRow = 26;
            blockColumn = 20;

            addNewBlock(makeBubbleSheet('A'));
            status_lbl.Text = "Name Block Added";
        }

        private void numeric_btn_Click(object sender, EventArgs e)
        {
            option = (int)option_type.ROLL;
            blockRow = 10;
            blockColumn = 4;

            addNewBlock(makeBubbleSheet('0'));
            status_lbl.Text = "Roll Block Added";
        }

        private void questions_btn_Click(object sender, EventArgs e)
        {
            option = (int)option_type.QUESTION;
            blockRow = 5;
            blockColumn = 4;

            addNewBlock(makeBubbleSheet(index, blockColumn, blockRow, 'A', 1));
            index += 5;
            status_lbl.Text = "Question Block Added";
        }






        private void delete_btn_Click_1(object sender, EventArgs e)
        {
            int count = blocks.Count;
            if (count <= 0)
                return;
            if (blocks.Count > 0)

                if (MessageBox.Show(this, "Do you realy want to delete the block? this cannot be undone.", "Notice", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                {
                    if (selectedIndex == 0 || selectedIndex >= blocks.Count)
                        selectedIndex = blocks.Count - 1;

                    int rows = 0;
                    if (type[selectedIndex] == (int)option_type.NAME)
                    {
                        data.delete_option(template_id, "Name");
                    }
                    else if (type[selectedIndex] == (int)option_type.ROLL)
                    {
                        data.delete_option(template_id, "Roll");
                    }
                    else if (type[selectedIndex] == (int)option_type.QUESTION)
                    {
                        rows = Convert.ToInt32(data.select_template_option_data(template_id, "Question", blocks[selectedIndex].Name, "rows"));
                        data.delete_question_option(template_id, "Question", blocks[selectedIndex].Name);
                    }

                    if (type[selectedIndex] == (int)option_type.QUESTION)
                        answers_option.Remove(blocks[selectedIndex].Name);

                    sheet.Controls.Remove((sender as Button).Parent);
                    blocks[selectedIndex].Dispose();
                    blocks.RemoveAt(selectedIndex);
                    type.RemoveAt(selectedIndex);

                    index -= rows;



                    status_lbl.Text = "Deleted Successfully";
                }
        }


        private void save_btn_Click_1(object sender, EventArgs e)
        {
            saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Image files | *.bmp";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;
                data.add_template_options(template_id);
                string name = saveFileDialog1.FileName;
                int width = imagePanel.Size.Width;
                int height = imagePanel.Size.Height;

                int width_new = imagePanel.BackgroundImage.Width;
                int height_new = imagePanel.BackgroundImage.Height;

                Bitmap bm = new Bitmap(width, height);

                SaveBitmap(imagePanel, name);


                bm.Dispose();

                Cursor = Cursors.Arrow;

                save_btn.Enabled = false;
                status_lbl.Text = "Please restart to make new template";
                MessageBox.Show("Please restart window to make new template", "Note",
    MessageBoxButtons.OK,
    MessageBoxIcon.Information,
    MessageBoxDefaultButton.Button1);
            }

        }
        private Bitmap ResizeImage(Bitmap originalImage, int maxWidth, int maxHeight)
        {

            Bitmap newImage = new Bitmap(originalImage, maxWidth, maxHeight);

            Graphics g = Graphics.FromImage(newImage);
            g.Clear(Color.White);
            g.SmoothingMode =
            SmoothingMode.AntiAlias;
            g.TextRenderingHint =
                TextRenderingHint.AntiAlias;
            g.InterpolationMode =
                InterpolationMode.HighQualityBicubic;

            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            g.DrawImage(originalImage, 0, 0, newImage.Width, newImage.Height);

            originalImage.Dispose();

            return newImage;
        }
        private void x_space_bar_ValueChanged_1(object sender, EventArgs e)
        {
            int count = blocks.Count;
            if (count <= 0)
                return;
            X_bubbleSpace = x_space_bar.Value;

            if (type[selectedIndex] == (int)option_type.NAME)
            {
                option = (int)option_type.NAME;
                blockRow = 26;
                blockColumn = 20;
                blocks[selectedIndex].BackgroundImage = makeBubbleSheet('A');
            }
            else if (type[selectedIndex] == (int)option_type.ROLL)
            {
                option = (int)option_type.ROLL;
                blockRow = 10;
                blockColumn = 4;
                blocks[selectedIndex].BackgroundImage = makeBubbleSheet('0');
            }
            else if (type[selectedIndex] == (int)option_type.QUESTION)
            {
                option = (int)option_type.QUESTION;
                blockColumn = 4;
                blockRow = Convert.ToInt32(data.select_template_option_data(template_id, "Question", blocks[selectedIndex].Name, "rows"));
                index -= blockRow;
                blocks[selectedIndex].BackgroundImage = makeBubbleSheet(index, blockColumn, blockRow, 'A', 1);
                index += blockRow;
            }

            blocks[selectedIndex].Size = new Size((int)Math.Round((double)blocks[selectedIndex].BackgroundImage.Width / drawingScale), (int)Math.Round((double)blocks[selectedIndex].BackgroundImage.Height / drawingScale));

            #region
            if (type[selectedIndex] == (int)option_type.NAME)
            {

                data.update_option(template_id, "Name", blocks[selectedIndex].Name, "x_cor", blocks[selectedIndex].Location.X.ToString());
                data.update_option(template_id, "Name", blocks[selectedIndex].Name, "y_cor", blocks[selectedIndex].Location.Y.ToString());
                data.update_option(template_id, "Name", blocks[selectedIndex].Name, "x_space", x_space_bar.Value.ToString());


            }
            else if (type[selectedIndex] == (int)option_type.ROLL)
            {

                data.update_option(template_id, "Roll", blocks[selectedIndex].Name, "x_cor", blocks[selectedIndex].Location.X.ToString());
                data.update_option(template_id, "Roll", blocks[selectedIndex].Name, "y_cor", blocks[selectedIndex].Location.Y.ToString());
                data.update_option(template_id, "Roll", blocks[selectedIndex].Name, "x_space", x_space_bar.Value.ToString());

            }
            else if (type[selectedIndex] == (int)option_type.QUESTION)
            {

                data.update_option(template_id, "Question", blocks[selectedIndex].Name, "x_cor", (blocks[selectedIndex].Location.X + indBlockWid).ToString());
                data.update_option(template_id, "Question", blocks[selectedIndex].Name, "y_cor", blocks[selectedIndex].Location.Y.ToString());
                data.update_option(template_id, "Question", blocks[selectedIndex].Name, "x_space", x_space_bar.Value.ToString());
            }
            #endregion
        }

        private void y_space_bar_ValueChanged_1(object sender, EventArgs e)
        {
            int count = blocks.Count;
            if (count <= 0)
                return;

            Y_bubbleSpace = y_space_bar.Value;
            if (type[selectedIndex] == (int)option_type.NAME)
            {
                option = (int)option_type.NAME;
                blockRow = 26;
                blockColumn = 20;
                blocks[selectedIndex].BackgroundImage = makeBubbleSheet('A');
            }
            else if (type[selectedIndex] == (int)option_type.ROLL)
            {
                option = (int)option_type.ROLL;
                blockRow = 10;
                blockColumn = 4;
                blocks[selectedIndex].BackgroundImage = makeBubbleSheet('0');
            }
            else if (type[selectedIndex] == (int)option_type.QUESTION)
            {
                option = (int)option_type.QUESTION;
                blockColumn = 4;
                blockRow = Convert.ToInt32(data.select_template_option_data(template_id, "Question", blocks[selectedIndex].Name, "rows"));
                index -= blockRow;
                blocks[selectedIndex].BackgroundImage = makeBubbleSheet(index, blockColumn, blockRow, 'A', 1);
                index += blockRow;
            }

            blocks[selectedIndex].Size = new Size((int)Math.Round((double)blocks[selectedIndex].BackgroundImage.Width / drawingScale), (int)Math.Round((double)blocks[selectedIndex].BackgroundImage.Height / drawingScale));

            #region
            if (type[selectedIndex] == (int)option_type.NAME)
            {
                data.update_option(template_id, "Name", blocks[selectedIndex].Name, "x_cor", blocks[selectedIndex].Location.X.ToString());
                data.update_option(template_id, "Name", blocks[selectedIndex].Name, "y_cor", blocks[selectedIndex].Location.Y.ToString());
                data.update_option(template_id, "Name", blocks[selectedIndex].Name, "y_space", y_space_bar.Value.ToString());
            }
            else if (type[selectedIndex] == (int)option_type.ROLL)
            {

                data.update_option(template_id, "Roll", blocks[selectedIndex].Name, "x_cor", blocks[selectedIndex].Location.X.ToString());
                data.update_option(template_id, "Roll", blocks[selectedIndex].Name, "y_cor", blocks[selectedIndex].Location.Y.ToString());
                data.update_option(template_id, "Roll", blocks[selectedIndex].Name, "y_space", y_space_bar.Value.ToString());
            }
            else if (type[selectedIndex] == (int)option_type.QUESTION)
            {

                data.update_option(template_id, "Question", blocks[selectedIndex].Name, "x_cor", (blocks[selectedIndex].Location.X + indBlockWid).ToString());
                data.update_option(template_id, "Question", blocks[selectedIndex].Name, "y_cor", blocks[selectedIndex].Location.Y.ToString());
                data.update_option(template_id, "Question", blocks[selectedIndex].Name, "y_space", y_space_bar.Value.ToString());
            }
            #endregion
        }

        private void row_bar_Scroll_1(object sender, EventArgs e)
        {
            int count = blocks.Count;
            if (count <= 0)
                return;
            try
            {
                if (type[selectedIndex] == (int)option_type.QUESTION)
                {
                    option = (int)option_type.QUESTION;
                    blockRow = row_bar.Value;
                    blockColumn = 4;
                    int rows = Convert.ToInt32(data.select_template_option_data(template_id, "Question", blocks[selectedIndex].Name, "rows"));
                    blocks[selectedIndex].BackgroundImage = makeBubbleSheet(index - rows, blockColumn, blockRow, 'A', 1);
                    index -= rows;
                    blockRow = row_bar.Value;
                    index += blockRow;
                    blocks[selectedIndex].Size = new Size((int)Math.Round((double)blocks[selectedIndex].BackgroundImage.Width / drawingScale), (int)Math.Round((double)blocks[selectedIndex].BackgroundImage.Height / drawingScale));


                    data.update_option(template_id, "Question", blocks[selectedIndex].Name, "x_cor", (blocks[selectedIndex].Location.X + indBlockWid).ToString());
                    data.update_option(template_id, "Question", blocks[selectedIndex].Name, "y_cor", blocks[selectedIndex].Location.Y.ToString());
                    data.update_option(template_id, "Question", blocks[selectedIndex].Name, "rows", blockRow.ToString());

                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void radius_bar_Scroll_1(object sender, EventArgs e)
        {
            int count = blocks.Count;
            if (count <= 0)
                return;
            bubbleWidth = Convert.ToInt32(radius_bar.Value) * 2;
            bubbleHeight = Convert.ToInt32(radius_bar.Value) * 2;
            if (type[selectedIndex] == (int)option_type.NAME)
            {
                option = (int)option_type.NAME;
                blockRow = 26;
                blockColumn = 20;
                blocks[selectedIndex].BackgroundImage = makeBubbleSheet('A');
            }
            else if (type[selectedIndex] == (int)option_type.ROLL)
            {
                option = (int)option_type.ROLL;
                blockRow = 10;
                blockColumn = 4;
                blocks[selectedIndex].BackgroundImage = makeBubbleSheet('0');
            }
            else if (type[selectedIndex] == (int)option_type.QUESTION)
            {
                option = (int)option_type.QUESTION;
                blockColumn = 4;
                blockRow = Convert.ToInt32(data.select_template_option_data(template_id, "Question", blocks[selectedIndex].Name, "rows"));
                index -= blockRow;
                blocks[selectedIndex].BackgroundImage = makeBubbleSheet(index, blockColumn, blockRow, 'A', 1);
                index += blockRow;
            }

            blocks[selectedIndex].Size = new Size((int)Math.Round((double)blocks[selectedIndex].BackgroundImage.Width / drawingScale), (int)Math.Round((double)blocks[selectedIndex].BackgroundImage.Height / drawingScale));

            #region
            if (type[selectedIndex] == (int)option_type.NAME)
            {

                data.update_option(template_id, "Name", blocks[selectedIndex].Name, "x_cor", blocks[selectedIndex].Location.X.ToString());
                data.update_option(template_id, "Name", blocks[selectedIndex].Name, "y_cor", blocks[selectedIndex].Location.Y.ToString());
                data.update_option(template_id, "Name", blocks[selectedIndex].Name, "radius", radius_bar.Value.ToString());

            }
            else if (type[selectedIndex] == (int)option_type.ROLL)
            {

                data.update_option(template_id, "Roll", blocks[selectedIndex].Name, "x_cor", blocks[selectedIndex].Location.X.ToString());
                data.update_option(template_id, "Roll", blocks[selectedIndex].Name, "y_cor", blocks[selectedIndex].Location.Y.ToString());
                data.update_option(template_id, "Roll", blocks[selectedIndex].Name, "radius", radius_bar.Value.ToString());

            }
            else if (type[selectedIndex] == (int)option_type.QUESTION)
            {
                data.update_option(template_id, "Question", blocks[selectedIndex].Name, "x_cor", blocks[selectedIndex].Location.X.ToString());
                data.update_option(template_id, "Question", blocks[selectedIndex].Name, "y_cor", blocks[selectedIndex].Location.Y.ToString());
                data.update_option(template_id, "Question", blocks[selectedIndex].Name, "radius", radius_bar.Value.ToString());
            }
            #endregion

        }

        private void imagePanel_MouseMove(object sender, MouseEventArgs e)
        {
            if (drag)
            {
                imagePanel.Location = new Point(e.X + imagePanel.Left - IMAGE_X, e.Y + imagePanel.Top - IMAGE_Y);
            }
        }

        private void imagePanel_MouseUp(object sender, MouseEventArgs e)
        {
            drag = false;
        }

        private void imagePanel_MouseDown(object sender, MouseEventArgs e)
        {
            drag = true;
            IMAGE_X = e.X;
            IMAGE_Y = e.Y;
        }

        /// <summary>
        /// MAIN DATA PROCESSING
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>


        private void addNewBlock(Image img)
        {
            #region
            int i = blocks.Count;
            blocks.Add(new Panel());
            sheet.Controls.Add(blocks[i]);
            blocks[i].Name = "customControlOP" + i;
            blocks[i].MouseDown += new MouseEventHandler(block_MouseDown);
            blocks[i].MouseMove += new MouseEventHandler(block_MouseMove);
            blocks[i].MouseUp += new MouseEventHandler(block_MouseUp);
            blocks[i].MouseClick += new MouseEventHandler(block_mouseClick);
            blocks[i].BackgroundImage = img;
            blocks[i].BackgroundImageLayout = ImageLayout.Stretch;
            blocks[i].Size = new Size((int)Math.Round((double)blocks[i].BackgroundImage.Width / drawingScale), (int)Math.Round((double)blocks[i].BackgroundImage.Height / drawingScale));

            blocks[i].BackColor = Color.FromArgb(0, 0, 0, 0);
            blocks[i].Parent = imagePanel;
            blocks[i].BackColor = Color.Transparent;
            blocks[i].Location = new Point(0, 0);
            blocks[i].BringToFront();
            type.Add((int)option);

            if (type[i] == (int)option_type.QUESTION)
                answers_option[blocks[i].Name] = blockRow;


            if (type[i] == (int)option_type.NAME)
            {
                data.insert_new_option(template_id, "Name", blocks[i].Name);
                data.update_option(template_id, "Name", blocks[i].Name, "x_cor", blocks[i].Location.X.ToString());
                data.update_option(template_id, "Name", blocks[i].Name, "y_cor", blocks[i].Location.Y.ToString());
                data.update_option(template_id, "Name", blocks[i].Name, "radius", (bubbleWidth / 2).ToString());
                data.update_option(template_id, "Name", blocks[i].Name, "x_space", X_bubbleSpace.ToString());
                data.update_option(template_id, "Name", blocks[i].Name, "y_space", Y_bubbleSpace.ToString());

            }
            else if (type[i] == (int)option_type.ROLL)
            {
                data.insert_new_option(template_id, "Roll", blocks[i].Name);
                data.update_option(template_id, "Roll", blocks[i].Name, "x_cor", blocks[i].Location.X.ToString());
                data.update_option(template_id, "Roll", blocks[i].Name, "y_cor", blocks[i].Location.Y.ToString());
                data.update_option(template_id, "Roll", blocks[i].Name, "radius", (bubbleWidth / 2).ToString());
                data.update_option(template_id, "Roll", blocks[i].Name, "x_space", X_bubbleSpace.ToString());
                data.update_option(template_id, "Roll", blocks[i].Name, "y_space", Y_bubbleSpace.ToString());

            }
            else if (type[i] == (int)option_type.QUESTION)
            {
                data.insert_new_option(template_id, "Question", blocks[i].Name);
                data.update_option(template_id, "Question", blocks[i].Name, "x_cor", (blocks[i].Location.X + indBlockWid).ToString());
                data.update_option(template_id, "Question", blocks[i].Name, "y_cor", blocks[i].Location.Y + indBlockWid.ToString());
                data.update_option(template_id, "Question", blocks[i].Name, "radius", (bubbleWidth / 2).ToString());
                data.update_option(template_id, "Question", blocks[i].Name, "x_space", X_bubbleSpace.ToString());
                data.update_option(template_id, "Question", blocks[i].Name, "y_space", Y_bubbleSpace.ToString());
                data.update_option(template_id, "Question", blocks[i].Name, "rows", blockRow.ToString());
                data.update_option(template_id, "Question", blocks[i].Name, "columns", blockColumn.ToString());
            }
            #endregion

        }

        private void block_mouseClick(object sender, MouseEventArgs e)
        {
            for (int i = 0; i <= blocks.Count; i++)
            {
                Control ctrl = sender as Control;
                if (blocks[i].Name == ctrl.Name)
                {
                    selectedIndex = i;
                    break;
                }
            }
        }

        private void block_MouseUp(object sender, MouseEventArgs e)
        {
            mouseClicked = false;
        }


        private void block_MouseDown(object sender, MouseEventArgs e)
        {
            for (int i = 0; i <= blocks.Count; i++)
            {
                Control ctrl = sender as Control;
                if (blocks[i].Name == ctrl.Name)
                {
                    selectedIndex = i;
                    break;
                }
            }
            mouseClicked = true;
            X_COR = e.X;
            Y_COR = e.Y;
        }

        private void block_MouseMove(object sender, MouseEventArgs e)
        {

            if (mouseClicked)
            {
                blocks[selectedIndex].Location = new Point(e.X + this.blocks[selectedIndex].Left - X_COR, e.Y + this.blocks[selectedIndex].Top - Y_COR);

                if (type[selectedIndex] == (int)option_type.NAME)
                {
                    data.update_option(template_id, "Name", blocks[selectedIndex].Name, "x_cor", blocks[selectedIndex].Location.X.ToString());
                    data.update_option(template_id, "Name", blocks[selectedIndex].Name, "y_cor", blocks[selectedIndex].Location.Y.ToString());
                }
                else if (type[selectedIndex] == (int)option_type.ROLL)
                {
                    data.update_option(template_id, "Roll", blocks[selectedIndex].Name, "x_cor", blocks[selectedIndex].Location.X.ToString());
                    data.update_option(template_id, "Roll", blocks[selectedIndex].Name, "y_cor", blocks[selectedIndex].Location.Y.ToString());
                }
                else if (type[selectedIndex] == (int)option_type.QUESTION)
                {
                    data.update_option(template_id, "Question", blocks[selectedIndex].Name, "x_cor", (blocks[selectedIndex].Location.X + indBlockWid).ToString());
                    data.update_option(template_id, "Question", blocks[selectedIndex].Name, "y_cor", blocks[selectedIndex].Location.Y.ToString());
                }
            }

        }

        //C
        //
        /// <summary>
        /// draws image block bound to number of info.
        /// if indexing start is zero, block will be converted to reg Block
        /// </summary>
        /// <param name="indexingStart"></param>
        /// <param name="NumberOfChoices"></param>
        /// <param name="NumberOfLines"></param>
        /// <param name="OptionChar"></param>
        /// <param name="Scale"></param>
        /// <returns></returns>
        private Image makeBubbleSheet(char OptionChar)
        {
            #region

            int indBlockXSpacing = 0, indBlockWid = 0, circleSize = bubbleWidth; float lineWidth = 1.55f;
            var brush = new SolidBrush(Color.FromArgb(231,48,128));
            Font labelFont = new Font("Arial", (circleSize * 55) / 100);
            Bitmap tBmp = new Bitmap(1, 1);
            tBmp.MakeTransparent();
            Graphics g2 = Graphics.FromImage(tBmp);
            

            int maxWid = 0;


            indBlockWid = maxWid + indBlockXSpacing;

            tBmp = new Bitmap(indBlockWid + blockColumn * (circleSize + X_bubbleSpace) - X_bubbleSpace + 1, blockRow * (circleSize + Y_bubbleSpace) - Y_bubbleSpace + 1);

            g2 = Graphics.FromImage(tBmp);

            g2.Clear(Color.White);
            g2.SmoothingMode =
            SmoothingMode.AntiAlias;
            g2.TextRenderingHint =
                TextRenderingHint.AntiAlias;
            g2.InterpolationMode =
                InterpolationMode.HighQualityBicubic;
            
            g2.CompositingMode = CompositingMode.SourceOver;
            

            //g2.DrawRectangle(new Pen(brush, lineWidth), new Rectangle(indBlockWid, 0, tBmp.Width - indBlockWid, tBmp.Height));

            //Only used in case of boxes
            //int boxWid = (int)Math.Round((double)(tBmp.Width - indBlockWid) / blockColumn);
            //double lineToMargin = 0.6;
            //int lineMargin = (int)Math.Round(boxWid * (1 - lineToMargin)) / 2;
            // int boxHei = (int)Math.Round((double)tBmp.Height / blockRow);
            // int lineHei = (int)Math.Round(boxHei * lineToMargin);

            for (int j = 0; j < blockRow; j++)
            {
                for (int i = 0; i < blockColumn; i++)
                {
                    g2.DrawEllipse(new Pen(brush, lineWidth),
                           new Rectangle(indBlockWid + (circleSize + X_bubbleSpace) * i, (circleSize + Y_bubbleSpace) * j, circleSize, circleSize));
                    if (OptionChar != 0 && OptionChar != 'N')
                    {
                        string num = "";
                        if (option == (int)(option_type.NAME) || option == (int)(option_type.ROLL))
                            num = ((char)(OptionChar + j)).ToString();
                        else
                            num = ((char)(OptionChar + i)).ToString();

                        int numX = (circleSize - (int)g2.MeasureString(num, labelFont).Width) / 2;
                        Point numLoc = new Point(indBlockWid + (circleSize + X_bubbleSpace) * i + numX, (circleSize + Y_bubbleSpace) * j + numX);
                        if (OptionChar == 'a' || OptionChar == '1' || OptionChar == '0')
                            numLoc = new Point(numLoc.X, numLoc.Y - circleSize / 15);
                        g2.DrawString(num, labelFont, brush, numLoc);
                    }

                }
            }

            g2.Dispose();
            tBmp.MakeTransparent(Color.White);
            return tBmp;
            #endregion


        }


        private Bitmap makeBubbleSheet(int indexingStart, int NumberOfChoices, int NumberOfLines, char OptionChar, double Scale)
        {
            int indBlockXSpacing = 5, spacing = Y_bubbleSpace, circleSize = bubbleWidth; float lineWidth = 1.55f;
            var brush = new SolidBrush(Color.FromArgb(231, 48, 128));
            Font labelFont = new Font("Arial", (circleSize * 50) / 100);
            Font indexingFont = new Font("Arial", (circleSize * 50) / 70);
            Bitmap tBmp = new Bitmap(1, 1);
            Graphics g2 = Graphics.FromImage(tBmp);
            int maxWid = 0;
            if (indexingStart > 0)
            {
                for (int i = 0; i < NumberOfLines; i++)
                {
                    int tw = (int)g2.MeasureString((i + indexingStart).ToString(), indexingFont).Width;
                    if (tw > maxWid)
                        maxWid = tw;
                }
            }
            indBlockWid = maxWid + indBlockXSpacing + 15;


            tBmp = new Bitmap(indBlockWid + blockColumn * (circleSize + X_bubbleSpace) - X_bubbleSpace + 1, blockRow * (circleSize + Y_bubbleSpace) - Y_bubbleSpace + 1);

            g2 = Graphics.FromImage(tBmp);

            g2.Clear(Color.White);
            g2.SmoothingMode =
            SmoothingMode.AntiAlias;
            g2.TextRenderingHint =
                TextRenderingHint.AntiAlias;
            g2.InterpolationMode =
                InterpolationMode.HighQualityBicubic;

            g2.CompositingMode = CompositingMode.SourceOver;

            //g2.DrawRectangle(new Pen(brush, lineWidth * 2), new Rectangle(indBlockWid, 0, tBmp.Width - indBlockWid, tBmp.Height));

            //Only used in case of boxes
            int boxWid = (int)Math.Round((double)(tBmp.Width - indBlockWid) / NumberOfChoices);
            double lineToMargin = 0.6;
            int lineMargin = (int)Math.Round(boxWid * (1 - lineToMargin)) / 2;
            int boxHei = (int)Math.Round((double)tBmp.Height / NumberOfLines);
            int lineHei = (int)Math.Round(boxHei * lineToMargin);

            for (int j = 0; j < NumberOfLines; j++)
            {

                for (int i = 0; i < NumberOfChoices; i++)
                {

                    g2.DrawEllipse(new Pen(brush, lineWidth),
                       new Rectangle(indBlockWid + (circleSize + X_bubbleSpace) * i, (circleSize + Y_bubbleSpace) * j, circleSize, circleSize));
                    if (OptionChar != 0 && OptionChar != 'N')
                    {
                        string num = ((char)(OptionChar + i)).ToString();
                       
                        int numX = (circleSize - (int)g2.MeasureString(num, labelFont).Width) / 2;
                        Point numLoc = new Point(indBlockWid + (circleSize + X_bubbleSpace) * i + numX, (circleSize + Y_bubbleSpace) * j + numX);
                        if (OptionChar == 'a' || OptionChar == '1' || OptionChar == '0')
                            numLoc = new Point(numLoc.X, numLoc.Y - circleSize / 15);
                        g2.DrawString(num, labelFont, brush, numLoc);
                    }

                }
                if (indexingStart > 0)
                {
                    string ansNum = (indexingStart + j).ToString();
                    int length = (int)g2.MeasureString(ansNum, indexingFont).Width;
                    int indX = (circleSize - (int)g2.MeasureString(ansNum, indexingFont).Width) / 2;
                    int indTextY = (circleSize + Y_bubbleSpace) * j + indX;
                    g2.DrawString(ansNum, indexingFont, brush, new Point(indX, indTextY));
                }
            }



            g2.Dispose();
            tBmp.MakeTransparent(Color.White);
            return tBmp;
        }

        public void SaveBitmap(System.Windows.Forms.Panel CtrlToSave, string fileName)
        {
            Point oldPosition = new Point(this.HorizontalScroll.Value, this.VerticalScroll.Value);

            CtrlToSave.PerformLayout();

            ComposedImage ci = new ComposedImage(new Size(CtrlToSave.DisplayRectangle.Width, CtrlToSave.DisplayRectangle.Height));

            int visibleWidth = CtrlToSave.Width - (CtrlToSave.VerticalScroll.Visible ? SystemInformation.VerticalScrollBarWidth : 0);
            int visibleHeightBuffer = CtrlToSave.Height - (CtrlToSave.HorizontalScroll.Visible ? SystemInformation.HorizontalScrollBarHeight : 0);

            //int Iteration = 0;

            for (int x = CtrlToSave.DisplayRectangle.Width - visibleWidth; x >= 0; x -= visibleWidth)
            {

                int visibleHeight = visibleHeightBuffer;

                for (int y = CtrlToSave.DisplayRectangle.Height - visibleHeight; y >= 0; y -= visibleHeight)
                {
                    CtrlToSave.HorizontalScroll.Value = x;
                    CtrlToSave.VerticalScroll.Value = y;

                    CtrlToSave.PerformLayout();

                    Bitmap bmp = new Bitmap(visibleWidth, visibleHeight);

                    CtrlToSave.DrawToBitmap(bmp, new Rectangle(0, 0, visibleWidth, visibleHeight));
                    ci.images.Add(new ImagePart(new Point(x, y), bmp));

                    ///Show image parts
                    //using (Graphics grD = Graphics.FromImage(bmp))
                    //{
                    //    Iteration++;
                    //    grD.DrawRectangle(Pens.Blue,new Rectangle(0,0,bmp.Width-1,bmp.Height-1));
                    //grD.DrawString("x:"+x+",y:"+y+",W:"+visibleWidth+",H:"+visibleHeight + " I:" + Iteration,new Font("Segoe UI",9f),Brushes.Red,new Point(2,2));                        
                    //}

                    if (y - visibleHeight < (CtrlToSave.DisplayRectangle.Height % visibleHeight))
                        visibleHeight = CtrlToSave.DisplayRectangle.Height % visibleHeight;

                    if (visibleHeight == 0)
                        break;
                }

                if (x - visibleWidth < (CtrlToSave.DisplayRectangle.Width % visibleWidth))
                    visibleWidth = CtrlToSave.DisplayRectangle.Width % visibleWidth;
                if (visibleWidth == 0)
                    break;
            }

            Bitmap img = ci.composeImage();
            img.Save(fileName, System.Drawing.Imaging.ImageFormat.Bmp);

            CtrlToSave.HorizontalScroll.Value = oldPosition.X;
            CtrlToSave.VerticalScroll.Value = oldPosition.Y;
        }
        
    }



    public class ComposedImage
    {
        public Size dimensions;
        public List<ImagePart> images;

        public ComposedImage(Size dimensions)
        {
            this.dimensions = dimensions;
            this.images = new List<ImagePart>();
        }

        public ComposedImage(Size dimensions, List<ImagePart> images)
        {
            this.dimensions = dimensions;
            this.images = images;
        }

        public Bitmap composeImage()
        {
            if (dimensions == null || images == null)
                return null;

            Bitmap fullbmp = new Bitmap(dimensions.Width, dimensions.Height);
            using (Graphics grD = Graphics.FromImage(fullbmp))
            {
                grD.SmoothingMode = SmoothingMode.AntiAlias;
                grD.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                grD.InterpolationMode = InterpolationMode.HighQualityBilinear;
                foreach (ImagePart bmp in images)
                {
                    grD.DrawImage(bmp.image, bmp.location.X, bmp.location.Y);
                }
            }
            return fullbmp;
        }
    }

    public class ImagePart
    {
        public Point location;
        public Bitmap image;

        public ImagePart(Point location, Bitmap image)
        {
            this.location = location;
            this.image = image;
        }
    }

}
