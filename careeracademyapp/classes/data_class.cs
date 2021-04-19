using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace omrapplication
{
    class data_class
    {
        private string connection;
        SqlConnection con;
        public SqlDataAdapter da;

        public data_class()
        {
            this.connection = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\database\database.mdf;Integrated Security=True;Connect Timeout=30";
        }

        public bool open_connection()
        {
            if (this.connection.Length < 1)
            {
                throw new Exception("App not initialized correctly");
            }
            else
            {
                con = new SqlConnection(this.connection);
                try
                {
                    con.Open();
                }
                catch (SqlException e)
                {
                    throw new Exception("Error in connection", e);
                }
                return true;
            }
        }
        public void close_connection()
        {
            if (this.con.State == ConnectionState.Open)
                this.con.Close();
            return;
        }
        public DataTable select_options_all(int id)
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();
            try
            {
                query = "SELECT * FROM options WHERE template_id = @id";
                cmd.CommandText = query;
                cmd.Parameters.AddWithValue("@id", id);

                da = new SqlDataAdapter(cmd);
                da.Fill(s);
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }
        public DataTable select_template()
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();
            try
            {
                query = "SELECT template_id, template_name from template where type = 'O'";
                cmd.CommandText = query;

                SqlDataAdapter r = new SqlDataAdapter(cmd);
                r.Fill(s);
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }

        public DataTable select_custom_template()
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();
            try
            {
                query = "SELECT template_id, template_name from template where type = 'S'";
                cmd.CommandText = query;

                SqlDataAdapter r = new SqlDataAdapter(cmd);
                r.Fill(s);
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }
        public string select_template_data(int id, string option_name)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "select " + option_name + " from template where template_id = @id";
            NewCmd.Parameters.AddWithValue("@id", id);
            try
            {
                string a = NewCmd.ExecuteScalar().ToString();
                return a;
            }
            catch (Exception e)
            {
                MessageBox.Show("Error Inserting data : " + e.Message, "Note",
    MessageBoxButtons.OK,
    MessageBoxIcon.Error,
    MessageBoxDefaultButton.Button1);
                return "";
            }
        }
        public string select_template_option_data(int id, string option_name, string name, string data)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "select " + data + " from template_options where template_id = @id AND option_name = @option AND name = @name";
            NewCmd.Parameters.AddWithValue("@id", id);
            NewCmd.Parameters.AddWithValue("@option", option_name);
            NewCmd.Parameters.AddWithValue("@name", name);
            try
            {
                string a = NewCmd.ExecuteScalar().ToString();
                return a;
            }
            catch (Exception e)
            {
                MessageBox.Show("Error Inserting data : " + e.Message, "Note",
    MessageBoxButtons.OK,
    MessageBoxIcon.Error,
    MessageBoxDefaultButton.Button1);
                return "";
            }
        }
        public DataTable select_template_options(int id, string option_name)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();

            NewCmd.CommandText = "select * from template_options where template_id = @id AND option_name = @option";
            NewCmd.Parameters.AddWithValue("@id", id);
            NewCmd.Parameters.AddWithValue("@option", option_name);
            try
            {

                SqlDataAdapter r = new SqlDataAdapter(NewCmd);
                r.Fill(s);
                return s;
            }
            catch (Exception e)
            {
                MessageBox.Show("Error Inserting data : " + e.Message, "Note",
    MessageBoxButtons.OK,
    MessageBoxIcon.Error,
    MessageBoxDefaultButton.Button1);
                return s;
            }
        }
        public void delete_template(int id)
        {
            SqlTransaction transaction;
            SqlCommand command = this.con.CreateCommand();
            // Start a local transaction.
            transaction = this.con.BeginTransaction();

            // Must assign both transaction object and connection
            // to Command object for a pending local transaction
            command.Connection = this.con;
            command.Transaction = transaction;

            try
            {

                command.CommandText = "delete from template where template_id = @id";
                command.Parameters.AddWithValue("@id", id);

                command.ExecuteNonQuery();

                command.CommandText = "delete from template_options where template_id = @id1";
                command.Parameters.AddWithValue("@id1", id);
                command.ExecuteNonQuery();

                command.CommandText = "delete from options where template_id = @id2";
                command.Parameters.AddWithValue("@id2", id);
                command.ExecuteNonQuery();

                command.CommandText = "delete from batch where template_id = @id3";
                command.Parameters.AddWithValue("@id3", id);
                command.ExecuteNonQuery();

                command.CommandText = "delete from batch_options where template_id = @id4";
                command.Parameters.AddWithValue("@id4", id);
                command.ExecuteNonQuery();
                


                // Attempt to commit the transaction.
                transaction.Commit();
            }
            catch (Exception e)
            {
                transaction.Rollback();
                MessageBox.Show(e.Message, "Note",
    MessageBoxButtons.OK,
    MessageBoxIcon.Error,
    MessageBoxDefaultButton.Button1);
            }
        }


        public int create_template(string name, char opt)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "insert into template (template_name, type) OUTPUT INSERTED.IDENTITYCOL values(@temp_name, @type)";
            NewCmd.Parameters.AddWithValue("@temp_name", name);
            NewCmd.Parameters.AddWithValue("@type", opt);
            try
            {
                int a = (int)NewCmd.ExecuteScalar();
                if (a > 0)
                    return a;
                else return 0;
            }
            catch (Exception)
            {
                throw new Exception("Cannot insert duplicate template for same name.");
            }
        }
        public void update_template(int id, int width, int height)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "update template set width = @width, height = @height where template_id = @id";
            NewCmd.Parameters.AddWithValue("@width", width);
            NewCmd.Parameters.AddWithValue("@height", height);
            NewCmd.Parameters.AddWithValue("@id", id);
            try
            {
                int a = NewCmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                MessageBox.Show("Error Inserting data : " + e.Message, "Note",
    MessageBoxButtons.OK,
    MessageBoxIcon.Error,
    MessageBoxDefaultButton.Button1);
            }
        }
        public void insert_new_option(int id, string option, string name)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "insert into template_options (template_id, option_name, name) values (@id, @option_name, @name)";
            NewCmd.Parameters.AddWithValue("@option_name", option);
            NewCmd.Parameters.AddWithValue("@name", name);
            NewCmd.Parameters.AddWithValue("@id", id);
            try
            {
                int a = NewCmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                MessageBox.Show("Error Inserting data : " + e.Message, "Note",
    MessageBoxButtons.OK,
    MessageBoxIcon.Error,
    MessageBoxDefaultButton.Button1);
            }
        }
        public void update_option(int id, string option_name, string name, string data_name, string data_value)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "update template_options set " + data_name + " = @data where template_id = @id AND name = @name AND option_name=@option_name";
            NewCmd.Parameters.AddWithValue("@data", data_value);
            NewCmd.Parameters.AddWithValue("@name", name);
            NewCmd.Parameters.AddWithValue("@option_name", option_name);
            NewCmd.Parameters.AddWithValue("@id", id);
            try
            {
                int a = NewCmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                MessageBox.Show("Error Inserting data : " + e.Message, "Note",
    MessageBoxButtons.OK,
    MessageBoxIcon.Error,
    MessageBoxDefaultButton.Button1);
            }
        }

        public void delete_option(int id, string option_name)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "delete from template_options where template_id = @id AND option_name = @option_name";
            NewCmd.Parameters.AddWithValue("@option_name", option_name);
            NewCmd.Parameters.AddWithValue("@id", id);
            try
            {
                NewCmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                MessageBox.Show("Error Inserting data : " + e.Message, "Note",
    MessageBoxButtons.OK,
    MessageBoxIcon.Error,
    MessageBoxDefaultButton.Button1);
            }
        }

        public void delete_question_option(int id, string option_name, string name)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "delete from template_options where template_id = @id AND option_name = @option_name AND name = @name";
            NewCmd.Parameters.AddWithValue("@option_name", option_name);
            NewCmd.Parameters.AddWithValue("@name", name);
            NewCmd.Parameters.AddWithValue("@id", id);
            try
            {
                NewCmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                MessageBox.Show("Error Inserting data : " + e.Message, "Note",
    MessageBoxButtons.OK,
    MessageBoxIcon.Error,
    MessageBoxDefaultButton.Button1);
            }
        }

        public void save_test(int batch_id, string name, DataTable dt)
        {
            SqlTransaction transaction;
            SqlCommand command = this.con.CreateCommand();
            // Start a local transaction.
            transaction = this.con.BeginTransaction();

            // Must assign both transaction object and connection
            // to Command object for a pending local transaction
            command.Connection = this.con;
            command.Transaction = transaction;

            try
            {

                command.CommandText = "insert into test (test_name, batch_id) OUTPUT INSERTED.IDENTITYCOL values(@name_test, @batch_id)";
                command.Parameters.AddWithValue("@batch_id", batch_id);
                command.Parameters.AddWithValue("@name_test", name);


                int id = (int)command.ExecuteScalar();
                int count = 0;
                foreach(DataRow row in dt.Rows)
                {
                    command.CommandText = "insert into test_details (test_id, name, roll, t_right, t_wrong, t_total, questions, total_marks) values (" + id + ", @name" + count + ", @roll" + count + ", @right_ans" + count + ", @wrong" + count + ", @total" + count + ", @questions" + count + ", @total_marks" + count + ")";
                    
                    command.Parameters.AddWithValue("@name" + count, row[0]);
                    command.Parameters.AddWithValue("@roll" + count, Convert.ToInt32(row[1]));
                    command.Parameters.AddWithValue("@right_ans" + count, Convert.ToInt32(row[2]));
                    command.Parameters.AddWithValue("@wrong" + count, Convert.ToInt32(row[3]));
                    command.Parameters.AddWithValue("@total" + count, Convert.ToInt32(row[4]));
                    command.Parameters.AddWithValue("@questions" + count, row[5]);
                    command.Parameters.AddWithValue("@total_marks" + count, row[6]);
                    command.ExecuteNonQuery();

                    count++;
                }



                // Attempt to commit the transaction.
                transaction.Commit();
            }
            catch (Exception e)
            {
                transaction.Rollback();

                throw new Exception(e.Message);
            }
        }

        public void save_custom_test(int batch_id, string name, DataTable dt)
        {
            SqlTransaction transaction;
            SqlCommand command = this.con.CreateCommand();
            // Start a local transaction.
            transaction = this.con.BeginTransaction();

            // Must assign both transaction object and connection
            // to Command object for a pending local transaction
            command.Connection = this.con;
            command.Transaction = transaction;

            try
            {

                command.CommandText = "insert into test (test_name, batch_id) OUTPUT INSERTED.IDENTITYCOL values(@name_test, @batch_id)";
                command.Parameters.AddWithValue("@batch_id", batch_id);
                command.Parameters.AddWithValue("@name_test", name);


                int id = (int)command.ExecuteScalar();
                int count = 0;
                foreach (DataRow row in dt.Rows)
                {
                    command.CommandText = "insert into test_details (test_id, name, roll, t_total, questions, total_marks) values (" + id + ", @name" + count + ", @roll" + count + ", @total" + count + ", @questions" + count + ", @total_marks" + count + ")";

                    command.Parameters.AddWithValue("@name" + count, row[0]);
                    command.Parameters.AddWithValue("@roll" + count, Convert.ToInt32(row[1]));
                    command.Parameters.AddWithValue("@total" + count, Convert.ToDouble(row[2]));
                    command.Parameters.AddWithValue("@questions" + count, row[3]);
                    command.Parameters.AddWithValue("@total_marks" + count, row[4]);
                    command.ExecuteNonQuery();

                    count++;
                }



                // Attempt to commit the transaction.
                transaction.Commit();
            }
            catch (Exception e)
            {
                transaction.Rollback();
                throw new Exception(e.Message);
            }
        }

        public DataTable select_test()
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();
            try
            {
                query = "SELECT test_id, test_name, batch_id from test where batch_id = (select batch_id from batch where template_id = (select template_id from template where type = 'O'))";
                cmd.CommandText = query;

                SqlDataAdapter r = new SqlDataAdapter(cmd);
                r.Fill(s);
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }
        public DataTable select_custom_test()
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();
            try
            {
                query = "SELECT test_id, test_name, batch_id from test where batch_id = (select batch_id from batch where template_id = (select template_id from template where type = 'S'))";
                cmd.CommandText = query;

                SqlDataAdapter r = new SqlDataAdapter(cmd);
                r.Fill(s);
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }
        public int select_highest_marks(int id)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "select MAX(CAST(t_total AS int)) from test_details where test_id = @id";
            NewCmd.Parameters.AddWithValue("@id", id);
            try
            {
                int a = (int)NewCmd.ExecuteScalar();
                if (a > 0)
                    return a;
                else return 0;
            }
            catch (Exception)
            {
                return 0;
            }
        }
        public int select_average_marks(int id)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "select AVG(CAST(t_total AS int)) from test_details where test_id = @id";
            NewCmd.Parameters.AddWithValue("@id", id);
            try
            {
                int a = (int)NewCmd.ExecuteScalar();
                if (a > 0)
                    return a;
                else return 0;
            }
            catch (Exception)
            {
                return 0;
            }
        }
        

        public void delete_test(int id)
        {
            SqlTransaction transaction;
            SqlCommand command = this.con.CreateCommand();
            // Start a local transaction.
            transaction = this.con.BeginTransaction();

            // Must assign both transaction object and connection
            // to Command object for a pending local transaction
            command.Connection = this.con;
            command.Transaction = transaction;

            try
            {

                command.CommandText = "delete from test where test_id = @id";
                command.Parameters.AddWithValue("@id", id);

                command.ExecuteNonQuery();

                command.CommandText = "delete from test_details where test_id = @id1";
                command.Parameters.AddWithValue("@id1", id);
                command.ExecuteNonQuery();


                // Attempt to commit the transaction.
                transaction.Commit();
            }
            catch (Exception e)
            {
                transaction.Rollback();
                MessageBox.Show(e.Message);
            }
        }

        public DataTable select_test_details(int id)
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();
            try
            {
                query = "SELECT name, roll, t_right, t_wrong, t_total, questions, total_marks from test_details where test_id = @id";
                cmd.CommandText = query;
                cmd.Parameters.AddWithValue("@id", id);

                SqlDataAdapter r = new SqlDataAdapter(cmd);
                r.Fill(s);
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }

        public DataTable select_test_custom_details(int id)
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();
            try
            {
                query = "SELECT name, roll, t_total, questions, total_marks from test_details where test_id = @id";
                cmd.CommandText = query;
                cmd.Parameters.AddWithValue("@id", id);

                SqlDataAdapter r = new SqlDataAdapter(cmd);
                r.Fill(s);
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }

        public DataTable select_subjects()
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();
            try
            {
                query = "SELECT subject_id, subject_name from subjects";
                cmd.CommandText = query;

                SqlDataAdapter r = new SqlDataAdapter(cmd);
                r.Fill(s);
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }
        public string subject_name(int id)
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            string s = "";
            try
            {
                query = "SELECT subject_name from subjects where subject_id = @id";
                cmd.CommandText = query;
                cmd.Parameters.AddWithValue("@id", id);

                s = (cmd.ExecuteScalar()).ToString();
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }

        public DataTable select_topics(int subject_id)
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();
            try
            {
                query = "SELECT topic_id, topic_name from topics where subject_id = @id";
                cmd.CommandText = query;
                cmd.Parameters.AddWithValue("@id", subject_id);

                SqlDataAdapter r = new SqlDataAdapter(cmd);
                r.Fill(s);
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }

        public DataTable select_sub_topics(int topic_id)
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();
            try
            {
                query = "SELECT sub_id, sub_name from sub_topics where topic_id = @id";
                cmd.CommandText = query;
                cmd.Parameters.AddWithValue("@id", topic_id);

                SqlDataAdapter r = new SqlDataAdapter(cmd);
                r.Fill(s);
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }
        
        public int insert_batch(int id, int subject_id, string name)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "insert into batch (template_id, batch_name, subject_id) OUTPUT INSERTED.IDENTITYCOL values(@id, @name, @sub_id)";
            NewCmd.Parameters.AddWithValue("@id", id);
            NewCmd.Parameters.AddWithValue("@name", name);
            NewCmd.Parameters.AddWithValue("@sub_id", subject_id);
            try
            {
                int a = (int)NewCmd.ExecuteScalar();
                if (a > 0)
                    return a;
                else return 0;
            }
            catch (Exception)
            {
                return 0;
            }
        }
        public int check_batch(int id, string name)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "select batch_id from batch where template_id = @id and batch_name = @sub_id";
            NewCmd.Parameters.AddWithValue("@id", id);
            NewCmd.Parameters.AddWithValue("@sub_id", name);
            try
            {
                int a = (int)NewCmd.ExecuteScalar();
                if (a > 0)
                    return a;
                else return 0;
            }
            catch (Exception)
            {
                return 0;
            }
        }


        public int count_batch_options(int id)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "select count(*) from batch_options where batch_id = @id";
            NewCmd.Parameters.AddWithValue("@id", id);
            try
            {
                int a = (int)NewCmd.ExecuteScalar();
                if (a > 0)
                    return a;
                else return 0;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        public int count_batch_custom_options(int id)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "select count(*) from batch_custom_options where batch_id = @id";
            NewCmd.Parameters.AddWithValue("@id", id);
            try
            {
                int a = (int)NewCmd.ExecuteScalar();
                if (a > 0)
                    return a;
                else return 0;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        public void insert_batch_option(int id, int template_id, int name, int answer, int marks, int neg, int topic, int sub_topic)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "insert into batch_options (batch_id, option_number, option_answer, option_marks, option_neg, topic, sub_topic, template_id) values(@id, @name, @answer, @marks, @neg, @topic, @sub_topic, @template_id)";
            NewCmd.Parameters.AddWithValue("@name", name);
            NewCmd.Parameters.AddWithValue("@answer", answer);
            NewCmd.Parameters.AddWithValue("@marks", marks);
            NewCmd.Parameters.AddWithValue("@neg", neg);
            NewCmd.Parameters.AddWithValue("@topic", topic);
            NewCmd.Parameters.AddWithValue("@sub_topic", sub_topic);
            NewCmd.Parameters.AddWithValue("@template_id", template_id);
            NewCmd.Parameters.AddWithValue("@id", id);
            try
            {
                NewCmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                MessageBox.Show("Error Inserting data : " + e.Message);
            }
        }
        public void insert_batch_custom_option(int id, int template_id, int name, int marks, int topic, int sub_topic)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "insert into batch_custom_options (batch_id, option_number, option_marks, topic, sub_topic, template_id) values(@id, @name, @marks, @topic, @sub_topic, @template_id)";
            NewCmd.Parameters.AddWithValue("@name", name);
            NewCmd.Parameters.AddWithValue("@marks", marks);
            NewCmd.Parameters.AddWithValue("@topic", topic);
            NewCmd.Parameters.AddWithValue("@sub_topic", sub_topic);
            NewCmd.Parameters.AddWithValue("@template_id", template_id);
            NewCmd.Parameters.AddWithValue("@id", id);
            try
            {
                NewCmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                MessageBox.Show("Error Inserting data : " + e.Message);
            }
        }


        public void add_template_options(int id)
        {

            DataTable s = new DataTable();

            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "select * from template_options where ( option_name = 'Question' or option_name = 'Custom' ) AND template_id = @id";
            NewCmd.Parameters.AddWithValue("@id", id);
            try
            {
                SqlDataAdapter r = new SqlDataAdapter(NewCmd);
                r.Fill(s);

                foreach (DataRow row in s.Rows)
                {
                    int length = Convert.ToInt32(row["rows"]);
                    int x = Convert.ToInt32(row["x_cor"]);
                    int y = Convert.ToInt32(row["y_cor"]);
                    int x_space = Convert.ToInt32(row["x_space"]);
                    int y_space = Convert.ToInt32(row["y_space"]);
                    int height = Convert.ToInt32(row["radius"]) * 2;
                    int columns = Convert.ToInt32(row["columns"]);
                    for (int i = 0; i < length; i++)
                    {
                        insert_template_option(id, i + 1, x, y, x_space, y_space, height, columns);
                        y += y_space;
                        y += height;
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        public DataTable select_batch()
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();
            try
            {
                query = "SELECT * from batch where template_id = (select template_id from template where type = 'O')";
                cmd.CommandText = query;

                SqlDataAdapter r = new SqlDataAdapter(cmd);
                r.Fill(s);
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }
        public DataTable select_custom_batch()
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();
            try
            {
                query = "SELECT * from batch where template_id = (select template_id from template where type = 'S')";
                cmd.CommandText = query;

                SqlDataAdapter r = new SqlDataAdapter(cmd);
                r.Fill(s);
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }
        public DataTable select_batch(int id)
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();
            try
            {
                query = "SELECT * from batch where template_id = @id";
                cmd.CommandText = query;
                cmd.Parameters.AddWithValue("@id", id);

                SqlDataAdapter r = new SqlDataAdapter(cmd);
                r.Fill(s);
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }

        public int select_subject_id_from_batch(int id)
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            int s = 0;
            try
            {
                query = "SELECT subject_id from batch where batch_id = @id";
                cmd.CommandText = query;
                cmd.Parameters.AddWithValue("@id", id);
                
                s = Convert.ToInt32(cmd.ExecuteScalar());
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }    

        public void delete_batch(int id)
        {
            SqlTransaction transaction;
            SqlCommand command = this.con.CreateCommand();
            // Start a local transaction.
            transaction = this.con.BeginTransaction();

            // Must assign both transaction object and connection
            // to Command object for a pending local transaction
            command.Connection = this.con;
            command.Transaction = transaction;

            try
            {

                command.CommandText = "delete from batch where batch_id = @id";
                command.Parameters.AddWithValue("@id", id);

                command.ExecuteNonQuery();

                command.CommandText = "delete from batch_options where batch_id = @id1";
                command.Parameters.AddWithValue("@id1", id);
                command.ExecuteNonQuery();


                command.CommandText = "delete from batch_custom_options where batch_id = @id2";
                command.Parameters.AddWithValue("@id2", id);
                command.ExecuteNonQuery();


                // Attempt to commit the transaction.
                transaction.Commit();
            }
            catch (Exception e)
            {
                transaction.Rollback();
                MessageBox.Show(e.Message);
            }
        }

        public DataSet select_batch_options(int id)
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataSet s = new DataSet();
            try
            {
                query = "SELECT option_id, option_number, option_answer, option_marks, option_neg, topic, sub_topic FROM batch_options WHERE batch_id = @id";
                cmd.CommandText = query;
                cmd.Parameters.AddWithValue("@id", id);

                da = new SqlDataAdapter(cmd);
                da.Fill(s, "options");
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }
        public DataSet select_batch_custom_options(int id)
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataSet s = new DataSet();
            try
            {
                query = "SELECT option_id, option_number, option_marks, topic, sub_topic FROM batch_custom_options WHERE batch_id = @id";
                cmd.CommandText = query;
                cmd.Parameters.AddWithValue("@id", id);

                da = new SqlDataAdapter(cmd);
                da.Fill(s, "options");
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }

        public DataTable select_batch_options_all(int id)
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();
            try
            {
                query = "SELECT * FROM batch_options WHERE batch_id = @id";
                cmd.CommandText = query;
                cmd.Parameters.AddWithValue("@id", id);

                da = new SqlDataAdapter(cmd);
                da.Fill(s);
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }
        public DataTable select_batch_custom_options_all(int id)
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();
            try
            {
                query = "SELECT * FROM batch_custom_options WHERE batch_id = @id";
                cmd.CommandText = query;
                cmd.Parameters.AddWithValue("@id", id);

                da = new SqlDataAdapter(cmd);
                da.Fill(s);
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }

        private void insert_template_option(int id, int name, int x_cor, int y_cor, int x_space, int y_space, int radius, int columns)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "insert into options (template_id, option_number, option_x, option_y, x_space, y_space, radius, columns) values(@id, @number, @x, @y, @x_space, @y_space, @radius, @columns)";
            NewCmd.Parameters.AddWithValue("@number", name);
            NewCmd.Parameters.AddWithValue("@x", x_cor);
            NewCmd.Parameters.AddWithValue("@y", y_cor);
            NewCmd.Parameters.AddWithValue("@x_space", x_space);
            NewCmd.Parameters.AddWithValue("@y_space", y_space);
            NewCmd.Parameters.AddWithValue("@radius", radius);
            NewCmd.Parameters.AddWithValue("@columns", columns);
            NewCmd.Parameters.AddWithValue("@id", id);
            try
            {
                NewCmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                MessageBox.Show("Error Inserting data : " + e.Message);
            }
        }



        public int insert_subject(String name)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "insert into subjects (subject_name) OUTPUT INSERTED.IDENTITYCOL values(@name)";
            NewCmd.Parameters.AddWithValue("@name", name);
            try
            {
                int a = (int)NewCmd.ExecuteScalar();
                if (a > 0)
                    return a;
                else return 0;
            }
            catch (Exception)
            {
                return 0;
            }
        }
        public void delete_subject(int id)
        {
            SqlTransaction transaction;
            SqlCommand command = this.con.CreateCommand();
            // Start a local transaction.
            transaction = this.con.BeginTransaction();

            // Must assign both transaction object and connection
            // to Command object for a pending local transaction
            command.Connection = this.con;
            command.Transaction = transaction;

            try
            {
                command.CommandText = "delete from batch where subject_id = @id";
                command.Parameters.AddWithValue("@id", id);
                command.ExecuteNonQuery();
                

                command.CommandText = "delete from subjects where subject_id = @id1";
                command.Parameters.AddWithValue("@id1", id);
                command.ExecuteNonQuery();

                command.CommandText = "select topic_id from topics where subject_id = @id_top";
                command.Parameters.AddWithValue("@id_top", id);
                SqlDataAdapter r = new SqlDataAdapter(command);
                DataTable s = new DataTable();
                r.Fill(s);
                int _count = 0;
                foreach (DataRow row in s.Rows)
                {

                    command.CommandText = "delete from sub_topics where topic_id = @id_" + _count;
                    command.Parameters.AddWithValue("@id_" + _count, row["topic_id"]);

                    command.ExecuteNonQuery();
                    _count++;
                }


                command.CommandText = "delete from topics where subject_id = @id2";
                command.Parameters.AddWithValue("@id2", id);
                command.ExecuteNonQuery();


                // Attempt to commit the transaction.
                transaction.Commit();
            }
            catch (Exception e)
            {
                transaction.Rollback();
                MessageBox.Show(e.Message);
            }
        }
        
        public int insert_topic(int topic_id, String name)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "insert into topics (subject_id, topic_name) OUTPUT INSERTED.IDENTITYCOL values(@id, @name)";
            NewCmd.Parameters.AddWithValue("@id", topic_id);
            NewCmd.Parameters.AddWithValue("@name", name);
            try
            {
                int a = (int)NewCmd.ExecuteScalar();
                if (a > 0)
                    return a;
                else return 0;
            }
            catch (Exception)
            {
                return 0;
            }
        }
        public void delete_topic(int id)
        {
            SqlTransaction transaction;
            SqlCommand command = this.con.CreateCommand();
            // Start a local transaction.
            transaction = this.con.BeginTransaction();

            // Must assign both transaction object and connection
            // to Command object for a pending local transaction
            command.Connection = this.con;
            command.Transaction = transaction;

            try
            {
                
                command.CommandText = "delete from batch_options where topic = @id_1";
                command.Parameters.AddWithValue("@id_1",id);

                command.ExecuteNonQuery();

                command.CommandText = "delete from topics where topic_id = @id1";
                command.Parameters.AddWithValue("@id1", id);
                command.ExecuteNonQuery();

                command.CommandText = "delete from sub_topics where topic_id = @id3";
                command.Parameters.AddWithValue("@id3", id);
                command.ExecuteNonQuery();


                // Attempt to commit the transaction.
                transaction.Commit();
            }
            catch (Exception e)
            {
                transaction.Rollback();
                MessageBox.Show(e.Message);
            }
        }

        public DataTable select_topics_all()
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();
            try
            {
                query = "SELECT topic_id, topic_name FROM topics";
                cmd.CommandText = query;

                da = new SqlDataAdapter(cmd);
                da.Fill(s);
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }


        public int insert_sub_topic(int topic_id, String name)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "insert into sub_topics (topic_id, sub_name) OUTPUT INSERTED.IDENTITYCOL values(@id, @name)";
            NewCmd.Parameters.AddWithValue("@id", topic_id);
            NewCmd.Parameters.AddWithValue("@name", name);
            try
            {
                int a = (int)NewCmd.ExecuteScalar();
                if (a > 0)
                    return a;
                else return 0;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        public void delete_sub_topic(int id)
        {
            SqlTransaction transaction;
            SqlCommand command = this.con.CreateCommand();
            // Start a local transaction.
            transaction = this.con.BeginTransaction();

            // Must assign both transaction object and connection
            // to Command object for a pending local transaction
            command.Connection = this.con;
            command.Transaction = transaction;

            try
            {
                command.CommandText = "update batch_options set sub_topic = 0 where sub_topic = @id";
                command.Parameters.AddWithValue("@id", id);
                
                command.ExecuteNonQuery();
         

                command.CommandText = "delete from sub_topics where sub_id = @id3";
                command.Parameters.AddWithValue("@id3", id);
                command.ExecuteNonQuery();


                // Attempt to commit the transaction.
                transaction.Commit();
            }
            catch (Exception e)
            {
                transaction.Rollback();
                MessageBox.Show(e.Message);
            }
        }

        public DataTable select_sub_topics_all()
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();
            try
            {
                query = "SELECT sub_id, sub_name FROM sub_topics";
                cmd.CommandText = query;

                da = new SqlDataAdapter(cmd);
                da.Fill(s);
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }
        public DataTable select_batch_questions_numbers(int id)
        {
            string query = "";
            SqlCommand cmd = this.con.CreateCommand();
            cmd.Connection = this.con;
            cmd.CommandType = CommandType.Text;
            DataTable s = new DataTable();
            try
            {
                query = "SELECT option_number, topic, sub_topic FROM batch_options where batch_id = @id";
                cmd.CommandText = query;
                cmd.Parameters.AddWithValue("@id", id);

                da = new SqlDataAdapter(cmd);
                da.Fill(s);
                return s;
            }
            catch (Exception)
            {
                return s;
            }
        }

        public string get_topic(int id, int number)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "select topic_name from topics where topic_id = @id";
            NewCmd.Parameters.AddWithValue("@id", id);
            try
            {
                string  a = NewCmd.ExecuteScalar().ToString();
                if (a != "" || a != null)
                    return a;
                else return "None";
            }
            catch (Exception)
            {
                return "";
            }
        }
        public string get_sub_topic(int id, int number)
        {
            SqlCommand NewCmd = this.con.CreateCommand();
            NewCmd.Connection = this.con;
            NewCmd.CommandType = CommandType.Text;
            NewCmd.CommandText = "select sub_name from sub_topics where sub_id = @id";
            NewCmd.Parameters.AddWithValue("@id", id);
            try
            {
                string a = NewCmd.ExecuteScalar().ToString();
                if (a != "" || a != null)
                    return a;
                else return "None";
            }
            catch (Exception)
            {
                return "";
            }
        }


        public void ExportToExcel(DataGridView gridviewID, string excelFilename)
        {
            Microsoft.Office.Interop.Excel.Application objexcelapp = new Microsoft.Office.Interop.Excel.Application();
            objexcelapp.Application.Workbooks.Add(Type.Missing);
            objexcelapp.Columns.ColumnWidth = 25;
            for (int i = 1; i < gridviewID.Columns.Count + 1; i++)
            {
                objexcelapp.Cells[1, i] = gridviewID.Columns[i - 1].HeaderText;
            }
            /*For storing Each row and column value to excel sheet*/
            for (int i = 0; i < gridviewID.Rows.Count; i++)
            {
                for (int j = 0; j < gridviewID.Columns.Count; j++)
                {
                    if (gridviewID.Rows[i].Cells[j].Value != null)
                    {
                        objexcelapp.Cells[i + 2, j + 1] = gridviewID.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
            objexcelapp.ActiveWorkbook.SaveAs(excelFilename);
            objexcelapp.ActiveWorkbook.Saved = true;
            objexcelapp.ActiveWorkbook.Close();
        }
       
    }
}
