using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace CS01
{
    class DBHelper
    {

        //private static readonly DBHelper dbhelper = new DBHelper();

        private MySqlConnection connection;
        public DBConn _dbCon;

        public DBHelper(DBConn conStr)
        {
            this._dbCon = conStr;

            MySqlConnectionStringBuilder b = new MySqlConnectionStringBuilder();
            b.Server = conStr.Host;
            b.Port = (uint)conStr.Port;
            b.UserID = conStr.Username;
            b.Password = conStr.Password;
            b.Database = conStr.Dbname;


            connection = new MySqlConnection(b.ToString());
        }

        public bool openConnection()
        {
            try
            {
                connection.Open();
                Console.WriteLine("DB: Connected");
                return true;
            }
            catch (MySqlException ex)
            {
                switch (ex.Number)
                {
                    case 0:
                        Console.WriteLine("ERROR: Cannot connect to server");
                        break;
                    case 1045:
                        Console.WriteLine("ERROR: Invalid username and password");
                        break;
                }
                return false;
            }
        }

        public void closeConnection()
        {
            try
            {
                connection.Close();
                Console.WriteLine("DB: Connection Close");
            }
            catch (MySqlException ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        // CRUD
        public void SelectAll()
        {

            string sql = "SELECT * FROM tbl_products";

            // open connection
            if (this.openConnection())
            {
                // Create Command
                MySqlCommand cmd = new MySqlCommand(sql, connection);
                // Create Data Reader
                MySqlDataReader dataReader = cmd.ExecuteReader();

                // Read data from first row
                while (dataReader.Read())
                {
                    Console.WriteLine("ID:{0} | Name:{1} | Size:{2} | Image:{3} ",
                        dataReader["id"],
                        dataReader["name"],
                        dataReader["size"],
                        dataReader["image"]);
                }

                // Close Data Reader
                dataReader.Close();
                // Close Connection
                this.closeConnection();
            }
            else
            {
                Console.WriteLine("WALA HAHAHAH XD");
            }
        }

        public bool Insert(string[] data)
        {
            try
            {
                if (this.openConnection())
                {
                    string sql = "INSERT INTO tbl_comelec (created_at, updated_at, voters_name, voters_address, bday, baranggay, city) VALUES (NOW(), NOW(), ";
                    for (int l = 0; l < data.Length; l++)
                    {
                        sql += "'" + data[l] + "'";
                        if (l <= 3)
                        {
                            sql += ", ";
                        }
                    }

                    sql += ")";

                    MySqlCommand cmd = new MySqlCommand(sql, connection);
                    // execute
                    cmd.ExecuteNonQuery();

                    this.closeConnection();

                    Console.WriteLine("Insert");

                    return true;
                }
                else
                {
                    Console.WriteLine("Connection Failed");
                    return false;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: "+ e.Message);
                return false;
            }            
        }

    }
}
