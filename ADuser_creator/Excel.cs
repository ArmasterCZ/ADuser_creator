using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADuser_creator
{
    class Excel
    {
        public Excel(string path)
        {
            this.path = path;
            if (!File.Exists(path))
            {
                throw new System.ArgumentException("Nepodařilo se soubor. /n " + path, "Excel Error");
            }
        }

        private string path;

        private DataSet tempTable;

        private string getConnectionString()
        {
            //create string for connection to excel file

            Dictionary<string, string> props = new Dictionary<string, string>();
            if (path == null | path == "" | path == "...")
            {
                throw new System.ArgumentException("Chybný typ cesty." + path);
            }
            else
            {
                //check if path exist
                if (!File.Exists(path))
                {
                    throw new System.ArgumentException("Nepodařilo se najít soubor." + path);
                }

            }

            // XLSX - Excel 2007, 2010, 2012, 2013
            props["Provider"] = "Microsoft.ACE.OLEDB.12.0";
            props["Extended Properties"] = "Excel 12.0 XML";
            props["Data Source"] = path; //"C:\\info.xlsx";

            // XLS - Excel 2003 and Older
            //props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
            //props["Extended Properties"] = "Excel 8.0";
            //props["Data Source"] = "C:\\info.xls";

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

        private DataSet getTable()
        {
            //load excel file to table

            DataSet ds = new DataSet();
            string connectionString = getConnectionString();
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    if (!sheetName.EndsWith("$"))
                        continue;

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    DataTable dt = new DataTable();
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    ds.Tables.Add(dt);
                }

                cmd = null;
                conn.Close();
            }
            if (ds.Tables.Count == 0)
            {
                throw new System.ArgumentException("Nepodařilo se načíst excel." + connectionString.ToString(), "Excel Error");
            }
            return ds;
        }

        private string getCompare(string toFindString, string column)
        {
            //search row (in tempTable) with "toFindString" and return string from "column"
            string finalString = "";
            DataRowCollection rowsAll = tempTable.Tables[0].Rows;
            foreach (DataRow row1 in rowsAll)
            {
                string columnOffice = row1["OfficeName"].ToString();
                string columnSearched = row1[column].ToString();
                if (columnOffice.Contains(toFindString))
                {
                    finalString = columnSearched;
                    break;
                }
            }
            return finalString;
        }

        public ADuser readFileRow(int rowNumber)
        {
            //load line from excel file
            ADuser user1 = new ADuser();
            DataSet ds = getTable();

            string sheetName = "List1$";
            if (ds.Tables.Contains(sheetName))
            {
                DataTable table1 = ds.Tables[sheetName];
                int table1RowsNumber = table1.Rows.Count;
                DataRow row = table1.Rows[rowNumber];

                string nameGiven = "";
                string nameSurn = "";
                string password = "";
                string emailAddress = "";
                string office = "";
                string department = "";
                string tel = "";
                string mob = "";
                string description = "";
                string title = "";
                string telOthers = "";
                string cardNumber = "";
                string manager = "";
                string path = "";
                string group = "";
                bool ChangePasswordAtLogon = false;
                bool CannotChangePassword = false;
                bool PasswordNeverExpires = false;
                bool Enabled = false;

                //load data (in rowNumber) by name of column
                string nameAcco = row["login"].ToString();
                string nameFull = row["jméno"].ToString();
                try { nameGiven = row["křestní"].ToString(); } catch { }
                try { nameSurn = row["příjmení"].ToString(); } catch { }
                try { password = row["heslo"].ToString(); } catch { }
                try { emailAddress = row["email"].ToString(); } catch { }
                try { office = row["kancelář"].ToString(); } catch { }
                try { department = row["středisko"].ToString(); } catch { }
                try { title = row["title"].ToString(); } catch { }
                try { description = row["description"].ToString(); } catch { }
                try { tel = row["tel"].ToString(); } catch { }
                try { mob = row["mob"].ToString(); } catch { }
                try { telOthers = row["otherTelephone"].ToString(); } catch { }
                try { cardNumber = row["karta"].ToString(); } catch { }
                try { manager = row["manager"].ToString(); } catch { }
                try { path = row["path"].ToString(); } catch { }
                try { group = row["group"].ToString(); } catch { }
                try { ChangePasswordAtLogon = Convert.ToBoolean(row["ChangePasswordAtLogon"].ToString()); } catch { }
                try { CannotChangePassword = Convert.ToBoolean(row["CannotChangePassword"].ToString()); } catch { }
                try { PasswordNeverExpires = Convert.ToBoolean(row["PasswordNeverExpires"].ToString()); } catch { }
                try { Enabled = Convert.ToBoolean(row["Enabled"].ToString()); } catch { }


                //create instance of ADuser
                user1.nameGiven = nameGiven;
                user1.nameSurn = nameSurn;
                user1.nameFull = nameFull;
                user1.nameAcco = nameAcco;
                user1.password = password;
                user1.emailAddress = emailAddress;
                user1.office = office;
                user1.department = department;
                user1.tel = tel;
                user1.mob = mob;
                user1.description = description;
                user1.title = title;
                user1.telOthers = telOthers;
                user1.cardNumber = cardNumber;
                user1.manager = manager;
                user1.path = path;
                user1.group = group;
                user1.ChangePasswordAtLogon = ChangePasswordAtLogon;
                user1.CannotChangePassword = CannotChangePassword;
                user1.PasswordNeverExpires = PasswordNeverExpires;
                user1.Enabled = Enabled;
            }
            else
            {
                throw new System.ArgumentException("Nepodařilo se najít záložku. " + sheetName, "Excel Error");
            }
            return user1;
            //method end
        }

        public string loadAndCompare(string toFindString, string column)
        {
            //load table from excel to local tempTable and compare input string in getCompare
            string finalString = "";

            if (tempTable == null)
            {
                tempTable = getTable();
            }
            finalString = getCompare(toFindString, column);

            return finalString;
        }

    }
}


