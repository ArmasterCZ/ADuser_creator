using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.IO;
using System.Data.Odbc;
using System.Data.OleDb;
using ExtensionMethods;
using System.Globalization;

namespace ADuser_creator
{
    public partial class Form1 : Form
    {
        public string verze = "0.00.44";
        public Form1()
        {
            InitializeComponent();
            //show table
            dataGridView1_ResetTable();
            //show actual version
            this.Text = "ADuser Creator V" + verze;
            //path for automatic load
            ts_ExTextBoxPath.Text = string.Format("{0}\\{1}", Directory.GetCurrentDirectory(), "info.xlsx");
            //path for automatic completion
            ts_ExTextBoxPathAC.Text = string.Format("{0}\\{1}", Directory.GetCurrentDirectory(), "Office.xlsx");
        }


        //TODO: 
        //create method for Compare User (and PS script to change only diferent item)
        //feature: add log (in text file)
        //feature: show powershell script, first line date and time.
        //bug: ctrl+v (get nuber of row from datagridrow. Bad when sort happend)
        //bug: check sort (add id column ? = for reorder back without data delete?)
        //bug: for automatic completion is enought part of string (wrong?)
        //better way of store excel paths "ts_ExTextBoxPath" & "ts_ExTextBoxPathAC" (right click menu with file names?)
        //incease size of table when ... max or increase window size.
        //Automatic replenishment for Path from outside source. excel | app.config

        /*Version info
         0.00.01 - Basic design idea
         0.00.05 - method for create new AD user
         0.00.23 - more parameters in ADuser class
         0.00.24 - code cleanup 1/2
         0.00.25 - custom patch for excel load
         0.00.27 - added move user function
         0.00.28 - add metod addQuotes, PS_SearchUser_Identity
         0.00.29 - more ADuser parameters (otherTelephone, EmailAddress, manager)
         0.00.30 - better PS script for write user
         0.00.31 - bug fixed, additional parameters to load from excel
         0.00.32 - automatic replenishment of the table
         0.00.33 - bugfix: copy to cell path only container names. delete name of profile. 
         0.00.34 - added right click menu for container move
         0.00.35 - remove diacritic from username (in autocomplete), added split from full name to first and second cell.
         0.00.36 - added icons
         0.00.37 - project rename. "PowerShell2_CreateADuserFromExcel" -> "ADuser_creator"
         0.00.38 - redesign 10% (form)
         0.00.39 - redesign 80% (functions) add clear for column if user was not found
         0.00.40 - redesign 100%(excel and ctrl+V)
         0.00.41 - excel class
         0.00.42 - group integration (ADuser [ADgroup], PS script [PS_GetUserGroups, PS_AddUserToGroups],table, + copy functions)
         0.00.43 - excel class redesign, add automatic completion for department,description,title,manager,group (from office), added cloneToTable method
         0.00.44 - BugHunt: (search group for non existional user) => crash. + add Automatic replenishment for Path.
         */

        #region junkFromDesign_0.00.37

        /*private void PS_EditExistUser(ADuser user1)
        {
            //pokusí se vytvořit uživatele 1 část (bez hesla a disablovaného)
            if (user1.nameAcco != "")
            {
                MessageBox.Show("bude vytvořen nový uživatel " + user1.QnameAcco);
                //TODO: check all data here

                using (var runspace = RunspaceFactory.CreateRunspace())
                {
                    using (var powerShell = PowerShell.Create())
                    {
                        log("Zápis uživatele do AD");
                        powerShell.Runspace = runspace;
                        powerShell.Runspace.Open();

                        //vkládání proměných
                        powerShell.AddScript("$nameAccout = " + user1.QnameAcco);
                        powerShell.AddScript("$nameGiven = " + user1.QnameGiven);
                        powerShell.AddScript("$nameSurn = " + user1.QnameSurn);
                        powerShell.AddScript("$nameFull = " + user1.QnameFull);
                        powerShell.AddScript("$office = " + user1.Qoffice);
                        powerShell.AddScript("$description = " + user1.Qdescription);
                        //powerShell.AddScript("$department = "" );
                        //powerShell.AddScript("$vedouci =  "" );
                        powerShell.AddScript("$password = " + user1.Qpassword);
                        powerShell.AddScript("$tel = " + user1.Qtel);
                        powerShell.AddScript("$mob = " + user1.Qmob);
                        powerShell.AddScript("$mob = " + user1.Qmob);
                        //nezapisuje kartu
                        //powerShell.AddScript("$pager" + @"""k""");
                        //powerShell.AddScript("$otherPager" + user1.cardNumber);

                        //založi uživatele
                        powerShell.AddScript("New-ADUser " +
                            "-name $nameFull " +
                            "-displayName $nameFull " +
                            "-sAMAccountName $nameAccout " +
                            "-givenName $nameGiven " +
                            "-Surname $nameSurn " +
                            "-title $description " +
                            "-department $department " +
                            "-Description $description " +
                            "-Office $office " +
                            "-officephone $tel " +
                            "-mobile $mob ");

                        //nastaví heslo, enabluje uživatele, nepotřebuje měnit heslo.
                        powerShell.AddScript("Set-ADUser -Identity $nameAccout -ChangePasswordAtLogon $false");
                        powerShell.AddScript("Set-ADAccountPassword -Identity $nameAccout -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $password -Force)");
                        powerShell.AddScript("Enable-ADAccount -Identity $nameAccout");

                        powerShell.Invoke();

                        if (powerShell.HadErrors == true)
                        {
                            log("Error. Chyba při vytváření uživatele.");
                            MessageBox.Show("Error. Chyba při vytváření uživatele.");
                            var test = powerShell.Streams.Error.ElementAt(0).Exception.Message;
                            MessageBox.Show("" + test);
                        }
                        else
                        {
                            log("Hotovo. Uživatel vytvořen.");
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Error. Nebylo možné vytvořit uživatele. Prázdné jméno!");
                log("Error. Nebylo možné vytvořit uživatele. Prázdné jméno!");
            }
        }*/

        /*private void PS_CompareUser(ADuser user1)
        {
            //pokusí se vytvořit uživatele 1 část (bez hesla a disablovaného)
            if (user1.nameAcco != "")
            {
                MessageBox.Show("bude vytvořen nový uživatel " + user1.QnameAcco);
                //TODO: check all data here

                using (var runspace = RunspaceFactory.CreateRunspace())
                {
                    using (var powerShell = PowerShell.Create())
                    {
                        log("Zápis uživatele do AD");
                        powerShell.Runspace = runspace;
                        powerShell.Runspace.Open();

                        //vkládání proměných
                        powerShell.AddScript("$nameAccout = " + user1.QnameAcco);
                        powerShell.AddScript("$nameGiven = " + user1.QnameGiven);
                        powerShell.AddScript("$nameSurn = " + user1.QnameSurn);
                        powerShell.AddScript("$nameFull = " + user1.QnameFull);
                        powerShell.AddScript("$office = " + user1.Qoffice);
                        powerShell.AddScript("$description = " + user1.Qdescription);
                        //powerShell.AddScript("$department = "" );
                        //powerShell.AddScript("$vedouci =  "" );
                        powerShell.AddScript("$password = " + user1.Qpassword);
                        powerShell.AddScript("$tel = " + user1.Qtel);
                        powerShell.AddScript("$mob = " + user1.Qmob);
                        powerShell.AddScript("$mob = " + user1.Qmob);
                        //nezapisuje kartu
                        //powerShell.AddScript("$pager" + @"""k""");
                        //powerShell.AddScript("$otherPager" + user1.cardNumber);

                        //založi uživatele
                        powerShell.AddScript("New-ADUser " +
                            "-name $nameFull " +
                            "-displayName $nameFull " +
                            "-sAMAccountName $nameAccout " +
                            "-givenName $nameGiven " +
                            "-Surname $nameSurn " +
                            "-title $description " +
                            "-department $department " +
                            "-Description $description " +
                            "-Office $office " +
                            "-officephone $tel " +
                            "-mobile $mob ");

                        //nastaví heslo, enabluje uživatele, nepotřebuje měnit heslo.
                        powerShell.AddScript("Set-ADUser -Identity $nameAccout -ChangePasswordAtLogon $false");
                        powerShell.AddScript("Set-ADAccountPassword -Identity $nameAccout -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $password -Force)");
                        powerShell.AddScript("Enable-ADAccount -Identity $nameAccout");

                        powerShell.Invoke();

                        if (powerShell.HadErrors == true)
                        {
                            log("Error. Chyba při vytváření uživatele.");
                            MessageBox.Show("Error. Chyba při vytváření uživatele.");
                            var test = powerShell.Streams.Error.ElementAt(0).Exception.Message;
                            MessageBox.Show("" + test);
                            /*foreach (var error in PowerShell.Streams.Error)
                            {
                            }//*
                        }
                        else
                        {
                            log("Hotovo. Uživatel vytvořen.");
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Error. Nebylo možné vytvořit uživatele. Prázdné jméno!");
                log("Error. Nebylo možné vytvořit uživatele. Prázdné jméno!");
            }
        }*/

        /*private void dataTableInicial()
        {
            dataTable1 = new DataTable();
            //table.Columns.Add("ID", typeof(int));
            //table.Columns.Add("Čas", typeof(DateTime));

            dataTable1.Columns.Add("-zdroj-", typeof(string));
            dataTable1.Columns.Add("Křestní", typeof(string));
            dataTable1.Columns.Add("Příjmení", typeof(string));
            dataTable1.Columns.Add("Plné jméno", typeof(string));
            dataTable1.Columns.Add("Username", typeof(string));
            dataTable1.Columns.Add("Kancelář", typeof(string));
            dataTable1.Columns.Add("Středisko", typeof(string));
            dataTable1.Columns.Add("Tel", typeof(string));
            dataTable1.Columns.Add("TelOstatní", typeof(string));
            dataTable1.Columns.Add("Mob", typeof(string));
            dataTable1.Columns.Add("Popis", typeof(string));
            dataTable1.Columns.Add("Pozice", typeof(string));
            dataTable1.Columns.Add("Karta", typeof(string));
            dataTable1.Columns.Add("Email", typeof(string));
            dataTable1.Columns.Add("Manager", typeof(string));
            dataTable1.Columns.Add("Heslo", typeof(string));
            dataTable1.Columns.Add("Path", typeof(string));
            dataTable1.Columns.Add("ChangePasswordAtLogon", typeof(bool));
            dataTable1.Columns.Add("CannotChangePassword", typeof(bool));
            dataTable1.Columns.Add("PasswordNeverExpires", typeof(bool));
            dataTable1.Columns.Add("Enabled", typeof(bool));

            dataTable1.Rows.Add();

            dataGridView1.DataSource = dataTable1;

            dataGridView1.Columns[0].Width = 35;
            dataGridView1.Columns[1].Width = 88;
            dataGridView1.Columns[2].Width = 94;
            dataGridView1.Columns[3].Width = 108;
            dataGridView1.Columns[4].Width = 76;
            dataGridView1.Columns[5].Width = 119;
            dataGridView1.Columns[6].Width = 56;
            dataGridView1.Columns[7].Width = 59;
            dataGridView1.Columns[8].Width = 106;
            dataGridView1.Columns[9].Width = 115;
            dataGridView1.Columns[10].Width = 95;
        }
        */

        /*private ADuser createUserFromRow(int row)
        {
            ADuser aduser1 = new ADuser("");

            if (dataTable1.Rows.Count > row & dataGridView1.Rows[row].Cells[4].Value.ToString() != "")
            {
                aduser1 = new ADuser();

                
                aduser1.nameGiven = dataGridView1.Rows[row].Cells[1].Value.ToString();
                aduser1.nameSurn = dataGridView1.Rows[row].Cells[2].Value.ToString();
                aduser1.nameFull = dataGridView1.Rows[row].Cells[3].Value.ToString();
                aduser1.nameAcco = dataGridView1.Rows[row].Cells[4].Value.ToString();
                aduser1.office = dataGridView1.Rows[row].Cells[5].Value.ToString();
                aduser1.description = dataGridView1.Rows[row].Cells[6].Value.ToString();
                aduser1.tel = dataGridView1.Rows[row].Cells[7].Value.ToString();
                aduser1.mob = dataGridView1.Rows[row].Cells[8].Value.ToString();
                aduser1.cardNumber = dataGridView1.Rows[row].Cells[9].Value.ToString();
                aduser1.password = dataGridView1.Rows[row].Cells[10].Value.ToString();
                

                aduser1.nameGiven = dataTable1.Rows[row]["Křestní"].ToString();
                aduser1.nameSurn = dataTable1.Rows[row]["Příjmení"].ToString();
                aduser1.nameFull = dataTable1.Rows[row]["Plné jméno"].ToString();
                aduser1.nameAcco = dataTable1.Rows[row]["Username"].ToString();
                aduser1.office = dataTable1.Rows[row]["Kancelář"].ToString();
                aduser1.department = dataTable1.Rows[row]["Středisko"].ToString();
                aduser1.tel = dataTable1.Rows[row]["Tel"].ToString();
                aduser1.telOthers = dataTable1.Rows[row]["TelOstatní"].ToString();
                aduser1.mob = dataTable1.Rows[row]["Mob"].ToString();
                aduser1.description = dataTable1.Rows[row]["Popis"].ToString();
                aduser1.title = dataTable1.Rows[row]["Pozice"].ToString();
                aduser1.cardNumber = dataTable1.Rows[row]["Karta"].ToString();
                aduser1.emailAddress = dataTable1.Rows[row]["Email"].ToString();
                aduser1.manager = dataTable1.Rows[row]["Manager"].ToString();
                aduser1.password = dataTable1.Rows[row]["Heslo"].ToString();
                aduser1.path = dataTable1.Rows[row]["Path"].ToString();
                try
                {
                    if (Convert.ToBoolean(dataTable1.Rows[row]["ChangePasswordAtLogon"]) == true)
                    {
                        aduser1.ChangePasswordAtLogon = true;
                    }
                }
                catch (Exception e)
                {
                    aduser1.ChangePasswordAtLogon = false;
                }

                try
                {
                    if (Convert.ToBoolean(dataTable1.Rows[row]["CannotChangePassword"]) == true)
                    {
                        aduser1.CannotChangePassword = true;
                    }
                }
                catch (Exception e)
                {
                    aduser1.CannotChangePassword = false;
                }

                try
                {
                    if (Convert.ToBoolean(dataTable1.Rows[row]["PasswordNeverExpires"]) == true)
                    {
                        aduser1.PasswordNeverExpires = true;
                    }
                }
                catch (Exception e)
                {
                    aduser1.PasswordNeverExpires = false;
                }

                try
                {
                    if (Convert.ToBoolean(dataTable1.Rows[row]["Enabled"]) == true)
                    {
                        aduser1.Enabled = true;
                    }
                }
                catch (Exception e)
                {
                    aduser1.Enabled = false;
                }
            }
            else
            {
                MessageBox.Show("Error. Nebylo možné převést data z tabulky do ADuser formátu.");
                log("Error. Nebylo možné převést data z tabulky do ADuser formátu.");
            }
            return aduser1;
        }*/

        /*private void setDataGridView()
        {
            //Nastavení omezení + 2.řádek
            if (dataGridView1.Rows.Count == 1)
            {
                dataTable1.Rows.Add();
            }
            dataGridView1.Rows[1].ReadOnly = true;
            dataGridView1.Rows[1].DefaultCellStyle.BackColor = System.Drawing.Color.LightGray;

            dataGridView1.Refresh();
        }
        */

        /*private void CopyRow(DataGridView dataGV, int SourceRow, int DestinationRow)
        {
            for (int i = 0; i < dataGV.Rows[SourceRow].Cells.Count; i++)
                dataGV.Rows[DestinationRow].Cells[i].Value = dataGV.Rows[SourceRow].Cells[i].Value;
        }*/

        /*private void bTest_Click(object sender, EventArgs e)
        {
            log(PS_SearchUser_Identity("ftester"));

            //log("$koment = " + @"""" + (DateTime.Now.ToString("yyyy/MM/dd") + "_created_by_script1") + @"""");

            //log("" + DateTime.Now.ToString("h:mm:ss tt"));
            
            dataGridView1.Focus();
            dataGridView1.CurrentCell = dataGridView1[0, 0];
            string s = Clipboard.GetText();
            string[] lines = s.Replace("\n", "").Split('\r');
            string[] fields;
            int row = 0;
            int column = 0;
            foreach (string l in lines)
            {
                fields = l.Split('\t');
                foreach (string f in fields)
                    dataGridView1[column++, row].Value = f;
                row++;
                column = 0;
            }

            label_Actual.Text = "...";
            dataTableInicial();
            ADuser user1 = new ADuser("Jaroslav", "Prchlík", "Prchlík Jaroslav", "jprchlik", "11230", "Tábor", "100", "+420 766 234 776", "81AE04C300000CBD", "123456Aa");
            ADuser user2 = new ADuser("Jiří", "Hlavatý", "Hlavatý Jiří", "jhlavaty", "11020", "Projekty UPC", "3 217", "+420 000 000 000", "81AE04C300000957", "Z1A9P7M3I8X2U6D4G5T!");
            addToTable("AD", user1);
            addToTable("EX", user2);

            ADuser userDocasny = new ADuser("jvaldauf");
            ADuser user4 = PS_SearchUser_UserName(userDocasny);
            addToTable("AD", user4);

            dataGridView1.DataSource = dataTable1;
            displayRefresh();

            
            //ADuser userDocasny = new ADuser("jvaldauf");
            //userDocasny = PS_SearchUser_UserName(userDocasny);
            //ADuser userNew = new ADuser("Křestní", "Příjmení", "Test1", "testAccount", "číslo střediska", "zaměstnání", "tel vnitřní", "tel vnější", "", "123456Aa");
            
            //PS_CreateNewUser(userNew);
        }*/

        /**private void dataGridView1_KeyUp(object sender, KeyEventArgs e)
        {
            //při odkliknutí klávesy se pokusí vložit data z clipboardu
            if ((e.Shift && e.KeyCode == Keys.Insert) || (e.Control && e.KeyCode == Keys.V))
            {
                dataGridView1.Focus();
                dataGridView1.CurrentCell = dataGridView1[1, 0];
                string clipboardString = Clipboard.GetText();
                //rozdělí clipboard na řádky
                string[] lines = clipboardString.Replace("\n", "").Split('\r');
                int row = 0;
                int column = 0;
                string[] cells = lines[0].Split('\t');
                int tableCollums = dataGridView1.ColumnCount;
                foreach (string f in cells)
                {
                    ++column;
                    if (tableCollums > column)
                        dataGridView1[column, row].Value = f;
                }
                row++;
                column = 0;
            }

            //při odkliknutí klávesy načte data z excelu
            if ((e.Control && e.KeyCode == Keys.Q))
            {
                //MessageBox.Show("zmačknuto ctrl+Q");
                ts_loadExcel_Click(sender, e);
            }

            //při odkliknutí klávesy uloží uživatele bez upozornění
            if ((e.Control && e.KeyCode == Keys.S))
            {
                //MessageBox.Show("zmačknuto ctrl+S");
                b_Write_Click(sender, e);
            }
        }*/

        /*public static string removeDiacritics(string s)
        {
            s = s.Normalize(NormalizationForm.FormD);
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < s.Length; i++)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(s[i]) != UnicodeCategory.NonSpacingMark) sb.Append(s[i]);
            }

            return sb.ToString();
        }*/

        /*private void excel_ReadFile(int rowNumber)
        {
            //load line from excel file

            DataSet ds = new DataSet();
            string connectionString = excel_GetConnectionString();
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

            //check if load work
            if (ds.Tables.Count == 0)
            {
                log("Error. Nepodařilo se načíst excel");
                MessageBox.Show("Nepodařilo se načíst excel (in ReadExcelFile). Info: " + connectionString.ToString(), "Error");
            }
            else
            {
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
                    try { ChangePasswordAtLogon = Convert.ToBoolean(row["ChangePasswordAtLogon"].ToString()); } catch { }
                    try { CannotChangePassword = Convert.ToBoolean(row["CannotChangePassword"].ToString()); } catch { }
                    try { PasswordNeverExpires = Convert.ToBoolean(row["PasswordNeverExpires"].ToString()); } catch { }
                    try { Enabled = Convert.ToBoolean(row["Enabled"].ToString()); } catch { }


                    //create instance of ADuser
                    ADuser user1 = new ADuser(nameAcco);
                    user1.nameGiven = nameGiven;
                    user1.nameSurn = nameSurn;
                    user1.nameFull = nameFull;
                    //user1.nameAcco = nameAcco;
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
                    user1.ChangePasswordAtLogon = ChangePasswordAtLogon;
                    user1.CannotChangePassword = CannotChangePassword;
                    user1.PasswordNeverExpires = PasswordNeverExpires;
                    user1.Enabled = Enabled;

                    //get index of Columns
                    int collumnNewUserID = dataTable1.Columns.IndexOf("NEW_User");

                    //write data to table
                    dataTable1_ShowADuser(collumnNewUserID, "EXC", user1);
                }
                else
                {
                    log("Error. Nepodařilo se najít záložku. " + sheetName);
                }



            }
        }*/

        /*private string excel_GetConnectionString()
        {
            //create string for connection to excel file

            Dictionary<string, string> props = new Dictionary<string, string>();

            //check if path in textbox exist
            string textboxText = ts_TextBox2.Text;
            string filePath = "";
            if (textboxText == "" | textboxText == "...")
            {
                filePath = string.Format("{0}\\{1}", Directory.GetCurrentDirectory(), "info.xlsx");
                MessageBox.Show("Nebyla vyplněna cesta k souboru.", "Otázka", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            else
            {
                filePath = textboxText;
                if (!File.Exists(filePath))
                {
                    MessageBox.Show("Zadaný soubor nebyl nalezen.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            // XLSX - Excel 2007, 2010, 2012, 2013
            props["Provider"] = "Microsoft.ACE.OLEDB.12.0";
            props["Extended Properties"] = "Excel 12.0 XML";
            props["Data Source"] = filePath; //"C:\\info.xlsx";

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
        }*/

        #endregion junkFromDesign_0.00.37

        /* ----- NEW design ----- */

        DataTable dataTable1;
        
        private void b_Delete_Click(object sender, EventArgs e)
        {
            //button - reset table
            log("Spuštěn: " + sender);
            label_Actual.Text = "...";

            dataGridView1_ResetTable();
        }

        private void b_Clone_Click(object sender, EventArgs e)
        {
            //clone Actual_User to NEW_User
            log("Spuštěn: " + sender);
            label_Actual.Text = "...";

            try
            {
                //get index of Columns
                int collumnNewUserID = dataTable1.Columns.IndexOf("NEW_User");
                int collumnActualUserID = dataTable1.Columns.IndexOf("Actual_User");
                //get row index of path
                DataRow row = dataTable1.Select("type = 'Path'").FirstOrDefault();
                int rowID = dataTable1.Rows.IndexOf(row);

                //actual Columns clone
                dataTable1_CopyColumn(collumnActualUserID, collumnNewUserID);

                //repair path (remove account name from it)
                //get text from that row adn remove accountName
                string path = dataTable1.Rows[rowID][collumnActualUserID].ToString();
                int charLocation = path.IndexOf(",") + 1;
                int maxLenght = path.Length;
                if (charLocation > 0)
                {
                    string textFinal = path.Substring(charLocation, maxLenght - charLocation);
                    dataTable1.Rows[rowID][collumnNewUserID] = textFinal;
                }

            }
            catch (Exception ee)
            {
                MessageBox.Show("Error. Nepodařilo se zkopírovat řádky!");
            }
        }

        private void b_Search_Click(object sender, EventArgs e)
        {
            //button - search user based on NEW_User userName

            log("Spuštěn: " + sender);
            label_Actual.Text = "...";

            //get index of Columns
            int collumnNewUserID = dataTable1.Columns.IndexOf("NEW_User");
            int collumnActualUserID = dataTable1.Columns.IndexOf("Actual_User");
            //get index of row
            int rowID = dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Username'").FirstOrDefault());

            ADuser userNew = new ADuser(dataTable1.Rows[rowID][collumnNewUserID].ToString());
            ADuser userAD = PS_SearchUser_UserName(userNew);
            userAD = PS_GetUserGroups(userAD);

            if (userAD.nameAcco != "")
            {
                dataTable1.Columns[collumnActualUserID].ReadOnly = false;
                dataTable1_ShowADuser(collumnActualUserID, "AD", userAD);
                dataTable1.Columns[collumnActualUserID].ReadOnly = true;
            }
        }

        private void b_Write_Click(object sender, EventArgs e)
        {
            //button - write NEW_User to AD

            log("Spuštěn: " + sender);
            label_Actual.Text = "...";

            //get index of Columns
            int collumnNewUserID = dataTable1.Columns.IndexOf("NEW_User");
            int collumnActualUserID = dataTable1.Columns.IndexOf("Actual_User");
            //get row index of path
            DataRow row = dataTable1.Select("type = 'Username'").FirstOrDefault();
            int rowID = dataTable1.Rows.IndexOf(row);
            
            //search if user exist
            string nameNewUser = dataTable1.Rows[rowID][collumnNewUserID].ToString();
            ADuser existUser = PS_SearchUser_UserName(new ADuser(nameNewUser));
            string nameExistUser = existUser.nameAcco;

            label_Actual.Text = "..."; //smazani erroru o nenalezeni uživatele.

            //if not create user
            if (nameNewUser != "")
            {
                if (nameExistUser == nameNewUser)
                {
                    //přepsání uživatele
                    dataTable1.Columns[collumnActualUserID].ReadOnly = false;
                    dataTable1_ShowADuser(collumnActualUserID, "AD", existUser);
                    dataTable1.Columns[collumnActualUserID].ReadOnly = true;
                    log("Uživatel existuje. Pozor není doděláno přepsání uživatele!!");
                    //MessageBox.Show("Error. Není doděláno přepsání uživatele!");
                }
                else
                {
                    //založení uživatele
                    //TODO: delete data from row 1
                    ADuser newUser1 = dataTable1_createADuser("NEW_User");
                    PS_CreateNewUser(newUser1);
                    PS_AddUserToGroups(newUser1);
                }
            }
            else
            {
                MessageBox.Show("Error. Nový užival nemůže mít prázdné jméno!");
                log("Error. Nový užival nemůže mít prázdné jméno!");
            }

        }

        private void ts_getPath_Click(object sender, EventArgs e)
        {
            //ToolStripMenuItem - get current Directory to textBox (for excel load)
            ts_ExTextBoxPath.Text = string.Format("{0}\\{1}", Directory.GetCurrentDirectory(), "info.xlsx");
        }

        private void ts_getPathAC_Click(object sender, EventArgs e)
        {
            //ToolStripMenuItem - get current Directory to textBox (for excel autocompletion)
            ts_ExTextBoxPathAC.Text = string.Format("{0}\\{1}", Directory.GetCurrentDirectory(), "Office.xlsx");
        }

        private void ts_loadExcel_Click(object sender, EventArgs e)
        {
            //ToolStripMenuItem - load single line from excel (patch and line number is in textbox)
            log("Spuštěn: " + sender);
            label_Actual.Text = "...";

            int excelRow = 1;
            try
            {
                excelRow = Convert.ToInt32(ts_TextBox1.Text);
                ts_TextBox1.Text = "" + (excelRow + 1);
            }
            catch
            {
                MessageBox.Show("Zadej číslo řádku.");
                log("Error. Nebylo zadáno číslo řádku.");
            }

            excel_readLine(excelRow);
        }

        private void ts_moveUser_Click(object sender, EventArgs e)
        {
            //ToolStripMenuItem - move (new) user to diferent container
            log("Spuštěn: " + sender);
            label_Actual.Text = "...";

            //get index of Columns
            int collumnNewUserID = dataTable1.Columns.IndexOf("NEW_User");
            //int collumnActualUserID = dataTable1.Columns.IndexOf("Actual_User");
            //get row index of path
            int rowID = dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Username'").FirstOrDefault());

            string nameNewUser = dataTable1.Rows[rowID][collumnNewUserID].ToString();
            ADuser existUser = PS_SearchUser_UserName(new ADuser(nameNewUser));
            string nameExistUser = existUser.nameAcco;
            //Clear error. canot find user. Will be added more accurate error.
            label_Actual.Text = "...";

            if (nameNewUser != "")
            {
                //check if user exist
                if (nameExistUser == nameNewUser)
                {
                    //move user
                    ADuser newUser = dataTable1_createADuser("NEW_User");
                    PS_MoveUser(newUser, ts_TextBoxPath.Text);
                }
                else
                {
                    MessageBox.Show("Error. Uživatel k přesunu neexistuje.");
                    log("Error. Uživatel k přesunu neexistuje!");
                }
            }
            else
            {
                MessageBox.Show("Error. Přesunovaný uživatel nemůže mít prázdné jméno!");
                log("Error. Přesunovaný uživatel nemůže mít prázdné jméno!");
            }
        }

        private void ts_TextBoxPath_MouseDown(object sender, MouseEventArgs e)
        {
            //Show menu for add apecific text to textbox (for move user)

            if (e.Button == MouseButtons.Right)
            {
                //create menu for quick proposal
                ContextMenuStrip cmenu_path = new ContextMenuStrip();
                cmenu_path.Items.Add("User Container");
                cmenu_path.Items.Add("Test Container");
                cmenu_path.Items.Add("Clone to Table");
                cmenu_path.Items.Add("Clear");

                cmenu_path.Items[0].Click += new System.EventHandler(this.tsPath_userContainer_Click);
                cmenu_path.Items[1].Click += new System.EventHandler(this.tsPath_testContainer_Click);
                cmenu_path.Items[2].Click += new System.EventHandler(this.tsPath_cloneToTable_Click);
                cmenu_path.Items[3].Click += new System.EventHandler(this.tsPath_clear_Click);
                cmenu_path.Show(MousePosition.X, MousePosition.Y);

                ts_TextBoxPath.Control.ContextMenuStrip = cmenu_path;
            }

        }

        private void ts_Test_Click(object sender, EventArgs e)
        {
            // ToolStripMenuItem - for test
        }

        private void ts_createTestUser_Click(object sender, EventArgs e)
        {
            //create specific testUser10 in ADuser and put it in AD
            ADuser user1 = new ADuser();
            user1.nameGiven = "test";
            user1.nameSurn = "User 10";
            user1.nameFull = "test User 10";
            user1.nameAcco = "testUser10";
            user1.password = "1a2z3A4Z5b";
            user1.emailAddress = "testEmail10.cz";
            user1.office = "10000";
            user1.department = "10011";
            user1.tel = "000";
            user1.mob = "000 000 000";
            user1.description = "testovací účet";
            user1.title = "Pozice záložní Tester";
            user1.telOthers = "000;001;002";
            user1.cardNumber = "001";
            user1.manager = "ftester";
            user1.ChangePasswordAtLogon = false;
            user1.CannotChangePassword = true;
            user1.PasswordNeverExpires = true;
            user1.Enabled = true;
            user1.path = "OU=Test,OU=Service,OU=Company,DC=sitel,DC=cz";

            PS_CreateNewUser(user1);
        }

        private void tsPath_userContainer_Click(object sender, EventArgs e)
        {
            //set path in textbox (for move user)
            ts_TextBoxPath.Text = "OU=Users,OU=People,OU=Company,DC=sitel,DC=cz";
            ts_userSetting.ShowDropDown();
        }

        private void tsPath_testContainer_Click(object sender, EventArgs e)
        {
            //set path in textbox (for move user)
            ts_TextBoxPath.Text = "OU=Test,OU=Service,OU=Company,DC=sitel,DC=cz";
            ts_userSetting.ShowDropDown();
        }

        private void tsPath_cloneToTable_Click(object sender, EventArgs e)
        {
            //clone text from textbox to line "path" in table

            string pathText = ts_TextBoxPath.Text;

            //get column index
            int collumnNewUserID = dataTable1.Columns.IndexOf("NEW_User");
            //get row index of path
            DataRow row = dataTable1.Select("type = 'Path'").FirstOrDefault();
            int rowID = dataTable1.Rows.IndexOf(row);

            dataTable1.Rows[rowID][collumnNewUserID] = pathText;
        }

        private void tsPath_clear_Click(object sender, EventArgs e)
        {
            //set path in textbox (for move user)
            ts_TextBoxPath.Text = "";
            ts_userSetting.ShowDropDown();
        }

        private void dataGridView1_ResetTable()
        {
            //reset dataTable1 display
            dataTable1 = new DataTable();

            dataTable1.Columns.Add("type", typeof(string));
            dataTable1.Columns.Add("NEW_User", typeof(string));
            dataTable1.Columns.Add("Actual_User", typeof(string));

            dataTable1.Rows.Add("zdroj", "Local");
            dataTable1.Rows.Add("Křestní", "");
            dataTable1.Rows.Add("Příjmení", "");
            dataTable1.Rows.Add("Plné jméno", "");
            dataTable1.Rows.Add("Username", "");
            dataTable1.Rows.Add("Email", "");
            dataTable1.Rows.Add("Kancelář", "");
            dataTable1.Rows.Add("Středisko", "");
            dataTable1.Rows.Add("Popis", "");
            dataTable1.Rows.Add("Pozice", "");
            dataTable1.Rows.Add("Vedoucí", "");
            dataTable1.Rows.Add("Skupina", "");
            dataTable1.Rows.Add("Tel", "");
            dataTable1.Rows.Add("TelOstatní", "");
            dataTable1.Rows.Add("Mob", "");
            dataTable1.Rows.Add("Karta", "");
            dataTable1.Rows.Add("Heslo", "");
            dataTable1.Rows.Add("Path", "", "");
            dataTable1.Rows.Add("ChangePasswordAtLogon", "False", "False");
            dataTable1.Rows.Add("CannotChangePassword", "False", "False");
            dataTable1.Rows.Add("PasswordNeverExpires", "False", "False");
            dataTable1.Rows.Add("Enabled", "True", "False");
            dataTable1.Columns[0].ReadOnly = true;
            dataTable1.Columns[2].ReadOnly = true;
            dataGridView1.DataSource = dataTable1;

            dataGridView1.Columns[0].Width = 150;
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].Width = 150;
            dataGridView1.Columns[0].DefaultCellStyle.BackColor = System.Drawing.Color.LightGray;
            dataGridView1.Columns[2].DefaultCellStyle.BackColor = System.Drawing.Color.LightGray;
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            //will try to fill box on the basis of the other boxes

            int row1 = e.RowIndex;
            int collum1 = e.ColumnIndex;

            //get index of Columns
            int collumnNewUserID = dataTable1.Columns.IndexOf("NEW_User");
            //get index of row
            int rowFullNameID = dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Plné jméno'").FirstOrDefault());
            int rowAccountNameID = dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Username'").FirstOrDefault());
            int rowFirstNameID = dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Křestní'").FirstOrDefault());
            int rowSecondNameID = dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Příjmení'").FirstOrDefault());
            int rowEmailID = dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Email'").FirstOrDefault());
            int rowOfficeID = dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Kancelář'").FirstOrDefault());
            int rowDepartmentID = dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Středisko'").FirstOrDefault());
            int rowDescriptionID = dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Popis'").FirstOrDefault());
            int rowPositionID = dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Pozice'").FirstOrDefault());
            int rowGroup = dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Skupina'").FirstOrDefault());
            int rowManager = dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Vedoucí'").FirstOrDefault());
            int rowPath = dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Path'").FirstOrDefault());

            //automatic replenishment full name
            if (row1 == rowFullNameID & collum1 == collumnNewUserID)
            {

                if (dataTable1.Rows[row1][collum1].ToString() == "")
                {
                    string first = dataTable1.Rows[rowFirstNameID][collumnNewUserID].ToString();
                    string second = dataTable1.Rows[rowSecondNameID][collumnNewUserID].ToString();
                    if (first != "" & second != "")
                    {
                        dataTable1.Rows[row1][collum1] = second + " " + first;
                    }

                }
            }

            //automatic replenishment First and Second name (from full name)
            if (row1 == rowFullNameID & collum1 == collumnNewUserID)
            {
                if (dataTable1.Rows[row1][collum1].ToString() != "")
                {
                    string fullName = dataTable1.Rows[rowFullNameID][collumnNewUserID].ToString();
                    string[] stringSeparators = new string[] { " " };
                    //separate input name
                    string[] result = fullName.Split(stringSeparators, StringSplitOptions.None);
                    //check if can insert name
                    if (result.Length >= 2)
                    {
                        string firstNameOld = dataTable1.Rows[rowFirstNameID][collumnNewUserID].ToString();
                        string secondNameOld = dataTable1.Rows[rowSecondNameID][collumnNewUserID].ToString();

                        string firstName = result[1];
                        string secondName = result[0];

                        if (firstNameOld == "")
                        {
                            dataTable1.Rows[rowFirstNameID][collumnNewUserID] = firstName;
                        }
                        if (secondNameOld == "")
                        {
                            dataTable1.Rows[rowSecondNameID][collumnNewUserID] = secondName;
                        }
                    }
                }
            }

            //automatic replenishment Username
            if (row1 == rowAccountNameID & collum1 == collumnNewUserID)
            {

                if (dataTable1.Rows[row1][collum1].ToString() == "")
                {
                    string first = dataTable1.Rows[rowFirstNameID][collumnNewUserID].ToString();
                    string second = dataTable1.Rows[rowSecondNameID][collumnNewUserID].ToString();

                    first = first.removeDiacritics();
                    second = second.removeDiacritics();

                    if (first != "" & second != "")
                    {
                        dataTable1.Rows[row1][collum1] = first.ToLower()[0] + second.ToLower();
                    }

                }
            }

            //automatic replenishment Email
            if (row1 == rowEmailID & collum1 == collumnNewUserID)
            {
                string userName = dataTable1.Rows[rowAccountNameID][collumnNewUserID].ToString();
                if (dataTable1.Rows[row1][collum1].ToString() == "")
                {
                    if (userName != "")
                    {
                        dataTable1.Rows[row1][collum1] = userName + "@sitel.cz"; 
                    }
                }
            }

            //automatic replenishment office, department, Description, position, Group (from excel table)
            if (row1 == rowOfficeID & collum1 == collumnNewUserID)
            {
                if (dataTable1.Rows[row1][collum1].ToString() != "")
                {
                    string currentOffice = dataTable1.Rows[row1][collum1].ToString();
                    string loadedDepartmen = "";
                    string loadedDescription = "";
                    string loadedPosition = "";
                    string loadedGroup = "";
                    string loadedManager = "";

                    //load excel data
                    try
                    {
                        string pathAR = ts_ExTextBoxPathAC.Text;
                        Excel myExcel = new Excel(pathAR);

                        loadedDepartmen = myExcel.loadAndCompare(currentOffice, "Department");
                        loadedDescription = myExcel.loadAndCompare(currentOffice, "Description");
                        loadedPosition = myExcel.loadAndCompare(currentOffice, "Title");
                        loadedGroup = myExcel.loadAndCompare(currentOffice, "Group");
                        loadedManager = myExcel.loadAndCompare(currentOffice, "Manager");

                    }
                    catch (Exception ee) { }

                    if (dataTable1.Rows[rowDepartmentID][collum1].ToString() == "")
                    {
                        dataTable1.Rows[rowDepartmentID][collum1] = loadedDepartmen;
                    }

                    if (dataTable1.Rows[rowDescriptionID][collum1].ToString() == "")
                    {
                        dataTable1.Rows[rowDescriptionID][collum1] = loadedDescription;
                    }

                    if (dataTable1.Rows[rowPositionID][collum1].ToString() == "")
                    {
                        dataTable1.Rows[rowPositionID][collum1] = loadedPosition;
                    }

                    if (dataTable1.Rows[rowGroup][collum1].ToString() == "")
                    {
                        dataTable1.Rows[rowGroup][collum1] = loadedGroup;
                    }

                    if (dataTable1.Rows[rowManager][collum1].ToString() == "")
                    {
                        dataTable1.Rows[rowManager][collum1] = loadedManager;
                    }
                }
            }

            //automatic replenishment Path
            if (row1 == rowPath & collum1 == collumnNewUserID)
            {
                if (dataTable1.Rows[row1][collum1].ToString() == "")
                {
                    //add it from outside?
                    dataTable1.Rows[row1][collum1] = "OU=Users,OU=People,OU=Company,DC=sitel,DC=cz";
                }
            }

        }

        private void dataGridView1_KeyUp(object sender, KeyEventArgs e)
        {
            //multiple key shortcut (try insert data from clipboard [excel])

            //ctrl+Q load excel line
            if ((e.Control && e.KeyCode == Keys.Q))
            {
                //MessageBox.Show("zmačknuto ctrl+Q");
                ts_loadExcel_Click(sender, e);
            }

            //ctrl+S write User
            if ((e.Control && e.KeyCode == Keys.S))
            {
                //MessageBox.Show("zmačknuto ctrl+S");
                b_Write_Click(sender, e);
            }

            //Del erase data (selected cells)
            if (e.KeyCode == Keys.Delete)
            {
                //MessageBox.Show("zmačknuto del");
                DataGridViewSelectedCellCollection cells = dataGridView1.SelectedCells;
                if (cells != null)
                {

                    //create rows field
                    List<int> rows = new List<int>();
                    foreach (DataGridViewTextBoxCell cell in cells)
                    {
                        rows.Add(cell.RowIndex);
                    }
                    rows.Reverse();
                    rows.ToArray();

                    //is selected NEW_User collumn
                    int collum = dataTable1.Columns.IndexOf("NEW_User");
                    int collum2 = cells[0].ColumnIndex;
                    if (collum == collum2)
                    {
                        //erase selected cells
                        int loader = 0;
                        foreach (int rowID in rows)
                        {
                            dataTable1.Rows[rowID][collum] = "";
                            loader++;
                        }
                    }
                }
            }

            //ctrl+V insert data from clipboard (separate excel lines)
            if ((e.Shift && e.KeyCode == Keys.Insert) || (e.Control && e.KeyCode == Keys.V))
            {
                //dataGridView1.Focus();

                //separate clipboard on lines
                string clipboardString = Clipboard.GetText();
                string[] lines = clipboardString.Replace("\r", "").Replace("\n", "").Split('\t');
                int linesCount = lines.Length;

                //get idea what to do with clipboard data (based on celected cell/s)
                int collum = dataTable1.Columns.IndexOf("NEW_User");
                DataGridViewSelectedCellCollection cells = dataGridView1.SelectedCells;
                if (cells != null) {

                    //create rows field
                    List<int> rows = new List<int>();
                    foreach (DataGridViewTextBoxCell cell in cells)
                    {
                        rows.Add(cell.RowIndex);
                    }
                    rows.Reverse();
                    rows.ToArray();

                    //selected NEW_User collumn
                    //int collum2 = cells[0].ColumnIndex;
                    //Boolean collumRightOne = false;
                    //if (collum == collum2)
                    //{
                    //    collumRightOne = true;
                    //}

                    //selected multile cells
                    //int rowsCount = rows.Count();
                    //Boolean rowsMultipleSelect = false;
                    //if (rowsCount > 1)
                    //{
                    //    rowsMultipleSelect = true;
                    //}

                    //fill cell with data from clipboard
                    if (cells[0].RowIndex == 0 & cells.Count == 1)
                    {
                        //default setup (from top)

                        //get all rows (-4 bool variable)
                        int rowsCount = rows.Count();
                        rowsCount = dataTable1.Rows.Count -4;

                        //insert all data to cells (loader + 1 -> skip first line, rowsCount-2 -> count+1 skipline+1)
                        for (int loader = 0; loader <= rowsCount-2; loader++)
                        {
                            if (loader <= linesCount - 1)
                            {
                                //insert Data from clipboard
                                dataTable1.Rows[loader + 1][collum] = lines[loader];
                            }
                            else
                            {
                                //erase data outside clipboard data range
                                dataTable1.Rows[loader + 1][collum] = "";
                            }
                        }
                    }
                    else
                    {
                        //insert to specific selected cells
                        int loader = 0;
                        foreach (int rowID in rows)
                        {
                            if (loader <= linesCount - 1)
                            {
                                //insert data to table from clipboard
                                dataTable1.Rows[rowID][collum] = lines[loader];
                            }
                            else
                            {
                                //erase data outside clipboard data range
                                dataTable1.Rows[rowID][collum] = "";
                            }
                            loader++;
                        }
                    }

                } else
                {
                    //no cell selected
                }
            }

        }

        private void dataTable1_CopyColumn(int SourceColumn, int DestinationColumn)
        {
            //clone column data to diferent column
            for (int i = 0; i < dataTable1.Rows.Count; i++)
                dataTable1.Rows[i][DestinationColumn] = dataTable1.Rows[i][SourceColumn].ToString();
        }

        private void dataTable1_ShowADuser(int column, string ADorEXC, ADuser aduser1)
        {
            //write ADuser data to table

            //get index of row and write data
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'zdroj'").FirstOrDefault())][column] = ADorEXC;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Křestní'").FirstOrDefault())][column] = aduser1.nameGiven;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Příjmení'").FirstOrDefault())][column] = aduser1.nameSurn;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Plné jméno'").FirstOrDefault())][column] = aduser1.nameFull;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Username'").FirstOrDefault())][column] = aduser1.nameAcco;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Kancelář'").FirstOrDefault())][column] = aduser1.office;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Středisko'").FirstOrDefault())][column] = aduser1.department;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Tel'").FirstOrDefault())][column] = aduser1.tel;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'TelOstatní'").FirstOrDefault())][column] = aduser1.telOthers;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Mob'").FirstOrDefault())][column] = aduser1.mob;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Popis'").FirstOrDefault())][column] = aduser1.description;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Pozice'").FirstOrDefault())][column] = aduser1.title;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Karta'").FirstOrDefault())][column] = aduser1.cardNumber;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Email'").FirstOrDefault())][column] = aduser1.emailAddress;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Vedoucí'").FirstOrDefault())][column] = aduser1.manager;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Heslo'").FirstOrDefault())][column] = aduser1.password;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Path'").FirstOrDefault())][column] = aduser1.path;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Skupina'").FirstOrDefault())][column] = aduser1.group;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'ChangePasswordAtLogon'").FirstOrDefault())][column] = aduser1.ChangePasswordAtLogon;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'CannotChangePassword'").FirstOrDefault())][column] = aduser1.CannotChangePassword;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'PasswordNeverExpires'").FirstOrDefault())][column] = aduser1.PasswordNeverExpires;
            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Enabled'").FirstOrDefault())][column] = aduser1.Enabled;
            log("-vyplnění řádku v tabulce dokončeno");
        }

        private ADuser dataTable1_createADuser(string columnName)
        {
            //create ADuser class from table
            ADuser aduser1 = new ADuser("");

            //get index of Columns
            int column = -1;
            try
            {
                column = dataTable1.Columns.IndexOf(columnName);
            }
            catch { }

            if (column != -1)
            {
                aduser1 = new ADuser();

                aduser1.nameGiven = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Křestní'").FirstOrDefault())][column].ToString();
                aduser1.nameSurn = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Příjmení'").FirstOrDefault())][column].ToString();
                aduser1.nameFull = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Plné jméno'").FirstOrDefault())][column].ToString();
                aduser1.nameAcco = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Username'").FirstOrDefault())][column].ToString();
                aduser1.office = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Kancelář'").FirstOrDefault())][column].ToString();
                aduser1.department = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Středisko'").FirstOrDefault())][column].ToString();
                aduser1.tel = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Tel'").FirstOrDefault())][column].ToString();
                aduser1.telOthers = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'TelOstatní'").FirstOrDefault())][column].ToString();
                aduser1.mob = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Mob'").FirstOrDefault())][column].ToString();
                aduser1.description = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Popis'").FirstOrDefault())][column].ToString();
                aduser1.title = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Pozice'").FirstOrDefault())][column].ToString();
                aduser1.cardNumber = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Karta'").FirstOrDefault())][column].ToString();
                aduser1.emailAddress = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Email'").FirstOrDefault())][column].ToString();
                aduser1.manager = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Vedoucí'").FirstOrDefault())][column].ToString();
                aduser1.password = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Heslo'").FirstOrDefault())][column].ToString();
                aduser1.path = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Path'").FirstOrDefault())][column].ToString();
                aduser1.group = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Skupina'").FirstOrDefault())][column].ToString();

                string ChangePasswordAtLogon = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'ChangePasswordAtLogon'").FirstOrDefault())][column].ToString();
                //aduser1.ChangePasswordAtLogon = Convert.ToBoolean(ChangePasswordAtLogon);
                string CannotChangePassword = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'CannotChangePassword'").FirstOrDefault())][column].ToString();
                //aduser1.CannotChangePassword = Convert.ToBoolean(CannotChangePassword);
                string PasswordNeverExpires = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'PasswordNeverExpires'").FirstOrDefault())][column].ToString();
                //aduser1.PasswordNeverExpires = Convert.ToBoolean(PasswordNeverExpires);
                string Enabled = dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Enabled'").FirstOrDefault())][column].ToString();
                //aduser1.Enabled = Convert.ToBoolean(Enabled);

                try
                {
                    if (Convert.ToBoolean(ChangePasswordAtLogon) == true)
                    {
                        aduser1.ChangePasswordAtLogon = true;
                    }
                }
                catch (Exception e)
                {
                    aduser1.ChangePasswordAtLogon = false;
                }

                try
                {
                    if (Convert.ToBoolean(CannotChangePassword) == true)
                    {
                        aduser1.CannotChangePassword = true;
                    }
                }
                catch (Exception e)
                {
                    aduser1.CannotChangePassword = false;
                }

                try
                {
                    if (Convert.ToBoolean(PasswordNeverExpires) == true)
                    {
                        aduser1.PasswordNeverExpires = true;
                    }
                }
                catch (Exception e)
                {
                    aduser1.PasswordNeverExpires = false;
                }

                try
                {
                    if (Convert.ToBoolean(Enabled) == true)
                    {
                        aduser1.Enabled = true;
                    }
                }
                catch (Exception e)
                {
                    aduser1.Enabled = false;
                }
            }
            else
            {
                MessageBox.Show("Error. Nebylo možné převést data z tabulky do ADuser formátu. Nenalezen název sloupce: " + columnName);
                log("Error. Nebylo možné převést data z tabulky do ADuser formátu.");
            }
            return aduser1;
        }

        private void PS_CreateNewUser(ADuser user1)
        {
            //try to seach informations about manager
            if (user1.manager != "")
            {
                user1.managerFull = PS_SearchUser_Identity(user1.manager);
            }

            //try to create user
            if (user1.nameAcco != "")
            {
                string mainScript;

                using (var runspace = RunspaceFactory.CreateRunspace())
                {
                    using (var powerShell = PowerShell.Create())
                    {
                        log("Zápis uživatele do AD");
                        powerShell.Runspace = runspace;
                        powerShell.Runspace.Open();

                        //build a script to create a user (not add empty variables)
                        mainScript = "$user = @{ ";
                        mainScript += "    Name = " + user1.nameFull.addQuotes() + ";";
                        mainScript += "    SamAccountName = " + user1.nameAcco.addQuotes() + ";";
                        mainScript += "    UserPrincipalName = " + user1.NamePrincipal.addQuotes() + ";";
                        mainScript += "    displayName = " + user1.nameFull.addQuotes() + ";";

                        if (user1.nameGiven != "")
                        {
                            mainScript += "    GivenName = " + user1.nameGiven.addQuotes() + ";";
                        }
                        if (user1.nameSurn != "")
                        {
                            mainScript += "    Surname = " + user1.nameSurn.addQuotes() + ";";
                        }
                        if (user1.password != "")
                        {
                            mainScript += "    AccountPassword = (" + user1.password.addQuotes() + " | ConvertTo-SecureString -AsPlainText -Force)" + ";";
                        }
                        if (user1.emailAddress != "")
                        {
                            mainScript += "    EmailAddress = " + user1.emailAddress.addQuotes() + ";";
                        }
                        if (user1.office != "")
                        {
                            mainScript += "    Office = " + user1.office.addQuotes() + ";";
                        }
                        if (user1.department != "")
                        {
                            mainScript += "    Department = " + user1.department.addQuotes() + ";";
                        }
                        if (user1.tel != "")
                        {
                            mainScript += "    officephone = " + user1.tel.addQuotes() + ";";
                        }
                        if (user1.mob != "")
                        {
                            mainScript += "    mobile = " + user1.mob.addQuotes() + ";";
                        }
                        if (user1.title != "")
                        {
                            mainScript += "    Title = " + user1.description.addQuotes() + ";";
                        }
                        if (user1.description != "")
                        {
                            mainScript += "    Description = " + user1.description.addQuotes() + ";";
                        }
                        if (user1.managerFull != "")
                        {
                            mainScript += "    Manager = " + user1.managerFull.addQuotes() + ";";
                        }
                        if (user1.path != "")
                        {
                            mainScript += "    Path = " + user1.path.addQuotes() + ";";
                        }
                        if (user1.ChangePasswordAtLogon)
                        {
                            mainScript += "    ChangePasswordAtLogon = " + user1.ChangePasswordAtLogon.addQuotes() + ";";
                        }
                        if (user1.CannotChangePassword)
                        {
                            mainScript += "    CannotChangePassword = " + user1.CannotChangePassword.addQuotes() + ";";
                        }
                        if (user1.PasswordNeverExpires)
                        {
                            mainScript += "    PasswordNeverExpires = " + user1.PasswordNeverExpires.addQuotes() + ";";
                        }
                        if (user1.Enabled)
                        {
                            mainScript += "    Enabled = " + user1.Enabled.addQuotes() + ";";
                        }
                        if (true)
                        {
                            mainScript += "    OtherAttributes = @{";
                            mainScript += "        'Comment' = " + (DateTime.Now.ToString("yyyy/MM/dd") + "_created_by_script2").addQuotes() + ";";
                            if (user1.telOthers != "")
                            {
                                mainScript += "        'otherTelephone'= " + user1.telOthers.addQuotes() + ";";
                            }
                            if (user1.cardNumber != "")
                            {
                                mainScript += "'Pager' = 'k'" + ";";
                                mainScript += "'otherPager' = " + user1.cardFullNumber.addQuotes() + ";";
                            }
                            mainScript += "    }" + ";";
                        }
                        mainScript += "}" + ";";

                        //create user
                        powerShell.AddScript(mainScript);
                        powerShell.AddScript("New-ADUser @User");
                        powerShell.Invoke();

                        if (powerShell.HadErrors == true)
                        {
                            log("Error. Chyba při vytváření uživatele.");
                            MessageBox.Show("Error. Chyba při vytváření uživatele.");
                            //var test = powerShell.Streams.Error.ElementAt(0).Exception.Message;
                            //MessageBox.Show("" + test);
                            foreach (var error in powerShell.Streams.Error)
                            {
                                DialogResult nextmessage = MessageBox.Show("" + error, "Error v Powershellu. Detailed message.", MessageBoxButtons.OKCancel);
                                if (nextmessage == DialogResult.Cancel)
                                    break;
                            }
                        }
                        else
                        {
                            log("Hotovo. Uživatel vytvořen.");
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Error. Nebylo možné vytvořit uživatele. Prázdné jméno!");
                log("Error. Nebylo možné vytvořit uživatele. Prázdné jméno!");
            }
        }

        private void PS_AddUserToGroups(ADuser user1)
        {
            //add user to groups
            if (user1.nameAcco != "" & user1.group != "")
            {
                //start powershell
                using (var runspace = RunspaceFactory.CreateRunspace())
                {
                    using (var powerShell = PowerShell.Create())
                    {
                        //log("Přidávám uživatele do skupin");
                        string script = @"$groups = " + user1.ADgroup + " ; Foreach ($group in $groups){Add-ADGroupMember ($group) " + user1.nameAcco.addQuotes() + "}";

                        //start powershell script
                        powerShell.Runspace = runspace;
                        powerShell.Runspace.Open();
                        powerShell.AddScript(script);
                        powerShell.Invoke();

                        //check for errors
                        if (powerShell.HadErrors == true)
                        {
                            log("Error. Během přidávání uživatele do skupin.");
                            MessageBox.Show("Error. Během přidávání uživatele do skupin.");
                            foreach (var error in powerShell.Streams.Error)
                            {
                                DialogResult nextmessage = MessageBox.Show("" + error, "Error", MessageBoxButtons.OKCancel);
                                if (nextmessage == DialogResult.Cancel)
                                    break;
                            }
                        }
                        else
                        {
                            log("Hotovo. Uživatel přidán do skupiny.");
                        }
                    }
                }
            }
            else
            {
                if (user1.nameAcco == "")
                {
                    //no account name
                    MessageBox.Show("Error. Nebylo možné přidat uživatele do skupiny. Prázdné jméno!");
                    log("Error. Nebylo možné přidat uživatele do skupiny. Prázdné jméno!");
                }
                else
                {
                    //no group
                }

            }
        }

        private void PS_MoveUser(ADuser user1, string nameContainer)
        {
            //move user in diferent container in AD

            if (user1.nameAcco != "")
            {
                using (var runspace = RunspaceFactory.CreateRunspace())
                {
                    using (var powerShell = PowerShell.Create())
                    {
                        log("přesouvám uživatele do jiného kontejneru.");
                        powerShell.Runspace = runspace;
                        powerShell.Runspace.Open();

                        //skript
                        //Get-ADUser ftester | Move-ADObject -TargetPath 'OU=Users,OU=People,OU=Company,DC=sitel,DC=cz'

                        //move user
                        powerShell.AddScript("Get-ADUser " + user1.nameAcco + " | Move-ADObject -TargetPath '" + nameContainer + "'");

                        powerShell.Invoke();

                        if (powerShell.HadErrors == true)
                        {
                            log("Error. Chyba při přesouvání uživatele.");
                            MessageBox.Show("Error. Chyba při přesouvání uživatele.");
                            foreach (var error in powerShell.Streams.Error)
                            {
                                DialogResult nextmessage = MessageBox.Show("" + error, "Error", MessageBoxButtons.OKCancel);
                                if (nextmessage == DialogResult.Cancel)
                                    break;
                            }
                        }
                        else
                        {
                            log("Hotovo. Uživatel přesunut.");
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Error. Nebylo možné přesunout uživatele. Prázdné jméno!");
                log("Error. Nebylo možné přesunout uživatele. Prázdné jméno!");
            }
        }

        private string PS_SearchUser_Identity(string nameSamAccount)
        {
            //Launches PowerShell script to look for user identity based on name (use for manager assignment)
            string identity = "";

            if (nameSamAccount != "")
            {
                using (var runspace = RunspaceFactory.CreateRunspace())
                {
                    using (var powerShell = PowerShell.Create())
                    {
                        log("Hledám identitu manažera: " + nameSamAccount);
                        powerShell.Runspace = runspace;
                        powerShell.Runspace.Open();
                        powerShell.AddScript("$managerFull = Get-ADUser -Identity " + nameSamAccount.addQuotes());
                        powerShell.AddScript("$managerFull.DistinguishedName");
                        PSObject[] results = powerShell.Invoke().ToArray();
                        identity = results[0].BaseObject.ToString();
                        if (identity == "")
                        {
                            log("Error. Manažer nenalezen.");
                        }
                        else
                        {
                            log("Hledání Manažera dokončeno.");
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Error. Nebylo možné vyhledat manažera. Prázdné jméno!");
                log("Error. Nebylo možné vyhledat manažera. Prázdné jméno!");
            }

            return identity;
        }

        private ADuser PS_GetUserGroups(ADuser user1)
        {
            //powershell script to get user groups (to ADuser)
            string group = "";

            //add user to groups
            if (user1.nameAcco != "")
            {
                using (var runspace = RunspaceFactory.CreateRunspace())
                {
                    using (var powerShell = PowerShell.Create())
                    {
                        //log("Hledám skupiny uživatele");
                        string script = "$groups = Get-ADPrincipalGroupMembership " + user1.nameAcco.addQuotes() + @" | select name | Foreach-object {$finalstring+= $_.name + "",""}; $finalstring";

                        //start powershell script
                        powerShell.Runspace = runspace;
                        powerShell.Runspace.Open();
                        powerShell.AddScript(script);
                        PSObject[] results = powerShell.Invoke().ToArray();
                        group = results[0]?.BaseObject.ToString();

                        //check for errors
                        if (powerShell.HadErrors == true)
                        {
                            log("Error. Během hledání skupin uživatele.");
                            MessageBox.Show("Error. Během hledání skupin uživatele.");
                            foreach (var error in powerShell.Streams.Error)
                            {
                                DialogResult nextmessage = MessageBox.Show("" + error, "Error", MessageBoxButtons.OKCancel);
                                if (nextmessage == DialogResult.Cancel)
                                    break;
                            }
                        }
                        else
                        {
                            //log("Hotovo. Skupiny nalezeny.");
                        }
                    }
                }
                user1.ADgroup = group;
            }
            else
            {
                //MessageBox.Show("Error. Nebylo možné získat skupiny uživatele. Prázdné jméno!");
                //log("Error. Nebylo možné získat skupiny uživatele. Prázdné jméno!");
            }
            return user1;
        }

        private ADuser PS_SearchUser_UserName(ADuser user1)
        {
            //Launches PowerShell script to find a user based on the sAMAccountusername
            if (user1.nameAcco != "")
            {
                using (var runspace = RunspaceFactory.CreateRunspace())
                {
                    using (var powerShell = PowerShell.Create())
                    {
                        log("Hledám uživatele podle jména");
                        //string script = @"Get-ADUser -filter 'sAMAccountName -like """ + user1.nameAcco + @"""' -Properties givenName, Surname, name, sAMAccountName, title, PhysicalDeliveryOfficeName, pager, otherPager, mobile, TelephoneNumber | select givenName, Surname, name, sAMAccountName, title, PhysicalDeliveryOfficeName, pager, mobile, TelephoneNumber, @{name=""otherPager"";expression={$_.otherPager -join "";""}}";
                        string script = @"Get-ADUser -filter 'sAMAccountName -like """ + user1.nameAcco + @"""' -Properties * | select givenName, Surname, name, sAMAccountName, title, PhysicalDeliveryOfficeName, pager, mobile, TelephoneNumber,EmailAddress,description,CannotChangePassword,PasswordNeverExpires,Enabled,Path,department,distinguishedName,pwdlastset, @{name='otherPager';expression={$_.otherPager -join ';'}},@{name='otherTelephone';expression={$_.otherTelephone -join ';'}},@{N='Manager';E={(Get-ADUser $_.Manager).sAMAccountName}}";

                        powerShell.Runspace = runspace;
                        powerShell.Runspace.Open();
                        powerShell.AddScript(script);
                        PSObject[] results = powerShell.Invoke().ToArray();
                        user1.cleanse();
                        foreach (PSObject result in results)
                        {
                            try { user1.nameGiven = result.Members["givenName"].Value.ToString(); } catch { }
                            try { user1.nameSurn = result.Members["Surname"].Value.ToString(); } catch { }
                            try { user1.nameFull = result.Members["name"].Value.ToString(); } catch { }
                            try { user1.nameAcco = result.Members["sAMAccountName"].Value.ToString(); } catch { }
                            try { user1.cardFullNumber = result.Members["otherPager"].Value.ToString(); } catch { }
                            try { user1.office = result.Members["PhysicalDeliveryOfficeName"].Value.ToString(); } catch { }
                            try { user1.title = result.Members["title"].Value.ToString(); } catch { }
                            try { user1.description = result.Members["description"].Value.ToString(); } catch { }
                            try { user1.tel = result.Members["TelephoneNumber"].Value.ToString(); } catch { }
                            try { user1.mob = result.Members["mobile"].Value.ToString(); } catch { }
                            try { user1.emailAddress = result.Members["EmailAddress"].Value.ToString(); } catch { }
                            try
                            { //ChangePasswordAtLogon
                                if (Convert.ToInt32(result.Members["pwdlastset"].Value) == 0)
                                {
                                    user1.ChangePasswordAtLogon = true;
                                }
                            }
                            catch { }
                            try { if (Convert.ToBoolean(result.Members["CannotChangePassword"].Value) == true) { user1.CannotChangePassword = true; } } catch { }
                            try { if (Convert.ToBoolean(result.Members["PasswordNeverExpires"].Value) == true) { user1.PasswordNeverExpires = true; } } catch { }
                            try { if (Convert.ToBoolean(result.Members["Enabled"].Value) == true) { user1.Enabled = true; } } catch { }
                            try { user1.path = result.Members["distinguishedName"].Value.ToString(); } catch { }
                            try { user1.department = result.Members["department"].Value.ToString(); } catch { }
                            try { user1.telOthers = result.Members["otherTelephone"].Value.ToString(); } catch { }
                            try { user1.manager = result.Members["Manager"].Value.ToString(); } catch { }
                        }
                        if (user1.nameAcco == "")
                        {
                            //clear Actual_User column
                            int collumnActualUserID = dataTable1.Columns.IndexOf("Actual_User");
                            int numberOfRows = dataTable1.Rows.Count -1;
                            dataTable1.Columns[collumnActualUserID].ReadOnly = false;
                            for (int rowNumber = 0; rowNumber <= numberOfRows; rowNumber++)
                            {
                                dataTable1.Rows[rowNumber][collumnActualUserID] = "";
                            }
                            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'ChangePasswordAtLogon'").FirstOrDefault())][collumnActualUserID] = "False";
                            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'CannotChangePassword'").FirstOrDefault())][collumnActualUserID] = "False";
                            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'PasswordNeverExpires'").FirstOrDefault())][collumnActualUserID] = "False";
                            dataTable1.Rows[dataTable1.Rows.IndexOf(dataTable1.Select("type = 'Enabled'").FirstOrDefault())][collumnActualUserID] = "True";
                            dataTable1.Columns[collumnActualUserID].ReadOnly = true;
                            log("Error. Uživatel nenalezen.");
                        }
                        else
                        {
                            log("Hledani uživatele podle jména dokončeno");
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Error. Nebylo možné vyhledat uživatele. Prázdné jméno!");
                log("Error. Nebylo možné vyhledat uživatele. Prázdné jméno!");
            }
            return user1;
        }

        private void excel_readLine(int excelRow)
        {
            //read excel line and insert in column

            //get index of Columns
            int collumnNewUserID = dataTable1.Columns.IndexOf("NEW_User");

            //excel
            string path = ts_ExTextBoxPath.Text;
            Excel myExcel = new Excel(path);
            ADuser user1 = new ADuser();
            try
            {
                //read data
                user1 = myExcel.readFileRow(excelRow);

                //write data to table
                dataTable1_ShowADuser(collumnNewUserID, "EXC", user1);
            }
            catch (Exception e)
            {
                MessageBox.Show("Excel Error. " + e.Message);
            }

        }

        private void log(string message)
        {
            //show progres on form1 in text
            if (!label_Actual.Text.Contains("Error"))
            {
                label_Actual.Text = message;
            }

        }

    }
}


namespace ExtensionMethods
{
    public static class StringExtensions
    {
        public static string addQuotes(this String input)
        {
            return @"""" + input + @"""";
        }

        public static string addQuotes(this Boolean input)
        {
            string final = "";
            if (input)
            {
                final = "$true";
            }
            else
            {
                final = "$false";
            }
            return final;
        }

        public static string removeDiacritics(this string input)
        {
            // remove diacritics (change string "Jiří" to "Jiri")
            input = input.Normalize(NormalizationForm.FormD);
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < input.Length; i++)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(input[i]) != UnicodeCategory.NonSpacingMark) sb.Append(input[i]);
            }

            return sb.ToString();
        }
    }
}