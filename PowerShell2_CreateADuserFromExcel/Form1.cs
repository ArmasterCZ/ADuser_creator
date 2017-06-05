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

namespace PowerShell2_CreateADuserFromExcel
{
    public partial class Form1 : Form
    {
        public string verze = "0.00.25";
        public string filePath;
        public Form1()
        {
            InitializeComponent();
            dataTableInicial();
            this.Text = "PowerShell - AD User Creator V" + verze;
        }

        //TODO: 
        //user canot sort lines.
        //create Compare User
        //excel imput
        //better test table
        //

        /*Info o verzi
         0.00.01 - Basic design idea
         0.00.05 - method for create new AD user
         0.00.23 - more parameters in ADuser class
         0.00.24 - code cleanup 1/2
         */

        /* ----- test things ----- */

        static DataTable GetTable()
        {
            DataTable table = new DataTable();
            //table.Columns.Add("ID", typeof(int));
            //table.Columns.Add("Čas", typeof(DateTime));

            table.Columns.Add("-zdroj-", typeof(string));
            table.Columns.Add("Křestní", typeof(string));
            table.Columns.Add("Příjmení", typeof(string));
            table.Columns.Add("Plné jméno", typeof(string));
            table.Columns.Add("Username", typeof(string));
            table.Columns.Add("Středisko", typeof(string));
            table.Columns.Add("Pozice", typeof(string));
            table.Columns.Add("Tel", typeof(string));
            table.Columns.Add("Mob", typeof(string));
            table.Columns.Add("Karta", typeof(string));
            table.Columns.Add("Heslo", typeof(string));


            table.Rows.Add("AD", "Jaroslav", "Prchlík", "Prchlík Jaroslav", "jprchlik", "11230", "Tábor", "Zamestnanec", "100", "766 234 776", "123456Aa");
            table.Rows.Add("Excel", "Jiří", "Hlavatý", "Hlavatý Jiří", "jhlavaty", "11020", "Projekty UPC", "Neucetni", "217", "", "");
            //table.Rows.Add(3, "Hydralazine", "Christoff","", DateTime.Now);
            return table;

        }

        private void createUser() //test
        {
            ADuser user1 = new ADuser("Jaroslav", "Prchlík", "Prchlík Jaroslav", "jprchlik", "11230", "Tábor", "Zamestnanec", "100", "766 234 776", "123456Aa");
            ADuser user2 = new ADuser("Jiří", "Hlavatý", "Hlavatý Jiří", "jhlavaty", "11020", "Projekty UPC", "Neucetni", "217", "", "");
        }

        private void displayRefresh() //zakázání editace řádků
        {
            DataGridViewCellStyle style = new DataGridViewCellStyle();
            style.BackColor = System.Drawing.Color.LightGray;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToOrderColumns = false;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                for (int citac = 0; citac < 1; citac++)
                {
                    row.Cells[citac].Style = style;
                    row.Cells[citac].ReadOnly = true;
                }

                if (row.Cells[1].RowIndex == 0)
                {
                    foreach (DataGridViewColumn Columns in dataGridView1.Columns)
                    {
                        row.Cells[Columns.Index].Style = style;
                        row.Cells[Columns.Index].ReadOnly = true;
                    }
                }
            }

            //dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            //dataGridView1.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;

            dataGridView1.Refresh();
        }

        private void addToTable(string dataFrom, ADuser user1)
        {
            dataTableShow1.Rows.Add(
                dataFrom,
                user1.nameGiven,
                user1.nameSurn,
                user1.nameFull,
                user1.nameAcco,
                user1.description,
                user1.office,
                //user1.leader,
                user1.tel,
                user1.mob,
                user1.cardNumber,
                //user1.cardHave,
                user1.password
                );
        }

        /* ----- (powershell script) ----- */

        private ADuser PS_SearchUser_UserName(ADuser user1)
        {
            //spustí PowerShell script na vyhledání uživatele na základě jména
            if (user1.nameAcco != "")
            {
                using (var runspace = RunspaceFactory.CreateRunspace())
                {
                    using (var powerShell = PowerShell.Create())
                    {
                        log("Hledám uživatele podle jména");
                        string script = @"Get-ADUser -filter 'sAMAccountName -like """ + user1.nameAcco + @"""' -Properties givenName, Surname, name, sAMAccountName, title, PhysicalDeliveryOfficeName, pager, otherPager, mobile, TelephoneNumber | select givenName, Surname, name, sAMAccountName, title, PhysicalDeliveryOfficeName, pager, mobile, TelephoneNumber, @{name=""otherPager"";expression={$_.otherPager -join "";""}}";

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
                            try { user1.cardNumber = result.Members["otherPager"].Value.ToString(); } catch { }
                            try { user1.office = result.Members["PhysicalDeliveryOfficeName"].Value.ToString(); } catch { }
                            try { user1.description = result.Members["title"].Value.ToString(); } catch { }
                            try { user1.tel = result.Members["TelephoneNumber"].Value.ToString(); } catch { }
                            try { user1.mob = result.Members["mobile"].Value.ToString(); } catch { }
                        }
                        if (user1.nameAcco == "")
                        {
                            log("Error. Uživatel nenalezen.");
                        } else
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

        private void PS_CreateNewUser(ADuser user1)
        {
            /*
            $nameAccout = "testUser7"
            $nameFull = "test User 7"
            $nameLogin = "testUser7@sitel.cz"
            $nameGiven = "test"
            $nameSurn = "User 7"
            $description = "tester"
            $office = "10000"
            $tel = "000 000 000"
            $mob = "000 000 000"

            New-ADUser -name $nameFull -displayName $nameFull -sAMAccountName $nameAccout -userPrincipalName $nameLogin -givenName $nameGiven -Surname $nameSurn -title $description -Description $description -Office $office -officephone $tel -mobile $mob
            */

            //pokusí se vytvořit uživatele 1 část (bez hesla a disablovaného)
            if (user1.nameAcco != "")
            {
                //MessageBox.Show("bude vytvořen nový uživatel " + user1.QnameAcco);
                //TODO: check all data here

                string mainScript;

                using (var runspace = RunspaceFactory.CreateRunspace())
                {
                    using (var powerShell = PowerShell.Create())
                    {
                        log("Zápis uživatele do AD");
                        powerShell.Runspace = runspace;
                        powerShell.Runspace.Open();

                        //vkládání proměných
                        powerShell.AddScript("$nameAccout = " + user1.QnameAcco); 
                        powerShell.AddScript("$nameFull = " + user1.QnameFull);
                        powerShell.AddScript(@"$nameLogin = """ + user1.nameAcco + @"@sitel.cz""");
                        
                        //sestavení skriptu pro vytvoření uživatele (ošetření prázdných proměných)
                        mainScript = "New-ADUser -name $nameFull -displayName $nameFull -sAMAccountName $nameAccout -userPrincipalName $nameLogin ";
                        if (user1.nameGiven != "")
                        {
                            mainScript += "-givenName $nameGiven ";
                            powerShell.AddScript("$nameGiven = " + user1.QnameGiven);
                        }
                        if (user1.nameSurn != "")
                        {
                            mainScript += "-Surname $nameSurn ";
                            powerShell.AddScript("$nameSurn = " + user1.QnameSurn);
                        }
                        if (user1.description != "")
                        {
                            mainScript += "-title $description ";
                            mainScript += "-Description $description ";
                            powerShell.AddScript("$description = " + user1.Qdescription);
                        }
                        if (user1.office != "")
                        {
                            mainScript += "-Office $office ";
                            powerShell.AddScript("$office = " + user1.Qoffice);
                        }
                        if (user1.tel != "")
                        {
                            mainScript += "-officephone $tel ";
                            powerShell.AddScript("$tel = " + user1.Qtel);
                        }
                        if (user1.mob != "")
                        {
                            mainScript += "-mobile $mob ";
                            powerShell.AddScript("$mob = " + user1.Qmob);
                        }

                        //založi uživatele
                        powerShell.AddScript(mainScript);

                        //nastaví heslo, enabluje uživatele, potřeba změnit heslo při dalším přihlášení
                        if (user1.password != "")
                        {
                            powerShell.AddScript("$password = " + user1.Qpassword);
                            powerShell.AddScript("Set-ADUser -Identity $nameAccout -ChangePasswordAtLogon $true -PasswordNeverExpires $false");
                            powerShell.AddScript("Set-ADAccountPassword -Identity $nameAccout -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $password -Force)");
                            powerShell.AddScript("Enable-ADAccount -Identity $nameAccout");
                        }

                        //nastaví katru
                        if (user1.cardNumber != "")
                        {
                            powerShell.AddScript("$pager = " + @"""k""");
                            powerShell.AddScript("$otherPager =" + user1.QcardFullNumber);
                            powerShell.AddScript("Set-ADUser $nameAccout -Replace @{pager=$pager;otherPager=$otherPager}");
                            //powerShell.AddScript("Set-ADUser $nameAccout -Replace @{pager=$pager;otherPager=$otherPager}");
                            //powerShell.AddScript("Set-ADUser " + '"' + user1.userNameAcco + '"' + " -Replace @{pager= " + '"' + "k" + '"' + ";otherPager=" + '"' + user1.cardNumberFull + '"' + "}");
                        }

                        //nastaví koment
                        if (true)
                        {
                            //mainScript += "-comment $koment ";// neobsahuje příkaz comment
                            powerShell.AddScript("$koment = " + @"""" + (DateTime.Now.ToString("yyyy/MM/dd") + "_created_by_script1") + @"""");
                            powerShell.AddScript("Set-ADuser -identity $nameAccout -Replace @{comment= $koment}");
                        }

                        powerShell.Invoke();

                        if (powerShell.HadErrors == true)
                        {
                            log("Error. Chyba při vytváření uživatele.");
                            MessageBox.Show("Error. Chyba při vytváření uživatele.");
                            //var test = powerShell.Streams.Error.ElementAt(0).Exception.Message;
                            //MessageBox.Show("" + test);
                            foreach (var error in powerShell.Streams.Error)
                            {
                                DialogResult nextmessage = MessageBox.Show("" + error,"Error",MessageBoxButtons.OKCancel);
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

        /* ----- ( ostatní) ----- */

        DataTable dataTableShow1;

        private void dataTableInicial()
        {
            dataTableShow1 = new DataTable();
            //table.Columns.Add("ID", typeof(int));
            //table.Columns.Add("Čas", typeof(DateTime));

            dataTableShow1.Columns.Add("-zdroj-", typeof(string));
            dataTableShow1.Columns.Add("Křestní", typeof(string));
            dataTableShow1.Columns.Add("Příjmení", typeof(string));
            dataTableShow1.Columns.Add("Plné jméno", typeof(string));
            dataTableShow1.Columns.Add("Username", typeof(string));
            dataTableShow1.Columns.Add("Pozice", typeof(string));
            dataTableShow1.Columns.Add("Středisko", typeof(string));
            dataTableShow1.Columns.Add("Tel", typeof(string));
            dataTableShow1.Columns.Add("Mob", typeof(string));
            dataTableShow1.Columns.Add("Karta", typeof(string));
            dataTableShow1.Columns.Add("Heslo", typeof(string));

            dataTableShow1.Rows.Add();

            dataGridView1.DataSource = dataTableShow1;

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

        private void log(string message)
        {
            if (!label_Actual.Text.Contains("Error"))
            {
                label_Actual.Text = message;
            }

        }

        private void dataTableWriteUser(int row,string ADorEXC, ADuser user1)
        {
            if (dataTableShow1.Rows.Count > row)
            {
                log("-vyplnění řádku v tabulce");
                dataTableShow1.Rows[row][0] = ADorEXC;
                dataTableShow1.Rows[row][1] = user1.nameGiven;
                dataTableShow1.Rows[row][2] = user1.nameSurn;
                dataTableShow1.Rows[row][3] = user1.nameFull;
                dataTableShow1.Rows[row][4] = user1.nameAcco;
                dataTableShow1.Rows[row][5] = user1.description;
                dataTableShow1.Rows[row][6] = user1.office;
                //dataTableShow1.Rows[row][0] =user1.leader;
                dataTableShow1.Rows[row][7] = user1.tel;
                dataTableShow1.Rows[row][8] = user1.mob;
                dataTableShow1.Rows[row][9] = user1.cardNumber;
                //dataTableShow1.Rows[row][0] =user1.cardHave;
                dataTableShow1.Rows[row][10] = user1.password;
                log("-vyplnění řádku v tabulce dokončeno");
            } else
            {
                log("Error. Nelze zapsat do řádku v tabulce z důvodu čísla řádky ");
            }
        }

        private void setDataGridView()
        {
            //Nastavení omezení + 2.řádek
            if (dataGridView1.Rows.Count == 1)
            {
                dataTableShow1.Rows.Add();
            }
            dataGridView1.Rows[1].ReadOnly = true;
            dataGridView1.Rows[1].DefaultCellStyle.BackColor = System.Drawing.Color.LightGray;

            dataGridView1.Refresh();
        }
        
        private ADuser createUserFromRow(int row)
        {
            ADuser user1 = new ADuser("");



            if (dataTableShow1.Rows.Count > row & dataGridView1.Rows[row].Cells[4].Value.ToString() != "")
            {
                //todo vložit data do user
                //dataTableShow1.Columns.Add("-zdroj-", typeof(string));
                //dataTableShow1.Columns.Add("Křestní", typeof(string));
                //dataTableShow1.Columns.Add("Příjmení", typeof(string));
                //dataTableShow1.Columns.Add("Plné jméno", typeof(string));
                //dataTableShow1.Columns.Add("Username", typeof(string));
                //dataTableShow1.Columns.Add("Pozice", typeof(string));
                //dataTableShow1.Columns.Add("Středisko", typeof(string));
                //dataTableShow1.Columns.Add("Tel", typeof(string));
                //dataTableShow1.Columns.Add("Mob", typeof(string));
                //dataTableShow1.Columns.Add("Karta", typeof(string));
                //dataTableShow1.Columns.Add("Heslo", typeof(string));

                string nameGiven= dataGridView1.Rows[row].Cells[1].Value.ToString();
                string nameSurn= dataGridView1.Rows[row].Cells[2].Value.ToString();
                string nameFull= dataGridView1.Rows[row].Cells[3].Value.ToString();
                string nameAcco= dataGridView1.Rows[row].Cells[4].Value.ToString();
                string office= dataGridView1.Rows[row].Cells[5].Value.ToString();
                string description= dataGridView1.Rows[row].Cells[6].Value.ToString();
                string tel= dataGridView1.Rows[row].Cells[7].Value.ToString();
                string mob= dataGridView1.Rows[row].Cells[8].Value.ToString();
                string cardNumber= dataGridView1.Rows[row].Cells[9].Value.ToString();
                string password= dataGridView1.Rows[row].Cells[10].Value.ToString();

                user1 = new ADuser(nameGiven, nameSurn, nameFull, nameAcco, description, office, tel, mob, cardNumber, password);

            } else
            {
                MessageBox.Show("Error. Nebylo možné převést data z tabulky do ADuser formátu.");
                log("Error. Nebylo možné převést data z tabulky do ADuser formátu.");
            }
            return user1;
        }

        private void CopyRow(DataGridView dataGV, int SourceRow, int DestinationRow)
        {
            for (int i = 0; i < dataGV.Rows[SourceRow].Cells.Count; i++)
                dataGV.Rows[DestinationRow].Cells[i].Value = dataGV.Rows[SourceRow].Cells[i].Value;
        }


        /* ----- ( Excel ) ----- */

        private string GetConnectionString()
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            //zkontroluje zda cesta existuje

            string textboxText = ts_TextBox2.Text;

            if (textboxText == "" | textboxText == "...")
            {
                filePath = string.Format("{0}\\{1}", Directory.GetCurrentDirectory(), "info.xls");
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
        }

        private void ReadExcelFile(int rowNumber)
        {
            //načte řádek z excelu, a zapíše jej do první řádky
            //D:\Users\jvaldauf\Documents\Visual Studio 2015\Projects\PowerShell2_CreateADuserFromExcel\PowerShell2_CreateADuserFromExcel\bin\Debug
            DataSet ds = new DataSet();

            string connectionString = GetConnectionString();

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
                log("Error. Nepodařilo se načíst excel");
                MessageBox.Show("Nepodařilo se načíst excel (in ReadExcelFile). Info: " + connectionString.ToString(), "Error");
            } else
            {
                string sheetName = "List1$";
                if (ds.Tables.Contains(sheetName))
                {
                    DataTable table1 = ds.Tables[sheetName];
                    int table1RowsNumber = table1.Rows.Count;
                    DataRow row = table1.Rows[rowNumber];

                    string nameFirs = row["jmeno"].ToString();
                    string nameSeco = row["prijmeni"].ToString();
                    string nameFull = row["celeJmeno"].ToString();
                    string nameUser = row["uzivatel"].ToString();
                    string office = row["kmen_str"].ToString();
                    string password = row["heslo"].ToString();

                    ADuser user1 = new ADuser(nameUser);
                    user1.nameGiven = nameFirs;
                    user1.nameSurn = nameSeco;
                    user1.nameFull = nameFull;
                    user1.office = office;
                    user1.password = password;

                    dataTableWriteUser(0, "EXC", user1);
                }
                else
                {
                    log("Error. Nepodařilo se najít záložku. " + sheetName);
                }
                         
                
                
            }
        }

        /* ----- ( buttons ) ----- */

        private void bTest_Click(object sender, EventArgs e)
        {
            
            log("$koment = " + @"""" + (DateTime.Now.ToString("yyyy/MM/dd") + "_created_by_script1") + @"""");

            //log("" + DateTime.Now.ToString("h:mm:ss tt"));
            /*dataGridView1.Focus();
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
            }*/

            /*label_Actual.Text = "...";
            dataTableInicial();
            ADuser user1 = new ADuser("Jaroslav", "Prchlík", "Prchlík Jaroslav", "jprchlik", "11230", "Tábor", "100", "+420 766 234 776", "81AE04C300000CBD", "123456Aa");
            ADuser user2 = new ADuser("Jiří", "Hlavatý", "Hlavatý Jiří", "jhlavaty", "11020", "Projekty UPC", "3 217", "+420 000 000 000", "81AE04C300000957", "Z1A9P7M3I8X2U6D4G5T!");
            addToTable("AD", user1);
            addToTable("EX", user2);

            ADuser userDocasny = new ADuser("jvaldauf");
            ADuser user4 = PS_SearchUser_UserName(userDocasny);
            addToTable("AD", user4);

            dataGridView1.DataSource = dataTableShow1;
            displayRefresh();*/

            /*
            ADuser userDocasny = new ADuser("jvaldauf");
            userDocasny = PS_SearchUser_UserName(userDocasny);
            ADuser userNew = new ADuser("Křestní", "Příjmení", "Test1", "testAccount", "číslo střediska", "zaměstnání", "tel vnitřní", "tel vnější", "", "123456Aa");
            */
            //PS_CreateNewUser(userNew);
        }

        private void bSearch_Click(object sender, EventArgs e)
        {
            log("Spuštěn: " + sender);
            label_Actual.Text = "...";
            setDataGridView();
            //vyhledani uživatele podle userName v první řádce tabulky
            ADuser userNew = new ADuser(dataGridView1.Rows[0].Cells[4].Value.ToString());
            ADuser userAD = PS_SearchUser_UserName(userNew);
            if (userAD.nameAcco != "")
            {
                dataTableWriteUser(1,"AD", userAD);
            }
        }

        private void bDelete_Click(object sender, EventArgs e)
        {
            log("Spuštěn: " + sender);
            label_Actual.Text = "...";

            dataTableShow1 = new DataTable();
            dataTableInicial();
            setDataGridView();

            /*dataGridView1.SelectAll();
            foreach (DataGridViewRow item in this.dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.RemoveAt(item.Index);
            }*/
        }

        private void bClone_Click(object sender, EventArgs e)
        {
            log("Spuštěn: " + sender);
            label_Actual.Text = "...";

            try
            {
                //string nameGiven = dataGridView1.Rows[1].Cells[1].Value.ToString();
                //string nameSurn = dataGridView1.Rows[1].Cells[2].Value.ToString();
                //string nameFull = dataGridView1.Rows[1].Cells[3].Value.ToString();
                //string nameAcco = dataGridView1.Rows[1].Cells[4].Value.ToString();
                //string office = dataGridView1.Rows[1].Cells[5].Value.ToString();
                //string description = dataGridView1.Rows[1].Cells[6].Value.ToString();
                //string tel = dataGridView1.Rows[1].Cells[7].Value.ToString();
                //string mob = dataGridView1.Rows[1].Cells[8].Value.ToString();
                //string cardNumber = dataGridView1.Rows[1].Cells[9].Value.ToString();
                //string password = dataGridView1.Rows[1].Cells[10].Value.ToString();

                //dataGridView1.Rows[0].Cells[1].Value = nameGiven;
                //dataGridView1.Rows[0].Cells[2].Value = nameSurn;
                //dataGridView1.Rows[0].Cells[3].Value = nameFull;
                //dataGridView1.Rows[0].Cells[4].Value = nameAcco;
                //dataGridView1.Rows[0].Cells[5].Value = office;
                //dataGridView1.Rows[0].Cells[6].Value = description; 
                //dataGridView1.Rows[0].Cells[7].Value = tel;
                //dataGridView1.Rows[0].Cells[8].Value = mob;
                //dataGridView1.Rows[0].Cells[9].Value = cardNumber;
                //dataGridView1.Rows[0].Cells[10].Value = password;

                CopyRow(dataGridView1, 1, 0);


            } catch
            {
                MessageBox.Show("Error. Nepodařilo se zkopírovat řádky!");
            }
        }

        private void bWrite_Click(object sender, EventArgs e)
        {
            log("Spuštěn: " + sender);
            label_Actual.Text = "...";
            setDataGridView();

            string nameNewUser = dataGridView1.Rows[0].Cells[4].Value.ToString();
            ADuser existUser = PS_SearchUser_UserName(new ADuser(nameNewUser));
            string nameExistUser = existUser.nameAcco;

            label_Actual.Text = "..."; //smazani erroru o nenalezeni uživatele.

            if (nameNewUser != "")
            {
                if (nameExistUser == nameNewUser)
                {
                    //přepsání uživatele

                    dataTableWriteUser(1,"AD", existUser);
                    log("Uživatel existuje. Pozor není doděláno přepsání uživatele!!");
                    //MessageBox.Show("Error. Není doděláno přepsání uživatele!");

                } else
                {
                    //založení uživatele
                    //TODO: delete data from row 1
                    ADuser newUser = createUserFromRow(0);
                    PS_CreateNewUser(newUser);
                }

            } else
            {
                MessageBox.Show("Error. Nový užival nemůže mít prázdné jméno!");
                log("Error. Nový užival nemůže mít prázdné jméno!");
            }

        }

        private void TS_MenuItem_loadExcel_Click(object sender, EventArgs e)
        {
            log("Spuštěn: " + sender);
            label_Actual.Text = "...";

            int excelRow = 1;
            try
            {
                excelRow = Convert.ToInt32(ts_TextBox1.Text);
                ts_TextBox1.Text = "" + (excelRow +1);
            }
            catch
            {
                MessageBox.Show("Zadej číslo řádku.");
                log("Error. Nebylo zadáno číslo řádku.");
            }

            //MessageBox.Show("Error. Není doděláno načtení Excelu!");
            ReadExcelFile(excelRow);
            //TODO:load excel

        }

        private void dataGridView1_KeyUp(object sender, KeyEventArgs e)
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
                TS_MenuItem_loadExcel_Click(sender, e);
            }

            //při odkliknutí klávesy uloží uživatele bez upozornění
            if ((e.Control && e.KeyCode == Keys.S))
            {
                //MessageBox.Show("zmačknuto ctrl+S");
                bWrite_Click(sender, e);
            }
        }

        private void TS_getPath_Click(object sender, EventArgs e)
        {
            //vyplní cestu (z horního menu) aktuální pozicí souboru
            ts_TextBox2.Text = string.Format("{0}\\{1}", Directory.GetCurrentDirectory(), "data.xls");
        }
    }
}
