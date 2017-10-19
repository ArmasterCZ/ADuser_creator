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

namespace PowerShell2_CreateADuserFromExcel
{
    public partial class Form1 : Form
    {
        public string verze = "0.00.36";
        public string filePath;
        public Form1()
        {
            InitializeComponent();
            dataTableInicial();
            this.Text = "PowerShell - AD User Creator V" + verze;
        }

        //TODO: 
        //user canot sort lines (but can :/ ).
        //create Compare User
        //better test table
        //add log (uložení do texťáku), vypsat celý powershellový skript, první řádka datum a čas, dále odrážka kvůli čitelnosti.
        //opravit bugy -> funkci (ctrl+C,ctrl+V), 
        //

        /*Info o verzi
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
         0.00.32 - automatické doplňování tabulky
         0.00.32 - bugfix: kopírovat do horního řádku v kolonce path pouze cestu. odstranění začáteku textu 
         0.00.34 - added right click menu for container move
         0.00.35 - remove diacritic from username (in autocomplete), added split from full name to first and second cell.
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
            //ADuser user1 = new ADuser("Jaroslav", "Prchlík", "Prchlík Jaroslav", "jprchlik", "11230", "Tábor", "Zamestnanec", "100", "766 234 776", "123456Aa");
            //ADuser user2 = new ADuser("Jiří", "Hlavatý", "Hlavatý Jiří", "jhlavaty", "11020", "Projekty UPC", "Neucetni", "217", "", "");
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

        #region PowerShell
        /* ----- (powershell script) ----- */

        private string PS_SearchUser_Identity(string nameSamAccount)
        {
            //spustí PowerShell script na vyhledání identity uživatele na základě jména (využití pro přiřazení manažera)
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
                            catch { } //nefunguje
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

        private void PS_CreateNewUser(ADuser user1)
        {
            //pokusí se najít identifikační údaje k manažerovy
            if (user1.manager != "")
            {
                user1.managerFull = PS_SearchUser_Identity(user1.manager);
            }

            //pokusí se vytvořit uživatele
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

                        //sestavení skriptu pro vytvoření uživatele (ošetření prázdných proměných)
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

                        //založi uživatele
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

        private void PS_MoveUser(ADuser user1, string nameContainer)
        {
            //přesune uživatele do jiného kontejneru v AD

            if (user1.nameAcco != "")
            {
                using (var runspace = RunspaceFactory.CreateRunspace())
                {
                    using (var powerShell = PowerShell.Create())
                    {
                        log("přesouvám uživatele do jiného kontejneru.");
                        powerShell.Runspace = runspace;
                        powerShell.Runspace.Open();

                        //ukázka skriptu
                        //Get-ADUser ftester | Move-ADObject -TargetPath 'OU=Users,OU=People,OU=Company,DC=sitel,DC=cz'

                        //přesune uživatele
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

        #endregion

        #region Ostatní
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
            dataTableShow1.Columns.Add("Kancelář", typeof(string));
            dataTableShow1.Columns.Add("Středisko", typeof(string));
            dataTableShow1.Columns.Add("Tel", typeof(string));
            dataTableShow1.Columns.Add("TelOstatní", typeof(string));
            dataTableShow1.Columns.Add("Mob", typeof(string));
            dataTableShow1.Columns.Add("Popis", typeof(string));
            dataTableShow1.Columns.Add("Pozice", typeof(string));
            dataTableShow1.Columns.Add("Karta", typeof(string));
            dataTableShow1.Columns.Add("Email", typeof(string));
            dataTableShow1.Columns.Add("Manager", typeof(string));
            dataTableShow1.Columns.Add("Heslo", typeof(string));
            dataTableShow1.Columns.Add("Path", typeof(string));
            dataTableShow1.Columns.Add("ChangePasswordAtLogon", typeof(bool));
            dataTableShow1.Columns.Add("CannotChangePassword", typeof(bool));
            dataTableShow1.Columns.Add("PasswordNeverExpires", typeof(bool));
            dataTableShow1.Columns.Add("Enabled", typeof(bool));

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

        private void dataTableWriteUser(int row, string ADorEXC, ADuser aduser1)
        {
            if (dataTableShow1.Rows.Count > row)
            {
                dataTableShow1.Rows[row]["-zdroj-"] = ADorEXC;
                dataTableShow1.Rows[row]["Křestní"] = aduser1.nameGiven;
                dataTableShow1.Rows[row]["Příjmení"] = aduser1.nameSurn;
                dataTableShow1.Rows[row]["Plné jméno"] = aduser1.nameFull;
                dataTableShow1.Rows[row]["Username"] = aduser1.nameAcco;
                dataTableShow1.Rows[row]["Kancelář"] = aduser1.office;
                dataTableShow1.Rows[row]["Středisko"] = aduser1.department;
                dataTableShow1.Rows[row]["Tel"] = aduser1.tel;
                dataTableShow1.Rows[row]["TelOstatní"] = aduser1.telOthers;
                dataTableShow1.Rows[row]["Mob"] = aduser1.mob;
                dataTableShow1.Rows[row]["Popis"] = aduser1.description;
                dataTableShow1.Rows[row]["Pozice"] = aduser1.title;
                dataTableShow1.Rows[row]["Karta"] = aduser1.cardNumber;
                dataTableShow1.Rows[row]["Email"] = aduser1.emailAddress;
                dataTableShow1.Rows[row]["Manager"] = aduser1.manager;
                dataTableShow1.Rows[row]["Heslo"] = aduser1.password;
                dataTableShow1.Rows[row]["Path"] = aduser1.path;
                dataTableShow1.Rows[row]["ChangePasswordAtLogon"] = aduser1.ChangePasswordAtLogon;
                dataTableShow1.Rows[row]["CannotChangePassword"] = aduser1.CannotChangePassword;
                dataTableShow1.Rows[row]["PasswordNeverExpires"] = aduser1.PasswordNeverExpires;
                dataTableShow1.Rows[row]["Enabled"] = aduser1.Enabled;
                log("-vyplnění řádku v tabulce dokončeno");
            }
            else
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
            ADuser aduser1 = new ADuser("");

            if (dataTableShow1.Rows.Count > row & dataGridView1.Rows[row].Cells[4].Value.ToString() != "")
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

                aduser1.nameGiven = dataTableShow1.Rows[row]["Křestní"].ToString();
                aduser1.nameSurn = dataTableShow1.Rows[row]["Příjmení"].ToString();
                aduser1.nameFull = dataTableShow1.Rows[row]["Plné jméno"].ToString();
                aduser1.nameAcco = dataTableShow1.Rows[row]["Username"].ToString();
                aduser1.office = dataTableShow1.Rows[row]["Kancelář"].ToString();
                aduser1.department = dataTableShow1.Rows[row]["Středisko"].ToString();
                aduser1.tel = dataTableShow1.Rows[row]["Tel"].ToString();
                aduser1.telOthers = dataTableShow1.Rows[row]["TelOstatní"].ToString();
                aduser1.mob = dataTableShow1.Rows[row]["Mob"].ToString();
                aduser1.description = dataTableShow1.Rows[row]["Popis"].ToString();
                aduser1.title = dataTableShow1.Rows[row]["Pozice"].ToString();
                aduser1.cardNumber = dataTableShow1.Rows[row]["Karta"].ToString();
                aduser1.emailAddress = dataTableShow1.Rows[row]["Email"].ToString();
                aduser1.manager = dataTableShow1.Rows[row]["Manager"].ToString();
                aduser1.password = dataTableShow1.Rows[row]["Heslo"].ToString();
                aduser1.path = dataTableShow1.Rows[row]["Path"].ToString();
                try
                {
                    if (Convert.ToBoolean(dataTableShow1.Rows[row]["ChangePasswordAtLogon"]) == true)
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
                    if (Convert.ToBoolean(dataTableShow1.Rows[row]["CannotChangePassword"]) == true)
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
                    if (Convert.ToBoolean(dataTableShow1.Rows[row]["PasswordNeverExpires"]) == true)
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
                    if (Convert.ToBoolean(dataTableShow1.Rows[row]["Enabled"]) == true)
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
        }

        private void CopyRow(DataGridView dataGV, int SourceRow, int DestinationRow)
        {
            for (int i = 0; i < dataGV.Rows[SourceRow].Cells.Count; i++)
                dataGV.Rows[DestinationRow].Cells[i].Value = dataGV.Rows[SourceRow].Cells[i].Value;
        }

        private void log(string message)
        {
            if (!label_Actual.Text.Contains("Error"))
            {
                label_Actual.Text = message;
            }

        }

        #endregion

        /* ----- ( Excel ) ----- */

        private string GetConnectionString()
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            //zkontroluje zda cesta existuje

            string textboxText = ts_TextBox2.Text;

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


                    //načte informace (v řádce řádku) podle názvu sloupce
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


                    //vytvoří ADuser z daných informací
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


                    //vypíše informace do řádky
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
            log(PS_SearchUser_Identity("ftester"));

            //log("$koment = " + @"""" + (DateTime.Now.ToString("yyyy/MM/dd") + "_created_by_script1") + @"""");

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
            ADuser userNew = new ADuser(dataTableShow1.Rows[0]["Username"].ToString());
            ADuser userAD = PS_SearchUser_UserName(userNew);
            if (userAD.nameAcco != "")
            {
                dataTableWriteUser(1, "AD", userAD);
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

                //oprava path
                string text = dataTableShow1.Rows[0]["Path"].ToString();
                int charLocation = text.IndexOf(",")+1;
                int maxLenght = text.Length;
                if (charLocation > 0)
                {
                    string textFinal = text.Substring(charLocation, maxLenght - charLocation);
                    dataTableShow1.Rows[0]["Path"] = textFinal;
                }

            }
            catch
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

                    dataTableWriteUser(1, "AD", existUser);
                    log("Uživatel existuje. Pozor není doděláno přepsání uživatele!!");
                    //MessageBox.Show("Error. Není doděláno přepsání uživatele!");

                }
                else
                {
                    //založení uživatele
                    //TODO: delete data from row 1
                    ADuser newUser = createUserFromRow(0);
                    PS_CreateNewUser(newUser);
                }

            }
            else
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
                ts_TextBox1.Text = "" + (excelRow + 1);
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
            ts_TextBox2.Text = string.Format("{0}\\{1}", Directory.GetCurrentDirectory(), "info.xlsx");
        }

        private void TS_MenuItem_moveUser_Click(object sender, EventArgs e)
        {
            log("Spuštěn: " + sender);
            label_Actual.Text = "...";
            //setDataGridView();
            //dataGridView1.Refresh();
            string nameNewUser = dataTableShow1.Rows[0]["Username"].ToString();
            //string nameNewUser = dataGridView1.Rows[0].Cells[4].Value.ToString();
            ADuser existUser = PS_SearchUser_UserName(new ADuser(nameNewUser));
            string nameExistUser = existUser.nameAcco;

            label_Actual.Text = "..."; //smazani erroru o nenalezeni uživatele. (kontrola přepsání)

            if (nameNewUser != "")
            {
                if (nameExistUser == nameNewUser)
                {
                    //move user
                    ADuser newUser = createUserFromRow(0);
                    PS_MoveUser(newUser, ts_TextBox3.Text);
                }

            }
            else
            {
                MessageBox.Show("Error. Přesunovaný uživatel nemůže mít prázdné jméno!");
                log("Error. Přesunovaný uživatel nemůže mít prázdné jméno!");
            }
        }

        private void TS_createTestUser_Click(object sender, EventArgs e)
        {
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

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            //metoda na doplňování kolonek v tabulce podle vyplněných údajů
            int row1 = e.RowIndex;
            int collum1 = e.ColumnIndex;

            //automatické doplnění full jména
            if (row1 == 0 & collum1 == dataTableShow1.Columns.IndexOf("Plné jméno"))
            {

                if (dataTableShow1.Rows[row1][collum1].ToString() == "")
                {
                    string first = dataTableShow1.Rows[row1][dataTableShow1.Columns.IndexOf("Křestní")].ToString();
                    string second = dataTableShow1.Rows[row1][dataTableShow1.Columns.IndexOf("Příjmení")].ToString();
                    if (first != "" & second != "")
                    {
                        dataTableShow1.Rows[row1][collum1] = second + " " + first;
                    }

                }
            }

            //automatické doplnění křestní a příjmení (z full jména)
            if (row1 == 0 & collum1 == dataTableShow1.Columns.IndexOf("Plné jméno"))
            {
                if (dataTableShow1.Rows[row1][collum1].ToString() != "")
                {
                    string fullName = dataTableShow1.Rows[row1][dataTableShow1.Columns.IndexOf("Plné jméno")].ToString();
                    string[] stringSeparators = new string[] { " " };
                    //rozdělí vstupní jméno
                    string[] result = fullName.Split(stringSeparators, StringSplitOptions.None);
                    //zkontroluje podmínky a přiřadí
                    if (result.Length >= 2)
                    {
                        string firstNameOld = dataTableShow1.Rows[row1][dataTableShow1.Columns.IndexOf("Křestní")].ToString();
                        string secondNameOld = dataTableShow1.Rows[row1][dataTableShow1.Columns.IndexOf("Příjmení")].ToString();

                        string firstName = result[1];
                        string secondName = result[0];

                        if (firstNameOld == "")
                        {
                            dataTableShow1.Rows[row1][dataTableShow1.Columns.IndexOf("Křestní")] = firstName;
                        }
                        if (secondNameOld == "")
                        {
                            dataTableShow1.Rows[row1][dataTableShow1.Columns.IndexOf("Příjmení")] = secondName;
                        }
                    }
                }
            }

            //automatické doplnění Username
            if (row1 == 0 & collum1 == dataTableShow1.Columns.IndexOf("Username"))
            {

                if (dataTableShow1.Rows[row1][collum1].ToString() == "")
                {
                    string first = dataTableShow1.Rows[row1][dataTableShow1.Columns.IndexOf("Křestní")].ToString();
                    string second = dataTableShow1.Rows[row1][dataTableShow1.Columns.IndexOf("Příjmení")].ToString();

                    first = RemoveDiacritics(first);
                    second = RemoveDiacritics(second);

                    if (first != "" & second != "")
                    {
                        dataTableShow1.Rows[row1][collum1] = first.ToLower()[0] + second.ToLower();
                    }

                }
            }


        }

        public static string RemoveDiacritics(string s)
        {
            s = s.Normalize(NormalizationForm.FormD);
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < s.Length; i++)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(s[i]) != UnicodeCategory.NonSpacingMark) sb.Append(s[i]);
            }

            return sb.ToString();
        }

        private void ts_TextBox3_MouseDown(object sender, MouseEventArgs e)
        {
            //vytvoří menu na vložení přednastavených věcí

            if (e.Button == MouseButtons.Right)
            {
                //contextMenuStrip1.(Cursor.Position.X, Cursor.Position.Y);
                TS_userSetting.ShowDropDown();
                contextMenuStrip1.Show(Cursor.Position.X - 150, Cursor.Position.Y);

                //ContextMenuStrip contexMenuuu = new ContextMenuStrip();
                //contexMenuuu.Items.Add("Edit ");
                //contexMenuuu.Items.Add("Delete ");
                //contexMenuuu.Show();
                //contexMenuuu.ItemClicked += new ToolStripItemClickedEventHandler(
                //    userContainerToolStripMenuItem_Click);
            }
        }

        private void userContainerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ts_TextBox3.Text = "OU=Users,OU=People,OU=Company,DC=sitel,DC=cz";
            TS_userSetting.ShowDropDown();
        }

        private void testContainerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ts_TextBox3.Text = "OU=Test,OU=Service,OU=Company,DC=sitel,DC=cz";
            TS_userSetting.ShowDropDown();
        }

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ts_TextBox3.Text = "";
            TS_userSetting.ShowDropDown();
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
    }
}