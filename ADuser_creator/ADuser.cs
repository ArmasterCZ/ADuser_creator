﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExtensionMethods;

namespace ADuser_creator
{
    class ADuser
    {

        public ADuser()
        {
            cleanse();
        }

        public ADuser(ADuser user1)
        {
            //create ADuser from Another ADuser
            this.nameGiven = user1.nameGiven;
            this.nameSurn = user1.nameSurn;
            this.nameFull = user1.nameFull;
            this.nameAcco = user1.nameAcco;
            this.namePrincipal = user1.namePrincipal;
            this.description = user1.description;
            this.title = user1.title;
            this.office = user1.office;
            this.department = user1.department;
            this.tel = user1.tel;
            this.telOthers = user1.telOthers;
            this.mob = user1.mob;
            this.cardNumber = user1.cardNumber;
            this.cardHave = user1.cardHave;
            this.password = user1.password;
            this.emailAddress = user1.emailAddress;
            this.manager = user1.manager;
            this.managerFull = user1.managerFull;
            this.path = user1.path;
            this.group = user1.group;
            this.ChangePasswordAtLogon = user1.CannotChangePassword;
            this.CannotChangePassword = user1.CannotChangePassword;
            this.PasswordNeverExpires = user1.PasswordNeverExpires;
            this.Enabled = user1.Enabled;
        }

        public ADuser(string nameAcco)
        {
            //create ADuser with sAMAccountname
            this.nameGiven = "";
            this.nameSurn = "";
            this.nameFull = "";
            this.nameAcco = nameAcco;
            this.namePrincipal = "";
            this.description = "";
            this.title = "";
            this.office = "";
            this.department = "";
            this.tel = "";
            this.telOthers = "";
            this.mob = "";
            this.cardNumber = "";
            this.cardHave = "";
            this.password = "";
            this.emailAddress = "";
            this.manager = "";
            this.managerFull = "";
            this.path = "";
            this.group = "";
            this.ChangePasswordAtLogon = false;
            this.CannotChangePassword = false;
            this.PasswordNeverExpires = false;
            this.Enabled = false;
        }

        public string nameGiven; //křestní
        public string nameSurn; //příjmení
        public string nameFull; //Příjmení + křestní
        public string nameAcco; //uživatelské jméno (jvaldauf)

        private string namePrincipal; //Sam + @sitel.cz
        public string description;  // pozice (ADMINISTRATOR) [= description] 
        public string title;        //pozice (ADMINISTRATOR) [= Title] 
        public string office;   //středisko (17120)
        public string department;   // občas např (17000)
        public string tel;      //telefon
        public string telOthers;//další telefony
        public string mob;      //mobil
        public string cardNumber;   //číslo karty
        public string cardHave;     //info ke katě
        public string password;     //heslo
        public string emailAddress; //email
        public string manager;      //vedoucí
        public string managerFull;  //vedoucí DistinguishedName (CN=test User 8,OU=Test,OU=Service,OU=Company,DC=sitel,DC=cz)
        public string path;         //cesta ke kontejneru
        public string group;         //skupina
        public bool ChangePasswordAtLogon;
        public bool CannotChangePassword;
        public bool PasswordNeverExpires;
        public bool Enabled;

        //public string QnameGiven
        //{
        //    get
        //    {
        //        return @"""" + this.nameGiven + @"""";
        //    }
        //}

        //public string QnameSurn
        //{
        //    get
        //    {
        //        return @"""" + this.nameSurn + @"""";
        //    }
        //}

        //public string QnameFull
        //{
        //    get
        //    {
        //        return @"""" + this.nameFull + @"""";
        //    }
        //}

        //public string QnameAcco
        //{
        //    get
        //    {
        //        return @"""" + this.nameAcco + @"""";
        //    }
        //}

        //public string Qdescription
        //{
        //    get
        //    {
        //        return @"""" + this.description + @"""";
        //    }
        //}

        //public string Qoffice
        //{
        //    get
        //    {
        //        return @"""" + this.office + @"""";
        //    }
        //}

        //public string Qtel
        //{
        //    get
        //    {
        //        return @"""" + this.tel + @"""";
        //    }
        //}

        //public string Qmob
        //{
        //    get
        //    {
        //        return @"""" + this.mob + @"""";
        //    }
        //}

        //public string QcardNumber
        //{
        //    get
        //    {
        //        return @"""" + this.cardNumber + @"""";
        //    }
        //}

        //public string QcardFullNumber
        //{
        //    get
        //    {
        //        return @"""" + converterToFullCardNumber(this.cardNumber) + @"""";
        //    }
        //}

        //public string QcardHave
        //{
        //    get
        //    {
        //        return @"""" + this.cardHave + @"""";
        //    }
        //}

        //public string Qpassword
        //{
        //    get
        //    {
        //        return @"""" + this.password + @"""";
        //    }
        //}

        public string cardFullNumber
        {
            set
            {
                this.cardNumber = converterFromFullCardNumber(value);
            }
            get
            {
                return converterToFullCardNumber(this.cardNumber);
            }
        }

        public string NamePrincipal
        {
            get
            {
                return (this.nameAcco + "@sitel.cz");
            }
        }

        public string ADgroup
        {
            set
            {
                //remove last character if string exist (for input from PS)
                string inputGroup = value;
                if (!(inputGroup == null | inputGroup == ""))
                {
                    this.group = inputGroup.Remove(inputGroup.Length - 1);
                }
                else
                {
                    this.group = "";
                }
            }
            get
            {
                //from ( group1,group2 ) create ( "group1","group2" )
                string startGroup = this.group;
                string finalGroup = "";

                if (startGroup != "" | startGroup != null)
                {
                    //split and insert quota
                    if (startGroup.Contains(","))
                    {
                        string[] separatedGroups = startGroup.Split(',');
                        foreach (string group in separatedGroups)
                        {
                            finalGroup += group.addQuotes() + ",";
                        }
                        finalGroup = finalGroup.Remove(finalGroup.Length - 1);
                    }
                    else
                    {
                        finalGroup = startGroup.addQuotes();
                    }
                } else
                {
                    finalGroup = "";
                }
                return (finalGroup);
            }
        }


        /*public string[] toField()
        {
            string[] field = { nameGiven, nameSurn, nameFull, nameAcco, office, description, tel, mob, cardNumber, password };
            return field;
        }*/

        public void cleanse()
        {
            //erase all data from ADuser
            this.nameGiven = "";
            this.nameSurn = "";
            this.nameFull = "";
            this.nameAcco = "";
            this.namePrincipal = "";
            this.description = "";
            this.title = "";
            this.office = "";
            this.department = "";
            this.tel = "";
            this.telOthers = "";
            this.mob = "";
            this.cardNumber = "";
            this.cardHave = "";
            this.password = "";
            this.emailAddress = "";
            this.manager = "";
            this.managerFull = "";
            this.path = "";
            this.group = "";
            this.ChangePasswordAtLogon = false;
            this.CannotChangePassword = false;
            this.PasswordNeverExpires = false;
            this.Enabled = false;
        }

        public override string ToString()
        {
            return "User: " + nameFull + " .";
        }

        private string converterToFullCardNumber(string number)
        {
            //převádí číslo z desítkové soustavy do Hexadecimální
            string finalNumber = "81AE04C300000";
            try
            {
                int number2 = Int32.Parse(number);
                finalNumber += number2.ToString("X");
            }
            catch (Exception e)
            {
                throw e;
            }

            return finalNumber;
        }

        private string converterFromFullCardNumber(string number)
        {
            //převádí číslo z Hexadecimální soustavy do desítkové
            string finalNumber = "";
            try
            {
                //if (number.Length == 16)
                if (number.Length >= 4)
                {
                    string lastCharacters = number.Substring(number.Length - 3);
                    finalNumber = "" + int.Parse(lastCharacters, System.Globalization.NumberStyles.HexNumber).ToString();
                }
            }
            catch (Exception e)
            {
                throw e;
            }

            return finalNumber;
        }
        
    }

}

