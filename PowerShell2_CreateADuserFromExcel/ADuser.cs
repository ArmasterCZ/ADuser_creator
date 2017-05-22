using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerShell2_CreateADuserFromExcel
{
    class ADuser
    {
        public string nameGiven; //křestní
        public string nameSurn; //příjmení
        public string nameFull; //Příjmení + křestní
        public string nameAcco; //uživatelské jméno (jvaldauf)
        public string description;  // pozice (ADMINISTRATOR) [= Title] 
        public string office;   //středisko (17120)
        public string tel;      //telefon
        public string mob;      //mobil
        public string cardNumber;   //číslo karty
        public string cardHave;     //info ke katě
        public string leader;       //vedoucí
        public string password;     //heslo
        //public string title;    //pozice (ADMINISTRATOR) [= description] 
        //public string department;   // občas např (17000)

        public string QnameGiven
        {
            get
            {
                return @"""" + this.nameGiven + @"""";
            }
        }

        public string QnameSurn
        {
            get
            {
                return @"""" + this.nameSurn + @"""";
            }
        }

        public string QnameFull
        {
            get
            {
                return @"""" + this.nameFull + @"""";
            }
        }

        public string QnameAcco
        {
            get
            {
                return @"""" + this.nameAcco + @"""";
            }
        }

        public string Qdescription
        {
            get
            {
                return @"""" + this.description + @"""";
            }
        }

        public string Qoffice
        {
            get
            {
                return @"""" + this.office + @"""";
            }
        }

        public string Qtel
        {
            get
            {
                return @"""" + this.tel + @"""";
            }
        }

        public string Qmob
        {
            get
            {
                return @"""" + this.mob + @"""";
            }
        }

        public string QcardNumber
        {
            get
            {
                return @"""" + this.cardNumber + @"""";
            }
        }

        public string QcardFullNumber
        {
            get
            {
                return @"""" + converterToCardNumber(this.cardNumber) + @"""";
            }
        }

        public string QcardHave
        {
            get
            {
                return @"""" + this.cardHave + @"""";
            }
        }

        public string Qleader
        {
            get
            {
                return @"""" + this.leader + @"""";
            }
        }

        public string Qpassword
        {
            get
            {
                return @"""" + this.password + @"""";
            }
        }

        // ----- metody ------

        public ADuser()
        {
            cleanse();

            /*
            $ADaccountName = "pokres"
            $krestni = "Vlachová"
            $prijmeni = "Regina"
            $displayName = "Vlachová Regina"
            $office = "17120"
            $pozice = "recepce"
            $description = "SafeQ"
            $department = "17000"
            $heslo = "123456Aa"
            $tel = "3 421"
            $mob = "+420 702 197 082"
             */
        }

        public ADuser(ADuser user1)
        {
            this.nameGiven = user1.nameGiven;
            this.nameSurn = user1.nameSurn;
            this.nameFull = user1.nameFull;
            this.nameAcco = user1.nameAcco;
            this.description = user1.description;
            this.office = user1.office;
            this.leader = user1.leader;
            this.tel = user1.tel;
            this.mob = user1.mob;
            this.cardNumber = user1.cardNumber;
            this.cardHave = user1.cardHave;
            this.password = user1.password;
        }

        public ADuser(string nameAcco)
        {
            this.nameGiven = "";
            this.nameSurn = "";
            this.nameAcco = nameAcco;
            this.nameFull = "";
            this.tel = "";
            this.mob = "";
            this.cardNumber = "";
            this.cardHave = "";
            this.leader = "";
            this.description = "";
            this.office = "";
            this.password = "";
        }

        public ADuser(string nameGiven, string nameSurn, string nameFull, string nameAcco, string office, string description, string tel, string mob, string cardNumber, string password)
        {
            this.nameGiven = nameGiven;
            this.nameSurn = nameSurn;
            this.nameAcco = nameAcco;
            this.nameFull = nameFull;
            this.tel = tel;
            this.mob = mob;
            this.cardNumber = cardNumber;
            this.cardHave = "";
            this.leader = "";
            this.description = description;
            this.office = office;
            this.password = password;
        }

        public string[] toField()
        {
            string[] field = { nameGiven, nameSurn, nameFull, nameAcco, office, description, tel, mob, cardNumber, password };
            return field;
        }

        public void cleanse()
        {
            this.nameGiven = "";
            this.nameSurn = "";
            this.nameFull = "";
            this.nameAcco = "";
            this.description = "";
            this.office = "";
            this.leader = "";
            this.tel = "";
            this.mob = "";
            this.cardNumber = "";
            this.cardHave = "";
            this.password = "";
        }

        public override string ToString()
        {
            return "User: " + nameFull + " .";
        }

        private string converterToCardNumber(string number)
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

    }

}

