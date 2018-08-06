using Nilsen.Framework.Objects.Interfaces;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;

namespace Nilsen.Framework.Objects.Class
{
    public sealed class Race : IRace
    {      
        public Race(string[] Fields)
        {
            Track = new Track(Fields, Convert.ToDecimal(Fields[5]));
            Name = Fields[2];
            DateText = Fields[1];
            Date = DateTime.ParseExact(Fields[1], "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None);
            Type = Fields[8];
            AgeOfRace = Fields[9];
            PostTime = Fields[1373];

            Horses = new List<IHorse>();
        }

        public string GetAgeOfRace()
        {
            List<char> chars = AgeOfRace.ToList();
            var sbAgeOfRace = new StringBuilder();

            switch (chars[0].ToString().ToUpper())
            {
                case "A":
                    sbAgeOfRace.Append("2 Year Olds ");
                    break;
                case "B":
                    sbAgeOfRace.Append("3 Year Olds ");
                    break;
                case "C":
                    sbAgeOfRace.Append("4 Year Olds ");
                    break;
                case "D":
                    sbAgeOfRace.Append("5 Year Olds ");
                    break;
                case "E":
                    sbAgeOfRace.Append("3 & 4 Year Olds ");
                    break;
                case "F":
                    sbAgeOfRace.Append("4 & 5 Year Olds ");
                    break;
                case "G":
                    sbAgeOfRace.Append("3, 4 & 5 Year Olds ");
                    break;
                case "H":
                    sbAgeOfRace.Append("All Ages ");
                    break;
            }

            switch (chars[1].ToString().ToUpper())
            {
                case "O":
                    sbAgeOfRace.Append("Only - ");
                    break;
                case "U":
                    sbAgeOfRace.Append("and Up - ");
                    break;
            }

            switch (chars[2].ToString().ToUpper())
            {
                case "N":
                    sbAgeOfRace.Append("No Sex Restrictions");
                    break;
                case "M":
                    sbAgeOfRace.Append("Mares and Fillies Only");
                    break;
                case "C":
                    sbAgeOfRace.Append("Colts and/or Geldings Only");
                    break;
                case "F":
                    sbAgeOfRace.Append("Fillies Only");
                    break;
            }

            return sbAgeOfRace.ToString();
        }

        private string AgeOfRace { get; set; }

        public string Name { get; set; }

        public string Type { get; set; }

        public string DateText { get; set; }

        public DateTime Date { get; set; }

        public List<IHorse> Horses { get; set; }

        public ITrack Track { get; set; }

        public string PostTime { get; set; }

        public void SortHorses()
        {
            //declares and assigns
            var dt = new DataTable();
            var iHorseIndex = 0;
            dt.Columns.Add(new DataColumn("Total", System.Type.GetType("System.Int32")));
            dt.Columns.Add(new DataColumn("Horse", System.Type.GetType("System.Object")));

            //process
            //add horses to the datatable
            foreach (var h in Horses)
            {
                dt.Rows.Add(dt.NewRow());
                dt.Rows[iHorseIndex][0] = h.Total;
                dt.Rows[iHorseIndex++][1] = h;
            }

            //clear all horses
            Horses.Clear();

            //sort the horses in the table
            dt.DefaultView.Sort = "Total desc";
            var dtHorses = dt.DefaultView.ToTable();

            //now readd the sorted horses.  
            foreach (DataRow dr in dtHorses.Rows)
            {
                Horses.Add((IHorse)dr[1]);
            }
        }

        public Decimal GetTop3Total()
        {
            //declares and assigns
            var Total = (Decimal)0.00;
            var iIndex = 0;

            //process
            foreach(var h in Horses)
            {
                if (iIndex < 3)
                {
                    Total += h.Total;
                    iIndex++;
                }  
            }

            //returns
            return Total;
        }

        public int GetGreatestKeyTrainerStatCategoryCount()
        {
            var greatestCount = 0;

            foreach(var horse in Horses){

                if (greatestCount < horse.KeyTrainerStatCategory.Count)
                {
                    greatestCount = horse.KeyTrainerStatCategory.Count;
                }
            }

            return greatestCount;
        }

    }
}
