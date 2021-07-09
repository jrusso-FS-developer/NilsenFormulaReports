using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace Nilsen.Framework.Objects.Interfaces
{
    public interface IRace
    {
        String GetAgeOfRace();
        String Name { get; set; }
        String Type { get; set; }
        String DateText { get; set; }
        DateTime Date { get; set; }
        List<IHorse> Horses { get; set; }
        ITrack Track { get; set; }
        String PostTime { get; set; }
        string PAR { get; set; }
        void SortHorses();
        Decimal GetTop3Total();
        int GetGreatestKeyTrainerStatCategoryCount();
    }
}
