using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Nilsen.Framework.Objects.Interfaces
{
    public interface IHorse
    {
        Decimal BCR { get; set; }
        Int16 Blinkers { get; set; }
        Decimal BSR { get; set; }
        Decimal ClaimingPrice { get; set; }
        Decimal ClaimingPriceLastRace { get; set; }
        Decimal CP { get; set; }
        Decimal CR { get; set; }
        String DIS { get; set; }
        Int32 Distance { get; set; } //#F Column
        Int32 DSLR { get; set; }
        Decimal DSR { get; set; }
        String ExtendedComment { get; set; }
        Decimal PPWR { get; set; }
        String HorseName { get; set; }
        List<String> KeyTrainerStatCategory { get; set; }
        Decimal LastPurse { get; set; }
        Int32 LP { get; set; }
        String MDC { get; set; } //Major Drop in Class
        String MJS { get; set; } //Major Jockey Switch
        decimal MJS1156 { get; set; }
        decimal MJS1157 { get; set; }
        decimal MJS1158 { get; set; }
        decimal MJS1159 { get; set; }
        decimal MJS1161 { get; set; }
        decimal MJS1162 { get; set; }
        decimal MJS1163 { get; set; }
        decimal MJS1164 { get; set; }
        Decimal MorningLine { get; set; }
        String Note { get; set; }
        String Note2 { get; set; }
        String Note3 { get; set; }
        Int32 Pace { get; set; }
        Int32 PostPoints { get; set; }
        String ProgramNumber { get; set; }
        Int32 Quirin { get; set; }
        Decimal RacePurse { get; set; }
        Int32 Rank { get; set; } //RK Column
        Decimal RBCPercent { get; set; }
        String RET { get; set; } //Start of Layoff
        Decimal RnkWrkrsPct { get; set; }
        Int32 RunStyle { get; set; }
        Decimal TB { get; set; }
        Decimal Total { get; set; }
        String Trk { get; set; }
        Decimal TSR { get; set; }
        Int32 Workers { get; set; } //Workers of Workout
        Int32 Workout { get; set; } //#W Column
    }
}
