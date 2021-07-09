using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Nilsen.Framework.Objects.Interfaces
{
    public interface IHorse
    {
        decimal BCR { get; set; }
        short Blinkers { get; set; }
        decimal BSR { get; set; }
        decimal ClaimingPrice { get; set; }
        decimal ClaimingPriceLastRace { get; set; }
        decimal CP { get; set; }
        decimal CR { get; set; }
        decimal CRF { get; set; }
        string DIS { get; set; }
        decimal DIS_Value { get; set; }
        int Distance { get; set; } //#F Column
        int DSLR { get; set; }
        decimal DSR { get; set; }
        decimal DST { get; set; }
        decimal? E2_1 { get; set; }
        decimal? E2_2 { get; set; }
        decimal Earnings { get; set; }
        string ExtendedComment { get; set; }
        int Field96 { get; set; }
        string HorseName { get; set; }
        decimal HT { get; set; }
        string HTDisplay { get; set; }
        int JockeyMeetStarts { get; set; }
        List<string> KeyTrainerStatCategory { get; set; }
        decimal LastPurse { get; set; }
        int LP { get; set; }
        decimal LR { get; set; }
        string MDC { get; set; } //Major Drop in Class
        string MJS { get; set; } //Major Jockey Switch
        decimal MJS1156 { get; set; }
        decimal MJS1157 { get; set; }
        decimal MJS1158 { get; set; }
        decimal MJS1159 { get; set; }
        decimal MJS1161 { get; set; }
        decimal MJS1162 { get; set; }
        decimal MJS1163 { get; set; }
        decimal MJS1164 { get; set; }
        decimal MorningLine { get; set; }
        int MountCount { get; set; }
        decimal MUD { get; set; }
        decimal MUD_SR { get; set; }
        decimal MUD_ST { get; set; }
        decimal MUD_W { get; set; }
        decimal NilsenRating { get; set; }
        string Note { get; set; }
        string Note2 { get; set; }
        string Note3 { get; set; }
        int Pace { get; set; }
        int Place { get; set; }
        int PostPoints { get; set; }
        decimal PPWR { get; set; }
        string ProgramNumber { get; set; }
        int Quirin { get; set; }
        decimal RacePurse { get; set; }
        int Rank { get; set; } //RK Column
        decimal RBCPercent { get; set; }
        string RET { get; set; } //Start of Layoff
        decimal RnkWrkrsPct { get; set; }
        int RunStyle { get; set; }
        int Show { get; set; }
        decimal SR { get; set; }
        decimal TB { get; set; }
        int TFW { get; set; }
        decimal Total { get; set; }
        string Trk { get; set; }
        decimal Trk_Value { get; set; }
        decimal TRF { get; set; }
        decimal TSR { get; set; }
        int TurfStarts { get; set; }
        bool TopCR { get; set; }
        int TurfPedigree { get; set; }
        string TurfPedigreeDisplay { get; set; }
        int Wins { get; set; }
        decimal WinPercent { get; set; }
        decimal WinPlacePercent { get; set; }
        decimal WinPlaceShowPercent { get; set; }
        int Workers { get; set; } //Workers of Workout
        int Workout { get; set; } //#W Column
    }
}
