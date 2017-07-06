using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Nilsen.Framework.Common
{
    [Serializable]
    public static class RaceFieldsFormat
    {

        public struct SortDirections
        {
            public const string Asc = "asc";
            public const string Desc = "desc";
        }

        public struct Text
        {
            public struct Style
            {
                public const string Bold = "Bold";
                public const string Regular = "Regular";
                public const string Italic = "Italic";
            }
        }

        public enum BasisTypes
        {
            BaseAmountOrHigher,
            BaseAmountOrLower,
            BetweenTwoValues,
            Equals,
            GreaterThanOrEqualTo,
            HighestValue,
            HighestValueWithinFloorRange,
            LessThanZero,
            LowestValue,
            None,
            SecondHighestValue,
            Top5,
            RnkWrkrsCustom,
            ValueExists,
            WithinRange,
            WithinRangeOfOtherFieldValue,
            WithinRangeOfLastHorseInTopFive
        }

        public enum FormatTypes
        {
            Text, Background, Borders
        }

        public struct Fields
        {
            public const string BCR = "B-CR";
            public const string BSR = "B-SR";
            public const string CR = "CR";
            public const string DSLR = "DSLR";
            public const string DSR = "DSR";
            public const string Distance = "F";
            public const string KeyTrainerStatCategory1 = "KeyTrainerStatCategories1";
            public const string KeyTrainerStatCategory2 = "KeyTrainerStatCategories2";
            public const string KeyTrainerStatCategory3 = "KeyTrainerStatCategories3";
            public const string LP = "LP";
            public const string MDC = "MDC";
            public const string MJS = "MJS";
            public const string ML = "ML";
            public const string PP = "PP";
            public const string Pace = "Pace";
            public const string PPWR = "PPwr";
            public const string RBC = "RBC";
            public const string RQ = "R/Q";
            public const string TotalPace = "Total";
            public const string TB = "TB";
            public const string TSR = "TSR";
            public const string Workout = "W";
            public const string RnkWrkrsPercentage1 = "RnkWrkrsPercentage1";
            public const string RnkWrkrsPercentage2 = "RnkWrkrsPercentage2";
        }

        public static Dictionary<string, Type> GetFieldList()
        {
            var list = new Dictionary<string, Type>() { 
                                             { Fields.BCR, System.Type.GetType("System.Decimal") }, 
                                             { Fields.BSR, System.Type.GetType("System.Decimal") }, 
                                             { Fields.CR, System.Type.GetType("System.Decimal") }, 
                                             { Fields.DSLR, System.Type.GetType("System.Decimal") },
                                             { Fields.DSR, System.Type.GetType("System.Decimal") },
                                             { Fields.Distance, System.Type.GetType("System.Int32") },
                                             { Fields.KeyTrainerStatCategory1, System.Type.GetType("System.String") },
                                             { Fields.KeyTrainerStatCategory2, System.Type.GetType("System.String") },
                                             { Fields.KeyTrainerStatCategory3, System.Type.GetType("System.String") },
                                             { Fields.MDC, System.Type.GetType("System.Decimal") },
                                             { Fields.MJS, System.Type.GetType("System.Decimal") },
                                             { Fields.LP, System.Type.GetType("System.Decimal") },
                                             { Fields.ML, System.Type.GetType("System.Decimal") },
                                             { Fields.Pace, System.Type.GetType("System.Int32") },
                                             { Fields.PP, System.Type.GetType("System.Int32") },
                                             { Fields.PPWR, System.Type.GetType("System.Decimal") },
                                             { Fields.RQ, System.Type.GetType("System.Int32") },
                                             { Fields.RBC, System.Type.GetType("System.Decimal") },
                                             { Fields.TB, System.Type.GetType("System.Decimal") },
                                             { Fields.TSR, System.Type.GetType("System.Decimal") },
                                             { Fields.TotalPace, System.Type.GetType("System.Decimal") } ,
                                             { Fields.RnkWrkrsPercentage1, System.Type.GetType("System.Boolean") },
                                             { Fields.RnkWrkrsPercentage2, System.Type.GetType("System.Boolean") },
                                             { Fields.Workout, System.Type.GetType("System.Int32") } };

            return list;
        }
    }

    [Serializable]
    public class FieldFormat
    {
        #region Public Members
        public List<int> AdditionalColumnsToAffect { get; set; }
        public XlRgbColor BackgroundColor { get; set; }
        public string SortDirection { get; set; }
        public XlRgbColor TextColor { get; set; }
        public List<string> TextStyles { get; set; }
        public RaceFieldsFormat.BasisTypes BasisType { get; set; }
        public RaceFieldsFormat.FormatTypes FormatType { get; set; }
        public string Field { get; set; }
        public int WsColumnIndex { get; set; }
        public List<decimal> EvaluationValues { get; set; }
        public List<decimal> HorseValues { get; set; }
        public RangeValues<decimal, decimal> EvaluationRangeValues { get; set; }
        #endregion

        #region Constructors
        public FieldFormat()
        {
            EvaluationRangeValues = new RangeValues<decimal, decimal>();
            EvaluationValues = new List<decimal>();
            HorseValues = new List<decimal>();
        }
        #endregion
    }

    [Serializable]
    public class RangeValues<TRangeStart, TRangeEnd>
    {
        public TRangeStart RangeStart { get; set; }
        public TRangeEnd RangeEnd { get; set; }
    }
}
