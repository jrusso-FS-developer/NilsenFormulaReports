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

        public struct NilsenCell
        {
            public struct BorderStyle
            {
                public struct Weight
                {
                    public const string Bold = "Bold";
                    public const string Regular = "Thin";
                }
                public struct Sides
                {
                    public const string Bottom = "Bottom";
                    public const string Top = "Top";
                    public const string Left = "Left";
                    public const string Right = "Right";
                    public const string All = "All";
                }
            }
        }

        public enum FormTypes
        {
            PaceForecasterFormula, TurfFormula
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
            IsTrue,
            LessThanZero,
            LowestValue,
            None,
            SecondHighestValue,
            Top4,
            Top5,
            RnkWrkrsCustom,
            ValueNotInStringList,
            WithinRange,
            WithinRangeOfOtherFieldValue,
            WithinRangeOfLastHorseInTopFive
        }

        public enum FormatTypes
        {
            Text, Background, Borders
        }

        public struct TurfFormulaFormatFields
        {
            public const string SR = "SR";
            public const string TurfPedigree = "Turf Ped.";
            public const string TFW = "TFW";
            public const string E2_1 = "E2_1";
            public const string E2_2 = "E2_2";
        }

        public struct PaceForecasterFormatFields
        {
            public const string BCR = "B-CR";
            public const string BSR = "B-SR";
            public const string CR = "CR";
            public const string CRF = "CRF";
            public const string CP = "CP";
            public const string DSLR = "DSLR";
            public const string DIS = "DIS";
            public const string DSR = "DSR";
            public const string DST = "DST";
            public const string Distance = "F";
            public const string KeyTrainerStatCategory1 = "KeyTrainerStatCategories1";
            public const string KeyTrainerStatCategory2 = "KeyTrainerStatCategories2";
            public const string KeyTrainerStatCategory3 = "KeyTrainerStatCategories3";
            public const string LP = "LP";
            public const string LR = "LR";
            public const string MDC = "MDC";
            public const string MJS = "MJS";
            public const string ML = "ML";
            public const string MUD = "MUD";
            public const string PP = "PP";
            public const string Pace = "Pace";
            public const string PPWR = "PPwr";
            public const string RBC = "RBC";
            public const string RQ = "R/Q";
            public const string TotalPace = "Total";
            public const string TB = "TB";
            public const string TRF = "TRF";
            public const string TRK = "TRK";
            public const string TSR = "TSR";
            public const string Workout = "W";
            public const string RnkWrkrsPercentage1 = "RnkWrkrsPercentage1";
            public const string RnkWrkrsPercentage2 = "RnkWrkrsPercentage2";
        }

        public static Dictionary<string, Type> GetFieldList(FormTypes formType)
        {
            switch(formType)
            {
                case FormTypes.PaceForecasterFormula:
                    return new Dictionary<string, Type>() {
                        { PaceForecasterFormatFields.BCR, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.BSR, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.CP, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.CR, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.CRF, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.DIS, Type.GetType("System.String") },
                        { PaceForecasterFormatFields.DSLR, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.DSR, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.DST, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.Distance, Type.GetType("System.Int32") },
                        { PaceForecasterFormatFields.KeyTrainerStatCategory1, Type.GetType("System.String") },
                        { PaceForecasterFormatFields.KeyTrainerStatCategory2, Type.GetType("System.String") },
                        { PaceForecasterFormatFields.KeyTrainerStatCategory3, Type.GetType("System.String") },
                        { PaceForecasterFormatFields.MDC, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.MJS, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.LP, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.LR, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.ML, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.MUD, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.Pace, Type.GetType("System.Int32") },
                        { PaceForecasterFormatFields.PP, Type.GetType("System.Int32") },
                        { PaceForecasterFormatFields.PPWR, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.RQ, Type.GetType("System.Int32") },
                        { PaceForecasterFormatFields.RBC, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.TB, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.TRF, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.TRK, Type.GetType("System.String") },
                        { PaceForecasterFormatFields.TSR, Type.GetType("System.Decimal") },
                        { PaceForecasterFormatFields.TotalPace, Type.GetType("System.Decimal") } ,
                        { PaceForecasterFormatFields.RnkWrkrsPercentage1, Type.GetType("System.Boolean") },
                        { PaceForecasterFormatFields.RnkWrkrsPercentage2, Type.GetType("System.Boolean") },
                        { PaceForecasterFormatFields.Workout, Type.GetType("System.Int32") }
                    };
                case FormTypes.TurfFormula:
                    return new Dictionary<string, Type>() {
                        { TurfFormulaFormatFields.SR, Type.GetType("System.Decimal") },
                        { TurfFormulaFormatFields.TurfPedigree, Type.GetType("System.Decimal") },
                        { TurfFormulaFormatFields.TFW, Type.GetType("System.Decimal") },
                        { TurfFormulaFormatFields.E2_1, Type.GetType("System.Decimal") },
                        { TurfFormulaFormatFields.E2_2, Type.GetType("System.Decimal") },
                    };
            }

            return new Dictionary<string, Type>();
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
        public List<string> CellStyles { get; set; }
        public RaceFieldsFormat.BasisTypes BasisType { get; set; }
        public RaceFieldsFormat.FormatTypes FormatType { get; set; }
        public string Field { get; set; }
        public int WsColumnIndex { get; set; }
        public List<string> EvaluationStringValues { get; set; }
        public List<decimal> EvaluationDecimalValues { get; set; }
        public List<decimal> HorseValues { get; set; }
        public RangeValues<decimal, decimal> EvaluationRangeValues { get; set; }
        #endregion

        #region Constructors
        public FieldFormat()
        {
            EvaluationDecimalValues = new List<decimal>();
            HorseValues = new List<decimal>();
        }
        #endregion
    }

    [Serializable]
    public class RangeValues<TRangeStart, TRangeEnd>
    {
        public RangeValues(TRangeStart rangeStart, TRangeEnd rangeEnd)
        {
            RangeStart = rangeStart;
            RangeEnd = rangeEnd;
        }
        public TRangeStart RangeStart { get; set; }
        public TRangeEnd RangeEnd { get; set; }
    }
}
