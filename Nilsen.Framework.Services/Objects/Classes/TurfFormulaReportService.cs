using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.FileIO;
using Nilsen.Framework.Common;
using Nilsen.Framework.Objects.Class;
using Nilsen.Framework.Objects.Interfaces;
using Nilsen.Framework.Services.Objects.Interfaces;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using static Nilsen.Framework.Common.RaceFieldsFormat;

namespace Nilsen.Framework.Services.Objects.Classes
{
    public class TurfFormulaReportService : IReportService
    {
        private ConsoleService consoleSvc;
        private String _SavePath = string.Format("{0}\\Flicker City Productions\\RacesCSVToExcel\\files", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

        public TurfFormulaReportService(System.Windows.Forms.TextBox consoleWindow, System.Windows.Forms.Button btnProcess)
        {
            consoleSvc = new ConsoleService(consoleWindow, btnProcess);
        }

        public void CreateExcelFile(FileInfo fi)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            var wb = ExcelApp.Workbooks.Add(Type.Missing);
            ExcelApp.DisplayAlerts = false;
            Worksheet ws = new Worksheet();
            StringBuilder sbFullFileName = new StringBuilder();
            StringBuilder sbFileName = new StringBuilder();
            StringBuilder sbPath = new StringBuilder();
            Int32 iLength = fi.Name.Split('.').GetLength(0);
            Boolean bContinue = true;

            consoleSvc.ToggleProcessButton(false);

            sbFileName.Append(fi.Name.Split('.')[0]);
            sbFullFileName.AppendFormat("{0}\\NilsenRatings_TurfFormula_{1}", _SavePath, sbFileName.ToString());

            //Build Worksheet
            consoleSvc.UpdateConsoleText("Creating Worksheet...", true);
            BuildWorksheet(wb.Sheets.Add(), fi);

            //Clear Extra Worksheet
            wb.Sheets[wb.Sheets.Count].Delete();

            //Save Workbook
            consoleSvc.UpdateConsoleText("Saving File...", false);
            if (!Directory.Exists(_SavePath))
            {
                DirectoryInfo di = Directory.CreateDirectory(_SavePath);
            } 
            else 
            {
                if (File.Exists(string.Format("{0}.xlsx", sbFullFileName.ToString())))
                {
                    var results = MessageBox.Show(string.Format("File '{0}' Exists.\n\nReplace?", string.Format("{0}.xlsx", sbFileName)), "File Exists", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

                    if (results.Equals(DialogResult.Yes))
                    {
                        File.Delete(string.Format("{0}.xlsx", sbFullFileName.ToString()));
                    }
                    else
                    {
                        bContinue = false;
                        consoleSvc.UpdateConsoleText("File Not Saved.", false);
                    }
                }
            }

            if (bContinue)
            {
                wb.SaveAs(sbFullFileName.ToString(), XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, false,
                        XlSaveAsAccessMode.xlShared, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                consoleSvc.UpdateConsoleText(string.Format("File Saved As: {0}.xlsx", sbFullFileName.ToString()), false);
            }

            //Cleanup
            ExcelApp.Application.Visible = true;
            Marshal.ReleaseComObject(ws);
            ExcelApp.Application.Visible = true;
            wb.Close();
            ExcelApp.Application.Quit();
            consoleSvc.UpdateConsoleText("Processing Completed", false);

            consoleSvc.ToggleProcessButton(true);
        }

        public void BuildWorksheet(Worksheet ws, FileInfo fi)
        {
            //declares and assigns
            var reader = new StreamReader(File.OpenRead(fi.FullName));
            string[] Lines;
            string[] Fields;
            var iRow = 1;
            Range rHeader;
            var iTop5Row = 0;
            var iHeaderRow = 0;
            var sAllHorses = new String[100, 2];
            var decAllHorses = new Decimal[100];
            var Top5Horses = new String[5, 2];
            IRace race = null;
            TextFieldParser lineParser = null;
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            ws.Name = "Nilsen Turf Rating";
            ws.get_Range("A1", "E1").Merge(Type.Missing);
            rHeader = ws.get_Range("A1", Type.Missing);
            rHeader.Value = "Nilsen Turf Rating Report";
            rHeader.Font.Bold = true;
            ws.Cells[1, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            Lines = Regex.Split(reader.ReadToEnd(), Environment.NewLine);

            //page style
            ws.PageSetup.Orientation = XlPageOrientation.xlPortrait;

            //race vars
            string sRaceDate = string.Empty;
            string sTrack = string.Empty;
            string sRaceName = string.Empty;

            consoleSvc.UpdateConsoleText("Reading CSV File...", false);

            if (Lines.GetLength(0) > 0)
            {
                foreach (var line in Lines)
                {
                    lineParser = new TextFieldParser(new StringReader(line));
                    lineParser.TextFieldType = FieldType.Delimited;
                    lineParser.SetDelimiters(new string[] { "," });
                    lineParser.HasFieldsEnclosedInQuotes = true;

                    while (!lineParser.EndOfData)
                    {
                        Fields = lineParser.ReadFields();

                        if (Fields[6].ToLower().Equals("t")) //Either 'T' or 't', per spec.  
                        {
                            Decimal furlongs;
                            var sbTurfPed = new StringBuilder();
                            var sbTurfPedChars = new StringBuilder();

                            if (!Fields[2].ToLower().Equals(sRaceName))
                            {
                                if (race != null)
                                {
                                    ws.Cells[iTop5Row, 1].Value = top5AndOr400PlusCalc(race);
                                    ws.get_Range(string.Format("A{0}", iTop5Row), string.Format("D{0}", iTop5Row)).Merge(Type.Missing);
                                    iRow = listHorses(race, ws, iHeaderRow++);
                                }

                                race = RaceService.GetRace(Fields);

                                consoleSvc.UpdateConsoleText(string.Format("Reading and building Race {0}...", Fields[2]), false);
                                iRow = iRow + 2;
                                sRaceDate = DateTime.ParseExact(Fields[1], "yyyyMMdd", CultureInfo.InvariantCulture).ToString("MM/dd/yyyy");
                                sTrack = Fields[0];
                                sRaceName = Fields[2];

                                sAllHorses = new String[100, 2];
                                Top5Horses = new String[5, 2];
                                furlongs = Math.Round(((Convert.ToDecimal(Regex.Replace(Fields[5], "[^.0-9]", "")) * 3) / 5280) * 8, 2);

                                ws.get_Range(string.Format("A{0}", iRow), string.Format("C{0}", iRow)).Merge(Type.Missing);
                                ws.Cells[iRow, 1].Value = sRaceDate;
                                ws.Cells[iRow, 1].Style.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                ws.Cells[++iRow, 1].Value = sTrack;
                                ws.get_Range(string.Format("A{0}", iRow), string.Format("C{0}", iRow)).Merge(Type.Missing);
                                ws.Cells[++iRow, 1].Value = string.Format("Race {0}", Fields[2]);
                                ws.get_Range(string.Format("A{0}", iRow), string.Format("C{0}", iRow)).Merge(Type.Missing);
                                ws.Cells[++iRow, 1].Value = string.Format("{0} Furlongs", furlongs);
                                ws.get_Range(string.Format("A{0}", iRow), string.Format("C{0}", iRow)).Merge(Type.Missing);

                                switch (Fields[8].ToLower())
                                {
                                    case "g1":
                                        ws.Cells[++iRow, 1].Value = "Grade I Stk/Hcp";
                                        break;
                                    case "g2":
                                        ws.Cells[++iRow, 1].Value = "Grade II Stk/Hcp";
                                        break;
                                    case "g3":
                                        ws.Cells[++iRow, 1].Value = "Grade III Stk/Hcp";
                                        break;
                                    case "n":
                                        ws.Cells[++iRow, 1].Value = "Nongraded Stake/Handicap";
                                        break;
                                    case "a":
                                        ws.Cells[++iRow, 1].Value = "Allowance";
                                        break;
                                    case "r":
                                        ws.Cells[++iRow, 1].Value = "Starter Alw";
                                        break;
                                    case "t":
                                        ws.Cells[++iRow, 1].Value = "Starter Hcp";
                                        break;
                                    case "c":
                                        ws.Cells[++iRow, 1].Value = "Claiming";
                                        break;
                                    case "co":
                                        ws.Cells[++iRow, 1].Value = "Optional Clmg";
                                        break;
                                    case "s":
                                        ws.Cells[++iRow, 1].Value = "Mdn Sp Wt";
                                        break;
                                    case "m":
                                        ws.Cells[++iRow, 1].Value = "Ndn Claimer";
                                        break;
                                    case "ao":
                                        ws.Cells[++iRow, 1].Value = "Alw Opt Clm";
                                        break;
                                    case "mo":
                                        ws.Cells[++iRow, 1].Value = "Mdn Opt Clm";
                                        break;
                                    case "no":
                                        ws.Cells[++iRow, 1].Value = "Opt Clm Stk";
                                        break;
                                }

                                ws.get_Range(string.Format("A{0}", iRow), string.Format("C{0}", iRow)).Merge(Type.Missing);
                                ws.Cells[++iRow, 1].Value = string.Format("{0}", Fields[1373]);
                                ws.get_Range(string.Format("A{0}", iRow), string.Format("C{0}", iRow)).Merge(Type.Missing);
                                ws.Cells[++iRow, 1].Value = string.Format("Purse: {0:C0}", Convert.ToDecimal(Fields[11]));
                                ws.get_Range(string.Format("A{0}", iRow), string.Format("C{0}", iRow)).Merge(Type.Missing);
                                ws.Cells[++iRow, 1].Value = string.Format("Race Type: {0}", Fields[8]);
                                ws.get_Range(string.Format("A{0}", iRow), string.Format("C{0}", iRow)).Merge(Type.Missing);

                                iTop5Row = ++iRow;
                                iHeaderRow = iTop5Row + 2;

                                //row headers
                                ws.Cells[iHeaderRow, 1].Value = "Prg #";
                                ws.Cells[iHeaderRow, 2].Value = "ML";
                                ws.Cells[iHeaderRow, 3].Value = "Horse Name";
                                ws.Cells[iHeaderRow, 4].Value = "Turf Rating";
                                ws.Cells[iHeaderRow, 5].Value = "Sts.";
                                ws.Cells[iHeaderRow, 6].Value = "Win";
                                ws.Cells[iHeaderRow, 7].Value = "Win%";
                                ws.Cells[iHeaderRow, 8].Value = "Place";
                                ws.Cells[iHeaderRow, 9].Value = "WP%";
                                ws.Cells[iHeaderRow, 10].Value = "Show";
                                ws.Cells[iHeaderRow, 11].Value = "WPS%";
                                ws.Cells[iHeaderRow, 12].Value = "Earnings";
                                ws.Cells[iHeaderRow, 13].Value = "AE";
                                ws.Cells[iHeaderRow, 14].Value = "SR";
                                ws.Cells[iHeaderRow, 15].Value = "Turf Ped.";
                                ws.Cells[iHeaderRow, 16].Value = "DSLR";
                                ws.Cells[iHeaderRow, 17].Value = string.Empty;
                                ws.Cells[iHeaderRow, 18].Value = "TFW";
                                ws.Cells[iHeaderRow, 19].Value = "E2 (1)";
                                ws.Cells[iHeaderRow, 20].Value = "E2 (2)";
                                ws.Cells[iHeaderRow, 21].Value = ""; // EMPTY COLUMN
                                ws.Cells[iHeaderRow, 22].Value = "MUD-SR";
                                ws.Cells[iHeaderRow, 23].Value = "MUD-ST";
                                ws.Cells[iHeaderRow, 24].Value = "MUD-W";

                                ws.Cells[iHeaderRow, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 2].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 3].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                ws.Cells[iHeaderRow, 4].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 5].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 6].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 7].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 8].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 9].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 10].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 11].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 12].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 13].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 14].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 15].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 16].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 17].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 18].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 19].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 20].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 21].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 22].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 23].HorizontalAlignment = XlHAlign.xlHAlignRight;
                                ws.Cells[iHeaderRow, 24].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            }
                            race.Horses.Add(new Horse(fi.FullName, Fields, race));
                        }
                    }
                }

                if (race != null)
                {
                    ws.Cells[iTop5Row, 1].Value = top5AndOr400PlusCalc(race);
                    ws.get_Range(string.Format("A{0}", iTop5Row), string.Format("D{0}", iTop5Row)).Merge(Type.Missing);
                    iRow = listHorses(race, ws, iHeaderRow++);
                }
            }

            //Column Widths
            consoleSvc.UpdateConsoleText("Auto-Fitting columns...", false);
            foreach (Range c in ws.get_Range("A1", "Q1"))
                c.EntireColumn.AutoFit();
        }

        private int listHorses(IRace race, Worksheet ws, int iRow)
        {
            //race.SortHorses();
            var iRangeStart = iRow;
            iRangeStart++;

            foreach (Horse horse in race.Horses)
            {
                iRow++;
                var sNote = horse.TurfStarts.Equals(0) && horse.DSLR > 0 ? "TF DEBUT" : horse.DSLR.Equals(0) ? "FTS" : string.Empty;
                sNote = horse.TurfStarts.Equals(1) && horse.DSLR > 0 ? "2ND TF" : sNote;

                ws.Cells[iRow, 1].Value = string.Format("{0})", horse.ProgramNumber);
                ws.Cells[iRow, 2].Value = horse.MorningLine;
                ws.Cells[iRow, 3].Value = horse.HorseName;
                ws.Cells[iRow, 4].Value = !horse.Note.Equals("M") ? Convert.ToInt32(Math.Round(horse.NilsenRating, MidpointRounding.AwayFromZero)).ToString() : "MTO";
                ws.Cells[iRow, 5].Value = horse.TurfStarts.ToString();
                ws.Cells[iRow, 6].Value = horse.Wins.ToString();
                ws.Cells[iRow, 7].Value = string.Format("{0}%", Convert.ToInt32(horse.WinPercent));
                ws.Cells[iRow, 8].Value = horse.Place.ToString();
                ws.Cells[iRow, 9].Value = string.Format("{0}%", Convert.ToInt32(horse.WinPlacePercent));
                ws.Cells[iRow, 10].Value = horse.Show.ToString();
                ws.Cells[iRow, 11].Value = string.Format("{0}%", Convert.ToInt32(horse.WinPlaceShowPercent));
                ws.Cells[iRow, 12].Value = string.Format("{0:C0}", horse.Earnings);
                ws.Cells[iRow, 13].Value = string.Format("{0:C0}", horse.AverageEarnings);
                ws.Cells[iRow, 14].Value = horse.SR.ToString();
                ws.Cells[iRow, 15].Value = horse.TurfPedigreeDisplay;
                ws.Cells[iRow, 16].Value = horse.DSLR.ToString();
                ws.Cells[iRow, 17].Value = sNote;
                ws.Cells[iRow, 18].Value = horse.TFW;
                ws.Cells[iRow, 19].Value = horse.E2_1.HasValue ? horse.E2_1.Value.ToString() : "N/A";
                ws.Cells[iRow, 20].Value = horse.E2_2.HasValue ? horse.E2_2.Value.ToString() : "N/A";
                ws.Cells[iRow, 21].Value = string.Empty;
                ws.Cells[iRow, 22].Value = horse.MUD_SR;
                ws.Cells[iRow, 23].Value = horse.MUD_ST;
                ws.Cells[iRow, 24].Value = horse.MUD_W;

                // Index used in cell should always be the index used for 'Note' cell above.  
                if ((horse.TurfStarts.Equals(0) || horse.TurfStarts.Equals(1)) && horse.DSLR > 0)
                    ws.Cells[iRow, 17].Interior.Color = XlRgbColor.rgbLightGray;

                ws.Cells[iRow, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 2].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 3].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                ws.Cells[iRow, 4].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 5].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 6].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 7].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 8].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 9].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 10].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 11].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 12].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 13].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 14].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 15].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 16].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 17].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 18].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 19].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 20].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 21].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 22].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 23].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 24].HorizontalAlignment = XlHAlign.xlHAlignRight;
            }
            return FormatFields(race.Horses, ws, iRangeStart, iRow);
        }

        public int FormatFields(List<IHorse> horses, Worksheet ws, int iRangeStart, int iRangeEnd)
        {
            //declares and assigns
            DataRow dr = null;
            List<FieldFormat> fieldFormats = null;
            var sortedHorses = new List<IHorse>();

            //process
            foreach (var f in GetFieldList(FormTypes.TurfFormula))
            {
                var dt = new System.Data.DataTable();
                System.Data.DataTable dtHorses = null;

                dt.Columns.Add(new DataColumn("Value", f.Value));
                dt.Columns.Add(new DataColumn("Horse", Type.GetType("System.Object")));

                foreach (var h in horses)
                {
                    fieldFormats = new List<FieldFormat>();

                    switch (f.Key)
                    {
                        case TurfFormulaFormatFields.E2_1:
                            var e21Styles = new List<string>();

                            e21Styles.Add(Text.Style.Bold);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = TurfFormulaFormatFields.E2_1,
                                BasisType = BasisTypes.Top4,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = e21Styles,
                                WsColumnIndex = 19
                            });

                            dr[0] = h.E2_1.HasValue ? h.E2_1.Value : (decimal)0.00;
                            dr[1] = h;
                            break;
                        case TurfFormulaFormatFields.E2_2:
                            var e22Styles = new List<string>();

                            e22Styles.Add(Text.Style.Bold);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = TurfFormulaFormatFields.E2_2,
                                BasisType = BasisTypes.Top4,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = e22Styles,
                                WsColumnIndex = 20
                            });

                            dr[0] = h.E2_2.HasValue ? h.E2_2.Value : (decimal)0.00;
                            dr[1] = h;
                            break;
                        case TurfFormulaFormatFields.SR:
                            var srStyles = new List<string>();

                            srStyles.Add(Text.Style.Bold);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = TurfFormulaFormatFields.SR,
                                BasisType = BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = srStyles,
                                WsColumnIndex = 14
                            });
                            dr[0] = h.SR;
                            dr[1] = h;
                            break;
                        case TurfFormulaFormatFields.TFW:
                            var tfwStyles = new List<string>();
                            var evaluationValues = new List<decimal>();

                            tfwStyles.Add(Text.Style.Regular);

                            evaluationValues.Add(3);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = TurfFormulaFormatFields.TFW,
                                BasisType = BasisTypes.GreaterThanOrEqualTo,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = tfwStyles,
                                WsColumnIndex = 18,
                                EvaluationDecimalValues = evaluationValues
                            });
                            dr[0] = h.TFW;
                            dr[1] = h;
                            break;
                        case TurfFormulaFormatFields.TurfPedigree:
                            var turfPedigreeStyles = new List<string>();

                            turfPedigreeStyles.Add(Text.Style.Bold);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = TurfFormulaFormatFields.TurfPedigree,
                                BasisType = BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = turfPedigreeStyles,
                                WsColumnIndex = 15
                            });
                            dr[0] = h.TurfPedigree;
                            dr[1] = h;
                            break;
                    }
                }

                foreach (var ff in fieldFormats)
                {
                    var val = new Object();
                    sortedHorses.Clear();

                    switch (ff.BasisType)
                    {
                        case BasisTypes.HighestValue:
                        case BasisTypes.HighestValueWithinFloorRange:
                        case BasisTypes.Top5:
                        case BasisTypes.Top4:
                            ff.SortDirection = SortDirections.Desc;
                            break;
                        case BasisTypes.LowestValue:
                            ff.SortDirection = SortDirections.Asc;
                            break;
                        case BasisTypes.SecondHighestValue:
                        case BasisTypes.WithinRangeOfLastHorseInTopFive:
                            ff.SortDirection = SortDirections.Desc;
                            break;
                    }

                    //process sort
                    if (!string.IsNullOrEmpty(ff.SortDirection))
                    {
                        dt.DefaultView.Sort = string.Format("Value {0}", ff.SortDirection);
                    }
                    dtHorses = dt.DefaultView.ToTable();

                    switch (f.Key)
                    {
                        case TurfFormulaFormatFields.SR:
                        case TurfFormulaFormatFields.TurfPedigree:
                            val = Convert.ToDecimal(dtHorses.Rows[0][0]);
                            break;
                    }

                    switch (ff.BasisType)
                    {
                        case BasisTypes.BaseAmountOrHigher:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (Convert.ToDecimal(r[0]) >= ff.EvaluationDecimalValues[0])
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case BasisTypes.BaseAmountOrLower:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (Convert.ToDecimal(r[0]) <= ff.EvaluationDecimalValues[0])
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case BasisTypes.BetweenTwoValues:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (Convert.ToDecimal(r[0]) >= ff.EvaluationDecimalValues[0] && Convert.ToDecimal(r[0]) <= ff.EvaluationDecimalValues[1])
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case BasisTypes.Equals:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (Convert.ToDecimal(r[0]) == ff.EvaluationDecimalValues[0])
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case BasisTypes.HighestValueWithinFloorRange:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                var highValue = (decimal)val;
                                var floorValue = (decimal)val - ff.EvaluationDecimalValues[0];

                                if ((Convert.ToDecimal(r[0]) <= highValue) &&
                                    (Convert.ToDecimal(r[0]) >= floorValue) &&
                                    (!Convert.ToDecimal(r[0]).Equals((decimal)0)))
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case BasisTypes.HighestValue:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (Convert.ToDecimal(r[0]).Equals(val) && (Convert.ToDecimal(r[0]) > ((decimal)0)))
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case BasisTypes.GreaterThanOrEqualTo:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                var horse = (IHorse)r[1];
                                for (var iIndex = 0; iIndex < ff.EvaluationDecimalValues.Count(); iIndex++)
                                {
                                    var evalValue = ff.EvaluationDecimalValues[iIndex];
                                    var horseValue = Convert.ToDecimal(r[0]);

                                    if (horseValue >= evalValue)
                                    {
                                        sortedHorses.Add(horse);
                                    }
                                }
                            }
                            break;
                        case BasisTypes.LowestValue:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (r[0].Equals(val))
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case BasisTypes.LessThanZero:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (Convert.ToDecimal(r[0]) < (decimal)0.00)
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case BasisTypes.Top5:
                        case BasisTypes.Top4:
                            var iHorseCount = 0;
                            var basis = ff.BasisType == BasisTypes.Top4 ? 4 : 5;
                            decimal lastValue = 0;

                            foreach (DataRow r in dtHorses.Rows)
                            {
                                var value = Convert.ToDecimal(r[0]);

                                if (iHorseCount < basis)
                                {
                                    if (value > 0)
                                    {
                                        iHorseCount++;
                                        lastValue = value;
                                        sortedHorses.Add((IHorse)r[1]);
                                    }
                                }
                                else
                                {
                                    if (lastValue.Equals(value))
                                        sortedHorses.Add((IHorse)r[1]);
                                    else
                                        break;
                                }
                            }
                            break;
                        case BasisTypes.RnkWrkrsCustom:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (Convert.ToBoolean(r[0]))
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case BasisTypes.SecondHighestValue:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (Convert.ToDecimal(r[0]) < Convert.ToDecimal(val))
                                {
                                    val = r[0];
                                    break;
                                }
                            }
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (r[0].Equals(val))
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case BasisTypes.ValueNotInStringList:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (!string.IsNullOrWhiteSpace(r[0].ToString()))
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case BasisTypes.WithinRange:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if ((Convert.ToDecimal(r[0]) <= Convert.ToDecimal(val) + ff.EvaluationDecimalValues[0]) && (Convert.ToDecimal(r[0]) >= Convert.ToDecimal(val) - ff.EvaluationDecimalValues[0]))
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case BasisTypes.WithinRangeOfLastHorseInTopFive:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (Convert.ToDecimal(r[0]) >= Convert.ToDecimal(val) - ff.EvaluationDecimalValues[0] && Convert.ToDecimal(r[0]) < Convert.ToDecimal(val))
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case BasisTypes.None:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                sortedHorses.Add((IHorse)r[1]);
                            }
                            break;
                    }

                    if (sortedHorses != null)
                    {
                        for (var iIndex = iRangeStart; iIndex <= iRangeEnd; iIndex++)
                        {
                            //Get the name cell
                            var cell = ws.Cells[iIndex, 3];
                            var row = ws.Rows[iIndex];
                            var name = cell.Value;

                            foreach (var h in sortedHorses)
                            {
                                //Check to see if it's our selected horse.  
                                if (name.Equals(h.HorseName))
                                {
                                    //now on the same row, find the cell, and format it.  
                                    cell = ws.Cells[iIndex, ff.WsColumnIndex];

                                    if ((cell.Value != null) && !string.IsNullOrWhiteSpace(cell.Value.ToString()))
                                    {
                                        if (!(ff.BackgroundColor.Equals(XlRgbColor.rgbWhite)))
                                        {
                                            cell.Interior.Color = ff.BackgroundColor;

                                            if (ff.BasisType == BasisTypes.RnkWrkrsCustom)
                                            {
                                                var indexAdjust = ff.WsColumnIndex;
                                                ws.Cells[iIndex, --indexAdjust].Interior.Color = ff.BackgroundColor;
                                                ws.Cells[iIndex, --indexAdjust].Interior.Color = ff.BackgroundColor;
                                                ws.Cells[iIndex, --indexAdjust].Interior.Color = ff.BackgroundColor;
                                                ws.Cells[iIndex, --indexAdjust].Interior.Color = ff.BackgroundColor;
                                            }
                                        }
                                        cell.Font.Color = ff.TextColor;
                                        foreach (var style in ff.TextStyles)
                                        {
                                            cell.Font.Bold = style.Equals(Text.Style.Bold);
                                            cell.Font.Italic = style.Equals(Text.Style.Italic);

                                            if (ff.BasisType == BasisTypes.RnkWrkrsCustom)
                                            {
                                                var indexAdjust = ff.WsColumnIndex;
                                                ws.Cells[iIndex, --indexAdjust].Font.Bold = style.Equals(Text.Style.Bold);
                                                ws.Cells[iIndex, indexAdjust].Font.Italic = style.Equals(Text.Style.Italic);
                                                ws.Cells[iIndex, --indexAdjust].Font.Bold = style.Equals(Text.Style.Bold);
                                                ws.Cells[iIndex, indexAdjust].Font.Italic = style.Equals(Text.Style.Italic);
                                                ws.Cells[iIndex, --indexAdjust].Font.Bold = style.Equals(Text.Style.Bold);
                                                ws.Cells[iIndex, indexAdjust].Font.Italic = style.Equals(Text.Style.Italic);
                                                ws.Cells[iIndex, --indexAdjust].Font.Bold = style.Equals(Text.Style.Bold);
                                                ws.Cells[iIndex, indexAdjust].Font.Italic = style.Equals(Text.Style.Italic);
                                            }
                                        }
                                    }
                                    break;
                                }
                            }

                            Marshal.ReleaseComObject(cell);
                            cell = null;

                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                    }
                }
            }

            return iRangeEnd;
        }

        private string top5AndOr400PlusCalc(IRace race)
        {
            var horseList = (((from h in race.Horses select h).OrderByDescending(x => x.NilsenRating).Take(5)).Concat
                             (from h in race.Horses where h.NilsenRating >= 400 select h).Distinct()).OrderByDescending(x => x.NilsenRating);
            var sbHorses = new StringBuilder();
            var lastHorseRanking = (decimal)0.00;

            foreach (var horse in horseList)
            {
                var greaterThan80Gap = (!horse.Equals(horseList.First()) && ((horse.NilsenRating - lastHorseRanking) >= 80 || ((horse.NilsenRating - lastHorseRanking) * -1) >= 80));
                var asterisk = (horse.MorningLine >= (decimal)6.0) ? "*" : string.Empty;
                var separator = (!horse.Equals(horseList.First()) ? ((greaterThan80Gap) ? " / " : " - ") : string.Empty);

                sbHorses.AppendFormat("{0}{1}{2}", separator, horse.ProgramNumber, asterisk);
                lastHorseRanking = horse.NilsenRating;
            }

            return string.Format("Turf:   {0}", sbHorses.ToString()); 
        }
    }
}
