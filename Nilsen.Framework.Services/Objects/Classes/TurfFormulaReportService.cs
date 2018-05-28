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
                wb.SaveAs(sbFullFileName.ToString(), XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false,
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
            var iHorse = 0;
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
                            Int32 iDSLR;
                            Decimal decTotalNilsenRating;
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
                            }
                            race.Horses.Add(new Horse(Fields, race));
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
            race.SortHorses();

            foreach(Horse horse in race.Horses)
            {
                iRow++;
                var sNote = horse.TurfStarts.Equals(0) && horse.DSLR > 0 ? "TURF DEBUT" : horse.DSLR.Equals(0) ? "FTS" : string.Empty;
                sNote = horse.TurfStarts.Equals(1) && horse.DSLR > 0 ? "2nd TF" : sNote;

                ws.Cells[iRow, 1].Value = string.Format("{0})", horse.ProgramNumber);
                ws.Cells[iRow, 2].Value = horse.MorningLine;
                ws.Cells[iRow, 3].Value = horse.HorseName;
                ws.Cells[iRow, 4].Value = !horse.Note.Equals("M") ? Convert.ToInt32(Math.Round(horse.NilsenRating, MidpointRounding.AwayFromZero)).ToString() : "MTO";
                ws.Cells[iRow, 5].Value = horse.TurfStarts.ToString();
                ws.Cells[iRow, 6].Value = horse.Wins.ToString();
                ws.Cells[iRow, 7].Value = string.Format("{0}%", horse.WinPercent.ToString());
                ws.Cells[iRow, 8].Value = horse.Place.ToString();
                ws.Cells[iRow, 9].Value = string.Format("{0}%", horse.WinPlacePercent.ToString());
                ws.Cells[iRow, 10].Value = horse.Show.ToString();
                ws.Cells[iRow, 11].Value = string.Format("{0}%", horse.WinPlaceShowPercent.ToString());
                ws.Cells[iRow, 12].Value = string.Format("{0:C0}", horse.Earnings.ToString());
                ws.Cells[iRow, 13].Value = string.Format("{0:C0}", horse.AverageEarnings.ToString());
                ws.Cells[iRow, 14].Value = horse.SR.ToString();
                ws.Cells[iRow, 15].Value = horse.TurfPedigreeDisplay;
                ws.Cells[iRow, 16].Value = horse.DSLR.ToString();
                ws.Cells[iRow, 17].Value = sNote;
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
            }
            return FormatFields(race.Horses, ws, iRow);
        }

        public int FormatFields(List<IHorse> horses, Worksheet ws, int iRow)
        {
            //declares and assigns
            DataRow dr = null;
            List<FieldFormat> fieldFormats = null;
            var sortedHorses = new List<IHorse>();
            var keyTrainerStatIndex = 0;

            //process
            foreach (var f in GetFieldList(FormTypes.TurfFormula))
            {
            }

            return iRow;
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
                var asterisk = (horse.MorningLine >= (decimal)6.1) ? "*" : string.Empty;
                var separator = (!horse.Equals(horseList.First()) ? ((greaterThan80Gap) ? "/" : "-") : string.Empty);

                sbHorses.AppendFormat("{0}{1}{2}", asterisk, separator, horse.ProgramNumber);
                lastHorseRanking = horse.NilsenRating;
            }

            return string.Format("Turf:   {0}", sbHorses.ToString()); 
        }
    }
}
