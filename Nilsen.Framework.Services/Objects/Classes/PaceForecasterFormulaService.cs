using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using System.Configuration;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.FileIO;
using Nilsen.Framework.Services;
using Nilsen.Framework.Objects.Class;
using Nilsen.Framework.Objects.Enums;
using Nilsen.Framework.Services.Objects.Interfaces;
using Nilsen.Framework.Objects.Interfaces;
using Nilsen.Framework.Common;

namespace Nilsen.Framework.Services.Objects.Classes
{
    public class PaceForecasterFormulaReportService : IReportService
    {
        private ConsoleService consoleSvc;
        private String _SavePath = string.Format("{0}\\Flicker City Productions\\RacesCSVToExcel\\files", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
        private string[] Fields;

        public PaceForecasterFormulaReportService(System.Windows.Forms.TextBox consoleWindow, System.Windows.Forms.Button btnProcess)
        {
            consoleSvc = new ConsoleService(consoleWindow, btnProcess);
        }

        public void CreateExcelFile(FileInfo fi)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = ExcelApp.Workbooks.Add(Type.Missing);
            Worksheet ws = wb.Sheets.Add();
            StringBuilder sbFullFileName = new StringBuilder();
            StringBuilder sbFileName = new StringBuilder();
            StringBuilder sbPath = new StringBuilder();
            Int32 iLength = fi.Name.Split('.').GetLength(0);
            Boolean bContinue = true;

            consoleSvc.ToggleProcessButton(false);

            sbFileName.Append(fi.Name.Split('.')[0]);
            sbFullFileName.AppendFormat("{0}\\NilsenRatings_PaceForecasterFormula_{1}", _SavePath, sbFileName.ToString());

            //Build Worksheet
            consoleSvc.UpdateConsoleText("Creating Worksheet...", true);
            BuildWorksheet(ws, fi);

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
                    var results = MessageBox.Show(string.Format("File '{0}' Exists.\n\nReplace?", string.Format("NilsenRatings_PaceForecasterFormula_{0}.xlsx", sbFileName)), "File Exists", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

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
                wb.SaveAs(sbFullFileName.ToString(), Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false,
                        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                consoleSvc.UpdateConsoleText(string.Format("File Saved As: {0}.xlsx", sbFullFileName.ToString()), false);
            }

            //Cleanup
            ExcelApp.Application.Visible = true;
            Marshal.ReleaseComObject(ws);

            wb.Close();
            Marshal.ReleaseComObject(wb);

            ExcelApp.Quit();
            Marshal.ReleaseComObject(ExcelApp);
            consoleSvc.UpdateConsoleText("Processing Completed", false);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            consoleSvc.ToggleProcessButton(true);
        }

        public void BuildWorksheet(Worksheet ws, FileInfo fi)
        {
            //declares and assigns
            var reader = new StreamReader(File.OpenRead(fi.FullName));
            string[] Lines;
            var iRow = 1;
            var iHeaderRow = 1;
            Range rHeader;
            var sAllHorses = new String[100, 2];
            var decAllHorses = new Decimal[100];
            var Top5Horses = new String[5, 2];
            var bInitialHeader = true;
            IRace race = null;
            Int16 iTotalRow = 0;
            TextFieldParser lineParser = null;
            ws.Name = "Nilsen Pace Rating";
            ws.get_Range("A1", "E1").Merge(Type.Missing);
            rHeader = ws.get_Range("A1", Type.Missing);
            rHeader.Value = "Nilsen Pace Rating Report";
            rHeader.Font.Bold = true;
            ws.Cells[1, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            Lines = Regex.Split(reader.ReadToEnd(), Environment.NewLine);

            //page style
            ws.PageSetup.Orientation = XlPageOrientation.xlPortrait;

            //race vars
            var sRaceDate = string.Empty;
            var sRaceName = string.Empty;

            consoleSvc.UpdateConsoleText("Reading CSV File...", false);

            if (Lines.GetLength(0) > 0)
            {
                foreach (var line in Lines)
                {
                    lineParser = new TextFieldParser(new StringReader(line));
                    lineParser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
                    lineParser.SetDelimiters(new string[] {","});
                    lineParser.HasFieldsEnclosedInQuotes = true;

                    while (!lineParser.EndOfData)
                    {
                        Fields = lineParser.ReadFields();
                        //New Race Record
                        if (!Fields[2].ToLower().Equals(sRaceName))
                        {
                            if (!(race == null))
                            {
                                iRow = ListHorses(race, ws, iRow, iHeaderRow);
                                ws.Cells[iTotalRow, 5].Value = race.GetTop3Total();
                            }

                            race = RaceService.GetRace(Fields);

                            consoleSvc.UpdateConsoleText(string.Format("Reading and building Race {0}...", Fields[2]), false);
                            iRow = iRow + 2;

                            //Only show the Race Date and Track
                            if (bInitialHeader)
                            {
                                sRaceDate = DateTime.ParseExact(race.DateText, "yyyyMMdd", CultureInfo.InvariantCulture).ToString("MM/dd/yyyy");
                                ws.get_Range(string.Format("A{0}", iRow), string.Format("C{0}", iRow)).Merge(Type.Missing);
                                ws.Cells[iRow, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                ws.Cells[iRow++, 1].Value = sRaceDate;
                                ws.get_Range(string.Format("A{0}", iRow), string.Format("C{0}", iRow)).Merge(Type.Missing);
                                ws.Cells[iRow, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                ws.Cells[iRow++, 1].Value = race.Track.TrackName;
                                bInitialHeader = false;
                            }

                            iRow++;

                            sRaceName = Fields[2];
                            ws.get_Range(string.Format("A{0}", iRow), string.Format("D{0}", iRow)).Merge(Type.Missing);
                            ws.Cells[iRow++, 1].Value = string.Format("Race {0}", race.Name);
                            ws.get_Range(string.Format("A{0}", iRow), string.Format("D{0}", iRow)).Merge(Type.Missing);
                            ws.Cells[iRow, 1].Value = string.Format("{0} Furlongs / {1}", race.Track.Furlongs, (race.Track.AllWeather) ? "All Weather" : race.Track.TrackType); //Distance / Turf
                            if (race.Track.TrackTypeShort.ToLower().Equals("t"))
                            {
                                ws.Cells[iRow, 1].Interior.Color = XlRgbColor.rgbGreen;
                                ws.Cells[iRow, 1].Font.Color = XlRgbColor.rgbWhite;
                                ws.Cells[iRow, 1].Font.Bold = true;
                            }
                            iRow++;
                            ws.get_Range(string.Format("A{0}", iRow), string.Format("D{0}", iRow)).Merge(Type.Missing);

                            switch (race.Type.ToLower())
                            {
                                case "g1":
                                    ws.Cells[iRow++, 1].Value = "Grade I Stk/Hcp";
                                    break;
                                case "g2":
                                    ws.Cells[iRow++, 1].Value = "Grade II Stk/Hcp";
                                    break;
                                case "g3":
                                    ws.Cells[iRow++, 1].Value = "Grade III Stk/Hcp";
                                    break;
                                case "n":
                                    ws.Cells[iRow++, 1].Value = "Nongraded Stake/Handicap";
                                    break;
                                case "a":
                                    ws.Cells[iRow++, 1].Value = "Allowance";
                                    break;
                                case "r":
                                    ws.Cells[iRow++, 1].Value = "Starter Alw";
                                    break;
                                case "t":
                                    ws.Cells[iRow++, 1].Value = "Starter Hcp";
                                    break;
                                case "c":
                                    ws.Cells[iRow++, 1].Value = "Claiming";
                                    break;
                                case "co":
                                    ws.Cells[iRow++, 1].Value = "Optional Clmg";
                                    break;
                                case "s":
                                    ws.Cells[iRow, 1].Interior.Color = XlRgbColor.rgbRed;
                                    ws.Cells[iRow, 1].Font.Color = XlRgbColor.rgbWhite;
                                    ws.Cells[iRow, 1].Font.Bold = true;
                                    ws.Cells[iRow++, 1].Value = "Mdn Sp Wt";
                                    break;
                                case "m":
                                    ws.Cells[iRow, 1].Interior.Color = XlRgbColor.rgbRed;
                                    ws.Cells[iRow, 1].Font.Color = XlRgbColor.rgbWhite;
                                    ws.Cells[iRow, 1].Font.Bold = true;
                                    ws.Cells[iRow++, 1].Value = "Ndn Claimer";
                                    break;
                                case "ao":
                                    ws.Cells[iRow++, 1].Value = "Alw Opt Clm";
                                    break;
                                case "mo":
                                    ws.Cells[iRow, 1].Interior.Color = XlRgbColor.rgbRed;
                                    ws.Cells[iRow, 1].Font.Color = XlRgbColor.rgbWhite;
                                    ws.Cells[iRow, 1].Font.Bold = true;
                                    ws.Cells[iRow++, 1].Value = "Mdn Opt Clm";
                                    break;
                                case "no":
                                    ws.Cells[iRow++, 1].Value = "Opt Clm Stk";
                                    break;
                            }

                            ws.get_Range(string.Format("A{0}", iRow), string.Format("D{0}", iRow)).Merge(Type.Missing);
                            ws.Cells[iRow++, 1].Value = string.Format("{0}", race.PostTime);
                            ws.get_Range(string.Format("A{0}", iRow), string.Format("G{0}", iRow)).Merge(Type.Missing);
                            ws.Cells[iRow++, 1].Value = race.GetAgeOfRace();
                            iTotalRow = Convert.ToInt16(iRow); 

                            iRow++;

                            iHeaderRow = iRow;

                            //Horse headers
                            ws.Cells[iHeaderRow, 1].Value = "Prg #";
                            ws.Cells[iHeaderRow, 2].Value = "ML";
                            ws.Cells[iHeaderRow, 3].Value = "BL";
                            ws.Cells[iHeaderRow, 4].Value = "Horse Name";
                            ws.Cells[iHeaderRow, 5].Value = "TOTAL";
                            ws.Cells[iHeaderRow, 6].Value = "R/Q";
                            ws.Cells[iHeaderRow, 7].Value = "PP";
                            ws.Cells[iHeaderRow, 8].Value = "Pace";
                            ws.Cells[iHeaderRow, 9].Value = "CP";
                            ws.Cells[iHeaderRow, 10].Value = "DSLR";
                            ws.Cells[iHeaderRow, 11].Value = "CR";
                            ws.Cells[iHeaderRow, 12].Value = "LP";
                            ws.Cells[iHeaderRow, 13].Value = "RBC%";
                            ws.Cells[iHeaderRow, 14].Value = "B-CR";
                            ws.Cells[iHeaderRow, 15].Value = "B-SR";
                            ws.Cells[iHeaderRow, 16].Value = "PPWR";
                            ws.Cells[iHeaderRow, 17].Value = "MJS";
                            ws.Cells[iHeaderRow, 18].Value = "RET";
                            ws.Cells[iHeaderRow, 19].Value = "MDC";
                            ws.Cells[iHeaderRow, 20].Value = "TB";
                            ws.Cells[iHeaderRow, 21].Value = "Trk";
                            ws.Cells[iHeaderRow, 22].Value = "Dis";
                            ws.Cells[iHeaderRow, 23].Value = "T-SR";
                            ws.Cells[iHeaderRow, 24].Value = "D-SR";
                            ws.Cells[iHeaderRow, 25].Value = "Note";
                            ws.Cells[iHeaderRow, 26].Value = "Note 2";
                            ws.Cells[iHeaderRow, 27].Value = "Note 3";
                            ws.Cells[iHeaderRow, 28].Value = "#W";
                            ws.Cells[iHeaderRow, 29].Value = "#F";
                            ws.Cells[iHeaderRow, 30].Value = "RK";
                            ws.Cells[iHeaderRow, 31].Value = "WKrs";
                            ws.Cells[iHeaderRow, 32].Value = string.Empty;
                            ws.Cells[iHeaderRow, 33].Value = "Extended Comment";

                            ws.Cells[iHeaderRow, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            ws.Cells[iHeaderRow, 2].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            ws.Cells[iHeaderRow, 3].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            ws.Cells[iHeaderRow, 4].HorizontalAlignment = XlHAlign.xlHAlignLeft;
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
                            ws.Cells[iHeaderRow, 25].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            ws.Cells[iHeaderRow, 26].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            ws.Cells[iHeaderRow, 27].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            ws.Cells[iHeaderRow, 28].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            ws.Cells[iHeaderRow, 29].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            ws.Cells[iHeaderRow, 30].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            ws.Cells[iHeaderRow, 31].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            ws.Cells[iHeaderRow, 32].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            ws.Cells[iHeaderRow, 33].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        }
                        race.Horses.Add(new Horse(Fields, race));
                    }
                }

                if (race != null)
                {
                    iRow = ListHorses(race, ws, iRow, iHeaderRow);
                    ws.Cells[iTotalRow, 5].Value = race.GetTop3Total();
                }
            }

            //Column Widths
            consoleSvc.UpdateConsoleText("Auto-Fitting columns...", false);
            foreach (Range c in ws.get_Range("A1", "AJ1"))
            {
                c.EntireColumn.AutoFit();
            }

            Marshal.ReleaseComObject(rHeader);
            Marshal.ReleaseComObject(ws);
        }

        private Int32 ListHorses(IRace race, Worksheet ws, Int32 iRowRangeStart, Int32 iHeaderRow)
        {
            var iRow = iRowRangeStart;
            var iRowRangeEnd = 0;
            var ktscFirstIndex = 34;
            var ktscLastIndex = race.GetGreatestKeyTrainerStatCategoryCount() + ktscFirstIndex - 1;
            var iColIndex = 1;
            race.SortHorses();

            for (var iColCount = ktscFirstIndex; iColCount <= ktscLastIndex; iColCount++)
            {
                ws.Cells[iHeaderRow, iColCount].Value = string.Format("Key Trainer Stat-{0}", iColIndex++);
            }

            foreach (Horse horse in race.Horses)
            {
                iRow++;
                ws.Cells[iRow, 1].Value = string.Format("{0})", horse.ProgramNumber); //Program Number
                ws.Cells[iRow, 2].Value = horse.MorningLine; //Morning Line
                ws.Cells[iRow, 3].Value = (horse.Blinkers.Equals(1)) ? "on" : (horse.Blinkers.Equals(2)) ? "OFF" : string.Empty; //Horse Name
                ws.Cells[iRow, 4].Value = horse.HorseName; //Horse Name
                ws.Cells[iRow, 5].Value = horse.Total; //TOTAL
                ws.Cells[iRow, 6].Value = (horse.RunStyle + horse.Quirin); //RQ 
                ws.Cells[iRow, 7].Value = horse.PostPoints; //PP
                ws.Cells[iRow, 8].Value = horse.Pace; //Pace
                ws.Cells[iRow, 9].Value = horse.CP; //CP
                ws.Cells[iRow, 10].Value = horse.DSLR; //DSLR
                ws.Cells[iRow, 11].Value = horse.CR; //CR
                ws.Cells[iRow, 12].Value = horse.LP; //LP
                ws.Cells[iRow, 13].Value = horse.RBCPercent; //RBCPercent
                ws.Cells[iRow, 14].Value = horse.BCR; //BCR
                ws.Cells[iRow, 15].Value = horse.BSR; //BSR
                ws.Cells[iRow, 16].Value = horse.PPWR; //PPWR
                ws.Cells[iRow, 17].Value = horse.MJS; //MJS
                ws.Cells[iRow, 18].Value = horse.RET; //RET
                ws.Cells[iRow, 19].Value = horse.MDC; //MDC
                ws.Cells[iRow, 20].Value = (horse.TB >= 110) ? "TB" : string.Empty; //TB
                ws.Cells[iRow, 21].Value = horse.Trk; //Trk
                ws.Cells[iRow, 22].Value = horse.DIS; //Dis
                ws.Cells[iRow, 23].Value = horse.TSR; //TSR
                ws.Cells[iRow, 24].Value = horse.DSR; //DSR
                ws.Cells[iRow, 25].Value = horse.Note; //Note
                ws.Cells[iRow, 26].Value = horse.Note2; //Note2
                ws.Cells[iRow, 27].Value = horse.Note3; //Note3
                ws.Cells[iRow, 28].Value = horse.Workout; //Workout
                ws.Cells[iRow, 29].Value = horse.Distance; //Distance
                ws.Cells[iRow, 30].Value = horse.Rank; //Rank
                ws.Cells[iRow, 31].Value = horse.Workers; //WKrs
                ws.Cells[iRow, 32].Value = (Math.Round(horse.RnkWrkrsPct, MidpointRounding.AwayFromZero)).ToString() + "%"; //PERCENTAGE
                ws.Cells[iRow, 33].Value = horse.ExtendedComment; //Extended Comment

                if (horse.Blinkers.Equals(1) || horse.Blinkers.Equals(2))
                {
                    ws.Cells[iRow, 3].Interior.Color = XlRgbColor.rgbLightGray;
                }

                if (horse.Note.ToLower().Equals("lasix") || horse.Note.ToLower().Equals("mts"))
                {
                    ws.Cells[iRow, 25].Interior.Color = XlRgbColor.rgbLightGray;
                    ws.Cells[iRow, 25].Font.Bold = true;
                }

                if (horse.Note2.ToLower().Equals("lasix") || horse.Note2.ToLower().Equals("mts"))
                {
                    ws.Cells[iRow, 26].Interior.Color = XlRgbColor.rgbLightGray;
                    ws.Cells[iRow, 26].Font.Bold = true;
                }

                if (horse.Note3.ToLower().Equals("lasix") || horse.Note3.ToLower().Equals("mts") || horse.Note3.ToLower().Equals("ts"))
                {
                    ws.Cells[iRow, 27].Interior.Color = XlRgbColor.rgbLightGray;
                }

                ws.Cells[iRow, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 2].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 3].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 4].HorizontalAlignment = XlHAlign.xlHAlignLeft;
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
                ws.Cells[iRow, 25].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 26].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 27].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 28].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 29].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 30].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 31].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 32].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 33].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                //format for decimal.  
                ws.Cells[iRow, 11].NumberFormat = "0.00";
                ws.Cells[iRow, 13].NumberFormat = "0.00";
                ws.Cells[iRow, 14].NumberFormat = "0.00";
                ws.Cells[iRow, 16].NumberFormat = "0.0";

                for (var iColumnIndex = ktscFirstIndex; iColumnIndex <= ktscLastIndex; iColumnIndex++)
                {
                    ws.Cells[iRow, iColumnIndex].Value = (horse.KeyTrainerStatCategory.Count > (iColumnIndex - ktscFirstIndex)) ? horse.KeyTrainerStatCategory[iColumnIndex - ktscFirstIndex] : " ";
                    ws.Cells[iRow, iColumnIndex].HorizontalAlignment = XlHAlign.xlHAlignRight;
                }
            }

            iRowRangeEnd = iRow;
            iRow = FormatFields(race.Horses, ws, iRowRangeStart, iRowRangeEnd, ktscFirstIndex, ktscLastIndex);

            return iRow;
        }

        public int FormatFields(List<IHorse> horses, Worksheet ws, int iRangeStart, int iRangeEnd, int ktscFirstIndex, int ktscLastIndex)
        {
            //declares and assigns
            DataRow dr = null;
            List<FieldFormat> fieldFormats = null;
            var sortedHorses = new List<IHorse>();
            var keyTrainerStatIndex = 0;

            //process
            foreach (var f in RaceFieldsFormat.GetFieldList())
            {
                var dt = new System.Data.DataTable();
                var ktscStyles = new List<string>();
                var ktsIndex = keyTrainerStatIndex + ktscFirstIndex;
                System.Data.DataTable dtHorses = null;

                dt.Columns.Add(new DataColumn("Value", f.Value));
                dt.Columns.Add(new DataColumn("Horse", System.Type.GetType("System.Object")));

                foreach (var h in horses)
                {
                    fieldFormats = new List<FieldFormat>();

                    switch (f.Key)
                    {
                        case RaceFieldsFormat.Fields.BCR: //BCR
                            var bcrStyles = new List<string>();
                            var bcrEvaluationValues = new List<decimal>();

                            bcrStyles.Add(RaceFieldsFormat.Text.Style.Bold);
                            bcrEvaluationValues.Add((decimal)2.5);
                            
                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.BCR,
                                BasisType = RaceFieldsFormat.BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = bcrStyles,
                                WsColumnIndex = 14
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.BCR,
                                BasisType = RaceFieldsFormat.BasisTypes.WithinRange,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = bcrStyles,
                                WsColumnIndex = 14,
                                EvaluationValues = bcrEvaluationValues
                            });
                            dr[0] = h.BCR;
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.BSR: //BSR
                            var bsrStyles = new List<string>();
                            var bsrEvaluationValues = new List<decimal>();

                            bsrStyles.Add(RaceFieldsFormat.Text.Style.Bold);
                            bsrEvaluationValues.Add(7);
                            
                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.BCR,
                                BasisType = RaceFieldsFormat.BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = bsrStyles,
                                WsColumnIndex = 15
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.BSR,
                                BasisType = RaceFieldsFormat.BasisTypes.WithinRange,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = bsrStyles,
                                WsColumnIndex = 15,
                                EvaluationValues = bsrEvaluationValues
                            });
                            dr[0] = h.BSR;
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.CR: //CR
                            var crStyles = new List<string>();
                            var crEvaluationValues = new List<decimal>();

                            crStyles.Add(RaceFieldsFormat.Text.Style.Bold);
                            crEvaluationValues.Add((decimal)1);
                            
                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.CR,
                                BasisType = RaceFieldsFormat.BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = crStyles,
                                WsColumnIndex = 11
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.CR,
                                BasisType = RaceFieldsFormat.BasisTypes.WithinRange,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = crStyles,
                                WsColumnIndex = 11,
                                EvaluationValues = crEvaluationValues
                            });
                            dr[0] = h.CR;
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.DSR: //DSR
                            var dsrStyles = new List<string>();
                            var dsrEvaluationValues = new List<decimal>();

                            dsrStyles.Add(RaceFieldsFormat.Text.Style.Bold);
                            dsrEvaluationValues.Add(3);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.DSR,
                                BasisType = RaceFieldsFormat.BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGray,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = dsrStyles,
                                WsColumnIndex = 24
                            });

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.DSR,
                                BasisType = RaceFieldsFormat.BasisTypes.HighestValueWithinFloorRange,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = dsrStyles,
                                WsColumnIndex = 24,
                                EvaluationValues = dsrEvaluationValues
                            });
                            dr[0] = (decimal)h.DSR;
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.DSLR: //DSLR
                            var dslrStyles = new List<string>();
                            var dslrStyles1 = new List<string>();
                            var dslrEvaluationValues = new List<decimal>();
                            var dslrEvaluationValues1 = new List<decimal>();
                            var dslrEvaluationValues2 = new List<decimal>();

                            dslrStyles.Add(RaceFieldsFormat.Text.Style.Bold);
                            dslrStyles1.Add(RaceFieldsFormat.Text.Style.Italic);
                            dslrEvaluationValues.Add((decimal)280);
                            dslrEvaluationValues1.Add((decimal)11);
                            dslrEvaluationValues2.Add((decimal)60);
                            dslrEvaluationValues2.Add((decimal)279);
                            
                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.DSLR,
                                BasisType = RaceFieldsFormat.BasisTypes.BaseAmountOrHigher,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbRed,
                                TextStyles = dslrStyles,
                                WsColumnIndex = 10,
                                EvaluationValues = dslrEvaluationValues
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.DSLR,
                                BasisType = RaceFieldsFormat.BasisTypes.BaseAmountOrLower,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = dslrStyles,
                                WsColumnIndex = 10,
                                EvaluationValues = dslrEvaluationValues1
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.DSLR,
                                BasisType = RaceFieldsFormat.BasisTypes.BetweenTwoValues,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = dslrStyles1,
                                WsColumnIndex = 10,
                                EvaluationValues = dslrEvaluationValues2
                            });
                            dr[0] = (decimal)h.DSLR;
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.KeyTrainerStatCategory1: //KeyTrainerStatCategory1
                            ktscStyles = new List<string>();
                            ktsIndex = ktscFirstIndex;

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.KeyTrainerStatCategory3,
                                BasisType = RaceFieldsFormat.BasisTypes.ValueExists,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = ktscStyles,
                                WsColumnIndex = ktsIndex
                            });

                            dr[0] = (h.KeyTrainerStatCategory.Count > 0) ? h.KeyTrainerStatCategory[0] : string.Empty;
                            dr[1] = h;

                            break;
                        case RaceFieldsFormat.Fields.KeyTrainerStatCategory2: //KeyTrainerStatCategory2
                            ktscStyles = new List<string>();
                            ktsIndex = ktscFirstIndex + 1;

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.KeyTrainerStatCategory3,
                                BasisType = RaceFieldsFormat.BasisTypes.ValueExists,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = ktscStyles,
                                WsColumnIndex = ktsIndex
                            });

                            dr[0] = (h.KeyTrainerStatCategory.Count > 1) ? h.KeyTrainerStatCategory[1] : string.Empty;
                            dr[1] = h;

                            keyTrainerStatIndex++;
                            break;
                        case RaceFieldsFormat.Fields.KeyTrainerStatCategory3: //KeyTrainerStatCategory3
                            ktscStyles = new List<string>();
                            ktsIndex = ktscFirstIndex + 2;

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.KeyTrainerStatCategory3,
                                BasisType = RaceFieldsFormat.BasisTypes.ValueExists,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = ktscStyles,
                                WsColumnIndex = ktsIndex
                            });

                            dr[0] = (h.KeyTrainerStatCategory.Count > 2) ? h.KeyTrainerStatCategory[2] : string.Empty;
                            dr[1] = h;

                            break;
                        case RaceFieldsFormat.Fields.LP: //LP
                            var lpStyles = new List<string>();

                            lpStyles.Add(RaceFieldsFormat.Text.Style.Bold);
                            
                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.LP,
                                BasisType = RaceFieldsFormat.BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = lpStyles,
                                WsColumnIndex = 12
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.LP,
                                BasisType = RaceFieldsFormat.BasisTypes.SecondHighestValue,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = lpStyles,
                                WsColumnIndex = 12
                            });
                            dr[0] = (decimal)h.LP;
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.MDC: //MDC
                            var mdcStyles = new List<string>();
                            
                            mdcStyles.Add(RaceFieldsFormat.Text.Style.Bold);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.MDC,
                                BasisType = RaceFieldsFormat.BasisTypes.GreaterThanOrEqualTo,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                EvaluationValues = new List<decimal>() { (decimal)1.40, h.ClaimingPrice + (h.ClaimingPrice * (decimal)0.33) },
                                HorseValues = new List<decimal>() { Math.Round((h.LastPurse / h.RacePurse), 2), h.ClaimingPriceLastRace },
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = mdcStyles,
                                WsColumnIndex = 19
                            });

                            dr[0] = Math.Round(h.LastPurse / h.RacePurse);
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.MJS: //MJS
                            var mjsStyles = new List<string>();

                            mjsStyles.Add(RaceFieldsFormat.Text.Style.Bold);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.MJS,
                                BasisType = RaceFieldsFormat.BasisTypes.GreaterThanOrEqualTo,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                EvaluationValues = new List<decimal>() { (decimal).38 },
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = new List<string>() { RaceFieldsFormat.Text.Style.Regular },
                                WsColumnIndex = 17
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.MJS,
                                BasisType = RaceFieldsFormat.BasisTypes.GreaterThanOrEqualTo,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                EvaluationValues = new List<decimal>() { (decimal).59 },
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = new List<string>() { RaceFieldsFormat.Text.Style.Bold },
                                WsColumnIndex = 17
                            });

                            var winPercentage = ((h.MJS1157 + h.MJS1162) > 0 && (h.MJS1156 + h.MJS1161) > 0) ? (h.MJS1157 + h.MJS1162) / (h.MJS1156 + h.MJS1161) : (decimal)0.00;
                            var top3FinishPercentage = ((h.MJS1157 + h.MJS1158 + h.MJS1159 + h.MJS1162 + h.MJS1163 + h.MJS1164) > 0 && (h.MJS1156 + h.MJS1161) > 0) ?
                                (h.MJS1157 + h.MJS1158 + h.MJS1159 + h.MJS1162 + h.MJS1163 + h.MJS1164) / (h.MJS1156 + h.MJS1161) : (decimal)0.00;
                            dr[0] = winPercentage + top3FinishPercentage;
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.ML: //ML
                            var mlStyles = new List<string>();

                            mlStyles.Add(RaceFieldsFormat.Text.Style.Regular);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.ML,
                                BasisType = RaceFieldsFormat.BasisTypes.LowestValue,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbRed,
                                TextStyles = mlStyles,
                                WsColumnIndex = 2
                            });
                            dr[0] = h.MorningLine;
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.PP: //PP
                            var ppStyles = new List<string>();

                            ppStyles.Add(RaceFieldsFormat.Text.Style.Regular);
                            
                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.PP,
                                BasisType = RaceFieldsFormat.BasisTypes.LessThanZero,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbRed,
                                TextStyles = ppStyles,
                                WsColumnIndex = 7
                            });
                            dr[0] = (decimal)h.PostPoints;
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.Pace: //Pace
                            var paceStyles = new List<string>();
                            var paceEvaluationValues = new List<decimal>();

                            paceStyles.Add(RaceFieldsFormat.Text.Style.Bold);
                            paceEvaluationValues.Add((decimal)2);
                            
                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.Pace,
                                BasisType = RaceFieldsFormat.BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = paceStyles,
                                WsColumnIndex = 8
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.Pace,
                                BasisType = RaceFieldsFormat.BasisTypes.WithinRange,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = paceStyles,
                                WsColumnIndex = 8,
                                EvaluationValues = paceEvaluationValues
                            });
                            dr[0] = (decimal)h.Pace;
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.PPWR: //PPWR
                            var ppwrStyles = new List<string>();
                            var ppwrStyles1 = new List<string>();
                            var ppwrStyles2 = new List<string>();
                            var ppwrEvaluationValues = new List<decimal>();

                            ppwrStyles.Add(RaceFieldsFormat.Text.Style.Regular);
                            ppwrStyles1.Add(RaceFieldsFormat.Text.Style.Bold);
                            ppwrStyles2.Add(RaceFieldsFormat.Text.Style.Italic);

                            ppwrEvaluationValues.Add((decimal)3.5);
                            
                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.PPWR,
                                BasisType = RaceFieldsFormat.BasisTypes.Top5,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = ppwrStyles,
                                WsColumnIndex = 16
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.PPWR,
                                BasisType = RaceFieldsFormat.BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = ppwrStyles1,
                                WsColumnIndex = 16
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.PPWR,
                                BasisType = RaceFieldsFormat.BasisTypes.WithinRangeOfLastHorseInTopFive,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = ppwrStyles2,
                                WsColumnIndex = 16,
                                EvaluationValues = ppwrEvaluationValues
                            });
                            dr[0] = Decimal.Round(h.PPWR, 1);
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.RBC: //RBC
                            var rbcStyles = new List<string>();
                            var rbcStyles1 = new List<string>();
                            var rbcStyles2 = new List<string>();
                            var rbcEvaluationValues = new List<decimal>();
                            var rbcEvaluationValues1 = new List<decimal>();
                            var rbcEvaluationValues2 = new List<decimal>();
                            var rbcEvaluationValues3 = new List<decimal>();

                            rbcStyles.Add(RaceFieldsFormat.Text.Style.Bold);
                            rbcEvaluationValues.Add((decimal)0.5);
                            rbcEvaluationValues1.Add((decimal)1);
                            rbcEvaluationValues2.Add((decimal)0.33);
                            rbcEvaluationValues3.Add((decimal)0.0);
                            
                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.RBC,
                                BasisType = RaceFieldsFormat.BasisTypes.BaseAmountOrHigher,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = rbcStyles,
                                WsColumnIndex = 13,
                                EvaluationValues = rbcEvaluationValues
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.RBC,
                                BasisType = RaceFieldsFormat.BasisTypes.Equals,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = rbcStyles,
                                WsColumnIndex = 13,
                                EvaluationValues = rbcEvaluationValues1
                            });
                            rbcStyles1.Add(RaceFieldsFormat.Text.Style.Italic);
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.RBC,
                                BasisType = RaceFieldsFormat.BasisTypes.Equals,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextStyles = rbcStyles1,
                                TextColor = XlRgbColor.rgbBlack,
                                WsColumnIndex = 13,
                                EvaluationValues = rbcEvaluationValues2
                            });
                            rbcStyles2.Add(RaceFieldsFormat.Text.Style.Regular);
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.RBC,
                                BasisType = RaceFieldsFormat.BasisTypes.Equals,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextStyles = rbcStyles2,
                                TextColor = XlRgbColor.rgbRed,
                                WsColumnIndex = 13,
                                EvaluationValues = rbcEvaluationValues3
                            });
                            dr[0] = Decimal.Round(h.RBCPercent, 2);
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.RQ: //RQ
                            var rqStyles = new List<string>();
                            var rqEvaluationValues = new List<decimal>();

                            rqStyles.Add(RaceFieldsFormat.Text.Style.Regular);
                            rqEvaluationValues.Add((decimal)70);
                            
                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.RQ,
                                BasisType = RaceFieldsFormat.BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = rqStyles,
                                WsColumnIndex = 6
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.RQ,
                                BasisType = RaceFieldsFormat.BasisTypes.BaseAmountOrHigher,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = rqStyles,
                                WsColumnIndex = 6,
                                EvaluationValues = rqEvaluationValues
                            });
                            dr[0] = (decimal)(h.RunStyle + h.Quirin);
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.TB: //TB
                            var tbStyles = new List<string>();
                            var tbEvaluationValues = new List<decimal>();

                            tbStyles.Add(RaceFieldsFormat.Text.Style.Bold);
                            tbEvaluationValues.Add((decimal)219.00);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.TB,
                                BasisType = RaceFieldsFormat.BasisTypes.GreaterThanOrEqualTo,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                EvaluationValues = tbEvaluationValues,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = tbStyles,
                                WsColumnIndex = 20
                            });
                            dr[0] = h.TB;
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.TotalPace: //TotalPace
                            var totalPaceStyles = new List<string>();

                            totalPaceStyles.Add(RaceFieldsFormat.Text.Style.Bold);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.TotalPace,
                                BasisType = RaceFieldsFormat.BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = totalPaceStyles,
                                WsColumnIndex = 5
                            });
                            dr[0] = h.Total;
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.TSR: //TSR
                            var tsrStyles = new List<string>();
                            var tsrEvaluationValues = new List<decimal>();

                            tsrStyles.Add(RaceFieldsFormat.Text.Style.Bold);
                            tsrEvaluationValues.Add(3);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.TSR,
                                BasisType = RaceFieldsFormat.BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGray,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = tsrStyles,
                                WsColumnIndex = 23
                            });

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.TSR,
                                BasisType = RaceFieldsFormat.BasisTypes.HighestValueWithinFloorRange,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = tsrStyles,
                                WsColumnIndex = 23,
                                EvaluationValues = tsrEvaluationValues
                            });
                            dr[0] = h.TSR;
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.RnkWrkrsPercentage1:
                            var rnkWrkrsStyles1 = new List<string>();
                            var rnkWrkrsEvaluationValues1 = new List<decimal>();

                            rnkWrkrsStyles1.Add(RaceFieldsFormat.Text.Style.Bold);
                            
                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.RnkWrkrsPercentage1,
                                BasisType = RaceFieldsFormat.BasisTypes.RnkWrkrsCustom,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = rnkWrkrsStyles1,
                                WsColumnIndex = 32,
                                EvaluationValues = rnkWrkrsEvaluationValues1
                            });
                            dr[0] = (h.RnkWrkrsPct < (decimal)16) && (h.Workers >= 40);
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.RnkWrkrsPercentage2:
                            var rnkWrkrsStyles2 = new List<string>();
                            var rnkWrkrsEvaluationValues2 = new List<decimal>();

                            rnkWrkrsStyles2.Add(RaceFieldsFormat.Text.Style.Bold);
                            
                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.RnkWrkrsPercentage2,
                                BasisType = RaceFieldsFormat.BasisTypes.RnkWrkrsCustom,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = rnkWrkrsStyles2,
                                WsColumnIndex = 32,
                                EvaluationValues = rnkWrkrsEvaluationValues2
                            });
                            dr[0] = (((h.RnkWrkrsPct >= (decimal)16) && (h.RnkWrkrsPct <= (decimal)30)) && (h.Workers >= 40)) || ((h.RnkWrkrsPct <= (decimal)30) && (h.RnkWrkrsPct > (decimal)0) && (h.Workers < 40));
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.Distance:
                            var distanceStyles = new List<string>();
                            var distanceEvaluationValues = new List<decimal>();

                            distanceStyles.Add(RaceFieldsFormat.Text.Style.Bold);
                            distanceEvaluationValues.Add(3500);
                            
                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.Distance,
                                BasisType = RaceFieldsFormat.BasisTypes.BaseAmountOrHigher,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = distanceStyles,
                                WsColumnIndex = 29,
                                EvaluationValues = distanceEvaluationValues
                            });
                            dr[0] = (decimal)h.Distance;
                            dr[1] = h;
                            break;
                        case RaceFieldsFormat.Fields.Workout:
                            var workoutStyles = new List<string>();
                            var workoutEvaluationValues = new List<decimal>();

                            workoutStyles.Add(RaceFieldsFormat.Text.Style.Bold);
                            workoutEvaluationValues.Add(4);
                            
                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = RaceFieldsFormat.Fields.Workout,
                                BasisType = RaceFieldsFormat.BasisTypes.BaseAmountOrHigher,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = workoutStyles,
                                WsColumnIndex = 28,
                                EvaluationValues = workoutEvaluationValues
                            });
                            dr[0] = (decimal)h.Workout;
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
                        case RaceFieldsFormat.BasisTypes.HighestValue:
                        case RaceFieldsFormat.BasisTypes.HighestValueWithinFloorRange:
                        case RaceFieldsFormat.BasisTypes.Top5:
                            ff.SortDirection = RaceFieldsFormat.SortDirections.Desc;
                            break;
                        case RaceFieldsFormat.BasisTypes.LowestValue:
                            ff.SortDirection = RaceFieldsFormat.SortDirections.Asc;
                            break;
                        case RaceFieldsFormat.BasisTypes.SecondHighestValue:
                        case RaceFieldsFormat.BasisTypes.WithinRangeOfLastHorseInTopFive:
                            ff.SortDirection = RaceFieldsFormat.SortDirections.Desc;
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
                        case RaceFieldsFormat.Fields.TotalPace: 
                        case RaceFieldsFormat.Fields.CR:
                        case RaceFieldsFormat.Fields.DSLR:
                        case RaceFieldsFormat.Fields.DSR: 
                        case RaceFieldsFormat.Fields.LP:
                        case RaceFieldsFormat.Fields.MDC:
                        case RaceFieldsFormat.Fields.MJS:
                        case RaceFieldsFormat.Fields.ML: 
                        case RaceFieldsFormat.Fields.BCR:
                        case RaceFieldsFormat.Fields.BSR:
                        case RaceFieldsFormat.Fields.PP:
                        case RaceFieldsFormat.Fields.Pace:
                        case RaceFieldsFormat.Fields.RQ:
                            val = Convert.ToDecimal(dtHorses.Rows[0][0]);
                            break;
                        case RaceFieldsFormat.Fields.TSR:
                        case RaceFieldsFormat.Fields.PPWR:  
                            if (f.Key == RaceFieldsFormat.Fields.PPWR && ff.BasisType == RaceFieldsFormat.BasisTypes.WithinRangeOfLastHorseInTopFive){
                                val = (decimal)dtHorses.Rows[4][0];
                            } else {
                                val = (decimal)dtHorses.Rows[0][0];
                            }
                            break;
                    }

                    switch (ff.BasisType)
                    {
                        case RaceFieldsFormat.BasisTypes.BaseAmountOrHigher:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (Convert.ToDecimal(r[0]) >= ff.EvaluationValues[0])
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case RaceFieldsFormat.BasisTypes.BaseAmountOrLower:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (Convert.ToDecimal(r[0]) <= ff.EvaluationValues[0])
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case RaceFieldsFormat.BasisTypes.BetweenTwoValues:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (Convert.ToDecimal(r[0]) >= ff.EvaluationValues[0] && Convert.ToDecimal(r[0]) <= ff.EvaluationValues[1])
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case RaceFieldsFormat.BasisTypes.Equals:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (Convert.ToDecimal(r[0]) == ff.EvaluationValues[0])
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case RaceFieldsFormat.BasisTypes.HighestValueWithinFloorRange:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                var highValue = (decimal)val;
                                var floorValue = (decimal)val - ff.EvaluationValues[0];

                                if ((Convert.ToDecimal(r[0]) <= highValue) && 
                                    (Convert.ToDecimal(r[0]) >= floorValue) && 
                                    (!Convert.ToDecimal(r[0]).Equals((decimal)0)))
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case RaceFieldsFormat.BasisTypes.HighestValue:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (Convert.ToDecimal(r[0]).Equals(val) && (Convert.ToDecimal(r[0]) > ((decimal)0)))
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case RaceFieldsFormat.BasisTypes.GreaterThanOrEqualTo:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                for (var iIndex = 0; iIndex < ff.EvaluationValues.Count(); iIndex++)
                                {
                                    var evalValue = ff.EvaluationValues[iIndex];
                                    var horseValue = (ff.HorseValues.Count() > 0) ? ff.HorseValues[iIndex] : r[0];

                                    if (Convert.ToDecimal(r[0]) > evalValue)
                                    {
                                        sortedHorses.Add((IHorse)r[1]);
                                    }
                                }
                            }
                            break;
                        case RaceFieldsFormat.BasisTypes.LowestValue:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (r[0].Equals(val))
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case RaceFieldsFormat.BasisTypes.LessThanZero:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (Convert.ToDecimal(r[0]) < (decimal)0.00)
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case RaceFieldsFormat.BasisTypes.Top5:
                            var iHorseCount = 0;
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (iHorseCount < 5)
                                {
                                    if (Convert.ToDecimal(r[0]) > (decimal)0)
                                    {
                                        sortedHorses.Add((IHorse)r[1]);
                                    }
                                }
                                else
                                {
                                    break;
                                }
                                iHorseCount++;
                            }
                            break;
                        case RaceFieldsFormat.BasisTypes.RnkWrkrsCustom:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (Convert.ToBoolean(r[0]))
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case RaceFieldsFormat.BasisTypes.SecondHighestValue:
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
                        case RaceFieldsFormat.BasisTypes.ValueExists:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (!string.IsNullOrWhiteSpace(r[0].ToString()))
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case RaceFieldsFormat.BasisTypes.WithinRange:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if ((Convert.ToDecimal(r[0]) <= Convert.ToDecimal(val) + ff.EvaluationValues[0]) && (Convert.ToDecimal(r[0]) >= Convert.ToDecimal(val) - ff.EvaluationValues[0]))
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case RaceFieldsFormat.BasisTypes.WithinRangeOfLastHorseInTopFive:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                if (Convert.ToDecimal(r[0]) >= Convert.ToDecimal(val) - ff.EvaluationValues[0] && Convert.ToDecimal(r[0]) < Convert.ToDecimal(val))
                                {
                                    sortedHorses.Add((IHorse)r[1]);
                                }
                            }
                            break;
                        case RaceFieldsFormat.BasisTypes.None:
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
                            var cell = ws.Cells[iIndex, 4];
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

                                            if (ff.BasisType == RaceFieldsFormat.BasisTypes.RnkWrkrsCustom)
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
                                            cell.Font.Bold = style.Equals(RaceFieldsFormat.Text.Style.Bold);
                                            cell.Font.Italic = style.Equals(RaceFieldsFormat.Text.Style.Italic);

                                            if (ff.BasisType == RaceFieldsFormat.BasisTypes.RnkWrkrsCustom)
                                            {
                                                var indexAdjust = ff.WsColumnIndex;
                                                ws.Cells[iIndex, --indexAdjust].Font.Bold = style.Equals(RaceFieldsFormat.Text.Style.Bold);
                                                ws.Cells[iIndex, indexAdjust].Font.Italic = style.Equals(RaceFieldsFormat.Text.Style.Italic);
                                                ws.Cells[iIndex, --indexAdjust].Font.Bold = style.Equals(RaceFieldsFormat.Text.Style.Bold);
                                                ws.Cells[iIndex, indexAdjust].Font.Italic = style.Equals(RaceFieldsFormat.Text.Style.Italic);
                                                ws.Cells[iIndex, --indexAdjust].Font.Bold = style.Equals(RaceFieldsFormat.Text.Style.Bold);
                                                ws.Cells[iIndex, indexAdjust].Font.Italic = style.Equals(RaceFieldsFormat.Text.Style.Italic);
                                                ws.Cells[iIndex, --indexAdjust].Font.Bold = style.Equals(RaceFieldsFormat.Text.Style.Bold);
                                                ws.Cells[iIndex, indexAdjust].Font.Italic = style.Equals(RaceFieldsFormat.Text.Style.Italic);
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

            //add break above horse with Total >= last horse's total + 8
            var lastHorseTotal = (decimal)-1;
            for (var iIndex = iRangeStart; iIndex <= iRangeEnd; iIndex++)
            {
                //Get the name cell
                var cell = ws.Cells[iIndex, 4];
                var row = ws.Rows[iIndex];

                foreach (var h in horses)
                {
                    //Check to see if it's our selected horse.  
                    if (cell.Value.Equals(h.HorseName))
                    {
                        if ((lastHorseTotal > -1) && (lastHorseTotal >= (h.Total + 8)))
                        {
                            row.Insert();
                            var newRow = ws.Rows[iIndex];
                            newRow.ClearFormats();
                            iIndex++;
                            iRangeEnd++;
                        }

                        lastHorseTotal = h.Total;
                    }
                }
            }

            return iRangeEnd;
            Marshal.ReleaseComObject(ws);
        }
    }
}
