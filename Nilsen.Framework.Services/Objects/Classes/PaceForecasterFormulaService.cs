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
    public class PaceForecasterFormulaReportService : IReportService
    {
        private ConsoleService consoleSvc;
        private string _SavePath = string.Format("{0}\\Flicker City Productions\\RacesCSVToExcel\\files", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
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
            var sbFullFileName = new StringBuilder();
            var sbFileName = new StringBuilder();
            var bContinue = true;

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
            var bInitialHeader = true;
            IRace race = null;
            var iTotalRow = 0;
            ws.Name = "Nilsen Pace Rating";
            ws.get_Range("A1", "F1").Merge(Type.Missing);
            rHeader = ws.get_Range("A1", Type.Missing);
            rHeader.Value = "Nilsen Pace Rating Report";
            rHeader.Font.Bold = true;
            ws.Cells[1, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            Lines = Regex.Split(reader.ReadToEnd(), Environment.NewLine);

            //page style
            ws.PageSetup.Orientation = XlPageOrientation.xlPortrait;
            var sRaceName = string.Empty;

            consoleSvc.UpdateConsoleText("Reading CSV File...", false);

            if (Lines.GetLength(0) > 0)
            {
                foreach (var line in Lines)
                {
                    var lineParser = new TextFieldParser(new StringReader(line));
                    lineParser.TextFieldType = FieldType.Delimited;
                    lineParser.SetDelimiters(new string[] {","});
                    lineParser.HasFieldsEnclosedInQuotes = true;

                    while (!lineParser.EndOfData)
                    {
                        Fields = lineParser.ReadFields();
                        //New Race Record
                        if (!Fields[2].ToLower().Equals(sRaceName))
                        {
                            if (race != null)
                            {
                                iRow = listHorses(race, ws, iRow, iHeaderRow);
                                ws.Cells[iTotalRow, 6].Value = race.GetTop3Total();

                                ws.Cells[iTotalRow, 6].Font.Bold = (race.GetTop3Total() >= 490);
                                if (race.GetTop3Total() >= 520)
                                    ws.Cells[iTotalRow, 6].Interior.Color = XlRgbColor.rgbRed;
                            }

                            race = RaceService.GetRace(Fields);

                            consoleSvc.UpdateConsoleText(string.Format("Reading and building Race {0}...", Fields[2]), false);
                            iRow = iRow + 2;

                            //Only show the Race Date and Track
                            if (bInitialHeader)
                            {
                                //race vars
                                string sRaceDate = DateTime.ParseExact(race.DateText, "yyyyMMdd", CultureInfo.InvariantCulture).ToString("MM/dd/yyyy");
                                ws.get_Range(string.Format("A{0}", iRow), string.Format("D{0}", iRow)).Merge(Type.Missing);
                                ws.Cells[iRow, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                ws.Cells[iRow++, 1].Value = sRaceDate;
                                ws.get_Range(string.Format("A{0}", iRow), string.Format("D{0}", iRow)).Merge(Type.Missing);
                                ws.Cells[iRow, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                ws.Cells[iRow++, 1].Value = race.Track.TrackName;
                                bInitialHeader = false;
                            }

                            iRow++;

                            sRaceName = Fields[2];
                            ws.get_Range(string.Format("A{0}", iRow), string.Format("E{0}", iRow)).Merge(Type.Missing);
                            ws.Cells[iRow++, 1].Value = string.Format("Race {0}", race.Name);
                            ws.get_Range(string.Format("A{0}", iRow), string.Format("E{0}", iRow)).Merge(Type.Missing);
                            ws.Cells[iRow, 1].Value = string.Format("{0} Furlongs / {1}", race.Track.Furlongs, (race.Track.AllWeather) ? "All Weather" : race.Track.TrackType); //Distance / Turf
                            if (race.Track.TrackTypeShort.ToLower().Equals("t"))
                            {
                                ws.Cells[iRow, 1].Interior.Color = XlRgbColor.rgbGreen;
                                ws.Cells[iRow, 1].Font.Color = XlRgbColor.rgbWhite;
                                ws.Cells[iRow, 1].Font.Bold = true;
                            }
                            iRow++;
                            ws.get_Range(string.Format("A{0}", iRow), string.Format("E{0}", iRow)).Merge(Type.Missing);

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

                            ws.get_Range(string.Format("A{0}", iRow), string.Format("E{0}", iRow)).Merge(Type.Missing);
                            ws.Cells[iRow++, 1].Value = string.Format("{0}", race.PostTime);
                            ws.get_Range(string.Format("A{0}", iRow), string.Format("G{0}", iRow)).Merge(Type.Missing);
                            ws.Cells[iRow++, 1].Value = race.GetAgeOfRace();
                            ws.get_Range(string.Format("A{0}", iRow), string.Format("H{0}", iRow)).Merge(Type.Missing);
                            ws.Cells[iRow++, 1].Value = string.Format("PAR = {0}", race.PAR);
                            iTotalRow = Convert.ToInt16(iRow); 

                            iRow++;

                            iHeaderRow = iRow;

                            //Horse headers
                            ws.Cells[iHeaderRow, 1].Value = "Prg #";
                            ws.Cells[iHeaderRow, 2].Value = "ML";
                            ws.Cells[iHeaderRow, 3].Value = "BL";
                            ws.Cells[iHeaderRow, 4].Value = "J";
                            ws.Cells[iHeaderRow, 5].Value = "Horse Name";
                            ws.Cells[iHeaderRow, 6].Value = "TOTAL";
                            ws.Cells[iHeaderRow, 7].Value = "R/Q";
                            ws.Cells[iHeaderRow, 8].Value = "PP";
                            ws.Cells[iHeaderRow, 9].Value = "Pace";
                            ws.Cells[iHeaderRow, 10].Value = "CP";
                            ws.Cells[iHeaderRow, 11].Value = "DSLR";
                            ws.Cells[iHeaderRow, 12].Value = "CR";
                            ws.Cells[iHeaderRow, 13].Value = "LP";
                            ws.Cells[iHeaderRow, 14].Value = "RBC%";
                            ws.Cells[iHeaderRow, 15].Value = "B-CR";
                            ws.Cells[iHeaderRow, 16].Value = "B-SR";
                            ws.Cells[iHeaderRow, 17].Value = "PPWR";
                            ws.Cells[iHeaderRow, 18].Value = "MJS";
                            ws.Cells[iHeaderRow, 19].Value = "RET";
                            ws.Cells[iHeaderRow, 20].Value = "MDC";
                            ws.Cells[iHeaderRow, 21].Value = "TB";
                            ws.Cells[iHeaderRow, 22].Value = "MUD";
                            ws.Cells[iHeaderRow, 23].Value = "TRF";
                            ws.Cells[iHeaderRow, 24].Value = "DST";
                            ws.Cells[iHeaderRow, 25].Value = "Trk";
                            ws.Cells[iHeaderRow, 26].Value = "Dis";
                            ws.Cells[iHeaderRow, 27].Value = "T-SR";
                            ws.Cells[iHeaderRow, 28].Value = "D-SR";
                            ws.Cells[iHeaderRow, 29].Value = "Note";
                            ws.Cells[iHeaderRow, 30].Value = "Note 2";
                            ws.Cells[iHeaderRow, 31].Value = "Note 3";
                            ws.Cells[iHeaderRow, 32].Value = "#W";
                            ws.Cells[iHeaderRow, 33].Value = "#F";
                            ws.Cells[iHeaderRow, 34].Value = "RK";
                            ws.Cells[iHeaderRow, 35].Value = "WKrs";
                            ws.Cells[iHeaderRow, 36].Value = string.Empty;
                            ws.Cells[iHeaderRow, 37].Value = "Extended Comment";

                            ws.Cells[iHeaderRow, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            ws.Cells[iHeaderRow, 2].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            ws.Cells[iHeaderRow, 3].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            ws.Cells[iHeaderRow, 4].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            ws.Cells[iHeaderRow, 5].HorizontalAlignment = XlHAlign.xlHAlignLeft;
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
                            ws.Cells[iHeaderRow, 33].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            ws.Cells[iHeaderRow, 34].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            ws.Cells[iHeaderRow, 35].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            ws.Cells[iHeaderRow, 36].HorizontalAlignment = XlHAlign.xlHAlignRight;
                            ws.Cells[iHeaderRow, 37].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        }
                        race.Horses.Add(new Horse(fi.FullName, Fields, race));
                    }
                }

                if (race != null)
                {
                    iRow = listHorses(race, ws, iRow, iHeaderRow);                    
                    ws.Cells[iTotalRow, 6].Value = race.GetTop3Total();

                    ws.Cells[iTotalRow, 6].Font.Bold = (race.GetTop3Total() >= 490);
                    if (race.GetTop3Total() >= 520)
                        ws.Cells[iTotalRow, 6].Interior.Color = XlRgbColor.rgbRed;
                }
            }

            //Column Widths
            consoleSvc.UpdateConsoleText("Auto-Fitting columns...", false);
            foreach (Range c in ws.get_Range("A1", "AZ1"))
            {
                c.EntireColumn.AutoFit();
            }

            Marshal.ReleaseComObject(rHeader);
            Marshal.ReleaseComObject(ws);
        }

        private int listHorses(IRace race, Worksheet ws, int iRowRangeStart, int iHeaderRow)
        {
            var iRow = iRowRangeStart;
            var ktscFirstIndex = 38;
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
                ws.Cells[iRow, 4].Value = horse.MountCount == 1 ? "J" : string.Empty; //1 Jockey Mount
                ws.Cells[iRow, 5].Value = horse.HorseName; //Horse Name
                ws.Cells[iRow, 6].Value = horse.Total; //TOTAL
                ws.Cells[iRow, 7].Value = (horse.RunStyle + horse.Quirin); //RQ 
                ws.Cells[iRow, 8].Value = horse.PostPoints; //PP
                ws.Cells[iRow, 9].Value = horse.Pace; //Pace
                ws.Cells[iRow, 10].Value = horse.CP; //CP
                ws.Cells[iRow, 11].Value = horse.DSLR; //DSLR
                ws.Cells[iRow, 12].Value = horse.CR; //CR
                ws.Cells[iRow, 13].Value = horse.LP; //LP
                ws.Cells[iRow, 14].Value = horse.RBCPercent; //RBCPercent
                ws.Cells[iRow, 15].Value = horse.BCR; //BCR
                ws.Cells[iRow, 16].Value = horse.BSR; //BSR
                ws.Cells[iRow, 17].Value = horse.PPWR; //PPWR
                ws.Cells[iRow, 18].Value = horse.MJS; //MJS
                ws.Cells[iRow, 19].Value = horse.RET; //RET
                ws.Cells[iRow, 20].Value = horse.MDC; //MDC
                ws.Cells[iRow, 21].Value = (horse.TB >= 110) ? "TB" : string.Empty; //TB
                ws.Cells[iRow, 22].Value = horse.MUD; //MUD
                ws.Cells[iRow, 23].Value = horse.TRF; //TRF
                ws.Cells[iRow, 24].Value = horse.DST; //DST
                ws.Cells[iRow, 25].Value = horse.Trk; //Trk
                ws.Cells[iRow, 26].Value = horse.DIS; //Dis
                ws.Cells[iRow, 27].Value = horse.TSR; //TSR
                ws.Cells[iRow, 28].Value = horse.DSR; //DSR
                ws.Cells[iRow, 29].Value = horse.Note; //Note
                ws.Cells[iRow, 30].Value = horse.Note2; //Note2
                ws.Cells[iRow, 31].Value = horse.Note3; //Note3
                ws.Cells[iRow, 32].Value = horse.Workout; //Workout
                ws.Cells[iRow, 33].Value = horse.Distance; //Distance
                ws.Cells[iRow, 34].Value = horse.Rank; //Rank
                ws.Cells[iRow, 35].Value = horse.Workers; //WKrs
                ws.Cells[iRow, 36].Value = (Math.Round(horse.RnkWrkrsPct, MidpointRounding.AwayFromZero)).ToString() + "%"; //PERCENTAGE
                ws.Cells[iRow, 37].Value = horse.ExtendedComment; //Extended Comment

                if (horse.Blinkers.Equals(1) || horse.Blinkers.Equals(2))
                {
                    ws.Cells[iRow, 3].Interior.Color = XlRgbColor.rgbLightGray;
                }

                if (horse.JockeyMeetStarts <= 7)
                {
                    ws.Cells[iRow, 4].Font.Bold = true;
                }

                if (horse.Note.ToLower().Equals("lasix") || horse.Note.ToLower().Equals("mts"))
                {
                    ws.Cells[iRow, 29].Interior.Color = XlRgbColor.rgbLightGray;
                    ws.Cells[iRow, 29].Font.Bold = true;
                }

                if (horse.Note2.ToLower().Equals("lasix") || horse.Note2.ToLower().Equals("mts"))
                {
                    ws.Cells[iRow, 30].Interior.Color = XlRgbColor.rgbLightGray;
                    ws.Cells[iRow, 30].Font.Bold = true;
                }

                if (horse.Note3.ToLower().Equals("lasix") || horse.Note3.ToLower().Equals("mts") || horse.Note3.ToLower().Equals("ts"))
                {
                    ws.Cells[iRow, 31].Interior.Color = XlRgbColor.rgbLightGray;
                    ws.Cells[iRow, 31].Font.Bold = true;
                }

                ws.Cells[iRow, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 2].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 3].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 4].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ws.Cells[iRow, 5].HorizontalAlignment = XlHAlign.xlHAlignLeft;
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
                ws.Cells[iRow, 33].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 34].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 35].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 36].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ws.Cells[iRow, 37].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                //format for decimal.  
                ws.Cells[iRow, 12].NumberFormat = "0.00";
                ws.Cells[iRow, 14].NumberFormat = "0.00";
                ws.Cells[iRow, 15].NumberFormat = "0.00";
                ws.Cells[iRow, 17].NumberFormat = "0.0";

                for (var iColumnIndex = ktscFirstIndex; iColumnIndex <= ktscLastIndex; iColumnIndex++)
                {
                    ws.Cells[iRow, iColumnIndex].Value = (horse.KeyTrainerStatCategory.Count > (iColumnIndex - ktscFirstIndex)) ? horse.KeyTrainerStatCategory[iColumnIndex - ktscFirstIndex] : " ";
                    ws.Cells[iRow, iColumnIndex].HorizontalAlignment = XlHAlign.xlHAlignRight;
                }
            }

            int iRowRangeEnd = iRow;
            iRow = FormatFields(race.Horses, ws, iRowRangeStart, iRowRangeEnd, ktscFirstIndex, ktscLastIndex);

            return iRow;
        }

        public int FormatFields(List<IHorse> horses, Worksheet ws, int iRangeStart, int iRangeEnd, int ktscFirstIndex, int ktscLastIndex)
        {
            List<FieldFormat> fieldFormats = null;
            var sortedHorses = new List<IHorse>();
            var keyTrainerStatIndex = 0;

            //process
            foreach (var f in GetFieldList(FormTypes.PaceForecasterFormula))
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

                    //declares and assigns
                    DataRow dr;
                    switch (f.Key)
                    {
                        case PaceForecasterFormatFields.BCR: //BCR
                            var bcrStyles = new List<string>();
                            var bcrEvaluationValues = new List<decimal>();

                            bcrStyles.Add(Text.Style.Bold);
                            bcrEvaluationValues.Add((decimal)2.5);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.BCR,
                                BasisType = BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = bcrStyles,
                                WsColumnIndex = 15
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.BCR,
                                BasisType = BasisTypes.WithinRange,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = bcrStyles,
                                WsColumnIndex = 15,
                                EvaluationDecimalValues = bcrEvaluationValues
                            });
                            dr[0] = h.BCR;
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.BSR: //BSR
                            var bsrStyles = new List<string>();
                            var bsrEvaluationValues = new List<decimal>();

                            bsrStyles.Add(Text.Style.Bold);
                            bsrEvaluationValues.Add(7);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.BCR,
                                BasisType = BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = bsrStyles,
                                WsColumnIndex = 16
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.BSR,
                                BasisType = BasisTypes.WithinRange,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = bsrStyles,
                                WsColumnIndex = 16,
                                EvaluationDecimalValues = bsrEvaluationValues
                            });
                            dr[0] = h.BSR;
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.CP: //CP
                            var cpStyles = new List<string>();
                            var cpStyles1 = new List<string>();
                            var evaluationValues = new List<decimal>();

                            cpStyles.Add(Text.Style.Bold);
                            cpStyles1.Add(Text.Style.Italic);
                            evaluationValues.Add(39);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.CP,
                                BasisType = BasisTypes.GreaterThanOrEqualTo,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = cpStyles,
                                WsColumnIndex = 10,
                                EvaluationDecimalValues = evaluationValues
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.CP,
                                BasisType = BasisTypes.BetweenTwoValues,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = cpStyles,
                                WsColumnIndex = 10,
                                EvaluationRangeValues = new RangeValues<decimal, decimal>(29, 39)
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.CP,
                                BasisType = BasisTypes.BetweenTwoValues,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = cpStyles1,
                                WsColumnIndex = 10,
                                EvaluationRangeValues = new RangeValues<decimal, decimal>(19, 29)
                            });

                            dr[0] = h.CP;
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.CR: //CR
                            var crStyles = new List<string>();
                            var crEvaluationValues = new List<decimal>();

                            crStyles.Add(Text.Style.Bold);
                            crEvaluationValues.Add((decimal)1);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.CR,
                                BasisType = BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = crStyles,
                                WsColumnIndex = 12
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.CR,
                                BasisType = BasisTypes.WithinRange,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = crStyles,
                                WsColumnIndex = 12,
                                EvaluationDecimalValues = crEvaluationValues
                            });

                            dr[0] = h.CR;
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.DSR: //DSR
                            var dsrStyles = new List<string>();
                            var dsrEvaluationValues = new List<decimal>();

                            dsrStyles.Add(Text.Style.Bold);
                            dsrEvaluationValues.Add(3);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.DSR,
                                BasisType = BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGray,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = dsrStyles,
                                WsColumnIndex = 28
                            });

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.DSR,
                                BasisType = BasisTypes.HighestValueWithinFloorRange,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = dsrStyles,
                                WsColumnIndex = 28,
                                EvaluationDecimalValues = dsrEvaluationValues
                            });
                            dr[0] = (decimal)h.DSR;
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.DST: //DST
                            var dstStyles = new List<string>();
                            var dstStyles1 = new List<string>();

                            dstStyles.Add(Text.Style.Regular);
                            dstStyles1.Add(Text.Style.Bold);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.DST,
                                BasisType = BasisTypes.Top4,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = dstStyles,
                                WsColumnIndex = 24
                            });

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.DST,
                                BasisType = BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = dstStyles1,
                                WsColumnIndex = 24
                            });

                            dr[0] = decimal.Round(h.DST, 1);
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.DSLR: //DSLR
                            var dslrStyles = new List<string>();
                            var dslrStyles1 = new List<string>();
                            var dslrStyles2 = new List<string>();
                            var dslrEvaluationRangeValues = new List<decimal>();
                            var dslrEvaluationRangeValues1 = new List<decimal>();
                            var dslrEvaluationRangeValues2 = new RangeValues<decimal, decimal>(60, 280);
                            var dslrEvaluationRangeValues3 = new List<decimal>();

                            dslrStyles.Add(Text.Style.Bold);
                            dslrStyles1.Add(Text.Style.Italic);
                            dslrStyles2.Add(XlUnderlineStyle.xlUnderlineStyleDouble.ToString());
                            dslrStyles2.Add(Text.Style.Bold);
                            dslrEvaluationRangeValues.Add(280);
                            dslrEvaluationRangeValues1.Add(11);
                            dslrEvaluationRangeValues3.Add(60);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.DSLR,
                                BasisType = BasisTypes.GreaterThanOrEqualTo,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = dslrStyles2,
                                WsColumnIndex = 11,
                                EvaluationDecimalValues = dslrEvaluationRangeValues3
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.DSLR,
                                BasisType = BasisTypes.BaseAmountOrHigher,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbRed,
                                TextStyles = dslrStyles,
                                WsColumnIndex = 11,
                                EvaluationDecimalValues = dslrEvaluationRangeValues
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.DSLR,
                                BasisType = BasisTypes.BaseAmountOrLower,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = dslrStyles,
                                WsColumnIndex = 11,
                                EvaluationDecimalValues = dslrEvaluationRangeValues1
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.DSLR,
                                BasisType = BasisTypes.BetweenTwoValues,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = dslrStyles1,
                                WsColumnIndex = 11,
                                EvaluationRangeValues = dslrEvaluationRangeValues2
                            });
                            dr[0] = (decimal)h.DSLR;
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.KeyTrainerStatCategory1: //KeyTrainerStatCategory1
                            ktscStyles = new List<string>();
                            ktsIndex = ktscFirstIndex;

                            ktscStyles.Add(Text.Style.Regular);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.KeyTrainerStatCategory1,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                BasisType = BasisTypes.ValueNotInStringList,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = ktscStyles,
                                WsColumnIndex = ktsIndex, 
                                EvaluationStringValues = new List<string>(new string[] {"Dirt Starts", "Claiming", "Routes",
                                                                            "Maiden Clming", "Sprints", "Allowance", "Nongrd Stk" })
                            });

                            dr[0] = (h.KeyTrainerStatCategory.Count > 0) ? h.KeyTrainerStatCategory[0] : string.Empty;
                            dr[1] = h;

                            break;
                        case PaceForecasterFormatFields.KeyTrainerStatCategory2: //KeyTrainerStatCategory2
                            ktscStyles = new List<string>();
                            ktsIndex = ktscFirstIndex + 1;

                            ktscStyles.Add(Text.Style.Regular);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.KeyTrainerStatCategory2,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                BasisType = BasisTypes.ValueNotInStringList,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = ktscStyles,
                                WsColumnIndex = ktsIndex,
                                EvaluationStringValues = new List<string>(new string[] {"Dirt Starts", "Claiming", "Routes",
                                                                            "Maiden Clming", "Sprints", "Allowance", "Nongrd Stk" })
                            });

                            dr[0] = (h.KeyTrainerStatCategory.Count > 1) ? h.KeyTrainerStatCategory[1] : string.Empty;
                            dr[1] = h;

                            keyTrainerStatIndex++;
                            break;
                        case PaceForecasterFormatFields.KeyTrainerStatCategory3: //KeyTrainerStatCategory3
                            ktscStyles = new List<string>();
                            ktsIndex = ktscFirstIndex + 2;

                            ktscStyles.Add(Text.Style.Regular);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.KeyTrainerStatCategory3,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                BasisType = BasisTypes.ValueNotInStringList,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = ktscStyles,
                                WsColumnIndex = ktsIndex,
                                EvaluationStringValues = new List<string>(new string[] {"Dirt Starts", "Claiming", "Routes",
                                                                            "Maiden Clming", "Sprints", "Allowance", "Nongrd Stk" })
                            });

                            dr[0] = (h.KeyTrainerStatCategory.Count > 2) ? h.KeyTrainerStatCategory[2] : string.Empty;
                            dr[1] = h;

                            break;
                        case PaceForecasterFormatFields.LP: //LP
                            var lpStyles = new List<string>();

                            lpStyles.Add(Text.Style.Bold);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.LP,
                                BasisType = BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = lpStyles,
                                WsColumnIndex = 13
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.LP,
                                BasisType = BasisTypes.SecondHighestValue,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = lpStyles,
                                WsColumnIndex = 13
                            });
                            dr[0] = (decimal)h.LP;
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.MDC: //MDC
                            var mdcStyles = new List<string>();

                            mdcStyles.Add(Text.Style.Bold);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.MDC,
                                BasisType = BasisTypes.GreaterThanOrEqualTo,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                EvaluationDecimalValues = new List<decimal>() { (decimal)1.40, h.ClaimingPrice + (h.ClaimingPrice * (decimal)0.33) },
                                HorseValues = new List<decimal>() { Math.Round((h.LastPurse / h.RacePurse), 2), h.ClaimingPriceLastRace },
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = mdcStyles,
                                WsColumnIndex = 20
                            });

                            dr[0] = Math.Round(h.LastPurse / h.RacePurse);
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.MJS: //MJS
                            var mjsStyles = new List<string>();

                            mjsStyles.Add(Text.Style.Bold);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.MJS,
                                BasisType = BasisTypes.GreaterThanOrEqualTo,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                EvaluationDecimalValues = new List<decimal>() { (decimal).38 },
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = new List<string>() { Text.Style.Regular },
                                WsColumnIndex = 18
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.MJS,
                                BasisType = BasisTypes.GreaterThanOrEqualTo,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                EvaluationDecimalValues = new List<decimal>() { (decimal).59 },
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = new List<string>() { Text.Style.Bold },
                                WsColumnIndex = 18
                            });

                            var winPercentage = ((h.MJS1157 + h.MJS1162) > 0 && (h.MJS1156 + h.MJS1161) > 0) ? (h.MJS1157 + h.MJS1162) / (h.MJS1156 + h.MJS1161) : (decimal)0.00;
                            var top3FinishPercentage = ((h.MJS1157 + h.MJS1158 + h.MJS1159 + h.MJS1162 + h.MJS1163 + h.MJS1164) > 0 && (h.MJS1156 + h.MJS1161) > 0) ?
                                (h.MJS1157 + h.MJS1158 + h.MJS1159 + h.MJS1162 + h.MJS1163 + h.MJS1164) / (h.MJS1156 + h.MJS1161) : (decimal)0.00;
                            dr[0] = winPercentage + top3FinishPercentage;
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.ML: //ML
                            var mlStyles = new List<string>();

                            mlStyles.Add(Text.Style.Regular);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.ML,
                                BasisType = BasisTypes.LowestValue,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbRed,
                                TextStyles = mlStyles,
                                WsColumnIndex = 2
                            });
                            dr[0] = h.MorningLine;
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.MUD: //MUD
                            var mudStyles = new List<string>();
                            var mudStyles1 = new List<string>();

                            mudStyles.Add(Text.Style.Regular);
                            mudStyles1.Add(Text.Style.Bold);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.MUD,
                                BasisType = BasisTypes.Top4,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = mudStyles,
                                WsColumnIndex = 22
                            });

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.MUD,
                                BasisType = BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = mudStyles1,
                                WsColumnIndex = 22
                            });

                            dr[0] = decimal.Round(h.MUD, 1);
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.PP: //PP
                            var ppStyles = new List<string>();

                            ppStyles.Add(Text.Style.Regular);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.PP,
                                BasisType = BasisTypes.LessThanZero,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbRed,
                                TextStyles = ppStyles,
                                WsColumnIndex = 8
                            });
                            dr[0] = (decimal)h.PostPoints;
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.Pace: //Pace
                            var paceStyles = new List<string>();
                            var paceEvaluationValues = new List<decimal>();

                            paceStyles.Add(Text.Style.Bold);
                            paceEvaluationValues.Add(2);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.Pace,
                                BasisType = BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = paceStyles,
                                WsColumnIndex = 9
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.Pace,
                                BasisType = BasisTypes.WithinRange,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = paceStyles,
                                WsColumnIndex = 9,
                                EvaluationDecimalValues = paceEvaluationValues
                            });
                            dr[0] = (decimal)h.Pace;
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.PPWR: //PPWR
                            var ppwrStyles = new List<string>();
                            var ppwrStyles1 = new List<string>();
                            var ppwrStyles2 = new List<string>();
                            var ppwrEvaluationValues = new List<decimal>();

                            ppwrStyles.Add(Text.Style.Regular);
                            ppwrStyles1.Add(Text.Style.Bold);
                            ppwrStyles2.Add(Text.Style.Italic);

                            ppwrEvaluationValues.Add((decimal)3.5);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.PPWR,
                                BasisType = BasisTypes.Top5,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = ppwrStyles,
                                WsColumnIndex = 17
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.PPWR,
                                BasisType = BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = ppwrStyles1,
                                WsColumnIndex = 17
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.PPWR,
                                BasisType = BasisTypes.WithinRangeOfLastHorseInTopFive,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = ppwrStyles2,
                                WsColumnIndex = 17,
                                EvaluationDecimalValues = ppwrEvaluationValues
                            });
                            dr[0] = Decimal.Round(h.PPWR, 1);
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.RBC: //RBC
                            var rbcStyles = new List<string>();
                            var rbcStyles1 = new List<string>();
                            var rbcStyles2 = new List<string>();
                            var rbcEvaluationValues = new List<decimal>();
                            var rbcEvaluationValues1 = new List<decimal>();
                            var rbcEvaluationValues2 = new List<decimal>();
                            var rbcEvaluationValues3 = new List<decimal>();

                            rbcStyles.Add(Text.Style.Bold);
                            rbcEvaluationValues.Add((decimal)0.5);
                            rbcEvaluationValues1.Add((decimal)1);
                            rbcEvaluationValues2.Add((decimal)0.33);
                            rbcEvaluationValues3.Add((decimal)0.0);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.RBC,
                                BasisType = BasisTypes.BaseAmountOrHigher,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = rbcStyles,
                                WsColumnIndex = 14,
                                EvaluationDecimalValues = rbcEvaluationValues
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.RBC,
                                BasisType = BasisTypes.Equals,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = rbcStyles,
                                WsColumnIndex = 14,
                                EvaluationDecimalValues = rbcEvaluationValues1
                            });
                            rbcStyles1.Add(Text.Style.Italic);
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.RBC,
                                BasisType = BasisTypes.Equals,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextStyles = rbcStyles1,
                                TextColor = XlRgbColor.rgbBlack,
                                WsColumnIndex = 14,
                                EvaluationDecimalValues = rbcEvaluationValues2
                            });
                            rbcStyles2.Add(Text.Style.Regular);
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.RBC,
                                BasisType = BasisTypes.Equals,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextStyles = rbcStyles2,
                                TextColor = XlRgbColor.rgbRed,
                                WsColumnIndex = 14,
                                EvaluationDecimalValues = rbcEvaluationValues3
                            });
                            dr[0] = decimal.Round(h.RBCPercent, 2);
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.RQ: //RQ
                            var rqStyles = new List<string>();
                            var rqEvaluationValues = new List<decimal>();

                            rqStyles.Add(Text.Style.Regular);
                            rqEvaluationValues.Add((decimal)70);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.RQ,
                                BasisType = BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = rqStyles,
                                WsColumnIndex = 7
                            });
                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.RQ,
                                BasisType = BasisTypes.BaseAmountOrHigher,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = rqStyles,
                                WsColumnIndex = 7,
                                EvaluationDecimalValues = rqEvaluationValues
                            });
                            dr[0] = (decimal)(h.RunStyle + h.Quirin);
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.TB: //TB
                            var tbStyles = new List<string>();
                            var tbEvaluationValues = new List<decimal>();

                            tbStyles.Add(Text.Style.Bold);
                            tbEvaluationValues.Add((decimal)219.00);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.TB,
                                BasisType = BasisTypes.GreaterThanOrEqualTo,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                EvaluationDecimalValues = tbEvaluationValues,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = tbStyles,
                                WsColumnIndex = 21
                            });
                            dr[0] = h.TB;
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.TotalPace: //TotalPace
                            var totalPaceStyles = new List<string>();

                            totalPaceStyles.Add(Text.Style.Bold);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.TotalPace,
                                BasisType = BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = totalPaceStyles,
                                WsColumnIndex = 6
                            });

                            dr[0] = h.Total;
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.TRF: //TRF
                            var trfStyles = new List<string>();
                            var trfStyles1 = new List<string>();

                            trfStyles.Add(Text.Style.Regular);
                            trfStyles1.Add(Text.Style.Bold);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.MUD,
                                BasisType = BasisTypes.Top4,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = trfStyles,
                                WsColumnIndex = 23
                            });

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.MUD,
                                BasisType = BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = trfStyles1,
                                WsColumnIndex = 23
                            });

                            dr[0] = decimal.Round(h.TRF, 1);
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.TSR: //TSR
                            var tsrStyles = new List<string>();
                            var tsrEvaluationValues = new List<decimal>();

                            tsrStyles.Add(Text.Style.Bold);
                            tsrEvaluationValues.Add(3);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.TSR,
                                BasisType = BasisTypes.HighestValue,
                                BackgroundColor = XlRgbColor.rgbLightGray,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = tsrStyles,
                                WsColumnIndex = 27
                            });

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.TSR,
                                BasisType = BasisTypes.HighestValueWithinFloorRange,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = tsrStyles,
                                WsColumnIndex = 27,
                                EvaluationDecimalValues = tsrEvaluationValues
                            });
                            dr[0] = h.TSR;
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.RnkWrkrsPercentage1:
                            var rnkWrkrsStyles1 = new List<string>();
                            var rnkWrkrsEvaluationValues1 = new List<decimal>();

                            rnkWrkrsStyles1.Add(Text.Style.Bold);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.RnkWrkrsPercentage1,
                                BasisType = BasisTypes.RnkWrkrsCustom,
                                BackgroundColor = XlRgbColor.rgbLightGrey,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = rnkWrkrsStyles1,
                                WsColumnIndex = 36,
                                EvaluationDecimalValues = rnkWrkrsEvaluationValues1
                            });
                            dr[0] = (h.RnkWrkrsPct < (decimal)16) && (h.Workers >= 40);
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.RnkWrkrsPercentage2:
                            var rnkWrkrsStyles2 = new List<string>();
                            var rnkWrkrsEvaluationValues2 = new List<decimal>();

                            rnkWrkrsStyles2.Add(Text.Style.Bold);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.RnkWrkrsPercentage2,
                                BasisType = BasisTypes.RnkWrkrsCustom,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = rnkWrkrsStyles2,
                                WsColumnIndex = 36,
                                EvaluationDecimalValues = rnkWrkrsEvaluationValues2
                            });
                            dr[0] = (((h.RnkWrkrsPct >= (decimal)16) && (h.RnkWrkrsPct <= (decimal)30)) && (h.Workers >= 40)) || ((h.RnkWrkrsPct <= (decimal)30) && (h.RnkWrkrsPct > (decimal)0) && (h.Workers < 40));
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.Distance:
                            var distanceStyles = new List<string>();
                            var distanceEvaluationValues = new List<decimal>();

                            distanceStyles.Add(Text.Style.Bold);
                            distanceEvaluationValues.Add(3500);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.Distance,
                                BasisType = BasisTypes.BaseAmountOrHigher,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = distanceStyles,
                                WsColumnIndex = 30,
                                EvaluationDecimalValues = distanceEvaluationValues
                            });
                            dr[0] = (decimal)h.Distance;
                            dr[1] = h;
                            break;
                        case PaceForecasterFormatFields.Workout:
                            var workoutStyles = new List<string>();
                            var workoutEvaluationValues = new List<decimal>();

                            workoutStyles.Add(Text.Style.Bold);
                            workoutEvaluationValues.Add(4);

                            dr = dt.NewRow();
                            dt.Rows.Add(dr);
                            dr = dt.Rows[dt.Rows.Count - 1];

                            fieldFormats.Add(new FieldFormat
                            {
                                Field = PaceForecasterFormatFields.Workout,
                                BasisType = BasisTypes.BaseAmountOrHigher,
                                BackgroundColor = XlRgbColor.rgbWhite,
                                TextColor = XlRgbColor.rgbBlack,
                                TextStyles = workoutStyles,
                                WsColumnIndex = 32,
                                EvaluationDecimalValues = workoutEvaluationValues
                            });
                            dr[0] = (decimal)h.Workout;
                            dr[1] = h;
                            break;
                    }
                }

                foreach (var ff in fieldFormats)
                {
                    var val = new Object();
                    var exitFor = false;
                    sortedHorses.Clear();

                    //set sort direction
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
                        case PaceForecasterFormatFields.TotalPace: 
                        case PaceForecasterFormatFields.CR:
                        case PaceForecasterFormatFields.CP:
                        case PaceForecasterFormatFields.DSLR:
                        case PaceForecasterFormatFields.DSR: 
                        case PaceForecasterFormatFields.LP:
                        case PaceForecasterFormatFields.MDC:
                        case PaceForecasterFormatFields.MJS:
                        case PaceForecasterFormatFields.MUD:
                        case PaceForecasterFormatFields.TRF:
                        case PaceForecasterFormatFields.DST:
                        case PaceForecasterFormatFields.ML: 
                        case PaceForecasterFormatFields.BCR:
                        case PaceForecasterFormatFields.BSR:
                        case PaceForecasterFormatFields.PP:
                        case PaceForecasterFormatFields.Pace:
                        case PaceForecasterFormatFields.RQ:
                            val = Convert.ToDecimal(dtHorses.Rows[0][0]);
                            break;
                        case PaceForecasterFormatFields.TSR:
                        case PaceForecasterFormatFields.PPWR:  
                            if (f.Key == PaceForecasterFormatFields.PPWR && 
                                ff.BasisType == BasisTypes.WithinRangeOfLastHorseInTopFive){
                                if (dtHorses.Rows.Count >= 5)
                                    val = Convert.ToDecimal(dtHorses.Rows[4][0]);
                                else
                                    exitFor = true;
                            } else {
                                val = Convert.ToDecimal(dtHorses.Rows[0][0]);
                            }
                            break;
                    }

                    if (exitFor) break;

                    switch (ff.BasisType)
                    {
                        case BasisTypes.BaseAmountOrHigher:
                            foreach (DataRow r in dtHorses.Rows)
                            {
                                var horse = (IHorse)r[1];
                                if (Convert.ToDecimal(r[0]) >= ff.EvaluationDecimalValues[0])
                                {                                    
                                    sortedHorses.Add(horse);
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
                                if (Convert.ToDecimal(r[0]) >= ff.EvaluationRangeValues.RangeStart && Convert.ToDecimal(r[0]) < ff.EvaluationRangeValues.RangeEnd)
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
                                if (Convert.ToDecimal(r[0]).Equals(val) && (Convert.ToDecimal(r[0]) > 0))
                                {
                                    var horse = (IHorse)r[1];
                                    sortedHorses.Add(horse);

                                    if (f.Key.Equals(PaceForecasterFormatFields.CR))
                                    {
                                        horse.TopCR = true;
                                    }
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
                                    var horseValue = (ff.HorseValues.Count() > 0) ? ff.HorseValues[iIndex] : r[0];

                                    if (Convert.ToDecimal(r[0]) >= evalValue)
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
                                        iHorseCount ++;
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
                                    var found = false;

                                    foreach(var v in ff.EvaluationStringValues)
                                    {
                                        if (r[0].ToString().ToLower() == v.ToLower())
                                        {
                                            found = true;
                                            break;
                                        }
                                    }

                                    if (!found)
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
                            var cell = ws.Cells[iIndex, 5];
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

                                            if (ff.Field.Equals(PaceForecasterFormatFields.PPWR) && h.MountCount == 1)
                                            {
                                                ws.Cells[iIndex, 4].Interior.Color = ff.BackgroundColor;
                                            }

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
                                            if (style.Equals(XlUnderlineStyle.xlUnderlineStyleDouble.ToString()))
                                            {
                                                var underline = addUnderline(style, ff, h);
                                                cell.Font.Underline = underline;

                                                if (underline && ff.Field.Equals(PaceForecasterFormatFields.DSLR))
                                                {
                                                    var peerCell = ws.Cells[iIndex, ff.WsColumnIndex + 1];

                                                    peerCell.Font.Underline = underline;
                                                }
                                            }
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

            //add break above horse with Total >= last horse's total + 8
            var lastHorseTotal = (decimal)-1;
            for (var iIndex = iRangeStart; iIndex <= iRangeEnd; iIndex++)
            {
                //Get the name cell
                var cell = ws.Cells[iIndex, 5];
                var row = ws.Rows[iIndex];

                foreach (var h in horses)
                {
                    //Check to see if it's our selected horse.  
                    if (cell.Value.Equals(h.HorseName))
                    {
                        if ((lastHorseTotal > -1) && (lastHorseTotal >= (h.Total + 7)))
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
        }

        private bool addUnderline(string style, FieldFormat ff, IHorse h)
        {
            var returnValue = false;

            if (ff.Field.Equals(PaceForecasterFormatFields.DSLR)) {
                returnValue = (style.Equals(XlUnderlineStyle.xlUnderlineStyleDouble.ToString()) && h.TopCR);
            }
            else
            {
                returnValue = (style.Equals(XlUnderlineStyle.xlUnderlineStyleDouble.ToString()));
            }

            return returnValue;
        }
    }
}
