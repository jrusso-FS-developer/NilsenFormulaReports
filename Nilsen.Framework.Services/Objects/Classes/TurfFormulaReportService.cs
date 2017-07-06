using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Text.RegularExpressions;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using System.Configuration;
using Microsoft.VisualBasic.FileIO;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Nilsen.Framework.Services.Objects.Interfaces;

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
                wb.SaveAs(sbFullFileName.ToString(), Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false,
                        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
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

        private void Top5Calc(Worksheet ws, Int32 iHorse, Int32 iTop5Row, 
                                String[,] sAllHorses, Int16 iHorseCount)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            System.Data.DataTable dtTop5 = new System.Data.DataTable();
            System.Data.DataColumn dc = new System.Data.DataColumn();
            StringBuilder sbTop5Horses = new StringBuilder();
            dt.Columns.Add(new DataColumn("NilsenRating", System.Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("HorseName", System.Type.GetType("System.String")));
            for (var jIndex = 1; jIndex <= iHorse; jIndex++)
            {
                dt.Rows.Add(dt.NewRow());

                dt.Rows[jIndex - 1][1] = sAllHorses[jIndex, 0];
                dt.Rows[jIndex - 1][0] = sAllHorses[jIndex, 1];
            }

            dt.DefaultView.Sort = "NilsenRating desc";
            dtTop5 = dt.DefaultView.ToTable();
            iHorseCount = 1;
            foreach (DataRow dr in dtTop5.Rows)
            {
                if (iHorseCount.Equals(5))
                {
                    sbTop5Horses.AppendFormat("{0}", dr[1].ToString());
                    break;
                }
                else
                {
                    sbTop5Horses.AppendFormat("{0}{1}", dr[1].ToString(), "-");
                }
                iHorseCount++;
            }
            ws.Cells[iTop5Row, 1].Value = string.Format("Turf:   {0}", sbTop5Horses.ToString());
            ws.get_Range(string.Format("A{0}", iTop5Row), string.Format("C{0}", iTop5Row)).Merge(Type.Missing);
        }

        public void BuildWorksheet(Worksheet ws, FileInfo fi)
        {
            //declares and assigns
            var reader = new StreamReader(File.OpenRead(fi.FullName));
            string[] Lines;
            string[] Fields;
            Int32 iRow = 1;
            Range rHeader;
            Int32 iTop5Row = 0;
            Int32 iHorse = 0;
            Int16 iHorseCount = 0;
            String[,] sAllHorses = new String[100, 2];
            Decimal[] decAllHorses = new Decimal[100];
            String[,] Top5Horses = new String[5, 2];
            TextFieldParser tfp = null;
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
                for (var iIndex = 0; iIndex < Lines.GetLength(0); iIndex++)
                {
                    tfp = new TextFieldParser(new StringReader(Lines[iIndex]));
                    tfp.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
                    tfp.SetDelimiters(new string[] { "," });
                    tfp.HasFieldsEnclosedInQuotes = true;

                    while (!tfp.EndOfData)
                    {
                        Fields = tfp.ReadFields();

                        if (Fields[6].ToLower().Equals("t")) //Either 'T' or 't', per spec.  
                        {
                            Int32 iWin;
                            Int32 iWP;
                            Int32 iWPS;
                            Int32 iSR;
                            Int32 iTurfPed;
                            Int32 iDSLR;
                            Decimal decAverageEarnings;
                            Decimal decTotalNilsenRating;
                            Decimal furlongs;
                            StringBuilder sbTurfPed = new StringBuilder();
                            StringBuilder sbTurfPedChars = new StringBuilder();

                            if (!Fields[2].ToLower().Equals(sRaceName))
                            {
                                if (!iHorse.Equals(0))
                                {
                                    Top5Calc(ws, iHorse, iTop5Row, sAllHorses, iHorseCount);
                                    iHorseCount = 0;
                                }

                                consoleSvc.UpdateConsoleText(string.Format("Reading and building Race {0}...", Fields[2]), false);
                                iRow = iRow + 2;
                                sRaceDate = DateTime.ParseExact(Fields[1], "yyyyMMdd", CultureInfo.InvariantCulture).ToString("MM/dd/yyyy");
                                sTrack = Fields[0];
                                sRaceName = Fields[2];

                                iHorse = 0;
                                sAllHorses = new String[100, 2];
                                decAllHorses = new Decimal[100];
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
                                iRow = iRow + 2;

                                //row headers
                                ws.Cells[iRow, 1].Value = "Prg #";
                                ws.Cells[iRow, 2].Value = "ML";
                                ws.Cells[iRow, 3].Value = "Horse Name";
                                ws.Cells[iRow, 4].Value = "Turf Rating";
                                ws.Cells[iRow, 5].Value = "Sts.";
                                ws.Cells[iRow, 6].Value = "Win";
                                ws.Cells[iRow, 7].Value = "Win%";
                                ws.Cells[iRow, 8].Value = "Place";
                                ws.Cells[iRow, 9].Value = "WP%";
                                ws.Cells[iRow, 10].Value = "Show";
                                ws.Cells[iRow, 11].Value = "WPS%";
                                ws.Cells[iRow, 12].Value = "Earnings";
                                ws.Cells[iRow, 13].Value = "AE";
                                ws.Cells[iRow, 14].Value = "SR";
                                ws.Cells[iRow, 15].Value = "Turf Ped.";
                                ws.Cells[iRow, 16].Value = "DSLR";

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
                            }

                            iRow++;

                            //calcs
                            foreach (Char c in Fields[1265].ToCharArray())
                            {
                                if (Char.IsNumber(c))
                                {
                                    sbTurfPed.Append(c);
                                }
                                else
                                {
                                    sbTurfPedChars.Append(c);
                                }
                            }

                            iWin = (Convert.ToInt32((string.IsNullOrEmpty(Fields[75]) ? "0" : Fields[75])).Equals(0) || Convert.ToInt32((string.IsNullOrEmpty(Fields[74]) ? "0" : Fields[74])).Equals(0)) ? 0 : Convert.ToInt32(Convert.ToDecimal((string.IsNullOrEmpty(Fields[75]) ? "0" : Fields[75])) / Convert.ToDecimal((string.IsNullOrEmpty(Fields[74]) ? "0" : Fields[74])) * 100);
                            iWP = ((Convert.ToDecimal((string.IsNullOrEmpty(Fields[76]) ? "0" : Fields[76])) + Convert.ToDecimal((string.IsNullOrEmpty(Fields[75]) ? "0" : Fields[75]))).Equals(0) || Convert.ToDecimal((string.IsNullOrEmpty(Fields[74]) ? "0" : Fields[74])).Equals(0)) ? 0 : Convert.ToInt32((Convert.ToDecimal((string.IsNullOrEmpty(Fields[76]) ? "0" : Fields[76])) + Convert.ToDecimal((string.IsNullOrEmpty(Fields[75]) ? "0" : Fields[75]))) / Convert.ToDecimal((string.IsNullOrEmpty(Fields[74]) ? "0" : Fields[74])) * 100);
                            iWPS = ((Convert.ToDecimal((string.IsNullOrEmpty(Fields[77]) ? "0" : Fields[77])) + Convert.ToDecimal((string.IsNullOrEmpty(Fields[76]) ? "0" : Fields[76])) + Convert.ToDecimal((string.IsNullOrEmpty(Fields[75]) ? "0" : Fields[75]))).Equals(0) || Convert.ToInt32((string.IsNullOrEmpty(Fields[74]) ? "0" : Fields[74])).Equals(0)) ? 0 : Convert.ToInt32((Convert.ToDecimal((string.IsNullOrEmpty(Fields[77]) ? "0" : Fields[77])) + Convert.ToDecimal((string.IsNullOrEmpty(Fields[76]) ? "0" : Fields[76])) + Convert.ToDecimal((string.IsNullOrEmpty(Fields[75]) ? "0" : Fields[75]))) / Convert.ToDecimal((string.IsNullOrEmpty(Fields[74]) ? "0" : Fields[74])) * 100);
                            decAverageEarnings = (Convert.ToInt32((string.IsNullOrEmpty(Fields[78]) ? "0" : Fields[78])).Equals(0) || Convert.ToInt32((string.IsNullOrEmpty(Fields[74]) ? "0" : Fields[74])).Equals(0)) ? 0 : (Convert.ToDecimal((string.IsNullOrEmpty(Fields[78]) ? "0" : Fields[78])) / Convert.ToDecimal((string.IsNullOrEmpty(Fields[74]) ? "0" : Fields[74])));
                            iSR = Convert.ToInt32((string.IsNullOrEmpty(Fields[1178]) ? "0" : Fields[1178]));
                            iTurfPed = Convert.ToInt32(sbTurfPed.ToString());
                            iDSLR = Convert.ToInt32((string.IsNullOrEmpty(Fields[223]) ? "0" : Fields[223]));
                            decTotalNilsenRating = (iWin * 5 + iWP * 2 + iWPS + decAverageEarnings / 200 + iSR + ((iTurfPed < 40) ? 110 : iTurfPed)) - ((Convert.ToDecimal(iDSLR / 1.3))) + (Convert.ToDecimal(Fields[74]) * Convert.ToDecimal(1.7));
                            iHorse++;
                            sAllHorses[iHorse, 0] = Fields[42];
                            sAllHorses[iHorse, 1] = decTotalNilsenRating.Equals(null) ? "0.00" : decTotalNilsenRating.ToString();

                            ws.Cells[iRow, 1].Value = string.Format("{0})", Fields[42]);
                            ws.Cells[iRow, 2].Value = Fields[43];
                            ws.Cells[iRow, 3].Value = textInfo.ToTitleCase(Fields[44].ToLower());
                            ws.Cells[iRow, 4].Value = Convert.ToInt32(Math.Round(decTotalNilsenRating, MidpointRounding.AwayFromZero));
                            ws.Cells[iRow, 5].Value = Fields[74];
                            ws.Cells[iRow, 6].Value = Fields[75];
                            ws.Cells[iRow, 7].Value = string.Format("{0}%", iWin);
                            ws.Cells[iRow, 8].Value = Fields[76];
                            ws.Cells[iRow, 9].Value = string.Format("{0}%", iWP);
                            ws.Cells[iRow, 10].Value = Fields[77];
                            ws.Cells[iRow, 11].Value = string.Format("{0}%", iWPS);
                            ws.Cells[iRow, 12].Value = string.Format("{0:C0}", Convert.ToDecimal(Fields[78]));
                            ws.Cells[iRow, 13].Value = string.Format("{0:C0}", decAverageEarnings);
                            ws.Cells[iRow, 14].Value = iSR;
                            ws.Cells[iRow, 15].Value = (iTurfPed < 40) ? string.Format("110{0}", sbTurfPedChars.ToString()) : string.Format("{0}{1}", sbTurfPed.ToString(), sbTurfPedChars.ToString());
                            ws.Cells[iRow, 16].Value = iDSLR;

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
                        }
                    }
                }

                if (!iHorse.Equals(0))
                {
                    Top5Calc(ws, iHorse, iTop5Row, sAllHorses, iHorseCount);
                }
            }

            //Column Widths
            consoleSvc.UpdateConsoleText("Auto-Fitting columns...", false);
            foreach (Range c in ws.get_Range("A1", "P1"))
            {
                c.EntireColumn.AutoFit();
            }
        }
    }
}
