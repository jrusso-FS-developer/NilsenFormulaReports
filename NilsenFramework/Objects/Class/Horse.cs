using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Threading;
using System.Threading.Tasks;
using Nilsen.Framework.Objects.Interfaces;
using Nilsen.Framework.Data.Factory.Objects.Classes;
using System.Configuration;
using System.Globalization;

namespace Nilsen.Framework.Objects.Class
{
    public sealed class Horse : IHorse
    {
        private const String RunStyleXmlFile = "Runstyle.xml";
        private const String QuirinSpeedPointsXmlFile = "QuirinSpeedPoints.xml";
        private const String TrackPostXmlFile = "TrackPost.xml";
        private const String CPXmlFile = "CP.xml";

        public Horse(String[] Fields, IRace race)
        {            
            //declares and assigns
            Decimal outDec = (Decimal)0.00;
            Int16 outInt16 = 0;
            Int32 outInt32 = 0;
            int outInt = 0;
            KeyTrainerStatCategory = new List<String>();

            ProgramNumber = Fields[42].Trim();
            MorningLine = (Decimal.TryParse(Fields[43].Trim(), out outDec)) ? outDec : (Decimal)0.00;
            Note = Fields[40].Trim();
            Note2 = (Fields[61].Trim().Equals("4") || Fields[61].Trim().Equals("5")) ? "LASIX" : string.Empty;
            Note3 = string.Empty;
            DSLR = (Int32.TryParse(Fields[223].Trim(), out outInt32)) ? outInt32 : 0;
            ExtendedComment = Fields[1382];
            PPWR = (Decimal.TryParse(Fields[250].Trim(), out outDec)) ? outDec : (Decimal)0.00;
            CR = (Decimal.TryParse(Fields[1145].Trim(), out outDec)) ? outDec : (Decimal)0.00;
            Trk = (int.TryParse(Fields[70].Trim(), out outInt)) ? (outInt >= 1) ? "T" : string.Empty : string.Empty;
            DIS = (int.TryParse(Fields[65].Trim(), out outInt)) ? (outInt >= 1) ? "D" : string.Empty : string.Empty;
            TSR = (Decimal.TryParse(Fields[1330].Trim(), out outDec)) ? outDec : (Decimal)0.00;
            DSR = (Decimal.TryParse(Fields[1180].Trim(), out outDec)) ? outDec : (Decimal)0.00; 
            HorseName = Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(Fields[44].Trim().ToLower());
            Blinkers = (Int16.TryParse(Fields[63], out outInt16)) ? outInt16 : Convert.ToInt16(0);

            CalcCP(Fields, race);
            CalcPace(Fields);
            CalcLP(Fields, race.Track);
            CalcBCR(Fields, race);
            CalcBSR(Fields, race.Track);
            CalcRBCPercent(Fields);
            CalcRET(Fields, race);
            CalcMDC(Fields);
            CalcMJS(Fields);
            CalcWorkouts(Fields, race);
            CalcTB(Fields);
            CalcTotal();
            CalcRnkWrksPct();
            ProcessKeyTrainerChange(Fields, race);
            ProcessKeyTrainerStatCategories(Fields, race);
        }
        
        #region IHorse Members
        public Int16 Blinkers { get; set; }

        public Decimal ClaimingPrice { get; set; }

        public Decimal ClaimingPriceLastRace { get; set; }

        public Decimal CP { get; set; }

        public Decimal CR { get; set; }

        public Decimal BCR { get; set; }

        public Decimal BSR { get; set; }

        public String DIS { get; set; }

        public Int32 Distance { get; set; }

        public Int32 DSLR { get; set; }

        public Decimal DSR { get; set; }

        public String ExtendedComment { get; set; }

        public String HorseName { get; set; }

        public List<String> KeyTrainerStatCategory { get; set; }

        public decimal LastPurse { get; set; }

        public Int32 LP { get; set; }

        public String MDC { get; set; }

        public String MJS { get; set; }

        public decimal MJS1156 { get; set; }

        public decimal MJS1157 { get; set; }

        public decimal MJS1158 { get; set; }

        public decimal MJS1159 { get; set; }

        public decimal MJS1161 { get; set; }

        public decimal MJS1162 { get; set; }

        public decimal MJS1163 { get; set; }

        public decimal MJS1164 { get; set; }

        public Decimal MorningLine { get; set; }

        public String Note { get; set; }

        public String Note2 { get; set; }

        public String Note3 { get; set; }

        public Int32 Pace { get; set; }

        public Int32 PostPoints { get; set; }

        public Decimal PPWR { get; set; }

        public String ProgramNumber { get; set; }

        public Int32 Quirin { get; set; }

        public decimal RBCPercent { get; set; }

        public decimal RacePurse { get; set; }

        public Int32 Rank { get; set; }

        public String RET { get; set; }

        public decimal RnkWrkrsPct { get; set; }

        public Int32 RunStyle { get; set; }

        public decimal TB { get; set; }

        public Decimal Total { get; set; }

        public String Trk { get; set; }

        public Decimal TSR { get; set; }

        public Int32 Workers { get; set; }

        public Int32 Workout { get; set; }
        #endregion

        #region Private Methods
        private void CalcLP(String[] Fields, ITrack track)
        {
            //declares and assigns
            var dtRaces = new DataTable();
            var dtAWRaces = new DataTable();
            var iIndex = 0;
            var iSurfaceIndex = 325;
            var iPaceRatingsIndex = 815;
            var iAWSurfaceIndex = 1402;
            Int32 iOutVal;
            LP = 0;

            dtRaces.Columns.Add(new DataColumn("PaceRating", System.Type.GetType("System.Decimal")));
            dtAWRaces.Columns.Add(new DataColumn("PaceRating", System.Type.GetType("System.Decimal")));

            //process
            while (iIndex < 10)
            {
                iOutVal = 0;

                //get all allweather races 
                if (Fields[iAWSurfaceIndex + iIndex].ToLower() == "a" && dtAWRaces.Rows.Count < 3)
                {
                    dtAWRaces.Rows.Add(dtAWRaces.NewRow());
                    dtAWRaces.Rows[dtAWRaces.Rows.Count - 1][0] = int.TryParse(Fields[iPaceRatingsIndex + iIndex], out iOutVal) ? iOutVal : 0;
                }

                //get all races of this track type
                if ((((Fields[iSurfaceIndex + iIndex].ToLower().Equals(track.TrackTypeShort.ToLower())) ||
                    ((track.AllWeather) && (Fields[iSurfaceIndex + iIndex].ToLower().Equals("t")))) &&
                    (!Fields[iAWSurfaceIndex + iIndex].ToLower().Equals("a"))) &&
                    (dtRaces.Rows.Count < 3))
                {
                    dtRaces.Rows.Add(dtRaces.NewRow());
                    dtRaces.Rows[dtRaces.Rows.Count - 1][0] = int.TryParse(Fields[iPaceRatingsIndex + iIndex], out iOutVal) ? iOutVal : 0;
                }

                //increment
                ++iIndex;
            }

            //if track is All Weather, and there are no races in the table to choose from, then use a Turf / Inner Turf race.
            if (track.AllWeather)
            {
                dtRaces.DefaultView.Sort = "PaceRating desc";
                dtAWRaces.DefaultView.Sort = "PaceRating desc";

                LP = (dtAWRaces.Rows.Count > 0) ? Convert.ToInt32(dtAWRaces.DefaultView.ToTable().Rows[0]["PaceRating"].ToString()) : LP;
                LP = (LP.Equals((Decimal)0.00)) && (dtRaces.Rows.Count > 0) ? Convert.ToInt32(dtRaces.DefaultView.ToTable().Rows[0]["PaceRating"].ToString()) : LP;
            }
            else
            {
                switch (track.TrackTypeShort.ToLower())
                {
                    case "t":
                        dtRaces.DefaultView.Sort = "PaceRating desc";
                        dtAWRaces.DefaultView.Sort = "PaceRating desc";

                        LP = (dtRaces.Rows.Count > 0) ? Convert.ToInt32(dtRaces.DefaultView.ToTable().Rows[0]["PaceRating"].ToString()) : LP;
                        LP = (LP.Equals((Decimal)0.00)) && (dtAWRaces.Rows.Count > 0) ? Convert.ToInt32(dtAWRaces.DefaultView.ToTable().Rows[0]["PaceRating"].ToString()) : LP;

                        //if track is Turf / Inner Turf, and there are no races in the table to choose from, then use an All Weather race.
                        break;
                    default: //All others
                        if (dtRaces.Rows.Count > 0)
                        {
                            dtRaces.DefaultView.Sort = "PaceRating desc";
                            LP = Convert.ToInt32(dtRaces.DefaultView.ToTable().Rows[0]["PaceRating"].ToString());
                        }
                        break;
                }
            }
        }

        private void CalcBCR(String[] Fields, IRace race)
        {
            //declares and assigns
            var dtRaces = new DataTable();
            var dtAWRaces = new DataTable();
            var iIndex = 0;
            var iSurfaceIndex = 325;
            var iRaceDateIndex = 255;
            var iClassRatingsIndex = 835;
            var iAWSurfaceIndex = 1402;
            Decimal dOutVal;
            BCR = (Decimal)0.00;

            dtRaces.Columns.Add(new DataColumn("ClassRating", System.Type.GetType("System.Decimal")));
            dtAWRaces.Columns.Add(new DataColumn("ClassRating", System.Type.GetType("System.Decimal")));

            //process
            while (iIndex < 10)
            {
                dOutVal = (Decimal)0.00;
                DateTime horseRaceDate;
                var raceDate = race.Date;

                if ((DateTime.TryParseExact(Fields[iRaceDateIndex + iIndex], "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out horseRaceDate)))
                {
                    var raceDays = (raceDate - horseRaceDate).TotalDays;

                    if (raceDays < 240)
                    {
                        //get all allweather races 
                        if (Fields[iAWSurfaceIndex + iIndex].ToLower() == "a" && dtAWRaces.Rows.Count < 4)
                        {
                            dtAWRaces.Rows.Add(dtAWRaces.NewRow());
                            dtAWRaces.Rows[dtAWRaces.Rows.Count - 1][0] = Decimal.TryParse(Fields[iClassRatingsIndex + iIndex], out dOutVal) ? Convert.ToDecimal(Fields[iClassRatingsIndex + iIndex]) : (Decimal)0.00;
                        }

                        //get all races of this track type
                        if ((((Fields[iSurfaceIndex + iIndex].ToLower().Equals(race.Track.TrackTypeShort.ToLower())) ||
                            ((race.Track.AllWeather) && (Fields[iSurfaceIndex + iIndex].ToLower().Equals("t")))) &&
                            (!Fields[iAWSurfaceIndex + iIndex].ToLower().Equals("a"))) &&
                            (dtRaces.Rows.Count < 4))
                        {
                            dtRaces.Rows.Add(dtRaces.NewRow());
                            dtRaces.Rows[dtRaces.Rows.Count - 1][0] = Decimal.TryParse(Fields[iClassRatingsIndex + iIndex], out dOutVal) ? Convert.ToDecimal(Fields[iClassRatingsIndex + iIndex]) : (Decimal)0.00;
                        }
                    }
                    else
                    {
                        iIndex = 10;
                    }
                }

                //increment
                ++iIndex;
            }

            //if track is All Weather, and there are no races in the table to choose from, then use a Turf / Inner Turf race.
            if (race.Track.AllWeather)
            {
                dtRaces.DefaultView.Sort = "ClassRating desc";
                dtAWRaces.DefaultView.Sort = "ClassRating desc";

                BCR = (dtAWRaces.Rows.Count > 0) ? Convert.ToDecimal(dtAWRaces.DefaultView.ToTable().Rows[0]["ClassRating"].ToString()) : BCR;
                BCR = (BCR.Equals((Decimal)0.00)) && (dtRaces.Rows.Count > 0) ? Convert.ToDecimal(dtRaces.DefaultView.ToTable().Rows[0]["ClassRating"].ToString()) : BCR;
            }
            else
            {
                switch (race.Track.TrackTypeShort.ToLower())
                {
                    case "t":
                        dtRaces.DefaultView.Sort = "ClassRating desc";
                        dtAWRaces.DefaultView.Sort = "ClassRating desc";

                        BCR = (dtRaces.Rows.Count > 0) ? Convert.ToDecimal(dtRaces.DefaultView.ToTable().Rows[0]["ClassRating"].ToString()) : BCR;
                        BCR = (BCR.Equals((Decimal)0.00)) && (dtAWRaces.Rows.Count > 0) ? Convert.ToDecimal(dtAWRaces.DefaultView.ToTable().Rows[0]["ClassRating"].ToString()) : BCR;

                        //if track is Turf / Inner Turf, and there are no races in the table to choose from, then use an All Weather race.
                        break;
                    default: //All others
                        if (dtRaces.Rows.Count > 0)
                        {
                            dtRaces.DefaultView.Sort = "ClassRating desc";
                            BCR = Convert.ToDecimal(dtRaces.DefaultView.ToTable().Rows[0]["ClassRating"].ToString());
                        }
                        break;
                }
            }
        }

        private void CalcBSR(String[] Fields, ITrack track)
        {
            var iBSRIndex = 845;
            var iDaysSinceIndex = 265;
            var dtBSR = new DataTable();
            var iDayMax = (DateTime.ParseExact(Fields[1], "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture) - DateTime.Now.AddMonths(-16)).TotalDays;
            var iDayCount = 0;
            BSR = (Decimal)0.00;

            dtBSR.Columns.Add(new DataColumn("SpeedRating", System.Type.GetType("System.Decimal")));

            for (var iIndex = 0; iIndex < 5; iIndex++)
            {
                DataRow nRow = null;
                var dslr = (iIndex < 1) ? DSLR : 0;

                iDayCount += (iIndex < 1) ? DSLR : (!string.IsNullOrWhiteSpace(Fields[iDaysSinceIndex + (iIndex - 1)])) ? Convert.ToInt32(Fields[iDaysSinceIndex + (iIndex - 1)]) : 0;

                if ((iDayCount <= iDayMax) && (!Fields[iBSRIndex + iIndex].Equals(string.Empty)))
                {
                    nRow = dtBSR.NewRow();
                    dtBSR.Rows.Add(nRow);
                    nRow[0] = Convert.ToDecimal(Fields[iBSRIndex + iIndex]);
                }
            }

            dtBSR.DefaultView.Sort = "SpeedRating desc";
            BSR = (dtBSR.Rows.Count > 0) ? Convert.ToDecimal(dtBSR.DefaultView.ToTable().Rows[0]["SpeedRating"].ToString()) : BSR;
        }

        private void CalcCP(String[] Fields, IRace race)
        {
            //Get Runstyle Value
            var RunStyleTable = (from elem in DataFactory.GetXml(RunStyleXmlFile).Elements((XName)"Options").Elements((XName)"Option")
                                 where (String)elem.Attribute("name").Value == Fields[209].Trim()
                                 select elem).ToList() ?? null;
            var furlongs = race.Track.Furlongs;
            List<XElement> PostTable = null;

            if (Fields[209].Trim().ToUpper().Equals("NA"))
            {
                RunStyle = 23;
            }
            else
            {
                RunStyle = (RunStyleTable == null) ? 0 : ((RunStyleTable.Count > 0) ? Convert.ToInt32(RunStyleTable.FirstOrDefault().Value) : 0);
            }

            //Get Quirin Value
            var QuirinTable = (from elem in DataFactory.GetXml(QuirinSpeedPointsXmlFile).Elements((XName)"Options").Elements((XName)"Option")
                               where (String)elem.Attribute("number").Value == Fields[210].Trim()
                               select elem).ToList() ?? null;

            Quirin = QuirinTable == null ? 0 : (QuirinTable.Count > 0) ? Convert.ToInt32(QuirinTable.FirstOrDefault().Value) : 0;

            //Get Post Value
            var PostElement = (from elem1 in
                                    ((from elem in DataFactory.GetXml(TrackPostXmlFile).Element((XName)"Options").Elements((XName)"Option")
                                    where (Int32)(elem.Attribute("post")) <= Convert.ToInt32(Fields[3].Trim())
                                    select elem).ToList().Last()).Element((XName)"Value").Elements((XName)"Track")
                                where elem1.Attribute("type").Value.ToString().IndexOf(Fields[6].Trim()) > -1
                                select elem1).ToList().FirstOrDefault();

            var PostElement1 = (from elem in PostElement.Elements((XName)"length")
                                where (!string.IsNullOrWhiteSpace(elem.Attribute("track-name").Value.ToString()) && (elem.Attribute("track-name").Value.ToString().Split(',').Contains(Fields[0].ToString().ToUpper())
                                        && ((elem.Attribute("distances").Value.ToString().IndexOf(furlongs.ToString()) > -1)
                                            || ((furlongs >= Convert.ToDecimal(elem.Attribute("distances").Value.ToString().Split((Char)',')[0])) &&
                                            (furlongs <= Convert.ToDecimal(elem.Attribute("distances").Value.ToString().Split((Char)',')[elem.Attribute("distances").Value.ToString().Split((Char)',').Length - 1])
                                                && furlongs >= Convert.ToDecimal(elem.Attribute("distances").Value.ToString().Split((Char)',')[0]))))))
                                select elem);

            if ((PostElement1 == null) || PostElement1.ToList().Count == 0)
            {
                var PostElement2 = (from elem in PostElement.Elements((XName)"length")
                                    where (string.IsNullOrWhiteSpace(elem.Attribute("track-name").Value.ToString())
                                            && ((elem.Attribute("distances").Value.ToString().IndexOf(furlongs.ToString()) > -1)
                                                || ((furlongs >= Convert.ToDecimal(elem.Attribute("distances").Value.ToString().Split((Char)',')[0])) &&
                                                (furlongs <= Convert.ToDecimal(elem.Attribute("distances").Value.ToString().Split((Char)',')[elem.Attribute("distances").Value.ToString().Split((Char)',').Length - 1])
                                                    && furlongs >= Convert.ToDecimal(elem.Attribute("distances").Value.ToString().Split((Char)',')[0])))))
                                    select elem);
                                
                PostTable = PostElement2 != null ? PostElement2.ToList() : null;
            }
            else
            {
                PostTable = PostElement1.ToList();
            }


            PostPoints = PostTable == null ? 0 : (PostTable.Count > 0) ? Convert.ToInt32(PostTable.FirstOrDefault().Value) : 0;

            //Get Contention Points. 
            Int32[] iCPArray = new Int32[3];
            var iParseOut = 0;
            var startCount = 0;

            for (var cpIndex = 575; cpIndex < 578; cpIndex++)
            {
                var mIndex = cpIndex;
                var valCount = 0;
                var iVal = 0;

                iCPArray[0] = (Fields[mIndex].Trim() != "") && (int.TryParse(Fields[mIndex].Trim(), out iParseOut)) ? Convert.ToInt16(Fields[mIndex].Trim()) : 0;

                mIndex += 10;
                iCPArray[1] = (Fields[mIndex].Trim() != "") && (int.TryParse(Fields[mIndex].Trim(), out iParseOut)) ? Convert.ToInt16(Fields[mIndex].Trim()) : 0;

                mIndex += 20;
                iCPArray[2] = (Fields[mIndex].Trim() != "") && (int.TryParse(Fields[mIndex].Trim(), out iParseOut)) ? Convert.ToInt16(Fields[mIndex].Trim()) : 0;

                for (var iIndex = 0; iIndex < iCPArray.Length; iIndex++)
                {
                    iVal += Convert.ToInt16((from el in DataFactory.GetXml(CPXmlFile).Element((XName)"Options").Elements((XName)"Option")
                                             where (Int16)el.Attribute("position") == iCPArray[iIndex]
                                             select (Int16)el.Attribute("value")).ToList().FirstOrDefault());
                    valCount += iCPArray[iIndex];
                }

                CP += iVal;

                startCount += (valCount > 0) ? 1 : 0;
            }

            CP = ((Decimal)CP * ((startCount == 1) ? (Decimal)3 : ((startCount == 2) ? (Decimal)1.5 : (Decimal)1.0)));
        }

        private void CalcPace(String[] Fields)
        {
            var iFieldLower = 765;
            var iCnt = 0;
            Pace = 0;
            System.Data.DataTable RacePaceDT = new System.Data.DataTable();
            RacePaceDT.Columns.Add(new System.Data.DataColumn("RacePace", System.Type.GetType("System.Int32")));

            for (var iIndex = iFieldLower; (iIndex <= iFieldLower + 9) && (iCnt < 4); iIndex++)
            {
                Int32 output;
                if (Int32.TryParse(Fields[iIndex].Trim(), out output))
                {
                    RacePaceDT.Rows.Add(RacePaceDT.NewRow());
                    RacePaceDT.Rows[iCnt++][0] = Convert.ToInt32(Fields[iIndex].Trim());
                }
            }

            RacePaceDT.DefaultView.Sort = "RacePace desc";
            var dt = RacePaceDT.DefaultView.ToTable();

            if (iCnt < 4)
            {
                foreach (System.Data.DataRow row in dt.Rows)
                {
                    Pace += Convert.ToInt32(row[0]);
                }
                if ((Pace > 0) && (iCnt > 0)) { Pace = (Pace / iCnt); }
            }
            else
            {
                Pace = Convert.ToInt32(dt.Rows[1][0]);
            }
        }

        private void CalcTotal()
        {
            Total = (((Decimal)RunStyle + (Decimal)Quirin) * (Decimal).60) + (Decimal)PostPoints + (Decimal)Pace + CP;

            //Add 7 points if horse is adding blinkers (Blinker == 1), deduct 5 if removing blinkers (Blinker == 2), or 0 if neither
            switch (Blinkers)
            {
                case 1:
                    Total += 7;
                    break;
                case 2:
                    Total -= 5;
                    break;
                default: //do nothing
                    break;
            }
        }

        private void CalcTB(String[] Fields)
        {
            var todayDistanceOut = (decimal)0.00;
            var lastDistanceOut = (decimal)0.00;
            var todayDistance = (decimal)0.00;
            var lastDistance = (decimal)0.00;

            todayDistance = (Decimal.TryParse(Fields[5].Trim(), out todayDistanceOut)) ? todayDistanceOut : (Decimal)0.00;
            lastDistance = (Decimal.TryParse(Fields[315].Trim(), out lastDistanceOut)) ? lastDistanceOut : (Decimal)0.00;

            TB = lastDistance - todayDistance;
        }

        private void CalcRBCPercent(String[] Fields)
        {
            const int secondCallIndex = 585;
            const int beatenLengthIndex = 685;
            const int finishPositionIndex = 615;
            const int beatenFinishPositionIndex = 745;

            var secondCallSums = new List<decimal>();
            var finishPositionSums = new List<decimal>();
            var evalCount = 0;
            var evalSum = (decimal)0.00;

            for (var iIndex = 0; iIndex < 3; iIndex++)
            {
                var secondCallNumber = (decimal)0.00;
                var beatenLengthNumber = (decimal)0.00;
                var secondCall = (decimal.TryParse(Fields[secondCallIndex + iIndex], out secondCallNumber)) ? secondCallNumber : 0;
                var beatenLength = (decimal.TryParse(Fields[beatenLengthIndex + iIndex], out beatenLengthNumber)) ? beatenLengthNumber : (decimal)0.00;

                secondCallSums.Add(secondCall + beatenLength);
            }

            for (var iIndex = 0; iIndex < 3; iIndex++)
            {
                var finishNumber = (decimal)0.00;
                var beatenLengthFinish = (decimal)0.00;
                var finishPosition = (decimal.TryParse(Fields[finishPositionIndex + iIndex], out finishNumber)) ? finishNumber : 0;
                var beatenLengthFinishPosition = (decimal.TryParse(Fields[beatenFinishPositionIndex + iIndex], out beatenLengthFinish)) ? beatenLengthFinish : 0;

                finishPositionSums.Add(finishPosition + beatenLengthFinishPosition);
            }

            for(var iIndex = 0; iIndex < 3; iIndex++)
            {
                var compareTotal = secondCallSums[iIndex] + finishPositionSums[iIndex];

                if (compareTotal > 0)
                {
                    evalCount++;
                    evalSum += (finishPositionSums[iIndex] <= secondCallSums[iIndex] + (decimal).5) ? 1 : 0;
                }
            }
            if ((evalCount > 0) && (evalSum > 0))
            {
                RBCPercent = (decimal)(evalSum / evalCount);
            }
        }

        private void CalcRET(String[] Fields, IRace race)
        {
            var raceDate = race.Date;
            var daysSinceLastRace1 = (!string.IsNullOrWhiteSpace(Fields[265]) ? Convert.ToInt32(Fields[265]) : 0);
            var daysSinceLastRace2 = (!string.IsNullOrWhiteSpace(Fields[266]) ? Convert.ToInt32(Fields[266]) : 0);
            var daysSinceLastRace3 = (!string.IsNullOrWhiteSpace(Fields[267]) ? Convert.ToInt32(Fields[267]) : 0);

            if (DSLR < 45)
            {
                if (daysSinceLastRace1 > 44)
                {
                    RET = "2L";
                }
                else
                {
                    if (daysSinceLastRace2 > 44)
                    {
                        RET = "3L";
                    }
                    else
                    {
                        if (daysSinceLastRace3 > 44)
                        {
                            RET = "4L";
                        }
                    }
                }               
            }
        }

        private void CalcMDC(String[] Fields)
        {
            var racePurse = 0;
            var stateBred = false;
            var lastRacePurse = 0;
            var lastRaceStateBred = false;
            var claimingPrice = (decimal)0.00;
            var claimingPriceLastRace = (decimal)0.00;

            int.TryParse(Fields[11], out racePurse);
            RacePurse = racePurse;
            stateBred = (Fields[238].ToLower().Equals("s"));
            int.TryParse(Fields[555], out lastRacePurse);
            LastPurse = lastRacePurse;
            lastRaceStateBred = (Fields[1105].ToLower().Equals("s"));
            decimal.TryParse(Fields[12], out claimingPrice);
            ClaimingPrice = claimingPrice;
            decimal.TryParse(Fields[1211], out claimingPriceLastRace);
            ClaimingPriceLastRace = claimingPriceLastRace;

            MDC = string.Empty;

            if ((lastRacePurse >= racePurse * 1.10) || (!lastRaceStateBred && stateBred)/* || (claimingPriceLastRace >= claimingPrice * (decimal)1.10)*/)
            {
                MDC = "MDC";
            }
        }

        private void CalcMJS(String[] Fields)
        {

            MJS1156 = (!string.IsNullOrWhiteSpace(Fields[1156])) ? Convert.ToDecimal(Fields[1156]) : (decimal)0.00;
            MJS1157 = (!string.IsNullOrWhiteSpace(Fields[1157])) ? Convert.ToDecimal(Fields[1157]) : (decimal)0.00;
            MJS1158 = (!string.IsNullOrWhiteSpace(Fields[1158])) ? Convert.ToDecimal(Fields[1158]) : (decimal)0.00;
            MJS1159 = (!string.IsNullOrWhiteSpace(Fields[1159])) ? Convert.ToDecimal(Fields[1159]) : (decimal)0.00;
            MJS1161 = (!string.IsNullOrWhiteSpace(Fields[1161])) ? Convert.ToDecimal(Fields[1161]) : (decimal)0.00;
            MJS1162 = (!string.IsNullOrWhiteSpace(Fields[1162])) ? Convert.ToDecimal(Fields[1162]) : (decimal)0.00;
            MJS1163 = (!string.IsNullOrWhiteSpace(Fields[1163])) ? Convert.ToDecimal(Fields[1163]) : (decimal)0.00;
            MJS1164 = (!string.IsNullOrWhiteSpace(Fields[1164])) ? Convert.ToDecimal(Fields[1164]) : (decimal)0.00;
            MJS = (Fields[32].ToLower() != Fields[1065].ToLower()) ? "MJS" : string.Empty;
        }

        private void CalcWorkouts(String[] Fields, IRace race)
        {
            var workoutIndex = 101;
            var distanceIndex = 137;
            var rankIndex = 197;
            var workersIndex = 185;

            for (var iIndex = 0; iIndex < 12; iIndex++)
            {
                DateTime workoutDate;
                var distance = 0;
                var rank = 0;
                var workers = 0;

                if (DateTime.TryParseExact(Fields[workoutIndex + iIndex], "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out workoutDate))
                {
                    if ((race.Date - workoutDate).TotalDays < 32)
                    {
                        Workout += 1;

                        if (int.TryParse(Fields[distanceIndex + iIndex], out distance))
                        {
                            Distance += distance;
                        }

                        if (int.TryParse(Fields[rankIndex + iIndex], out rank))
                        {
                            Rank += rank;
                        }

                        if (int.TryParse(Fields[workersIndex + iIndex], out workers))
                        {
                            Workers += workers;
                        }
                    }
                }
            }
        }

        private void ProcessKeyTrainerChange(String[] Fields, IRace race)
        {
            for (var keyTrainerChangeIndex = 1336; keyTrainerChangeIndex < 1365; keyTrainerChangeIndex += 5)
            {
                var keyTrainerChange = Fields[keyTrainerChangeIndex];

                if (keyTrainerChange.ToLower().Equals("1st after clm") || keyTrainerChange.ToLower().Equals("1st start w/trn"))
                {
                    double relativeWinPercent;
                    double numberOfStarts;

                    if ((double.TryParse(Fields[keyTrainerChangeIndex + 1], out numberOfStarts)) &&
                        (double.TryParse(Fields[keyTrainerChangeIndex + 2], out relativeWinPercent)))
                    {
                        Note3 = ((relativeWinPercent >= 19.00) && (numberOfStarts >= 20)) ? "MTS" : "TS";
                    }
                }
            }
        }

        private void ProcessKeyTrainerStatCategories(String[] Fields, IRace race)
        {
            for (var KeyTrainerStatCategoryIndex = 1336; KeyTrainerStatCategoryIndex < 1362; KeyTrainerStatCategoryIndex += 5)
            {
                var numberOfStarts = (!string.IsNullOrWhiteSpace(Fields[KeyTrainerStatCategoryIndex + 1])) ? Convert.ToInt32(Fields[KeyTrainerStatCategoryIndex + 1]) : 0;
                var winPercent = (!string.IsNullOrWhiteSpace(Fields[KeyTrainerStatCategoryIndex + 2])) ? Convert.ToDouble(Fields[KeyTrainerStatCategoryIndex + 2]) / 100 : (double)0.00;
                var twoDollarROI = (!string.IsNullOrWhiteSpace(Fields[KeyTrainerStatCategoryIndex + 4])) ? Convert.ToDouble(Fields[KeyTrainerStatCategoryIndex + 4]) / 100 : (double)0.00;
                var keyTrainerStatCategory = Fields[KeyTrainerStatCategoryIndex];

                if (numberOfStarts > 0)
                {
                    if (((winPercent >= .19) && (numberOfStarts > 19) && (twoDollarROI >= (double)0.00)) || 
                        ((winPercent >= (double).30) && (numberOfStarts > 49)) ||
                        (numberOfStarts >= 50 && winPercent >= .25) ||
                        (numberOfStarts >= 75 && winPercent >= .20) ||
                        (numberOfStarts >= 15 && winPercent >= .30 && twoDollarROI > (double)2.00) ||
                        (numberOfStarts >= 9 && winPercent >= .40 && twoDollarROI > (double)1.00))
                    {
                        KeyTrainerStatCategory.Add(keyTrainerStatCategory);
                    }
                }
            }
        }

        private void CalcRnkWrksPct()
        {
            RnkWrkrsPct = (Rank > 0 && Workers > 0) ? (decimal)(((double)Rank / (double)Workers) * 100) : (decimal)0.00;
        }
        #endregion
    }
}
