using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Nilsen.Framework.Objects.Interfaces;

namespace Nilsen.Framework.Objects.Class
{
    public sealed class Track : ITrack
    {
        #region "constructors"
        public Track(String[] Fields)
        {
            //declares and assigns
            string tt;
            AllWeather = (Fields[24].ToLower() == "a");
            TrackTypes.TryGetValue(Fields[6], out tt);
            TrackType = tt;
            TrackTypeShort = Fields[6];
            TrackName = Fields[0];
            Furlongs = Math.Round(((Convert.ToDecimal(Regex.Replace(Fields[5], "[^.0-9]", "")) * 3) / 5280) * 8, 2);
        }
        #endregion

        #region "properties"
        public Boolean AllWeather { get; set; }
        public String TrackName { get; set; }
        public String TrackType { get; set; }
        public String TrackTypeShort { get; set; }
        public Decimal Furlongs { get; set; }

        public IDictionary<String, String> TrackTypes = new Dictionary<String, String>() { 
            { "D", "Dirt" }, 
            { "d", "Inner Dirt" }, 
            { "T", "Turf" }, 
            { "t", "Inner Turf" } , 
            { "s", "Steeplechase" }, 
            { "h", "hunt" } 
        };
        #endregion
    }
}
