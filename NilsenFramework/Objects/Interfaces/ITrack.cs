using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Nilsen.Framework.Objects.Interfaces
{
    public interface ITrack
    {
        Boolean AllWeather { get; set; }
        String TrackName { get; set; }
        String TrackType { get; set; }
        String TrackTypeShort { get; set; }
        Decimal Furlongs { get; set; }
    }
}
