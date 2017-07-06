using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Nilsen.Framework.Services.Objects.Interfaces;
using Nilsen.Framework.Objects.Class;
using Nilsen.Framework.Objects.Interfaces;

namespace Nilsen.Framework.Services.Objects.Classes
{
    public static class RaceService
    {
        public static IRace GetRace(string[] fields)
        {
            return new Race(fields);
        }
    }
}
