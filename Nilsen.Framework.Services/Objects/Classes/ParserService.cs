using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Nilsen.Framework.Services.Objects.Classes
{
    public static class ParserService
    {
        public static string[] ParseCSV(string sCSV, char delimiter)
        {
            string[] fields = new string[1];
            string[] fieldsCopy = null;
            int iCnt = 0;
            int iFieldIndex = 0;

            foreach (char c in sCSV)
            {
                if (c.Equals('"'))
                {
                    iCnt++;
                }
                else
                {
                    if ((c.Equals(delimiter)) && ((iCnt % 2).Equals(0)))
                    {
                        fieldsCopy = fields;
                        fields = new string[fieldsCopy.Length + 1];
                        fieldsCopy.CopyTo(fields, 0);
                        iFieldIndex++;
                        fields[iFieldIndex] = string.Format("{0}{1}", fields[iFieldIndex], c.ToString());
                    }
                    else
                    {
                        fields[iFieldIndex] = string.Format("{0}{1}", fields[iFieldIndex], c.ToString());
                    }
                }
            }

            return fields;
        }
    }
}
