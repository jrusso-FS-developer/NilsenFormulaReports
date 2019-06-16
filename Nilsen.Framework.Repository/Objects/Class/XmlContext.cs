using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Threading.Tasks;
using Nilsen.Framework.Data.Repository.Objects.Interfaces;

namespace Nilsen.Framework.Data.Repository.Objects.Class
{
    public class XmlContext : IXmlContext
    {
        public string FileName { get; set; }

        public XmlContext(string fileName)
        {
            FileName = fileName;
        }

        public XElement GetXml()
        {
            var xmlElement = XElement.Load(Path.GetFullPath(string.Format("C:/Program Files/NilsenRaceFormulaReports/Data/Xml/{1}", Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory)))), FileName)));

            return xmlElement;
        }
    }
}
