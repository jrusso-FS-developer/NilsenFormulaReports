using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Threading.Tasks;
using Nilsen.Framework.Data.Repository.Objects.Interfaces;

namespace Nilsen.Framework.Data.Repository.Objects.Class
{
    public class ReportOptionsContext : IXmlContext
    {
        public string FileName { get; set; }

        public ReportOptionsContext(string fileName)
        {
            FileName = fileName;
        }

        public XElement GetXml()
        {
            var xmlElement = XElement.Load(FileName);

            return xmlElement;
        }
    }
}
