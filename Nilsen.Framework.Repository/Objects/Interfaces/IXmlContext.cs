using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Nilsen.Framework.Data.Repository.Objects.Interfaces
{
    public interface IXmlContext
    {
        string FileName { get; set; }
        XElement GetXml();
    }
}
