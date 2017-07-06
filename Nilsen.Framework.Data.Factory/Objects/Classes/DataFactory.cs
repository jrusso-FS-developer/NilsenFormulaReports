using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Threading.Tasks;
using Nilsen.Framework.Data.Repository;
using Nilsen.Framework.Data.Repository.Objects.Class;
using Nilsen.Framework.Data.Repository.Objects.Interfaces; 

namespace Nilsen.Framework.Data.Factory.Objects.Classes
{
    public static class DataFactory
    {
        public static XElement GetXml(String fileName)
        {
            return new XmlContext(fileName).GetXml();
        }
    }
}
