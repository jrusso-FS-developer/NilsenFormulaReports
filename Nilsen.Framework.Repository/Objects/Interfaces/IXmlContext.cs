using System.Xml.Linq;

namespace Nilsen.Framework.Data.Repository.Objects.Interfaces
{
    public interface IXmlContext
    {
        string FileName { get; set; }
        XElement GetXml();
    }
}
