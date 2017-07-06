using System;
using System.Linq;
using System.Linq.Expressions;

namespace Nilsen.Framework.Data.Repository.Objects.Interfaces
{
    public interface IRepository<T>
    {
        IXmlContext XmlContext { get; set; }
    }
}