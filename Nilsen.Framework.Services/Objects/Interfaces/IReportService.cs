using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace Nilsen.Framework.Services.Objects.Interfaces
{
    public interface IReportService
    {
        void CreateExcelFile(FileInfo fi);
        void BuildWorksheet(Worksheet ws, FileInfo fi);
    }
}
