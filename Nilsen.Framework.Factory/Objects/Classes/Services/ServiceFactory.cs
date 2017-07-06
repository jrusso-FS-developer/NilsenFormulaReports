using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading.Tasks;
using Nilsen.Framework.Services;
using Nilsen.Framework.Services.Objects.Enums;
using Nilsen.Framework.Services.Objects.Classes;
using Nilsen.Framework.Services.Objects.Interfaces;

namespace Nilsen.Framework.Factory.Objects.Classes.Services
{
    public static class ServiceFactory
    {
        public static IReportService GetReportService(TextBox consoleWindow, Button btnProcess, Int32 reportType)
        {
            IReportService service = null;
            
            switch (reportType)
            {
                case (Int32)ReportTypes.TurfFormula:
                    service = new TurfFormulaReportService(consoleWindow, btnProcess);
                    break;
                case (Int32)ReportTypes.PaceForecasterFormula:
                    service = new PaceForecasterFormulaReportService(consoleWindow, btnProcess);
                    break;
            }

            return service;
        }
    }
}
