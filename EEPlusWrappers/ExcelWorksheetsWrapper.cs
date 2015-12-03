using OfficeOpenXml;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace EEPlusWrappers
{
    public interface IExcelWorksheets : IEnumerable<IExcelWorksheets>
    {
        ExcelWorksheets Worksheets { get; }
        IExcelWorksheet this[string name] { get; }
    }

    public class ExcelWorksheetsWrapper : IExcelWorksheets
    {
        public ExcelWorksheets Worksheets { get; private set; }

        public ExcelWorksheetsWrapper(ExcelWorksheets worksheets)
        {
            Worksheets = worksheets;
        }

        public IExcelWorksheet this[string name]
        {
            get
            {
                return new ExcelWorksheetWrapper(Worksheets[name]);
            }
        }

        public IEnumerator<IExcelWorksheets> GetEnumerator()
        {
            return GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            var worksheets = Worksheets.Select(x => new ExcelWorksheetWrapper(x)).ToList();

            return worksheets.GetEnumerator();
        }
    }
}
