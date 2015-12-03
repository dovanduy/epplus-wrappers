using OfficeOpenXml;

namespace EEPlusWrappers
{
    public interface IExcelWorksheet
    {
        ExcelWorksheet Worksheet { get; }
        IExcelRange Cells { get; }
        IExcelAddress Dimension { get; }
    }

    public class ExcelWorksheetWrapper : IExcelWorksheet
    {
        public ExcelWorksheet Worksheet { get; private set; }

        public IExcelRange Cells { get; private set; }

        public IExcelAddress Dimension { get; private set; }

        public ExcelWorksheetWrapper(ExcelWorksheet worksheet)
        {
            Worksheet = worksheet;

            Cells = new ExcelRangeWrapper(Worksheet.Cells);

            Dimension = new ExcelAddressWrapper(Worksheet.Dimension);
        }

        public ExcelWorksheetWrapper(IExcelWorksheet worksheet)
            : this(worksheet.Worksheet)
        {

        }
    }
}
