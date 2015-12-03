using OfficeOpenXml;

namespace EEPlusWrappers
{
    public interface IExcelWorkbook
    {
        ExcelWorkbook Workbook { get; }
        IExcelWorksheets Worksheets { get; }
    }

    public class ExcelWorkbookWrapper : IExcelWorkbook
    {
        public ExcelWorkbook Workbook { get; private set; }

        public IExcelWorksheets Worksheets
        {
            get
            {
                return new ExcelWorksheetsWrapper(Workbook.Worksheets);
            }
        }

        public ExcelWorkbookWrapper(ExcelWorkbook workbook)
        {
            Workbook = workbook;
        }
    }
}
