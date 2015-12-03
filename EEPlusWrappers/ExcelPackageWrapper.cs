using OfficeOpenXml;
using System.IO;

namespace EEPlusWrappers
{
    public interface IExcelPackage
    {
        ExcelPackage Package { get; }
        IExcelWorkbook Workbook { get; }
    }

    public class ExcelPackageWrapper : IExcelPackage
    {
        public ExcelPackage Package { get; private set; }

        public IExcelWorkbook Workbook
        {
            get
            {
                return new ExcelWorkbookWrapper(Package.Workbook);
            }
        }

        public ExcelPackageWrapper(ExcelPackage package)
        {
            Package = package;
        }
    }

    public interface IExcelPackageFactory
    {
        IExcelPackage GetPackage(Stream stream);
    }

    public class ExcelPackageFactory
    {
        public IExcelPackage GetPackage(Stream stream)
        {
            return new ExcelPackageWrapper(new ExcelPackage(stream));
        }
    }
}
