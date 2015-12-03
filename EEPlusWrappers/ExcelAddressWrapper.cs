using OfficeOpenXml;

namespace EEPlusWrappers
{
    public interface IExcelAddress
    {
        ExcelAddressBase Address { get; }
        IExcelCellAddress Start { get; }
        IExcelCellAddress End { get; }
    }

    public class ExcelAddressWrapper : IExcelAddress
    {
        public ExcelAddressBase Address { get; private set; }

        public IExcelCellAddress Start
        {
            get
            {
                return new ExcelCellAddressWrapper(Address.Start);
            }
        }

        public IExcelCellAddress End
        {
            get
            {
                return new ExcelCellAddressWrapper(Address.End);
            }
        }

        public ExcelAddressWrapper(ExcelAddressBase address)
        {
            Address = address;
        }
    }
}
