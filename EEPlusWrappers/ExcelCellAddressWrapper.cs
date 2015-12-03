using OfficeOpenXml;

namespace EEPlusWrappers
{
    public interface IExcelCellAddress
    {
        ExcelCellAddress CellAddress { get; }
        int Row { get; }
        int Column { get; }
    }

    public class ExcelCellAddressWrapper : IExcelCellAddress
    {
        public ExcelCellAddress CellAddress { get; private set; }

        public int Row
        {
            get
            {
                return CellAddress.Row;
            }
        }

        public int Column
        {
            get
            {
                return CellAddress.Column;
            }
        }

        public ExcelCellAddressWrapper(ExcelCellAddress cellAddress)
        {
            CellAddress = cellAddress;
        }
    }
}
