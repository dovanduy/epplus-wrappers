using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;

namespace EEPlusWrappers
{
    public interface IExcelRange : IEnumerator, IEnumerator<IExcelRange>, IEnumerable, IEnumerable<IExcelRange>
    {
        ExcelRange Range { get; }
        IExcelRange this[int Row, int Col] { get; }
        string Text { get; }
        object Value { get; }
    }

    public class ExcelRangeWrapper : IExcelRange
    {
        public ExcelRange Range { get; private set; }

        private IDictionary<Tuple<int, int>, IExcelRange> _cells;

        public ExcelRangeWrapper(ExcelRange range)
        {
            Range = range;
        }

        [Obsolete]
        /// <summary>
        /// Unit testing constructor
        /// </summary>
        public ExcelRangeWrapper()
        {
            _cells = new Dictionary<Tuple<int, int>, IExcelRange>();
        }

        public IExcelRange this[int Row, int Col]
        {
            get
            {
                if (Range != null)
                {
                    var range = Range[Row, Col];

                    return new ExcelRangeWrapper(range);
                }

                return _cells[new Tuple<int, int>(Row, Col)];
            }
        }

        public void AddRange(IExcelRange range, int row, int column)
        {
            _cells[new Tuple<int, int>(row, column)] = range;
        }

        public string Text
        {
            get
            {
                return Range.Text;
            }
        }

        public object Value
        {
            get
            {
                return Range.Value;
            }
        }

        object IEnumerator.Current
        {
            get
            {
                return this.Current;
            }
        }

        public IExcelRange Current
        {
            get
            {
                return new ExcelRangeWrapper(Range);
            }
        }

        public bool MoveNext()
        {
            return Range.MoveNext();
        }

        public void Reset()
        {
            Range.Reset();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public IEnumerator<IExcelRange> GetEnumerator()
        {
            Reset();

            return this;
        }

        public void Dispose()
        {

        }
    }
}
