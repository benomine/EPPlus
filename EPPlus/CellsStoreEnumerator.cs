using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml
{
    public class CellsStoreEnumerator<T> : IEnumerable<T>, IEnumerator<T>
    {
        private readonly CellStore<T> _cellStore;
        private int _row, _colPos;
        private int[] _pagePos, _cellPos;
        private readonly int _startRow;
        private readonly int _startCol;
        private readonly int _endRow;
        private readonly int _endCol;
        private int _minRow, _minColPos, _maxRow, _maxColPos;

        public CellsStoreEnumerator(
            CellStore<T> cellStore,
            int startRow = 0,
            int startCol = 0,
            int endRow = ExcelPackage.MaxRows,
            int endCol = ExcelPackage.MaxColumns)
        {
            _cellStore = cellStore;
            _startRow=startRow;
            _startCol=startCol;
            _endRow=endRow;
            _endCol=endCol;

            Init();
        }

        internal void Init()
        {
            _minRow = _startRow;
            _maxRow = _endRow;

            _minColPos = _cellStore.GetPosition(_startCol);
            if (_minColPos < 0) _minColPos = ~_minColPos;
            _maxColPos = _cellStore.GetPosition(_endCol);
            if (_maxColPos < 0) _maxColPos = ~_maxColPos-1;
            _row = _minRow;
            _colPos = _minColPos - 1;

            var cols = _maxColPos - _minColPos + 1;
            _pagePos = new int[cols];
            _cellPos = new int[cols];
            for (var i = 0; i < cols; i++)
            {
                _pagePos[i] = -1;
                _cellPos[i] = -1;
            }
        }
        internal int Row => _row;

        internal int Column
        {
            get
            {
                if (_colPos == -1) MoveNext();
                return _colPos == -1 ? 0 : _cellStore.ColumnIndex[_colPos].Index;
            }
        }
        internal T Value
        {
            get
            {
                lock (_cellStore)
                {
                    return _cellStore.GetValue(_row, Column);
                }
            }
            set
            {
                lock (_cellStore)
                {
                    _cellStore.SetValue(_row, Column, value);
                }
            }
        }
        internal bool Next()
        {
            return _cellStore.GetNextCell(ref _row, ref _colPos, _minColPos, _maxRow, _maxColPos);
        }
        internal bool Previous()
        {
            lock (_cellStore)
            {
                return _cellStore.GetPrevCell(ref _row, ref _colPos, _minRow, _minColPos, _maxColPos);
            }
        }

        public string CellAddress => ExcelCellBase.GetAddress(Row, Column);

        public IEnumerator<T> GetEnumerator()
        {
            Reset();
            return this;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            Reset();
            return this;
        }

        public T Current => Value;

        public void Dispose()
        {

        }

        object IEnumerator.Current
        {
            get
            {
                Reset();
                return this;
            }
        }

        public bool MoveNext()
        {
            return Next();
        }

        public void Reset()
        {
            Init();
        }
    }
}