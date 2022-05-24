using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml
{
    public class CellsStoreEnumerator<T> : IEnumerable<T>, IEnumerator<T>
    {
        CellStore<T> _cellStore;
        int row, colPos;
        int[] pagePos, cellPos;
        int _startRow, _startCol, _endRow, _endCol;
        int minRow, minColPos, maxRow, maxColPos;
        public CellsStoreEnumerator(CellStore<T> cellStore) :
            this(cellStore, 0, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns)
        {
        }
        public CellsStoreEnumerator(CellStore<T> cellStore, int StartRow, int StartCol, int EndRow, int EndCol)
        {
            _cellStore = cellStore;

            _startRow=StartRow;
            _startCol=StartCol;
            _endRow=EndRow;
            _endCol=EndCol;

            Init();

        }

        internal void Init()
        {
            minRow = _startRow;
            maxRow = _endRow;

            minColPos = _cellStore.GetPosition(_startCol);
            if (minColPos < 0) minColPos = ~minColPos;
            maxColPos = _cellStore.GetPosition(_endCol);
            if (maxColPos < 0) maxColPos = ~maxColPos-1;
            row = minRow;
            colPos = minColPos - 1;

            var cols = maxColPos - minColPos + 1;
            pagePos = new int[cols];
            cellPos = new int[cols];
            for (int i = 0; i < cols; i++)
            {
                pagePos[i] = -1;
                cellPos[i] = -1;
            }
        }
        internal int Row
        {
            get => row;
        }
        internal int Column
        {
            get
            {
                if (colPos == -1) MoveNext();
                if (colPos == -1) return 0;
                return _cellStore._columnIndex[colPos].Index;
            }
        }
        internal T Value
        {
            get
            {
                lock (_cellStore)
                {
                    return _cellStore.GetValue(row, Column);
                }
            }
            set
            {
                lock (_cellStore)
                {
                    _cellStore.SetValue(row, Column, value);
                }
            }
        }
        internal bool Next()
        {
            return _cellStore.GetNextCell(ref row, ref colPos, minColPos, maxRow, maxColPos);
        }
        internal bool Previous()
        {
            lock (_cellStore)
            {
                return _cellStore.GetPrevCell(ref row, ref colPos, minRow, minColPos, maxColPos);
            }
        }

        public string CellAddress
        {
            get
            {
                return ExcelAddressBase.GetAddress(Row, Column);
            }
        }

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

        public T Current
        {
            get
            {
                return Value;
            }
        }

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