namespace OfficeOpenXml
{
    public class FlagCellStore : CellStore<byte>
    {
        internal void SetFlagValue(int row, int col, bool value, CellFlags cellFlags)
        {
            var currentValue = (CellFlags)GetValue(row, col);
            if (value)
            {
                SetValue(row, col, (byte)(currentValue | cellFlags));
            }
            else
            {
                SetValue(row, col, (byte)(currentValue & ~cellFlags));
            }
        }
        internal bool GetFlagValue(int row, int col, CellFlags cellFlags)
        {
            return ((byte)cellFlags & GetValue(row, col)) != 0;
        }
    }
}
