namespace OfficeOpenXml
{
    public class FlagCellStore : CellStore<byte>
    {
        internal void SetFlagValue(int Row, int Col, bool value, CellFlags cellFlags)
        {
            CellFlags currentValue = (CellFlags)GetValue(Row, Col);
            if (value)
            {
                SetValue(Row, Col, (byte)(currentValue | cellFlags));
            }
            else
            {
                SetValue(Row, Col, (byte)(currentValue & ~cellFlags));
            }
        }
        internal bool GetFlagValue(int Row, int Col, CellFlags cellFlags)
        {
            return !(((byte)cellFlags & GetValue(Row, Col)) == 0);
        }
    }
}
