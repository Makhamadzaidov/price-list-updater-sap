using SAPbouiCOM;
using System;
using System.Drawing;
using System.Globalization;

namespace PriceListUpdaterAddon.Matrix
{
    internal class MatrixHelper
    {
        public void FixDataTableValue(EditText editText, DataTable ratesDataTable, int row)
        {
            double result;
            if (!double.TryParse(editText.Value.Replace("'", ""), NumberStyles.Number, (IFormatProvider)CultureInfo.InvariantCulture, out result))
                throw new Exception("Неверный формат для процента.");
            ratesDataTable.SetValue((object)"Rate", row, (object)result);
        }

        public void FixRateCell(SAPbouiCOM.Matrix itemsMatrix, DataTable ratesDataTable)
        {
            CellPosition cellFocus = itemsMatrix.GetCellFocus();
            if (cellFocus == null || cellFocus.ColumnIndex != 8)
                return;
            double result;
            if (!double.TryParse(((IEditText)itemsMatrix.Columns.Item((object)"rate").Cells.Item((object)cellFocus.rowIndex).Specific).Value.Replace("'", ""), NumberStyles.Number, (IFormatProvider)CultureInfo.InvariantCulture, out result))
                throw new Exception("Неверный формат для процента.");
            ratesDataTable.SetValue((object)"Rate", cellFocus.rowIndex - 1, (object)result);
        }

        public void FixDataSourceValue(EditText editText, UserDataSource dataSource)
        {
            double result;
            if (!double.TryParse(editText.Value, NumberStyles.Number, (IFormatProvider)CultureInfo.InvariantCulture, out result))
                throw new Exception("Неверный формат процента.");
            dataSource.Value = result.ToString();
        }

        public void MarkDifference(SAPbouiCOM.Matrix itemsMatrix, DataTable ratesDataTable)
        {
            for (int rowIndex = 0; rowIndex < ratesDataTable.Rows.Count; ++rowIndex)
            {
                if (double.Parse(ratesDataTable.GetValue((object)"Difference", rowIndex).ToString()) >= 0.0)
                    itemsMatrix.CommonSetting.SetCellFontColor(rowIndex + 1, 10, (int)Color.Black.R);
                else
                    itemsMatrix.CommonSetting.SetCellFontColor(rowIndex + 1, 10, (int)Color.Red.R);
            }
        }

        public void ResetDifferenceMarking(SAPbouiCOM.Matrix itemsMatrix)
        {
            for (int index = 0; index < itemsMatrix.RowCount; ++index)
                itemsMatrix.CommonSetting.SetCellFontColor(index + 1, 10, (int)Color.Black.R);
        }
    }
}
