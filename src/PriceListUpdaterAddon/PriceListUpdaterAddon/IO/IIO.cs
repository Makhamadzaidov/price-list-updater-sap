using SAPbouiCOM;

namespace PriceListUpdaterAddon.IO
{
    internal interface IIO
    {
        bool SaveMatrixData(DataTable ratesTable, ComboBox priceList, string docEntry);

        bool GetMatrixData(DataTable ratesTable, ComboBox priceListBox, string docEntry);
    }
}
