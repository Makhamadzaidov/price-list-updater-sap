using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PriceListUpdaterAddon.IO
{
    internal class LocalIO : IIO
    {
        public bool GetMatrixData(DataTable ratesTable, ComboBox priceListBox, string docEntry)
        {
            int rowIndex = 0;
            string path = "Data\\" + docEntry;
            if (!File.Exists(path))
                return false;
            try
            {
                using (StreamReader streamReader = new StreamReader(path))
                {
                    priceListBox.Select((object)streamReader.ReadLine());
                    while (!streamReader.EndOfStream)
                    {
                        ratesTable.SetValue((object)"Rate", rowIndex, (object)double.Parse(streamReader.ReadLine()));
                        ++rowIndex;
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Ошибка при получении прошлых значений: [" + ex.Message + "]", BoMessageTime.bmt_Short);
                return false;
            }
        }

        public bool SaveMatrixData(DataTable ratesTable, ComboBox priceListBox, string docEntry)
        {
            try
            {
                using (StreamWriter streamWriter = new StreamWriter("Data\\" + docEntry))
                {
                    streamWriter.WriteLine(priceListBox.Value);
                    for (int rowIndex = 0; rowIndex < ratesTable.Rows.Count; ++rowIndex)
                        streamWriter.WriteLine(ratesTable.GetValue((object)"Rate", rowIndex));
                }
                return true;
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Ошибка при сохранении значений: [" + ex.Message + "]", BoMessageTime.bmt_Short);
                return false;
            }
        }
    }
}
