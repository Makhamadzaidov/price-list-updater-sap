using SAPbobsCOM;
using SAPbouiCOM;
using System.IO;
using System.Xml.Serialization;

namespace PriceListUpdaterAddon.PriceList
{
    internal class PricesManager
    {
        public void FillComboBox(Recordset oRS, ComboBox comboBox)
        {
            oRS.DoQuery("SELECT \"ListNum\", \"ListName\" FROM OPLN");
            PriceLists priceLists;
            using (StringReader stringReader = new StringReader(oRS.GetAsXML()))
                priceLists = (PriceLists)new XmlSerializer(typeof(PriceLists)).Deserialize((TextReader)stringReader);
            foreach (BOMBORow bomboRow in priceLists.BO.OPLN)
                comboBox.ValidValues.Add(bomboRow.ListName, bomboRow.ListName);
        }

        public void CalculateNewPrices(DataTable baseDataTable, DataTable ratesDataTable)
        {
            for (int rowIndex = 0; rowIndex < baseDataTable.Rows.Count; ++rowIndex)
            {
                string s1 = baseDataTable.GetValue((object)"PriceAtWH", rowIndex).ToString();
                string s2 = ratesDataTable.GetValue((object)"Rate", rowIndex).ToString();
                double three = (double.Parse(s1).NormalizeToThree() * (100.0 + double.Parse(s2)) / 100.0).NormalizeToThree();
                ratesDataTable.SetValue((object)"NewPrice", rowIndex, (object)three);
            }
        }

        public void CalculateDifference(DataTable baseDataTable, DataTable ratesDataTable)
        {
            for (int rowIndex = 0; rowIndex < baseDataTable.Rows.Count; ++rowIndex)
            {
                string s1 = baseDataTable.GetValue((object)"Price", rowIndex).ToString();
                string s2 = ratesDataTable.GetValue((object)"NewPrice", rowIndex).ToString();
                double three1 = double.Parse(s1).NormalizeToThree();
                double three2 = (double.Parse(s2) - three1).NormalizeToThree();
                ratesDataTable.SetValue((object)"Difference", rowIndex, (object)three2);
            }
        }
    }
}
