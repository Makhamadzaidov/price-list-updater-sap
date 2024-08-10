using System;
using System.ComponentModel;
using System.Xml.Serialization;

namespace PriceListUpdaterAddon.PriceList
{
    [DesignerCategory("code")]
    [XmlType(AnonymousType = true)]
    [Serializable]
    public class BOMBORow
    {
        private string listNumField;
        private string listNameField;

        public string ListNum
        {
            get => this.listNumField;
            set => this.listNumField = value;
        }

        public string ListName
        {
            get => this.listNameField;
            set => this.listNameField = value;
        }
    }
}
