using System;
using System.ComponentModel;
using System.Xml.Serialization;

namespace PriceListUpdaterAddon.PriceList
{
    [DesignerCategory("code")]
    [XmlType(AnonymousType = true)]
    [XmlRoot(Namespace = "", IsNullable = false, ElementName = "BOM")]
    [Serializable]
    public class PriceLists
    {
        private BOMBO boField;

        public BOMBO BO
        {
            get => this.boField;
            set => this.boField = value;
        }
    }
}
