using System;
using System.ComponentModel;
using System.Xml.Serialization;

namespace PriceListUpdaterAddon.PriceList
{
    [DesignerCategory("code")]
    [XmlType(AnonymousType = true)]
    [Serializable]
    public class BOMBOAdmInfo
    {
        private sbyte objectField;

        public sbyte Object
        {
            get => this.objectField;
            set => this.objectField = value;
        }
    }
}
