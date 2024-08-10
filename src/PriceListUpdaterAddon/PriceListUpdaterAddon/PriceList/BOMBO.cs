using System;
using System.ComponentModel;
using System.Xml.Serialization;

namespace PriceListUpdaterAddon.PriceList
{
    [DesignerCategory("code")]
    [XmlType(AnonymousType = true)]
    [Serializable]
    public class BOMBO
    {
        private BOMBOAdmInfo admInfoField;
        private BOMBORow[] oPLNField;

        public BOMBOAdmInfo AdmInfo
        {
            get => this.admInfoField;
            set => this.admInfoField = value;
        }

        [XmlArrayItem("row", IsNullable = false)]
        public BOMBORow[] OPLN
        {
            get => this.oPLNField;
            set => this.oPLNField = value;
        }
    }
}
