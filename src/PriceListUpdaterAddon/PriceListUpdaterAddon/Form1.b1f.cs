using System;
using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using PriceListUpdaterAddon.SBOL;
using PriceListUpdaterAddon.Matrix;
using PriceListUpdaterAddon.PriceList;
using PriceListUpdaterAddon.IO;
using System.Globalization;

namespace PriceListUpdaterAddon
{
    [FormAttribute("PriceListUpdaterAddon.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        private StaticText StaticText0;
        private StaticText StaticText1;
        private ComboBox PriceListsBox;
        private EditText GlobalRateEditText;
        private Button SetGlobalRateButton;
        private Button UpdatePriceListButton;
        private Button CalculateNewPricesButton;
        
        private Form oForm;
        private SAPbouiCOM.Matrix ItemsMatrix;
        private DataTable baseDT;
        private DataTable ratesDT;
        private UserDataSource globalRateDataSource;

        private readonly PricesManager pricesManager;
        private readonly MatrixHelper matrixHelper;
        private readonly IIO ioManager;

        private readonly string docEntry;
        private string currentPriceList;

        public Form1(string docEntry, Recordset oRS, IIO iio, PricesManager pm)
        {
            this.docEntry = docEntry;
            this.pricesManager = pm;
            this.matrixHelper = new MatrixHelper();
            this.ioManager = iio;
            this.currentPriceList = "";
            this.pricesManager.FillComboBox(oRS, this.PriceListsBox);
        }

        public override void OnInitializeComponent()
        {
            this.GetItem("glblrate").AffectsFormMode = false;
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.PriceListsBox = ((SAPbouiCOM.ComboBox)(this.GetItem("pricelst").Specific));
            this.PriceListsBox.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.PriceListsBox_ComboSelectAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.GlobalRateEditText = ((SAPbouiCOM.EditText)(this.GetItem("glblrate").Specific));
            this.SetGlobalRateButton = ((SAPbouiCOM.Button)(this.GetItem("Item_4").Specific));
            this.SetGlobalRateButton.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.SetGlobalRateButton_ClickBefore);
            this.ItemsMatrix = ((SAPbouiCOM.Matrix)(this.GetItem("itemsmtrx").Specific));
            this.UpdatePriceListButton = ((SAPbouiCOM.Button)(this.GetItem("Item_8").Specific));
            this.UpdatePriceListButton.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.UpdatePriceListButton_ClickBefore);
            this.CalculateNewPricesButton = ((SAPbouiCOM.Button)(this.GetItem("Item_9").Specific));
            this.CalculateNewPricesButton.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.CalculateNewPricesButton_ClickBefore);
            this.OnCustomInitialize();

        }
        public override void OnInitializeFormEvents()
        {
            this.VisibleAfter += Form1_VisibleAfter;
        }
        private void Form1_VisibleAfter(SBOItemEventArg pVal)
        {
            if (!(SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.TypeEx == "PriceListUpdaterAddon.Form1"))
                return;
            this.oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            this.baseDT = this.oForm.DataSources.DataTables.Item((object)"BaseDT");
            this.ratesDT = this.oForm.DataSources.DataTables.Item((object)"RatesDT");
            this.globalRateDataSource = this.oForm.DataSources.UserDataSources.Item((object)"GRateDS");
            this.baseDT.ExecuteQuery("SELECT \"LineNum\" + 1, \"ItemCode\", \"Dscription\", \"Quantity\", \"PriceFOB\", \"Cost\", " +
                "\"PriceAtWH\" FROM IPF1 WHERE \"DocEntry\" = " + this.docEntry + " ORDER BY \"LineNum\"");
            this.ratesDT.Rows.Add(this.baseDT.Rows.Count);
            this.ItemsMatrix.LoadFromDataSource();
        }
        private void OnCustomInitialize()
        {
            this.GetItem("itemsmtrx").AffectsFormMode = false;
            this.GetItem("pricelst").AffectsFormMode = false;
            this.GetItem("glblrate").AffectsFormMode = false;
        }
        private void PriceListsBox_ComboSelectAfter(object sboObject, SBOItemEventArg pVal)
        {
            this.currentPriceList = ((IComboBox)sboObject).Value;
            this.baseDT.ExecuteQuery("SELECT T0.\"LineNum\" + 1, T0.\"ItemCode\", T0.\"Dscription\", T0.\"Quantity\", T0.\"PriceFOB\", T0.\"Cost\", T0.\"PriceAtWH\", T2.\"Price\" FROM IPF1 T0 INNER JOIN OITM " +
                "T1 ON T0.\"ItemCode\" = T1.\"ItemCode\" INNER JOIN ITM1 T2 ON T0.\"ItemCode\" = T2.\"ItemCode\" " +
                "INNER JOIN OPLN T3 ON T2.\"PriceList\" = T3.\"ListNum\" WHERE T0.\"DocEntry\" = " + this.docEntry + " AND T3.\"ListName\" = '" + this.currentPriceList + "' ORDER BY T0.\"LineNum\"");
            this.ItemsMatrix.LoadFromDataSource();
        }
        private void SetGlobalRateButton_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                this.matrixHelper.FixDataSourceValue(this.GlobalRateEditText, this.globalRateDataSource);
                this.ItemsMatrix.FlushToDataSource();
                for (int rowIndex = 0; rowIndex < this.ratesDT.Rows.Count; ++rowIndex)
                {
                    this.ratesDT.SetValue((object)"Rate", rowIndex, (object)double.Parse(this.globalRateDataSource.Value));
                    this.ratesDT.SetValue((object)"NewPrice", rowIndex, (object)0);
                }
                this.ItemsMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Ошибка при выставлении глобального процента: [" + ex.Message + "]", BoMessageTime.bmt_Short);
            }

        }
        private void CalculateNewPricesButton_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this.CalculateNewPrices();
        }
        private bool CalculateNewPrices()
        {
            try
            {
                this.ItemsMatrix.FlushToDataSource();
                CellPosition cellFocus = this.ItemsMatrix.GetCellFocus();
                if (cellFocus != null && cellFocus.ColumnIndex == 8)
                    this.matrixHelper.FixDataTableValue((EditText)this.ItemsMatrix.Columns.Item("Rate").Cells.Item((object)cellFocus.rowIndex).Specific, this.ratesDT, cellFocus.rowIndex - 1);
                this.pricesManager.CalculateNewPrices(this.baseDT, this.ratesDT);
                this.ItemsMatrix.LoadFromDataSource();
                return true;
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Ошибка при подсчете цены: [" + ex.Message + "]", BoMessageTime.bmt_Short);
                return false;
            }
        }
        private void UpdatePriceListButton_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (this.currentPriceList == "")
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Выберите прайс-лист для редактирования.", BoMessageTime.bmt_Short);
            }
            else
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Началось обновление прайс-листа [" + this.currentPriceList + "]...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                CellPosition cellFocus = this.ItemsMatrix.GetCellFocus();
                this.ItemsMatrix.FlushToDataSource();
                if (cellFocus != null && cellFocus.ColumnIndex == 9)
                {
                    double result;
                    if (!double.TryParse(((IEditText)this.ItemsMatrix.Columns.Item((object)"Newprice").Cells.Item((object)cellFocus.rowIndex).Specific).Value.Replace("'", ""), NumberStyles.Number, (IFormatProvider)CultureInfo.InvariantCulture, out result))
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Неверный формат для процента.", BoMessageTime.bmt_Short);
                        return;
                    }
                    this.ratesDT.SetValue((object)"NewPrice", cellFocus.rowIndex - 1, (object)result);
                }

                SAPbobsCOM.Company oCom = (SAPbobsCOM.Company)SAPbouiCOM.Framework.Application.SBO_Application.Company.GetDICompany();
                SAPbobsCOM.Items businessObject = (SAPbobsCOM.Items)oCom.GetBusinessObject(BoObjectTypes.oItems);

                for (int rowIndex = 0; rowIndex < this.baseDT.Rows.Count; ++rowIndex)
                {
                    this.ItemsMatrix.SetCellFocus(rowIndex + 1, 9);
                    businessObject.GetByKey(this.baseDT.GetValue((object)"ItemCode", rowIndex).ToString());
                    for (int LineNum = 0; LineNum < businessObject.PriceList.Count; ++LineNum)
                    {
                        businessObject.PriceList.SetCurrentLine(LineNum);
                        if (businessObject.PriceList.PriceListName == this.currentPriceList)
                        {
                            businessObject.PriceList.Price = this.ratesDT.GetValue((object)"NewPrice", rowIndex).ToDouble();
                            break;
                        }
                    }
                    int errCode = businessObject.Update();
                    if (errCode != 0)
                    {
                        string errMsg;
                        SBOC.Company.GetLastError(out errCode, out errMsg);
                        SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(string.Format("Ошибка при обновлении цены для товара [{0}] [{1} - {2}].", (object)businessObject.ItemCode, (object)errCode, (object)errMsg), BoMessageTime.bmt_Short);
                        return;
                    }
                }
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Прайс-лист [" + this.currentPriceList + "] успешно обновлен.", Type: BoStatusBarMessageType.smt_Success);
            }
        }
    }
}