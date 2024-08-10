using PriceListUpdaterAddon.IO;
using PriceListUpdaterAddon.PriceList;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;

namespace PriceListUpdaterAddon
{
    class Program
    {
        private static SAPbobsCOM.Recordset oRS;
        private static SAPbobsCOM.Company oCom;
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    oApp = new Application(args[0]);
                }

                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;

                oCom = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                oRS = (SAPbobsCOM.Recordset)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.FormTypeEx == "992" & pVal.BeforeAction & pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            {
                SAPbouiCOM.Form form = Application.SBO_Application.Forms.Item((object)pVal.FormUID);
                SAPbouiCOM.Item tem1 = form.Items.Item((object)"2");
                SAPbouiCOM.Item tem2 = form.Items.Add("price", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                tem2.Left = tem1.Left + tem1.Width + 10;
                tem2.Top = tem1.Top;
                tem2.Width = 100;
                tem2.Height = tem1.Height;
                ((SAPbouiCOM.IButton)tem2.Specific).Caption = "Прайс-лист";
            }

            if (pVal.FormTypeEx == "992" && !pVal.BeforeAction && pVal.ItemUID == "price" 
                && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            {
                SAPbouiCOM.DBDataSource dBDataSource = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(pVal.FormUID).DataSources.DBDataSources.Item((object)"OIPF");
                if (dBDataSource.GetValue((object)"DocEntry", 0) == "")
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Документ не найден");
                }
                else
                {
                    LocalIO localIo = new LocalIO();
                    PricesManager pm = new PricesManager();
                    new Form1(dBDataSource.GetValue((object)"DocEntry", 0), Program.oRS, (IIO)localIo, pm).Show();
                }
            }
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }
    }
}
