using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;

namespace PriceListUpdaterAddon.SBOL
{
    public static class SBOC
    {
        private static SAPbouiCOM.Application oApp;
        private static SAPbobsCOM.Company oCom;

        public static SAPbobsCOM.Company Company => SBOC.oCom;

        public static SAPbouiCOM.Application App => SBOC.oApp;

        public static void Initialize()
        {
            SBOC.oApp = SAPbouiCOM.Framework.Application.SBO_Application;
            SBOC.oCom = (SAPbobsCOM.Company)SBOC.oApp.Company.GetDICompany();
        }

        public static Documents CreateDocument(BoObjectTypes type)
        {
            return (Documents)SBOC.oCom.GetBusinessObject(type);
        }

        public static StockTransfer CreateTransfer(BoObjectTypes type)
        {
            return (StockTransfer)SBOC.oCom.GetBusinessObject(type);
        }

        public static Recordset CreateRecordset()
        {
            return (Recordset)SBOC.oCom.GetBusinessObject(BoObjectTypes.BoRecordset);
        }

        public static int AddDocument<T>(T oDocument) where T : IDocuments
        {
            int errCode = oDocument.Add();
            if (errCode != 0)
            {
                string errMsg;
                SBOC.oCom.GetLastError(out errCode, out errMsg);
                SBOC.SetSystemMessage(string.Format("Ошибка при добавлении документа {0} {1}", (object)errCode, (object)errMsg), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
            }
            else
            {
                oDocument.GetByKey(int.Parse(SBOC.oCom.GetNewObjectKey()));
                SBOC.SetSystemMessage(string.Format("Документ [ {0} ] успешно создан.", (object)oDocument.DocNum), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
            }
            return errCode;
        }

        public static int AddStockTransfer<T>(T oTransfer) where T : IStockTransfer
        {
            int errCode = oTransfer.Add();
            if (errCode != 0)
            {
                string errMsg;
                SBOC.oCom.GetLastError(out errCode, out errMsg);
                SBOC.SetSystemMessage(string.Format("Ошибка при добавлении перевода {0} {1}", (object)errCode, (object)errMsg), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
            }
            else
            {
                oTransfer.GetByKey(int.Parse(SBOC.oCom.GetNewObjectKey()));
                SBOC.SetSystemMessage(string.Format("Перевод [ {0} ] успешно создан.", (object)oTransfer.DocNum), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
            }
            return errCode;
        }

        public static void SetSystemMessage(
          string msg,
          BoMessageTime length,
          BoStatusBarMessageType type)
        {
            SBOC.oApp.StatusBar.SetText(msg, length, type);
        }

        public static void SetFilters(List<string> forms, List<BoEventTypes> events)
        {
            EventFilters oFilters = (EventFilters)null;
            EventFilter oFilter = (EventFilter)null;
            oFilters = (EventFilters)new EventFiltersClass();
            events.ForEach((Action<BoEventTypes>)(e => oFilter = oFilters.Add(e)));
            forms.ForEach((Action<string>)(f => oFilter.AddEx(f)));
            SAPbouiCOM.Framework.Application.SBO_Application.SetFilter(oFilters);
        }
    }
}
