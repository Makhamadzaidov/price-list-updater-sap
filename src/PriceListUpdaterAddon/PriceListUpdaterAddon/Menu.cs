using System;
using SAPbouiCOM;


namespace PriceListUpdaterAddon
{
    class Menu
    {
        public void AddMenuItems()
        {
            Menus menus = (Menus)null;
            menus = SAPbouiCOM.Framework.Application.SBO_Application.Menus;
            MenuCreationParams CreationPackage = (MenuCreationParams)SAPbouiCOM.Framework.Application.SBO_Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
            MenuItem menuItem = SAPbouiCOM.Framework.Application.SBO_Application.Menus.Item((object)"43520");
            CreationPackage.Type = BoMenuType.mt_POPUP;
            CreationPackage.UniqueID = "IncomeItemPriceChanger";
            CreationPackage.String = "IncomeItemPriceChanger";
            CreationPackage.Enabled = true;
            CreationPackage.Position = -1;
            Menus subMenus1 = menuItem.SubMenus;
            try
            {
                subMenus1.AddEx(CreationPackage);
            }
            catch (Exception ex)
            {

            }
            try
            {
                Menus subMenus2 = SAPbouiCOM.Framework.Application.SBO_Application.Menus.Item((object)"IncomeItemPriceChanger").SubMenus;
                CreationPackage.Type = BoMenuType.mt_STRING;
                CreationPackage.UniqueID = "IncomeItemPriceChanger.Form1";
                CreationPackage.String = "Form1";
                subMenus2.AddEx(CreationPackage);
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", BoMessageTime.bmt_Short);
            }
        }

        public void SBO_Application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (!pVal.BeforeAction)
                    return;
                int num = pVal.MenuUID == "IncomeItemPriceChanger.Form1" ? 1 : 0;
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(ex.ToString());
            }
        }

    }
}
