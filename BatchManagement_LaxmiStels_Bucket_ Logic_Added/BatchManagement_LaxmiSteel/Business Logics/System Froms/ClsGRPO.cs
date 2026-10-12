using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BatchManagement_LaxmiSteel
{
    class ClsGRPO : Base
    {
       
        #region VARIABLE DECLARATION

        SAPbouiCOM.Item _ExistItem, _NewItem, _Items;
        SAPbouiCOM.Button _Button;
        SAPbouiCOM.ComboBox _ComboBox;
        SAPbouiCOM.StaticText _StaticText;
        private bool flag;
        public SAPbouiCOM.EditText _EditText;
        public SAPbouiCOM.Matrix _Matrix;
        string _DocEntry = "", _DocCurency = "", _DocNum = "", _BPCode = "", _BPName = "", _Branch = "", _PostingDate = "", _DueDate = "", _ConSignmentType = "", _CountryCode = "", _CountryName = "", _DocumentType = "",_DocType = "";
        double _DocTotal = 0;
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        SAPbobsCOM.Recordset _RS = null, _RS1 = null, _RSBatchManagement_LaxmiSteel = null, _Rs = null, _RSChk = null, _RsCal = null, _RSChkGetVal = null, _RSChk1 = null, _RSPriceList = null;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsGRPO()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsGRPO()
        {
        }
        #endregion

        #region ITEM EVENT

        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {

            #region BEFORE ACTION ACTION
            if (pVal.Before_Action == true)
            {
                switch (pVal.EventType)
                {
                    #region ITEM PRESS EVENT
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "BtnLandC")
                        {
                            SAPbobsCOM.Documents _oInventoryGenEntry = (SAPbobsCOM.Documents)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                            if (_oInventoryGenEntry.GetByKey(21944))
                            {
                                 //_oInventoryGenEntry.Lines.SetCurrentLine(0);

                                // _oInventoryGenEntry.Lines.BinAllocations.SetCurrentLine(0);
                                // _oInventoryGenEntry.Lines.BinAllocations.BinAbsEntry = 21944;
                                // _oInventoryGenEntry.Lines.BinAllocations.Quantity = 1;
                                //_oInventoryGenEntry.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = 0;
                                //_oInventoryGenEntry.Lines.BinAllocations.Add();
                                 _oInventoryGenEntry.Lines.Quantity = 1;
                               // _oInventoryGenEntry.Lines.BatchNumbers.SetCurrentLine(0);
                                _oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = "PPP";
                                _oInventoryGenEntry.Lines.BatchNumbers.Quantity = 1;
                                _oInventoryGenEntry.Lines.BatchNumbers.Add();
                                int i = _oInventoryGenEntry.Update();
                                string GetLastErrorDescription = Utilities.Application.Company.GetLastErrorDescription();
                            }
                        }
                        break;
                    #endregion
                }
            }
            #endregion

            #region After ACTION ACTION
            else if (pVal.Before_Action == false)
            {

                switch (pVal.EventType)
                {

                    #region ITEM PRESS EVENT
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "BtnLandC" && _Form.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                           
                        }

                        break;

                    #endregion

                }
            }

            #endregion

        }

        #endregion

        #region Form Default

        internal void FormDefault()
        {
            try
            {
                _ExistItem = _Form.Items.Item("2");
                _NewItem = _Form.Items.Add("BtnLandC", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                _NewItem.Left = _ExistItem.Left + _ExistItem.Width + 5;
                _NewItem.Width = _ExistItem.Width + 10;
                _NewItem.Top = _ExistItem.Top;
                _Button = (SAPbouiCOM.Button)_NewItem.Specific;
                _Button.Caption = "Landed Cost";

            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }

        #endregion

    }
}
