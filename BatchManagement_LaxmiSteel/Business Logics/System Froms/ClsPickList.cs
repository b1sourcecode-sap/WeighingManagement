using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BatchManagement_LaxmiSteel
{
    class ClsPickList : Base
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
        SAPbobsCOM.Recordset _RS = null, _RS1 = null, _RS2 = null, _RSBatchManagement_LaxmiSteel = null, _Rs = null, _RSChk = null, _RsCal = null, _RSChkGetVal = null, _RSChk1 = null, _RSPriceList = null;
        string _LastErrorMsg = " ", strcon = "";
        public string _DBName = string.Empty;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsPickList()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsPickList()
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
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("U_TruckNo").Specific).Value == "")
                            {
                                Utilities.Message("Truck No Can't left blank", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                            else if (((SAPbouiCOM.EditText)_Form.Items.Item("U_TruckNo").Specific).Value != "")
                            {
                                string _SqlQty = "select sum(\"RelQtty\") as \"RelQtty\" from PKL1 where \"AbsEntry\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value + "'";
                                Utilities.ExecuteSQL(ref _RS1, _SqlQty);


                                string _Sql = "update  \"@VHWTMGTH\" set \"U_LodWt\" = '" + _RS1.Fields.Item("RelQtty").Value + "' , \"U_PickLstNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value + "' where \"U_Status\" = 'Open' AND \"U_VehNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("U_TruckNo").Specific).Value.ToString() + "'";
                                Utilities.ExecuteSQL(ref _RS, _Sql);

                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.ItemUID == "U_TruckNo" && pVal.CharPressed == 9)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("U_TruckNo").Specific).Value == "")
                            {
                                string StrQuery = "select \"U_VehNo\",\"U_TranspName\",\"U_LdContNo\",\"U_DriverNo\",\"U_DriverNm\",\"U_DocNum\" from \"@VHWTMGTH\" ";
                                Utilities.ExecuteSQL(ref _RS1, StrQuery);

                                this._Object = new ClsChooseFromList();
                                ((ClsChooseFromList)this._Object).FormTitle = "Vehicle List";
                                ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref BubbleEvent, StrQuery, strcon, false, _DBName);

                            }
                        }
                        break;
                }
            }
            #endregion

            #region After ACTION ACTION
            else if (pVal.Before_Action == false)
            {

                switch (pVal.EventType)
                {

                    #region ITEM PRESS EVENT

                    //case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                    //    if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    //    {
                    //        if (((SAPbouiCOM.EditText)_Form.Items.Item("U_TruckNo").Specific).Value != "")
                    //        {
                    //            string _Sql = "update  \"@VHWTMGTH\" set \"U_PickLstNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value + "' where \"U_Status\" = 'Open' AND \"U_VehNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("U_TruckNo").Specific).Value.ToString() + "'"; 
                    //            Utilities.ExecuteSQL(ref _RS, _Sql);
                               
                    //        }
                    //    }

                    //    break;

                    #endregion

                }
            }

            #endregion

        }

        #endregion

        #region CHOOSE FROM LIST EVENT

        public override void ChooseFromList_Event()
        {
            try
            {
                _Form.Freeze(true);
                if (this._CFLEvent.ItemUID == "U_TruckNo")
                {
                    ((SAPbouiCOM.EditText)_Form.Items.Item("U_TruckNo").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][0].ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("U_TranspNm").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][1].ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("U_CounterNo").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][2].ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("U_ContactNo").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][3].ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("U_DriverNm").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][4].ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("U_VWMNo").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][5].ToString();
                }

            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.ToString());
            }
            finally
            {
                _Form.Freeze(false);
            }
        }

        #endregion

    }
}
