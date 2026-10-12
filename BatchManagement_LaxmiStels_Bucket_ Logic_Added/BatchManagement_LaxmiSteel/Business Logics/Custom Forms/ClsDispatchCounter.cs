using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;

namespace BatchManagement_LaxmiSteel 
{
    class ClsDispatchCounter : Base
    {

        #region Variable Declaration

        SAPbouiCOM.ComboBox _ComboBox;
        SAPbouiCOM.Matrix _Matrix;
        ChooseFromListEvent _SysCFLEvent;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource, _ChildDBDataSource;
        Conditions _Cons;
        Condition _Con;
        SAPbobsCOM.Recordset _RS = null, _Rs = null, _RSValiDate = null;
        public static int RowNo = 0;
        public static int NoOfPeacePer = 0;
        public static double VarianceAllowed = 0;
        public static string ItemGroupName = string.Empty;
        private bool Flags = true;
        private bool counter = true;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsDispatchCounter()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsDispatchCounter()
        {
        }

        #endregion

        #region Item Event

        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
        {
            #region Before Action
            if (pVal.Before_Action)
            {
                try
                {
                    switch (pVal.EventType)
                    {
                        #region et_ITEM_PRESSED

                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            if (pVal.ItemUID == "37")
                            {
                                _Form.PaneLevel = 1;
                            }
                            else if (pVal.ItemUID == "38")
                            {
                                _Form.PaneLevel = 2;
                            }

                            if (pVal.ItemUID == "1" && _Form.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                ((SAPbouiCOM.EditText)_Form.Items.Item("20").Specific).Value = Utilities.getMaxColumnValueNum("@WBLOADDTLH", "U_DocNum");
                                string _Sql = "select \"U_PickLstNo\" from \"@WBLOADDTLH\" where \"U_PickLstNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value + "'";
                                Utilities.ExecuteSQL(ref _RSValiDate, _Sql);
                                if (_RSValiDate.RecordCount > 0)
                                {
                                    Utilities.Message("PickList No. Already Exist", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    oBubbleEvent = false;
                                    return;
                                }

                                if (((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value == "")
                                {
                                    Utilities.Message("PickList can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    oBubbleEvent = false;
                                    return;
                                }
                                else if (((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value == "")
                                {
                                    Utilities.Message("Customer Code can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    oBubbleEvent = false;
                                    return;
                                }
                                if (Validation() == false)
                                {
                                    oBubbleEvent = false;
                                    return;
                                }
                            }
                            if (pVal.ItemUID == "1" && _Form.Mode == BoFormMode.fm_UPDATE_MODE)
                            {
                                //double TotalWeightInBundles = 0;
                                //for (int i = 1; i <= _Matrix.RowCount - 1; i++)
                                //{
                                //    TotalWeightInBundles = TotalWeightInBundles + Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_2").Cells.Item(i).Specific).Value);
                                //}
                                //string SQL = "update \"@VHWTMGTH\" set \"U_NTVarWt\" = \"U_NetWt\" - " + TotalWeightInBundles + ", \"U_WBDocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("20").Specific).Value + "',\"U_WBCounter\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("50").Specific).Value + "' where \"U_Status\" = 'Open' AND \"U_PickLstNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value + "' AND \"U_VehNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("8").Specific).Value + "' ";
                                //Utilities.ExecuteSQL(ref _Rs, SQL);
                            }
                          
                            break;

                        #endregion

                        #region et_CHOOSE_FROM_LIST

                        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                            if (pVal.ItemUID == "4")
                            {
                                if (((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value == "")
                                {
                                    this.FiltePickList("");
                                }
                            }
                            break;

                        #endregion

                        #region et_DOUBLE_CLICK
                        case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                            if (pVal.ColUID == "0")
                            {
                                try
                                {
                                    if (pVal.Row > 0)
                                    {
                                        if (_Form.Mode != BoFormMode.fm_ADD_MODE)
                                        {
                                            RowNo = pVal.Row;
                                            this._Object = new ClsWeighingBridgeDetails();
                                            this.OpenChildForm(Constants.Forms.WeighingBridgeDetails, ref pVal, true);
                                            string PickListNo = string.Empty, ItemCode = string.Empty, ItemDescription = string.Empty, TruckNo = string.Empty, LotNo = string.Empty, _Status = string.Empty;
                                            string QtyToDispatch = string.Empty;
                                            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("39").Specific;
                                            PickListNo = ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value.ToString();
                                            TruckNo = ((SAPbouiCOM.EditText)_Form.Items.Item("8").Specific).Value.ToString();
                                            ItemCode = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_10").Cells.Item(pVal.Row).Specific).Value.ToString().Trim();
                                            ItemDescription = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_9").Cells.Item(pVal.Row).Specific).Value.ToString().Trim();
                                            QtyToDispatch = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific).Value;
                                            ItemGroupName = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_12").Cells.Item(pVal.Row).Specific).Value.ToString().Trim();
                                            NoOfPeacePer = Convert.ToInt16(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_11").Cells.Item(pVal.Row).Specific).Value);
                                            LotNo = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_14").Cells.Item(pVal.Row).Specific).Value.ToString().Trim();
                                            VarianceAllowed = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_13").Cells.Item(pVal.Row).Specific).Value);
                                            _Status = _Form.DataSources.DBDataSources.Item("@WBLOADDTLH").GetValue("U_Status", 0);
                                            int TotalBundleCreated = 0, CurrentRowBundle = 0;
                                            double TotalWeightInTruck = 0;
                                            for (int i = 1; i <= _Matrix.RowCount; i++)
                                            {
                                                if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(i).Specific).Value != "")
                                                {
                                                    TotalBundleCreated = TotalBundleCreated + Convert.ToInt16(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(i).Specific).Value);
                                                }
                                                TotalWeightInTruck = TotalWeightInTruck + Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_2").Cells.Item(i).Specific).Value);
                                            }
                                            if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific).Value != "")
                                            {
                                                CurrentRowBundle = Convert.ToInt16(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific).Value);
                                            }
                                            TotalBundleCreated = TotalBundleCreated - CurrentRowBundle;
                                            TotalWeightInTruck = TotalWeightInTruck - Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific).Value);


                                            ((ClsWeighingBridgeDetails)_Object).FormDefault(PickListNo, ItemCode, ItemDescription, TruckNo, pVal.Row, NoOfPeacePer, QtyToDispatch, ItemGroupName, LotNo, TotalBundleCreated, TotalWeightInTruck, VarianceAllowed, _Status);
                                            oBubbleEvent = false;
                                        }
                                        else
                                        {
                                            Utilities.Message("Please first add data.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                            oBubbleEvent = false;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Utilities.ShowErrorMessage(ex.Message);
                                }
                            }
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.Message);
                }
            }
                        #endregion
            #endregion

            #region After Action
            else
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && _Form.Mode == BoFormMode.fm_OK_MODE)
                        {
                            _Form.Items.Item("4").Enabled = false;
                            EnableDisable();

                        }
                        if (pVal.ItemUID == "48")
                        {
                            _ComboBox.Select(2, BoSearchKey.psk_Index);
                            if (_Form.Mode == BoFormMode.fm_OK_MODE)
                            {
                                _Form.Mode = BoFormMode.fm_UPDATE_MODE;
                                _Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                            EnableDisable();
                        }
                        if (pVal.ItemUID == "45")
                        {
                            _ComboBox.Select(1, BoSearchKey.psk_Index);
                            ((SAPbouiCOM.EditText)_Form.Items.Item("36").Specific).Value = DateTime.Now.ToString("HH:mm tt");

                           int PickListNo =Convert.ToInt16(((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value);
                         //  UpdatePickedQuantity(PickListNo); //Commented By Manoj As discussed with Shahshank Standard SAP functionality handling the same after making the delivery
                           if (_Form.Mode == BoFormMode.fm_OK_MODE)
                               _Form.Mode = BoFormMode.fm_UPDATE_MODE;

                            //Start Vehicle management Screen data update 
                           double TotalWeightInBundles = 0;
                           for (int i = 1; i <= _Matrix.RowCount ; i++)
                           {
                               TotalWeightInBundles = TotalWeightInBundles + Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_2").Cells.Item(i).Specific).Value);
                           }
                           string SQL = "update \"@VHWTMGTH\" set \"U_NTVarWt\" = \"U_NetWt\" - " + TotalWeightInBundles + ",\"U_WBCounter\" = '" + TotalWeightInBundles + "' where \"U_Status\" = 'Open' AND \"U_PickLstNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value + "' AND \"U_VehNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("8").Specific).Value + "' ";
                           Utilities.ExecuteSQL(ref _Rs, SQL);
                           //End Vehicle management Screen data update

                          _Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                           _Form.Mode = BoFormMode.fm_VIEW_MODE;
                        }
                        if (pVal.ItemUID == "1" && _Form.Mode == BoFormMode.fm_ADD_MODE)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("20").Specific).Value == "")
                            {
                                ((SAPbouiCOM.EditText)_Form.Items.Item("20").Specific).Value = Utilities.getMaxColumnValueNum("@WBLOADDTLH", "U_DocNum");
                            }
                        }
                        break;

                    #region Set CFL Values

                    case BoEventTypes.et_CHOOSE_FROM_LIST:
                        _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                        if (pVal.ItemUID == "4" && _SysCFLEvent.SelectedObjects != null)
                        {
                            if(_Form.Mode != BoFormMode.fm_FIND_MODE)
                           _MasterDBDataSource.SetValue("U_PickLstNo", 0, _SysCFLEvent.SelectedObjects.GetValue("AbsEntry", 0).ToString().Trim());
                        }


                        break;

                    #endregion

                    #region et_KEY_DOWN

                    case BoEventTypes.et_KEY_DOWN:
                        if (pVal.ItemUID == "4" && pVal.CharPressed == 9)
                        {
                            if (((_Form.Items.Item("4").Specific).value != "") && _Form.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                _Form.Freeze(true);
                                
                                _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("39").Specific;
                                //string _Sql = "select T0.\"Remarks\",T0.\"Name\",T2.\"CardCode\",T2.\"CardName\",T2.\"DocNum\",T2.\"DocEntry\",T3.\"ItemCode\",T3.\"Dscription\",T3.\"Quantity\" " +
                                //              " ,T3.\"UomCode\",T4.\"OnHand\",T5.\"Quantity\" as \"Delivered Qty\",T1.\"RelQtty\" as \"PickQtty\",T6.\"U_NOPCPB\",T6.\"ItmsGrpNam\", " +
                                //              " T0.\"U_TruckNo\",T0.\"U_TranspNm\",T0.\"U_CounterNo\",T0.\"U_ContactNo\",T0.\"U_DriverNm\",\"U_VarAlWd\",T0.\"U_VehicleType\",T0.\"U_TareWeight\" from OPKL T0 " +
                                //              " INNER JOIN PKL1 T1 ON T0.\"AbsEntry\" = T1.\"AbsEntry\" " +
                                //              " INNER JOIN ORDR T2 ON T1.\"OrderEntry\" = T2.\"DocEntry\" " +
                                //              " INNER JOIN RDR1 T3 ON T2.\"DocEntry\" = T3.\"DocEntry\" AND T1.\"OrderLine\" = T3.\"LineNum\" " +
                                //              " INNER JOIN OITM T4 ON T3.\"ItemCode\" = T4.\"ItemCode\" " +
                                //              " INNER JOIN OITB T6 ON T6.\"ItmsGrpCod\" = T4.\"ItmsGrpCod\" " +
                                //              " LEFT OUTER JOIN DLN1 T5 ON T5.\"BaseEntry\" = T3.\"DocEntry\" AND T5.\"ItemCode\" = T3.\"ItemCode\" AND T5.\"BaseLine\" = T3.\"LineNum\" " +
                                //              " where T0.\"AbsEntry\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value.ToString().Trim() + "'";

                                string _Sql = "select CAST(T0.\"Remarks\" as VARCHAR) as \"Remarks\",T0.\"Name\",T2.\"CardCode\",T2.\"CardName\",T2.\"DocNum\",T2.\"DocEntry\",T3.\"ItemCode\",T3.\"Dscription\",T3.\"Quantity\" " +
                                            " ,T3.\"UomCode\",T4.\"OnHand\",SUM(T5.\"Quantity\") as \"Delivered Qty\",T1.\"RelQtty\" as \"PickQtty\",T6.\"U_NOPCPB\",T6.\"ItmsGrpNam\", " +
                                            " T0.\"U_TruckNo\",T0.\"U_TranspNm\",T0.\"U_CounterNo\",T0.\"U_ContactNo\",T0.\"U_DriverNm\",\"U_VarAlWd\",T0.\"U_VehicleType\",T0.\"U_TareWeight\" from OPKL T0 " +
                                            " INNER JOIN PKL1 T1 ON T0.\"AbsEntry\" = T1.\"AbsEntry\" " +
                                            " INNER JOIN ORDR T2 ON T1.\"OrderEntry\" = T2.\"DocEntry\" " +
                                            " INNER JOIN RDR1 T3 ON T2.\"DocEntry\" = T3.\"DocEntry\" AND T1.\"OrderLine\" = T3.\"LineNum\" " +
                                            " INNER JOIN OITM T4 ON T3.\"ItemCode\" = T4.\"ItemCode\" " +
                                            " INNER JOIN OITB T6 ON T6.\"ItmsGrpCod\" = T4.\"ItmsGrpCod\" " +
                                            " LEFT OUTER JOIN DLN1 T5 ON T5.\"BaseEntry\" = T3.\"DocEntry\" AND T5.\"ItemCode\" = T3.\"ItemCode\" AND T5.\"BaseLine\" = T3.\"LineNum\" " +
                                            " where T0.\"AbsEntry\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value.ToString().Trim() + "' "+
                                            " group by CAST(T0.\"Remarks\" as VARCHAR),T0.\"Name\",T2.\"CardCode\",T2.\"CardName\",T2.\"DocNum\",T2.\"DocEntry\",T3.\"ItemCode\", " +
                                            " T3.\"Dscription\",T3.\"Quantity\"  ,T3.\"UomCode\",T4.\"OnHand\",T1.\"RelQtty\",T6.\"U_NOPCPB\",T6.\"ItmsGrpNam\",  T0.\"U_TruckNo\",T0.\"U_TranspNm\", "+
                                            " T0.\"U_CounterNo\",T0.\"U_ContactNo\",T0.\"U_DriverNm\",\"U_VarAlWd\",T0.\"U_VehicleType\",T0.\"U_TareWeight\" ";


                                Utilities.ExecuteSQL(ref _RS, _Sql);
                                if (_RS.RecordCount > 0)
                                {
                                    if (counter == true)
                                    {
                                        ((SAPbouiCOM.EditText)_Form.Items.Item("6").Specific).Value = _RS.Fields.Item("U_TranspNm").Value.ToString();
                                        ((SAPbouiCOM.EditText)_Form.Items.Item("8").Specific).Value = _RS.Fields.Item("U_TruckNo").Value.ToString();
                                        ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value = _RS.Fields.Item("U_DriverNm").Value.ToString();
                                        ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value = _RS.Fields.Item("U_ContactNo").Value.ToString();
                                        ((SAPbouiCOM.EditText)_Form.Items.Item("50").Specific).Value = _RS.Fields.Item("U_CounterNo").Value.ToString();
                                        ((SAPbouiCOM.EditText)_Form.Items.Item("14").Specific).Value = DateTime.Today.AddDays(0).ToString("yyyyMMdd");
                                        ((SAPbouiCOM.EditText)_Form.Items.Item("16").Specific).Value = _RS.Fields.Item("U_TareWeight").Value.ToString();
                                        ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value = _RS.Fields.Item("CardCode").Value;
                                        ((SAPbouiCOM.EditText)_Form.Items.Item("24").Specific).Value = _RS.Fields.Item("CardName").Value;
                                        ((SAPbouiCOM.EditText)_Form.Items.Item("26").Specific).Value = _RS.Fields.Item("DocNum").Value.ToString();
                                        ((SAPbouiCOM.EditText)_Form.Items.Item("52").Specific).Value = _RS.Fields.Item("DocEntry").Value.ToString();
                                        ((SAPbouiCOM.EditText)_Form.Items.Item("30").Specific).Value = _RS.Fields.Item("Name").Value;
                                        ((SAPbouiCOM.EditText)_Form.Items.Item("32").Specific).Value = DateTime.Today.AddDays(0).ToString("yyyyMMdd");
                                        ((SAPbouiCOM.EditText)_Form.Items.Item("34").Specific).Value = DateTime.Now.ToString("HH:mm tt");
                                        ((SAPbouiCOM.EditText)_Form.Items.Item("60").Specific).Value = _RS.Fields.Item("U_VehicleType").Value;
                                        ((SAPbouiCOM.EditText)_Form.Items.Item("28").Specific).Value = _RS.Fields.Item("Remarks").Value;
                                      
                                        NoOfPeacePer = _RS.Fields.Item("U_NOPCPB").Value;
                                        ItemGroupName = _RS.Fields.Item("ItmsGrpNam").Value.ToString();
                                        VarianceAllowed = _RS.Fields.Item("U_VarAlWd").Value;

                                        for (int i = 1; i <= _RS.RecordCount; i++)
                                        {

                                            _Matrix.AddRow(1, -1);
                                            // _Matrix.FlushToDataSource();
                                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_13").Cells.Item(i).Specific).Value = _RS.Fields.Item("U_VarAlWd").Value.ToString();
                                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_10").Cells.Item(i).Specific).Value = _RS.Fields.Item("ItemCode").Value.ToString();
                                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_9").Cells.Item(i).Specific).Value = _RS.Fields.Item("Dscription").Value.ToString();
                                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_12").Cells.Item(i).Specific).Value = _RS.Fields.Item("ItmsGrpNam").Value.ToString();
                                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_11").Cells.Item(i).Specific).Value = _RS.Fields.Item("U_NOPCPB").Value.ToString();
                                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value = _RS.Fields.Item("Quantity").Value.ToString();
                                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_7").Cells.Item(i).Specific).Value = _RS.Fields.Item("UomCode").Value.ToString();
                                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_6").Cells.Item(i).Specific).Value = _RS.Fields.Item("OnHand").Value.ToString();
                                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_5").Cells.Item(i).Specific).Value = _RS.Fields.Item("Delivered Qty").Value.ToString();
                                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_4").Cells.Item(i).Specific).Value = _RS.Fields.Item("PickQtty").Value.ToString();
                                            _RS.MoveNext();

                                            //  _Matrix.LoadFromDataSource();
                                        }

                                        _Matrix.Columns.Item("V_8").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                                        _Matrix.Columns.Item("V_5").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                                        _Matrix.Columns.Item("V_4").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                                        _Matrix.Columns.Item("V_3").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                                        _Matrix.Columns.Item("V_2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                                        _Matrix.Columns.Item("V_1").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                                        _Matrix.Columns.Item("V_0").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                                    }
                                    Utilities.SerializedMartix(ref _Form, ref _Matrix);

                                    _Form.Freeze(false);
                                    counter = false;
                                }
                            }

                            EnableDisable();

                        }
                        //Key Down Enable Disable
                        if (_Form.Mode == BoFormMode.fm_UPDATE_MODE)
                        {
                            _Form.Items.Item("4").Enabled = false;
                            EnableDisable();
                        }

                        break;

                    #endregion


                }
            }
            #endregion
        }

        #endregion

        #region Menu Event

        public override void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool oBubbleEvent)
        {
            _Form.Freeze(true);
            try
            {
                #region Before Action

                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        #region mnu_DELETE_ROW

                        case Constants.System_Menus.mnu_DELETE_ROW:
                            {
                               
                            }
                            break;

                        #endregion
                    }
                }

                #endregion

                #region After Action

                else
                {
                    switch (pVal.MenuUID)
                    {
                        case Constants.System_Menus.mnu_FIND:
                            {
                                if(_Form.Mode != BoFormMode.fm_FIND_MODE)
                                { 
                                if (_Form.DataSources.DBDataSources.Item("@WBLOADDTLH").GetValue("U_Status", 0) == "Closed")
                                    _Form.Mode = BoFormMode.fm_VIEW_MODE;
                                else
                                    _Form.Mode = BoFormMode.fm_OK_MODE;
                                }

                                _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("39").Specific;
                                _Matrix.Columns.Item("V_10").Editable = false;
                                _Matrix.Columns.Item("V_9").Editable = false;
                                _Matrix.Columns.Item("V_8").Editable = false;
                                _Matrix.Columns.Item("V_7").Editable = false;
                                _Matrix.Columns.Item("V_6").Editable = false;
                                _Matrix.Columns.Item("V_5").Editable = false;
                                _Matrix.Columns.Item("V_4").Editable = false;
                                _Matrix.Columns.Item("V_3").Editable = false;
                                _Matrix.Columns.Item("V_2").Editable = false;
                                _Matrix.Columns.Item("V_1").Editable = false;
                                _Matrix.Columns.Item("V_13").Editable = false;
                                _Matrix.Columns.Item("V_0").Editable = false;
                                EnableDisable();
                            }
                            break;
                        case Constants.System_Menus.mnu_ADD:
                            {
                                if (_Form.Mode != BoFormMode.fm_ADD_MODE)
                                {
                                    if (_Form.DataSources.DBDataSources.Item("@WBLOADDTLH").GetValue("U_Status", 0) == "Closed")
                                        _Form.Mode = BoFormMode.fm_VIEW_MODE;
                                    else
                                        _Form.Mode = BoFormMode.fm_OK_MODE;
                                }
                                counter = true;
                                ((SAPbouiCOM.EditText)_Form.Items.Item("20").Specific).Value = Utilities.getMaxColumnValueNum("@WBLOADDTLH", "U_DocNum");
                              //  _Form.Items.Item("47").Enabled = true; ;
                                _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("47").Specific;
                                _ComboBox.Select(0, BoSearchKey.psk_Index);

                                _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("39").Specific;
                                _Matrix.Columns.Item("V_10").Editable = false;
                                _Matrix.Columns.Item("V_9").Editable = false;
                                _Matrix.Columns.Item("V_8").Editable = false;
                                _Matrix.Columns.Item("V_7").Editable = false;
                                _Matrix.Columns.Item("V_6").Editable = false;
                                _Matrix.Columns.Item("V_5").Editable = false;
                                _Matrix.Columns.Item("V_4").Editable = false;
                                _Matrix.Columns.Item("V_3").Editable = false;
                                _Matrix.Columns.Item("V_2").Editable = false;
                                _Matrix.Columns.Item("V_1").Editable = false;
                                _Matrix.Columns.Item("V_13").Editable = false;
                                _Matrix.Columns.Item("V_0").Editable = false;

                                EnableDisable();
                            }
                            break;
                        case Constants.System_Menus.mnu_FIRST:
                             {
                                 if (_Form.DataSources.DBDataSources.Item("@WBLOADDTLH").GetValue("U_Status", 0) == "Closed")
                                     _Form.Mode = BoFormMode.fm_VIEW_MODE;
                                 else
                                     _Form.Mode = BoFormMode.fm_OK_MODE;

                                 _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("39").Specific;
                                 _Matrix.Columns.Item("V_10").Editable = false;
                                 _Matrix.Columns.Item("V_9").Editable = false;
                                 _Matrix.Columns.Item("V_8").Editable = false;
                                 _Matrix.Columns.Item("V_7").Editable = false;
                                 _Matrix.Columns.Item("V_6").Editable = false;
                                 _Matrix.Columns.Item("V_5").Editable = false;
                                 _Matrix.Columns.Item("V_4").Editable = false;
                                 _Matrix.Columns.Item("V_3").Editable = false;
                                 _Matrix.Columns.Item("V_2").Editable = false;
                                 _Matrix.Columns.Item("V_1").Editable = false;
                                 _Matrix.Columns.Item("V_13").Editable = false;
                                 _Matrix.Columns.Item("V_0").Editable = false;

                                 _Form.Items.Item("4").Enabled = false;

                                 EnableDisable();
                            }
                            break;
                        case Constants.System_Menus.mnu_LAST:
                             {
                                 if (_Form.DataSources.DBDataSources.Item("@WBLOADDTLH").GetValue("U_Status", 0) == "Closed")
                                     _Form.Mode = BoFormMode.fm_VIEW_MODE;
                                 else
                                     _Form.Mode = BoFormMode.fm_OK_MODE;

                                 _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("39").Specific;
                                 _Matrix.Columns.Item("V_10").Editable = false;
                                 _Matrix.Columns.Item("V_9").Editable = false;
                                 _Matrix.Columns.Item("V_8").Editable = false;
                                 _Matrix.Columns.Item("V_7").Editable = false;
                                 _Matrix.Columns.Item("V_6").Editable = false;
                                 _Matrix.Columns.Item("V_5").Editable = false;
                                 _Matrix.Columns.Item("V_4").Editable = false;
                                 _Matrix.Columns.Item("V_3").Editable = false;
                                 _Matrix.Columns.Item("V_2").Editable = false;
                                 _Matrix.Columns.Item("V_1").Editable = false;
                                 _Matrix.Columns.Item("V_13").Editable = false;
                                 _Matrix.Columns.Item("V_0").Editable = false;

                                 _Form.Items.Item("4").Enabled = false;
                                 EnableDisable();
                            }
                            break;
                        case Constants.System_Menus.mnu_NEXT:
                             {
                                 if (_Form.DataSources.DBDataSources.Item("@WBLOADDTLH").GetValue("U_Status", 0) == "Closed")
                                     _Form.Mode = BoFormMode.fm_VIEW_MODE;
                                 else
                                     _Form.Mode = BoFormMode.fm_OK_MODE;

                                 _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("39").Specific;
                                 _Matrix.Columns.Item("V_10").Editable = false;
                                 _Matrix.Columns.Item("V_9").Editable = false;
                                 _Matrix.Columns.Item("V_8").Editable = false;
                                 _Matrix.Columns.Item("V_7").Editable = false;
                                 _Matrix.Columns.Item("V_6").Editable = false;
                                 _Matrix.Columns.Item("V_5").Editable = false;
                                 _Matrix.Columns.Item("V_4").Editable = false;
                                 _Matrix.Columns.Item("V_3").Editable = false;
                                 _Matrix.Columns.Item("V_2").Editable = false;
                                 _Matrix.Columns.Item("V_1").Editable = false;
                                 _Matrix.Columns.Item("V_13").Editable = false;
                                 _Matrix.Columns.Item("V_0").Editable = false;

                                 _Form.Items.Item("4").Enabled = false;
                                 EnableDisable();
                            }
                            break;
                        case Constants.System_Menus.mnu_PREVIOUS:
                            {
                                if (_Form.DataSources.DBDataSources.Item("@WBLOADDTLH").GetValue("U_Status", 0) == "Closed")
                                    _Form.Mode = BoFormMode.fm_VIEW_MODE;
                                else
                                    _Form.Mode = BoFormMode.fm_OK_MODE;

                                _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("39").Specific;
                                _Matrix.Columns.Item("V_10").Editable = false;
                                _Matrix.Columns.Item("V_9").Editable = false;
                                _Matrix.Columns.Item("V_8").Editable = false;
                                _Matrix.Columns.Item("V_7").Editable = false;
                                _Matrix.Columns.Item("V_6").Editable = false;
                                _Matrix.Columns.Item("V_5").Editable = false;
                                _Matrix.Columns.Item("V_4").Editable = false;
                                _Matrix.Columns.Item("V_3").Editable = false;
                                _Matrix.Columns.Item("V_2").Editable = false;
                                _Matrix.Columns.Item("V_1").Editable = false;
                                _Matrix.Columns.Item("V_13").Editable = false;
                                _Matrix.Columns.Item("V_0").Editable = false;

                                _Form.Items.Item("4").Enabled = false;
                                EnableDisable();
                            }
                            break;
                        //case Constants.System_Menus.mnu_DELETE_ROW:
                        //    {
                        //        if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        //        {
                        //            try
                        //            {
                        //                _Form.Freeze(true);
                        //                for (int i = 1; i <= _Matrix.VisualRowCount; i++)
                        //                {
                        //                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("0").Cells.Item(i).Specific).Value = i.ToString();
                        //                }
                        //            }
                        //            catch (Exception ex)
                        //            {
                        //                Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //            }
                        //            finally
                        //            {
                        //                _Form.Freeze(false);
                        //                _Form.Update();
                        //            }
                        //        }
                        //        else if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        //        {
                        //            try
                        //            {
                        //                _Form.Freeze(true);
                        //                for (int i = 1; i <= _Matrix.VisualRowCount; i++)
                        //                {

                        //                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("0").Cells.Item(i).Specific).Value = i.ToString();
                        //                }
                        //            }
                        //            catch (Exception ex)
                        //            {
                        //                Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //            }
                        //            finally
                        //            {
                        //                _Form.Freeze(false);
                        //                _Form.Update();
                        //            }
                        //        }
                        //    }
                        //    break;
                    }
                }
                #endregion
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

        #region FormData_Event

        public override void FormData_Event(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool oBubbleEvent)
        {

            if (BusinessObjectInfo.BeforeAction)
            {
                switch (BusinessObjectInfo.EventType)
                {
                    //case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                    //    if (BusinessObjectInfo.ActionSuccess == false)
                    //    {
                    //        ((SAPbouiCOM.EditText)_Form.Items.Item("20").Specific).Value = Utilities.getMaxColumnValueNum("@WBLOADDTLH", "U_DocNum");
                    //    }


                    //    break;
                }
            }
            else
            {
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                        string SQL = "update \"@VHWTMGTH\" set \"U_WBDocNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("20").Specific).Value + "',\"U_WBType\" = 'Dispatch Counter',\"U_WBStatus\" = 'Closed' where \"U_Status\" = 'Open' AND \"U_PickLstNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value + "' AND \"U_VehNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("8").Specific).Value + "' ";
                         Utilities.ExecuteSQL(ref _Rs, SQL);
                        break;
                }
            }
        }

        #endregion

        #region FormDefault

        public void FormDefault()
        {
            _MasterDBDataSource = Form.DataSources.DBDataSources.Item("@WBLOADDTLH");
            _ChildDBDataSource = Form.DataSources.DBDataSources.Item("@WBLOADDTLC");
            _Form.Items.Item("37").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            ((SAPbouiCOM.EditText)_Form.Items.Item("20").Specific).Value = Utilities.getMaxColumnValueNum("@WBLOADDTLH", "U_DocNum");
            _Form.DataBrowser.BrowseBy = "20";
          //  _Form.DataBrowser.BrowseBy = "4";
            _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("47").Specific;
           // _ComboBox.ValidValues.Add("Pending", "Pending");
            _ComboBox.ValidValues.Add("Open", "Open");
            _ComboBox.ValidValues.Add("Closed", "Closed");
            _ComboBox.ValidValues.Add("Draft", "Draft");
            _ComboBox.Select(0, BoSearchKey.psk_Index);

            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("39").Specific;
            _Matrix.Columns.Item("V_10").Editable = false;
            _Matrix.Columns.Item("V_9").Editable = false;
            _Matrix.Columns.Item("V_8").Editable = false;
            _Matrix.Columns.Item("V_7").Editable = false;
            _Matrix.Columns.Item("V_6").Editable = false;
            _Matrix.Columns.Item("V_5").Editable = false;
            _Matrix.Columns.Item("V_4").Editable = false;
            _Matrix.Columns.Item("V_3").Editable = false;
            _Matrix.Columns.Item("V_2").Editable = false;
            _Matrix.Columns.Item("V_1").Editable = false;
            _Matrix.Columns.Item("V_13").Editable = false;
            _Matrix.Columns.Item("V_0").Editable = false;

            _Form.Items.Item("6").Enabled = false;
            _Form.Items.Item("8").Enabled = false;
            _Form.Items.Item("10").Enabled = false;
            _Form.Items.Item("12").Enabled = false;
            _Form.Items.Item("50").Enabled = false;
            _Form.Items.Item("36").Enabled = false;
            _Form.Items.Item("28").Enabled = false;
            _Form.Items.Item("52").Enabled = false;
            _Form.Items.Item("54").Enabled = false;
            _Form.Items.Item("56").Enabled = false;
            _Form.Items.Item("14").Enabled = false;
            _Form.Items.Item("22").Enabled = false;
            _Form.Items.Item("24").Enabled = false;
            _Form.Items.Item("26").Enabled = false;
            _Form.Items.Item("30").Enabled = false;
            _Form.Items.Item("32").Enabled = false;
            _Form.Items.Item("34").Enabled = false;
            _Form.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            _Form.Items.Item("47").Enabled = false;
            _Form.Items.Item("20").Enabled = false;

            _Matrix.Columns.Item("V_8").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            _Matrix.Columns.Item("V_5").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            _Matrix.Columns.Item("V_4").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            _Matrix.Columns.Item("V_3").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            _Matrix.Columns.Item("V_2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            _Matrix.Columns.Item("V_1").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            _Matrix.Columns.Item("V_0").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
           
        }

        #endregion

        #region FILTER BP Type

        internal void FiltePickList(string CFLId)
        {

            _Form.Freeze(true);
            try
            {
                _Cons = new Conditions();
                _Con = _Cons.Add();
                _Con.BracketOpenNum = 1;
                _Con.Alias = "Status";
                _Con.Operation = BoConditionOperation.co_EQUAL;
                _Con.CondVal = "R";
                _Con.BracketCloseNum = 1;

                _Form.ChooseFromLists.Item("PickList").SetConditions(_Cons);

            }


            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }

        }

        #endregion

        #region GetDispatchValue

        public void GetDispatchValue(double TotalBundlesCreated, double TotalQuantityinBundles)
        {
            _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("39").Specific;
            _Form.Mode = BoFormMode.fm_UPDATE_MODE;
            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(RowNo).Specific).Value = TotalBundlesCreated.ToString();
            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_2").Cells.Item(RowNo).Specific).Value = TotalQuantityinBundles.ToString();
            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_1").Cells.Item(RowNo).Specific).Value = (Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_4").Cells.Item(RowNo).Specific).Value) - TotalQuantityinBundles).ToString();
            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(RowNo).Specific).Value = (Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_4").Cells.Item(RowNo).Specific).Value) - TotalQuantityinBundles).ToString();

            _Form.Items.Item("6").Enabled = false;
            _Form.Items.Item("8").Enabled = false;
            _Form.Items.Item("10").Enabled = false;
            _Form.Items.Item("12").Enabled = false;
            _Form.Items.Item("50").Enabled = false;
            _Form.Items.Item("36").Enabled = false;
            _Form.Items.Item("28").Enabled = false;

            _Form.Items.Item("14").Enabled = false;
            _Form.Items.Item("22").Enabled = false;
            _Form.Items.Item("24").Enabled = false;
            _Form.Items.Item("26").Enabled = false;
            _Form.Items.Item("30").Enabled = false;
            _Form.Items.Item("32").Enabled = false;
            _Form.Items.Item("34").Enabled = false;
            //_Form.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            _Form.Items.Item("18").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            _Form.Items.Item("47").Enabled = false;
            _Form.Items.Item("20").Enabled = false;
        }

        #endregion

        #region VALIDATION FUNCTION

        private bool Validation()
        {
            try
            {
                bool skipRow = false;
                for (int i = 0; i <= _ChildDBDataSource.Size - 1; i++)
                {
                    if ((i == _ChildDBDataSource.Size - 1 || Flags == false)
                        && _ChildDBDataSource.GetValue("U_ItemCode", _ChildDBDataSource.Size - 1).Trim() == ""
                        && _ChildDBDataSource.GetValue("U_ItemName", _ChildDBDataSource.Size - 1).Trim() == "")
                    {
                        skipRow = true;
                    }

                    if (skipRow == false)
                    {
                        if (_ChildDBDataSource.GetValue("U_ItemCode", i).Trim() == "")
                        {
                            Utilities.Message("Item Code can not left blank", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return false;
                        }
                        if (_ChildDBDataSource.GetValue("U_ItemName", i).Trim() == "")
                        {
                            Utilities.Message("Item Name can not left blank", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return false;
                        }

                    }
                }
                return true;
            }
            catch
            {
                throw;
            }
        }

        #endregion

        #region Update Picked Quantity

        public void UpdatePickedQuantity(int pickListAbsEntry)
        {

            PickLists opicklist = Utilities.Application.Company.GetBusinessObject(BoObjectTypes.oPickLists);
            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("39").Specific;

            if (opicklist.GetByKey(pickListAbsEntry))
            {
                int line = opicklist.Lines.Count;
                for (int i = 0; i < _Matrix.RowCount; i++)
                {
                    opicklist.Lines.SetCurrentLine(i);
                    if (_ChildDBDataSource.GetValue("U_NoOfBund", i).ToString() != "")
                    opicklist.Lines.PickedQuantity = Convert.ToDouble(_ChildDBDataSource.GetValue("U_NoOfBund", i).ToString());
                }
                int returnValue = opicklist.Update();
            }

        }

        #endregion

        #region Check Data in table

        private bool DATAEXIST()
        {
            try
            {
                string SQL = "Select Count(*)  from \"@WBLOADDTLH\" where \"U_PickLstN\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value.ToString() + "' ";
                Utilities.ExecuteSQL(ref _Rs, SQL);
                if (_Rs.Fields.Item(0).Value.ToString() != "" && Convert.ToInt32(_Rs.Fields.Item(0).Value.ToString()) > 0)
                {
                    return true;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }

        #endregion

        #region EnableDisable

        public void EnableDisable()
        {
            if (_Form.Mode != BoFormMode.fm_FIND_MODE)
            {
                if (_Form.Mode == BoFormMode.fm_ADD_MODE)
                {
                    _Form.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                _Form.Items.Item("20").Enabled = false;

            }
            _Form.Items.Item("6").Enabled = false;
            _Form.Items.Item("8").Enabled = false;
            _Form.Items.Item("10").Enabled = false;
            _Form.Items.Item("12").Enabled = false;
            _Form.Items.Item("50").Enabled = false;
            _Form.Items.Item("36").Enabled = false;
            _Form.Items.Item("28").Enabled = false;

            _Form.Items.Item("14").Enabled = false;
            _Form.Items.Item("16").Enabled = false;
            _Form.Items.Item("22").Enabled = false;
            _Form.Items.Item("24").Enabled = false;
            _Form.Items.Item("26").Enabled = false;
            _Form.Items.Item("30").Enabled = false;
            _Form.Items.Item("32").Enabled = false;
            _Form.Items.Item("34").Enabled = false;
            _Form.Items.Item("47").Enabled = false;
            _Form.Items.Item("52").Enabled = false;
            _Form.Items.Item("54").Enabled = false;
            _Form.Items.Item("56").Enabled = false;
            _Form.Items.Item("60").Enabled = false;

            if(_Form.Mode != BoFormMode.fm_FIND_MODE)
            {
                if (_Form.Mode == BoFormMode.fm_ADD_MODE )
                {
                    _Form.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                _Form.Items.Item("20").Enabled = false;

            }
        }

        #endregion

    }
}
