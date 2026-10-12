using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Threading;
using CrystalDecisions.CrystalReports.ViewerObjectModel;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Drawing.Printing;
using System.Net.Mail;
using System.IO;
using System.Threading;
using System.Text.RegularExpressions;
using CrystalDecisions.CrystalReports.Engine;
using System.Runtime.InteropServices;
using CrystalDecisions.Shared;
using System.Diagnostics;
using System.Configuration;
using System.Xml;
using CrystalDecisions.CrystalReports.Engine;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using System.IO.Ports;


namespace BatchManagement_LaxmiSteel
{
    class ClsWeighingBridgeDetails : Base
    {

        #region Variable Declaration

        private SAPbouiCOM.Matrix _Matrix;
        private SAPbouiCOM.DBDataSource _DBDataSource;
        private SAPbobsCOM.Recordset _Rs = null, _Rs1 = null;
        SAPbobsCOM.Recordset _RS = null;
        string SQL = string.Empty;
        int _LastErrorCode = 0;
        private bool Flags = true;
        int PvalCount = 0;
        string _LastErrorMsg = " ", strcon = "";
        SAPbouiCOM.ComboBox _ComboBox;
        private SAPbouiCOM.EditText _EditText;
        private int RowID { get; set; }
        private string _Code { get; set; }
        string _Activitytype = string.Empty;
        private static string _ItemCode = string.Empty;
        private static string _LotNo = string.Empty;
        private static string _ItemName = string.Empty;
        private static string _PickListNo = string.Empty;
        private static string _TruckNo = string.Empty;
        private static int _RowNo = 0, _NoOfPeacePer = 0, _TotalBundleCreated = 0;
        private static string _ItemGroupName = string.Empty;
        private static int RowNo = 0;
        private static double TotalBundlesCreated = 0,TotalWeight = 0;
        private static double TotalQuantityinBundles = 0, _TotalWeightInTruck = 0, _VarianceAllowed = 0;
        private static string _QtyToDispatch = string.Empty;
        public static int count = 1;
        string _DocEntry = "", _DocNum = "";
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        string _ImportFileName = "";
        SAPbouiCOM.Condition oCon;
        SAPbouiCOM.Conditions oCons;
        public string SubGrp1CodVal = string.Empty;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsWeighingBridgeDetails()
            : base()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsWeighingBridgeDetails()
        {
        }

        #endregion

        #region ITEM EVENT

        public override void Item_Event(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            #region BEFORE ACTION TRUE
            if (pVal.Before_Action)
            {
                switch (pVal.EventType)
                {
                   
                    #region et_ITEM_PRESSED Before Action

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && (Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                        {
                            try
                            {
                                
                                if (this.AddUpdateData() == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                                if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                _Matrix.FlushToDataSource();
                                if (Validation() == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                                else
                                {
                                    this.ConditionalReloadData(_PickListNo, _ItemCode);
                                    Utilities.SerializedMartix(ref _Form, ref _Matrix);
                                    Utilities.Application.SBO_Application.StatusBar.SetText("Operation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                    RowNo = 0;
                                }
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                        }
                        if (pVal.ItemUID == "7")
                        {
                            TotalBundlesCreated = 0;
                            TotalQuantityinBundles = 0;
                            TotalWeight = 0;
                            for (int i = 1; i <= _Matrix.RowCount-1; i++)
                            {
                                TotalQuantityinBundles = TotalQuantityinBundles + Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_4").Cells.Item(i).Specific).Value);
                            }

                            if (TotalQuantityinBundles < Convert.ToDouble(_QtyToDispatch) - Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("18").Specific).Value))
                            {
                                Utilities.Message("Weight exceeded more than allowed variance", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }

                            TotalBundlesCreated = _Matrix.RowCount - 1;
                            _Form.Close();
                            _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                            this._Object = new ClsDispatchCounter();
                            ((ClsDispatchCounter)this._Object).GetDispatchValue(TotalBundlesCreated, TotalQuantityinBundles);

                        }
                        if (pVal.ItemUID == "6")
                        {
                            _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("3").Specific;
                            string LotNo = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_6").Cells.Item(RowNo).Specific).Value.ToString().Trim();
                            string BundleNo = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_5").Cells.Item(RowNo).Specific).Value.ToString().Trim();
                            string Size = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_9").Cells.Item(RowNo).Specific).Value.ToString().Trim();
                            string NoOfPcs = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(RowNo).Specific).Value.ToString().Trim();
                            string Weight = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_4").Cells.Item(RowNo).Specific).Value.ToString().Trim();
                            string BatchNo = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_2").Cells.Item(RowNo).Specific).Value.ToString().Trim();
                            string ItemCode = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_8").Cells.Item(RowNo).Specific).Value.ToString();
                            string ItemCodeD = ItemCode.Substring(ItemCode.Length - 1);
                            if(ItemCodeD == "M")
                               ItemCodeD = "";

                            if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_10").Cells.Item(RowNo).Specific).Value != "Y")
                            {
                                ExportReportToPdfFile(LotNo, BundleNo, Size, NoOfPcs, Weight, BatchNo, ItemCodeD);
                            }
                           // Add update Data on Print Button
                            if (this.AddUpdateDataRowWise() == false)//this.AddUpdateDataRowWise
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                _Form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            _Matrix.FlushToDataSource();
                            if (Validation() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                                this.ConditionalReloadData(_PickListNo, _ItemCode);
                                Utilities.SerializedMartix(ref _Form, ref _Matrix);
                                RowNo = 0;
                            }
                            //Row Selection
                            _Matrix.Columns.Item("0").Cells.Item(_Matrix.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            //End Row Selection

                           // End Add update Data on Print Button

                        }
                        if (pVal.ItemUID == "12")
                        {
                            Weight(pVal.Row);
                        }
                        break;

                    #endregion

                    #region et_CLICK Before Action

                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ColUID == "0")
                        {
                            RowNo = pVal.Row;
                            _Matrix.CommonSetting.SetRowBackColor(count, -1);
                            count = pVal.Row ;
                            _Matrix.CommonSetting.SetRowBackColor(pVal.Row , 3198718);

                        }
                        if (pVal.ItemUID == "6")
                        {
                            if (RowNo <= 0)
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText("Please select any row for Printing.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                        }
                        if (pVal.ItemUID == "12")
                        {
                            if (RowNo <= 0)
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText("Please select any row for Printing.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                        }
                        break;

                    #endregion

                    #region et_KEY_DOWN Before Action

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.ColUID == "V_4" && pVal.CharPressed == 9)
                        {
                            double Weight = 0;
                           
                            for (int i = 1; i <= _Matrix.RowCount; i++)
                            {
                                Weight = Weight + Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_4").Cells.Item(i).Specific).Value);
                            }

                            if (Weight > Convert.ToDouble(_QtyToDispatch) + Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("18").Specific).Value))
                            {
                                Utilities.Message("Weight exceeded more than allowed variance : " + (Convert.ToDouble(_QtyToDispatch) + Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("18").Specific).Value) - Weight), SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific).Value = _PickListNo + "-" + _RowNo + "-B" + pVal.Row;
                            //((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_10").Cells.Item(pVal.Row).Specific).Value = "N";
                            if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                _Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                        if (pVal.ColUID == "V_4" && pVal.CharPressed != 9)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        if (pVal.ColUID == "V_5" && pVal.CharPressed != 9)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        if (pVal.ColUID == "V_10" && pVal.CharPressed != 9)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        if (pVal.ColUID == "V_0" && pVal.CharPressed == 9)
                        {
                            if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_6").Cells.Item(_Matrix.RowCount).Specific).Value.ToString() != "" && ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_2").Cells.Item(_Matrix.RowCount).Specific).Value.ToString() != "")
                            {
                                _Matrix.AddRow(1, -1);
                                ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_8").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _ItemCode.ToString();
                                ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_7").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _ItemName.ToString();
                                ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_6").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _LotNo.ToString();
                            
                                if (_ItemGroupName != "")
                                {
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_9").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _ItemGroupName.ToString();
                                }
                                else
                                {
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_9").Cells.Item(_Matrix.VisualRowCount).Specific).Value = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_9").Cells.Item(_Matrix.VisualRowCount - 1).Specific).Value;
                                }

                                if (pVal.Row < 9)
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_5").Cells.Item(_Matrix.VisualRowCount).Specific).Value = "0" + (pVal.Row + 1);
                                else
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_5").Cells.Item(_Matrix.VisualRowCount).Specific).Value = (pVal.Row + 1).ToString();

                                if (_NoOfPeacePer.ToString() != "0")
                                {
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _NoOfPeacePer.ToString();
                                }
                                else
                                {
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(_Matrix.VisualRowCount).Specific).Value = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(_Matrix.VisualRowCount - 1).Specific).Value;
                                }
                                ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_1").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _PickListNo.ToString();
                                ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _TruckNo.ToString();

                                //Row Selection
                                _Matrix.Columns.Item("0").Cells.Item(_Matrix.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                //End Row Selection

                                _Matrix.Columns.Item("V_9").Editable = true;
                                _Matrix.Columns.Item("V_4").Cells.Item(_Matrix.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                Utilities.SerializedMartix(ref _Form, ref _Matrix);
                            }
                        }
                        if (pVal.ColUID == "V_6" && pVal.CharPressed == 9)
                        {
                            //((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific).Value = _PickListNo + "-" + _RowNo + "-B" + pVal.Row;
                            //if( _Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            //_Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }

                        break;
                    #endregion

                }
            }

            #endregion

            #region After Action
            else
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:

                        if (pVal.ColUID == "V_4" && pVal.CharPressed == 9)
                        {
                            double Weight = 0;
                            double RemainingQty = 0;
                            for (int i = 1; i <= _Matrix.RowCount; i++)
                            {
                                Weight = Weight + Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_4").Cells.Item(i).Specific).Value);
                            }

                            RemainingQty = Convert.ToDouble(_QtyToDispatch) - Weight;
                            ((SAPbouiCOM.StaticText)_Form.Items.Item("11").Specific).Caption = RemainingQty.ToString();
                            ((SAPbouiCOM.StaticText)_Form.Items.Item("14").Specific).Caption = (_TotalBundleCreated + +_Matrix.RowCount).ToString();
                            ((SAPbouiCOM.StaticText)_Form.Items.Item("16").Specific).Caption = (_TotalWeightInTruck + Weight).ToString();
                        }
                        break;
                }
            }
            #endregion
        }

        #endregion

        #region FORM DEFAULT

        public void FormDefault(string PickListNo, string ItemCode, string ItemDescription, string TruckNo, int RowNo, int NoOfPeacePer, string QtyToDispatch, string ItemGroupName, string LotNo, int TotalBundleCreated, double TotalWeightInTruck, double VarianceAllowed,string _Status)
        {
            try
            {
                _Form.Freeze(true);
                _ItemCode = ItemCode;
                _ItemName = ItemDescription;
                _PickListNo = PickListNo;
                _TruckNo = TruckNo;
                _RowNo = RowNo;
                _NoOfPeacePer = NoOfPeacePer;
                _ItemGroupName = ItemGroupName;
                _QtyToDispatch = QtyToDispatch;
                _LotNo = LotNo;
                _TotalBundleCreated = TotalBundleCreated;
                _TotalWeightInTruck = TotalWeightInTruck;
                _VarianceAllowed = VarianceAllowed;

                count = 1;
                _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("3").Specific;
                _DBDataSource = this._Form.DataSources.DBDataSources.Item("@WBDETAIL");
                ((SAPbouiCOM.StaticText)_Form.Items.Item("5").Specific).Caption = ItemDescription.ToString();
                ((SAPbouiCOM.StaticText)_Form.Items.Item("9").Specific).Caption = _QtyToDispatch.ToString();
                ((SAPbouiCOM.StaticText)_Form.Items.Item("14").Specific).Caption = _TotalBundleCreated.ToString();

                if (dataexist() == true)
                {
                    _Matrix.AddRow(1, -1);

                    //Variance allowed value take data from database
                    SQL = "Select max(\"U_VarAlWd\") as \"U_VarAlWd\"  from \"@WBDETAIL\" where \"U_ItemCode\" = '" + _ItemCode.ToString() + "' AND \"U_PickLstN\" = '" + _PickListNo + "' ";
                    Utilities.ExecuteSQL(ref _Rs, SQL);
                    if (_Rs.RecordCount > 0)
                    {
                        ((SAPbouiCOM.EditText)_Form.Items.Item("18").Specific).Value = _Rs.Fields.Item("U_VarAlWd").Value.ToString();
                    }
                    //End Variance

                    this.ConditionalReloadData(_PickListNo, _ItemCode);
                    Utilities.SerializedMartix(ref _Form, ref _Matrix);
                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

                }
                else
                {
                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    _Matrix.AddRow(1, 1);

                    ((SAPbouiCOM.EditText)_Form.Items.Item("18").Specific).Value = _VarianceAllowed.ToString();
                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_8").Cells.Item(_Matrix.VisualRowCount).Specific).Value = ItemCode.ToString();
                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_7").Cells.Item(_Matrix.VisualRowCount).Specific).Value = ItemDescription.ToString();
                   // ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_10").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _VarianceAllowed.ToString();
                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_9").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _ItemGroupName.ToString();
                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_5").Cells.Item(_Matrix.VisualRowCount).Specific).Value = "01";
                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _NoOfPeacePer.ToString();
                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_1").Cells.Item(_Matrix.VisualRowCount).Specific).Value = PickListNo.ToString();
                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(_Matrix.VisualRowCount).Specific).Value = TruckNo.ToString();
                   ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_6").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _LotNo.ToString();
                   ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_10").Cells.Item(_Matrix.VisualRowCount).Specific).Value = "N";
                    Utilities.SerializedMartix(ref _Form, ref _Matrix);
                 
                }
                //remaining Quantity Calculation
                double Weight = 0;
                double RemainingQty = 0;
                _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("3").Specific;
                for (int i = 1; i <= _Matrix.RowCount; i++)
                {
                    Weight = Weight + Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_4").Cells.Item(i).Specific).Value);
                }

                RemainingQty = Convert.ToDouble(_QtyToDispatch) - Weight;
                ((SAPbouiCOM.StaticText)_Form.Items.Item("11").Specific).Caption = RemainingQty.ToString();
                ((SAPbouiCOM.StaticText)_Form.Items.Item("14").Specific).Caption = (_TotalBundleCreated + _Matrix.RowCount - 1).ToString();
                ((SAPbouiCOM.StaticText)_Form.Items.Item("16").Specific).Caption = (_TotalWeightInTruck + Weight).ToString();
                //End Remaining Quantity Calculation
                _Matrix.Columns.Item("V_5").Editable = true;
                _Matrix.Columns.Item("V_3").Editable = true;
                _Matrix.Columns.Item("V_4").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Matrix.Columns.Item("V_3").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                //Row Selection
                _Matrix.Columns.Item("0").Cells.Item(_Matrix.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                //End Row Selection

                string strSQL = "select \"U_VarAlwd\" from OUSR where \"USER_CODE\" = '" + Utilities.Application.Company.UserName + "'";
                Utilities.ExecuteSQL(ref _RS, strSQL);
                if (_RS.Fields.Item(0).Value.ToString().Trim() == "YES")
                {
                    _Form.Items.Item("18").Enabled = true;
                    _Form.Items.Item("1").Enabled = true;
                  
                }
                else
                {
                    _Form.Items.Item("1").Enabled = false;
                }
              if(_Status == "Closed")
              {
                  _Form.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE;
              }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Visible = true;
                _Form.Freeze(false);
            }


        }
        
        #endregion

        #region Check Data in table

        private bool dataexist()
        {
            try
            {
                SQL = "Select Count(*)  from \"@WBDETAIL\" where \"U_ItemCode\" = '" + _ItemCode.ToString() + "' AND \"U_PickLstN\" = '" + _PickListNo + "' AND \"U_BLineNum\" = '" + _RowNo + "' ";
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

        #region  Conditional Reload data in Form

        private void ConditionalReloadData(string _PickListNo, string _ItemCode)
        {
            try
            {
                _Form.Freeze(true);
                oCons = new SAPbouiCOM.Conditions();

                //oCon = oCons.Add();
                //oCon.BracketOpenNum = 2;
                //oCon.Alias = "U_PickLstN";
                //oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                //oCon.CondVal = _PickListNo.ToString();
                //oCon.BracketCloseNum = 1;
                //oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                //oCon = oCons.Add();
                //oCon.BracketOpenNum = 1;
                //oCon.Alias = "U_ItemCode";
                //oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                //oCon.CondVal = _ItemCode.ToString();
                //oCon.BracketCloseNum = 2;

                oCon = oCons.Add();
                oCon.BracketOpenNum = 2;
                oCon.Alias = "U_PickLstN";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = _PickListNo.ToString();
                oCon.BracketCloseNum = 1;
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCon = oCons.Add();
                oCon.Alias = "U_BLineNum";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal =_RowNo.ToString();
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                
                oCon = oCons.Add();
                oCon.BracketOpenNum = 1;
                oCon.Alias = "U_ItemCode";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = _ItemCode.ToString();
                oCon.BracketCloseNum = 2;
                
                _DBDataSource.Query(oCons);
                _Matrix.LoadFromDataSource();
                _DBDataSource.Clear();
                if (_Matrix.RowCount < 1)
                    _Matrix.AddRow(1, -1);

                if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_6").Cells.Item(_Matrix.RowCount).Specific).Value.ToString() != "")
                {
                    //    _Matrix.AddRow(1, -1);
                    _Matrix.AddRow(1, -1);
                   
                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_8").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _ItemCode.ToString();
                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_7").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _ItemName.ToString();
                  //  ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_10").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _VarianceAllowed.ToString();
                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_6").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _LotNo.ToString();
                    if (_ItemGroupName != "")
                    {
                        ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_9").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _ItemGroupName.ToString();
                    }
                    else
                    {
                        ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_9").Cells.Item(_Matrix.VisualRowCount).Specific).Value = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_9").Cells.Item(_Matrix.VisualRowCount - 1).Specific).Value;
                    }

                    if (_Matrix.RowCount <= 9)
                        ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_5").Cells.Item(_Matrix.VisualRowCount).Specific).Value = "0" + (_Matrix.RowCount);
                    else
                        ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_5").Cells.Item(_Matrix.VisualRowCount).Specific).Value = (_Matrix.RowCount).ToString();

                    if (_NoOfPeacePer.ToString() != "0")
                    {
                        ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _NoOfPeacePer.ToString();
                    }
                    else
                    {
                        ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(_Matrix.VisualRowCount).Specific).Value = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(_Matrix.VisualRowCount - 1).Specific).Value;
                    }
                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_1").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _PickListNo.ToString();
                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(_Matrix.VisualRowCount).Specific).Value = _TruckNo.ToString();
                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_10").Cells.Item(_Matrix.VisualRowCount).Specific).Value = "N";
                    _Matrix.Columns.Item("V_9").Editable = true;
                    _Matrix.Columns.Item("V_4").Cells.Item(_Matrix.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    Utilities.SerializedMartix(ref _Form, ref _Matrix);
                }

            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally { _Form.Freeze(false); }
        }

        #endregion

        #region ADD UPDATE DATA

        private bool AddUpdateData()
        {
            bool _Success = true;
            SAPbobsCOM.UserTable oUserTable = Utilities.Application.Company.UserTables.Item("WBDETAIL");

            try
            {
                _Form.Freeze(true);
                _Matrix.FlushToDataSource();


                for (int i = 0; i <= _DBDataSource.Size - 1; i++)
                {
                    string MaxValue = Utilities.getMaxColumnValue("@WBDETAIL", "Code");
                    string Code = _DBDataSource.GetValue("Code", i).ToString().Trim();
                    string LotNo = _DBDataSource.GetValue("U_LotNo", i).ToString().Trim();
                    string ChkBatchNo = _DBDataSource.GetValue("U_BatchNo", i).ToString().Trim();
                    if (LotNo != "" && ChkBatchNo != "" && oUserTable.GetByKey(Code)) 
                    {
                        oUserTable.UserFields.Fields.Item("U_ItemCode").Value = _DBDataSource.GetValue("U_ItemCode", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_ItemName").Value = _DBDataSource.GetValue("U_ItemName", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_VarAlWd").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("18").Specific).Value;
                        oUserTable.UserFields.Fields.Item("U_ItmsGrpN").Value = _DBDataSource.GetValue("U_ItmsGrpN", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_LotNo").Value = _DBDataSource.GetValue("U_LotNo", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_BundleNo").Value = _DBDataSource.GetValue("U_BundleNo", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_Weight").Value = _DBDataSource.GetValue("U_Weight", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_NOPCPB").Value = _DBDataSource.GetValue("U_NOPCPB", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_BatchNo").Value = _DBDataSource.GetValue("U_BatchNo", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_PickLstN").Value = _DBDataSource.GetValue("U_PickLstN", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_TruckNo").Value = _DBDataSource.GetValue("U_TruckNo", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_LineNum").Value = i.ToString();
                        oUserTable.UserFields.Fields.Item("U_BLineNum").Value = _RowNo.ToString();

                        int rtnVal = oUserTable.Update(); 
                        if (rtnVal != 0)
                        {
                            Utilities.Application.Company.GetLastError(out _LastErrorCode, out _LastErrorMsg);
                            return false;
                        }
                    }
                    else if (LotNo != "" && ChkBatchNo != "")
                    {
                        oUserTable.Code = MaxValue; 
                        oUserTable.Name = MaxValue;
                        oUserTable.UserFields.Fields.Item("U_ItemCode").Value = _DBDataSource.GetValue("U_ItemCode", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_ItemName").Value = _DBDataSource.GetValue("U_ItemName", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_VarAlWd").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("18").Specific).Value;
                        oUserTable.UserFields.Fields.Item("U_ItmsGrpN").Value = _DBDataSource.GetValue("U_ItmsGrpN", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_LotNo").Value = _DBDataSource.GetValue("U_LotNo", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_BundleNo").Value = _DBDataSource.GetValue("U_BundleNo", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_Weight").Value = _DBDataSource.GetValue("U_Weight", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_NOPCPB").Value = _DBDataSource.GetValue("U_NOPCPB", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_BatchNo").Value = _DBDataSource.GetValue("U_BatchNo", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_PickLstN").Value = _DBDataSource.GetValue("U_PickLstN", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_TruckNo").Value = _DBDataSource.GetValue("U_TruckNo", i).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_LineNum").Value = i.ToString();
                        oUserTable.UserFields.Fields.Item("U_BLineNum").Value = _RowNo.ToString();

                        int rtnVal = oUserTable.Add();
                        if (rtnVal != 0)
                        {
                            Utilities.Application.Company.GetLastError(out _LastErrorCode, out _LastErrorMsg);
                            return false;
                        }

                    }
                   
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
                return false;
            }
            finally
            {
                _Form.Freeze(false);
            }
            return _Success;
        }

        #endregion

        #region ADD UPDATE DATA

        private bool AddUpdateDataRowWise()
        {
            bool _Success = true;
            SAPbobsCOM.UserTable oUserTable = Utilities.Application.Company.UserTables.Item("WBDETAIL");

            try
            {
                _Form.Freeze(true);
                _Matrix.FlushToDataSource();


                //for (int i = 0; i <= _DBDataSource.Size - 1; i++)
                //{
                    string MaxValue = Utilities.getMaxColumnValue("@WBDETAIL", "Code");
                    string Code = _DBDataSource.GetValue("Code", RowNo-1).ToString().Trim();
                    string LotNo = _DBDataSource.GetValue("U_LotNo", RowNo-1).ToString().Trim();
                    string ChkBatchNo = _DBDataSource.GetValue("U_BatchNo", RowNo-1).ToString().Trim();
                    if (LotNo != "" && ChkBatchNo != "" && oUserTable.GetByKey(Code))
                    {
                        oUserTable.UserFields.Fields.Item("U_ItemCode").Value = _DBDataSource.GetValue("U_ItemCode", RowNo - 1).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_ItemName").Value = _DBDataSource.GetValue("U_ItemName", RowNo - 1).ToString().Trim();
                        //oUserTable.UserFields.Fields.Item("U_VarAlWd").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("18").Specific).Value;
                        oUserTable.UserFields.Fields.Item("U_ItmsGrpN").Value = _DBDataSource.GetValue("U_ItmsGrpN", RowNo - 1).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_LotNo").Value = _DBDataSource.GetValue("U_LotNo", RowNo - 1).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_BundleNo").Value = _DBDataSource.GetValue("U_BundleNo", RowNo - 1).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_Weight").Value = _DBDataSource.GetValue("U_Weight", RowNo - 1).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_NOPCPB").Value = _DBDataSource.GetValue("U_NOPCPB", RowNo - 1).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_BatchNo").Value = _DBDataSource.GetValue("U_BatchNo", RowNo - 1).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_PickLstN").Value = _DBDataSource.GetValue("U_PickLstN", RowNo - 1).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_TruckNo").Value = _DBDataSource.GetValue("U_TruckNo", RowNo - 1).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_PrntSt").Value = "Y";
                        oUserTable.UserFields.Fields.Item("U_LineNum").Value = (RowNo-1).ToString();
                        oUserTable.UserFields.Fields.Item("U_BLineNum").Value = _RowNo.ToString();

                        int rtnVal = oUserTable.Update();
                        if (rtnVal != 0)
                        {
                            Utilities.Application.Company.GetLastError(out _LastErrorCode, out _LastErrorMsg);
                            return false;
                        }
                    }
                    else if (LotNo != "" && ChkBatchNo != "")
                    {
                        oUserTable.Code = MaxValue;
                        oUserTable.Name = MaxValue;
                        oUserTable.UserFields.Fields.Item("U_ItemCode").Value = _DBDataSource.GetValue("U_ItemCode", RowNo - 1).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_ItemName").Value = _DBDataSource.GetValue("U_ItemName", RowNo - 1).ToString().Trim();
                        //oUserTable.UserFields.Fields.Item("U_VarAlWd").Value = ((SAPbouiCOM.EditText)_Form.Items.Item("18").Specific).Value;
                        oUserTable.UserFields.Fields.Item("U_ItmsGrpN").Value = _DBDataSource.GetValue("U_ItmsGrpN", RowNo - 1).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_LotNo").Value = _DBDataSource.GetValue("U_LotNo", RowNo - 1).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_BundleNo").Value = _DBDataSource.GetValue("U_BundleNo", RowNo - 1).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_Weight").Value = _DBDataSource.GetValue("U_Weight", RowNo - 1).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_NOPCPB").Value = _DBDataSource.GetValue("U_NOPCPB", RowNo - 1).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_BatchNo").Value = _DBDataSource.GetValue("U_BatchNo", RowNo - 1).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_PickLstN").Value = _DBDataSource.GetValue("U_PickLstN", RowNo - 1).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_TruckNo").Value = _DBDataSource.GetValue("U_TruckNo", RowNo - 1).ToString().Trim();
                        oUserTable.UserFields.Fields.Item("U_PrntSt").Value = "Y";
                        oUserTable.UserFields.Fields.Item("U_LineNum").Value = (RowNo-1).ToString();
                        oUserTable.UserFields.Fields.Item("U_BLineNum").Value = _RowNo.ToString();
                        int rtnVal = oUserTable.Add();
                        if (rtnVal != 0)
                        {
                            Utilities.Application.Company.GetLastError(out _LastErrorCode, out _LastErrorMsg);
                            return false;
                        }

                    }

                //}
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
                return false;
            }
            finally
            {
                _Form.Freeze(false);
            }
            return _Success;
        }

        #endregion

        #region Remove Record

        private bool RemoveRecord()
        {
            _Form.Freeze(true);
            SAPbobsCOM.UserTable oUserTable = Utilities.Application.Company.UserTables.Item("WBDETAIL");
            try
            {
                if (oUserTable.GetByKey(_Code))
                {
                    int lRetCode1 = oUserTable.Remove();
                    if (lRetCode1 != 0)
                    {
                        return false;
                    }
                    this.ConditionalReloadData(_PickListNo, _ItemCode);
                    Utilities.SerializedMartix(ref _Form, ref _Matrix);
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
                _Code = "";
            }
            return true;
        }

        #endregion

        #region VALIDATION FUNCTION

        private bool Validation()
        {
            try
            {
                bool skipRow = false;
                for (int i = 0; i <= _DBDataSource.Size - 1; i++)
                {
                    if ((i == _DBDataSource.Size - 1 || Flags == false)
                        && _DBDataSource.GetValue("U_ItemCode", _DBDataSource.Size - 1).Trim() == ""
                        && _DBDataSource.GetValue("U_ItemName", _DBDataSource.Size - 1).Trim() == "")
                    {
                        skipRow = true;
                    }

                    if (skipRow == false)
                    {
                        if (_DBDataSource.GetValue("U_ItemCode", i).Trim() == "")
                        {
                            Utilities.Message("Item Code can not left blank", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return false;
                        }
                        if (_DBDataSource.GetValue("U_ItemName", i).Trim() == "")
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

        //#region Export Report To PdfFile

        //private void ExportReportToPdfFile(string LotNo, string BundleNo, string Size, string NoOfPcs, string Weight, string BatchNo,string ItemD)
        //{
        //    DataSet ds = LoadCompanySetting();
        //    string DatabaseServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
        //    string dbUser = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
        //    string dbPassword = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();
        //    string Database = Utilities.Application.Company.CompanyDB;
        //    string currentUser = Utilities.Application.Company.UserName;
        //    CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
        //    CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();
        //    ParameterFields paramFields = new ParameterFields();
        //    ParameterField paramField = new ParameterField();
        //    ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
        //    ParameterRangeValue rangeVal = new ParameterRangeValue();

        //    using (CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument())
        //    {
        //        try
        //        {
        //            _ImportFileName = "";
        //            string ReportPath = string.Empty;

        //            ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "SteackerPrinting-Parameter.rpt";
        //            ReportDocument rpt = new ReportDocument();
        //            objCryRep.Load(ReportPath);
        //            objCryRep.DataSourceConnections.Clear();
        //            string strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY}", dbUser, dbPassword, DatabaseServerName, Database);
        //            NameValuePairs2 logonProps2 = objCryRep.DataSourceConnections[0].LogonProperties;
        //            logonProps2.Set("Provider", "B1CRHPROXY");
        //            logonProps2.Set("Server Type", "B1CRHPROXY");
        //            logonProps2.Set("Connection String", strConnection);
        //            objCryRep.DataSourceConnections[0].SetLogonProperties(logonProps2);
        //            objCryRep.DataSourceConnections[0].SetConnection(DatabaseServerName, Database, false);
        //            objCryRep.DataSourceConnections[0].SetLogon(dbUser, dbPassword);
        //            int n = objCryRep.Database.Tables.Count;
        //            //_DocEntry = _Form.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0).ToString();
        //            ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);
        //            objCryRep.Refresh();
        //            objCryRep.SetParameterValue("@LotNo", LotNo);
        //            objCryRep.SetParameterValue("@BundleNo", BundleNo);
        //            objCryRep.SetParameterValue("@Size", Size);
        //            objCryRep.SetParameterValue("@NoOfPcs", NoOfPcs);
        //            objCryRep.SetParameterValue("@Weight", Weight);
        //            objCryRep.SetParameterValue("@BatchNo", BatchNo);
        //            objCryRep.SetParameterValue("@ItemD", ItemD);

        //            _ImportFileName = "Test.pdf";
        //           // objCryRep.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, "C:" + "\\" + _ImportFileName);

        //            objCryRep.PrintToPrinter(1, true, 1, 1);

        //        }
        //        catch (Exception ex)
        //        {
        //            Utilities.WriteLogFile("ExportReportToPdfFile Error :" + ex.Message);
        //            throw ex;
        //        }
        //        finally
        //        {
        //            if (objCryRep != null)
        //            {
        //                objCryRep.Close();
        //                objCryRep.Dispose();
        //            }
        //            objConInfo = null;
        //            objTabLoginInfo = null;
        //            paramFields = null;
        //            paramField = null;
        //            discreteVal = null;
        //            rangeVal = null;
        //        }
        //    }
        //}

        //#endregion

        #region Export Report To PdfFile

        private void ExportReportToPdfFile(string LotNo, string BundleNo, string Size, string NoOfPcs, string Weight, string BatchNo, string ItemD)
        {
            string Database = Utilities.Application.Company.CompanyDB;
            string currentUser = Utilities.Application.Company.UserName;

            CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
            CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();

            ParameterFields paramFields = new ParameterFields();
            ParameterField paramField = new ParameterField();

            ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
            ParameterRangeValue rangeVal = new ParameterRangeValue();

            using (CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument())
            {
                try
                {
                    _ImportFileName = "";
                    string ReportPath = string.Empty;

                    ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "SteackerPrinting-Parameter.rpt";
                    objCryRep.Load(ReportPath);

                    objCryRep.Refresh();
                    objCryRep.SetParameterValue("@LotNo", LotNo);
                    objCryRep.SetParameterValue("@BundleNo", BundleNo);
                    objCryRep.SetParameterValue("@Size", Size);
                    objCryRep.SetParameterValue("@NoOfPcs", NoOfPcs);
                    objCryRep.SetParameterValue("@Weight", Weight);
                    objCryRep.SetParameterValue("@BatchNo", BatchNo);
                    objCryRep.SetParameterValue("@ItemD", ItemD);

                    _ImportFileName = "Test.pdf";
                    // objCryRep.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, "C:" + "\\" + _ImportFileName);

                    objCryRep.PrintToPrinter(1, true, 1, 1);

                }
                catch (Exception ex)
                {
                    Utilities.WriteLogFile("ExportReportToPdfFile Error :" + ex.Message);
                    throw ex;
                }
                finally
                {
                    if (objCryRep != null)
                    {
                        objCryRep.Close();
                        objCryRep.Dispose();
                    }
                    objConInfo = null;
                    objTabLoginInfo = null;
                    paramFields = null;
                    paramField = null;
                    discreteVal = null;
                    rangeVal = null;
                }
            }
        }

        #endregion

        #region LOAD COMPANY SETTINGS

        public DataSet LoadCompanySetting()
        {
            XmlDataDocument xmlDoc = new XmlDataDocument();
            DataSet dataSet = new DataSet();
            dataSet.ReadXml("AppConfig.xml");
            return dataSet;

        }

        #endregion

        #region Weight

        public void Weight(int RNo)
        {

            SerialPort mySerialPort = new SerialPort("COM1");
            mySerialPort.BaudRate = 9600;
            mySerialPort.Parity = Parity.Even;
            mySerialPort.StopBits = StopBits.One;
            mySerialPort.DataBits = 7;
            mySerialPort.Handshake = Handshake.None;
            if (!mySerialPort.IsOpen)
            {
                mySerialPort.Open();
            }
            mySerialPort.DataReceived += new SerialDataReceivedEventHandler(DataReceivedHandler);
            string aa = mySerialPort.ReadExisting();
            string aa1 = mySerialPort.ReadLine();
            string Test = mySerialPort.ReadLine().ToString();
            Regex digits = new Regex(@"^\D*?((-?(\d+(\.\d+)?))|(-?\.\d+)).*");
            Match mx = digits.Match(aa1);
            decimal strValue1 = mx.Success ? Convert.ToDecimal(mx.Groups[1].Value) : 0;
            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_4").Cells.Item(RowNo).Specific).Value = strValue1.ToString();
            if (mySerialPort.IsOpen)
            {
                mySerialPort.Close();
            }

        }

        #endregion

        #region DataReceivedHandler

        private static void DataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            
        }

        #endregion

    }
}
