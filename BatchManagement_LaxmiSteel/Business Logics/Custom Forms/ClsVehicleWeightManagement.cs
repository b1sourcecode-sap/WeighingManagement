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
using SAPbobsCOM;
using SAPbouiCOM;

namespace BatchManagement_LaxmiSteel
{
    class ClsVehicleWeightManagement : Base
    {

        #region Variable Declaration

        SAPbouiCOM.ComboBox _ComboBox;
        SAPbouiCOM.Matrix _Matrix;
        ChooseFromListEvent _SysCFLEvent;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource, _ChildDBDataSource;
        SAPbobsCOM.Recordset _RS = null, _Rs = null;
        public static int RowNo = 0;
        public static int NoOfPeacePer = 0;
        public static double VarianceAllowed = 0;
        public static string ItemGroupName = string.Empty;
        public static int count = 1;
         public string _DBName = string.Empty;
         string _DocEntry = "";
         string ReportPath = string.Empty;
         string _ImportFileName = "";

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsVehicleWeightManagement()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsVehicleWeightManagement()
        {
        }

        #endregion

        #region Item Event

        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
        {

            #region Before Action

            if (pVal.Before_Action)
            {
                switch (pVal.EventType)
                {
                    #region et_CLICK Before Action

                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ColUID == "0")
                        {
                            RowNo = pVal.Row;
                            _Matrix.CommonSetting.SetRowBackColor(count, -1);
                            count = pVal.Row;
                            _Matrix.CommonSetting.SetRowBackColor(pVal.Row, 3198718);

                        }
                        if (pVal.ItemUID == "44")
                        {
                            if (RowNo <= 0)
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText("Please select any row for Printing.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                        }

                        break;

                    #endregion

                    #region et_ITEM_PRESSED

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                        if (pVal.ItemUID == "1" && _Form.Mode == BoFormMode.fm_ADD_MODE)
                        {
                            ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value = Utilities.getMaxColumnValueNum("@VHWTMGTH", "U_DocNum");
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("52").Specific).Value == "")
                            {
                                Utilities.Message("Country Of Origin can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                            else if (((SAPbouiCOM.EditText)_Form.Items.Item("6").Specific).Value == "")
                            {
                                Utilities.Message("Vehicle No can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                            else if (((SAPbouiCOM.ComboBox)_Form.Items.Item("1000001").Specific).Selected.Value == "-")
                            {
                                Utilities.Message("Please select Visit Purpose.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                            else if (DATAEXIST() == true)
                            {
                                Utilities.Message("Vehicle no already exist with open status", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                        }
                        if (pVal.ItemUID == "63" && ((SAPbouiCOM.EditText)_Form.Items.Item("36").Specific).Value != "")
                        {
                            Utilities.Message("Weighing Bridge No. already allocated for this vehicle can not cancel the transaction ", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            oBubbleEvent = false;
                            return;
                        }

                        if(pVal.ItemUID == "44")
                        {
                             Weight(pVal.Row);
                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_7").Cells.Item(RowNo).Specific).Value = Utilities.Application.Company.UserName;
                            _Matrix.Columns.Item("V_5").Editable = true;
                            _Matrix.Columns.Item("V_4").Editable = true;
                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_5").Cells.Item(RowNo).Specific).Value = DateTime.Now.ToString("yyyyMMdd");
                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_4").Cells.Item(RowNo).Specific).Value = DateTime.Now.ToString("HH:mm tt");
                            _Matrix.Columns.Item("V_0").Cells.Item(RowNo).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            _Matrix.Columns.Item("V_5").Editable = false;
                            _Matrix.Columns.Item("V_4").Editable = false;
                        }

                        break;

                    #endregion

                    #region et_KEY_DOWN Before Action

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.ColUID == "V_1" && pVal.CharPressed != 9)
                        {
                            oBubbleEvent = false;
                            return;
                        }
                        if (pVal.ItemUID == "24" && pVal.CharPressed != 9)
                        {
                            oBubbleEvent = false;
                            return;
                        }
                        if (pVal.ItemUID == "52" && _Form.Mode != BoFormMode.fm_ADD_MODE && pVal.CharPressed != 9)
                        {
                            oBubbleEvent = false;
                            return;
                        }
                        if (pVal.ItemUID == "6" && _Form.Mode != BoFormMode.fm_ADD_MODE && pVal.CharPressed != 9)
                        {
                            oBubbleEvent = false;
                            return;
                        }
                        break;

                   #endregion

                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        if (pVal.ItemUID == "46" && _Form.Mode != BoFormMode.fm_ADD_MODE && pVal.CharPressed != 9)
                        {
                            oBubbleEvent = false;
                            return;
                        }

                        break;

                }
            }

            #endregion

            #region After Action

            else
            {
                switch (pVal.EventType)
                {

                    #region et_LOST_FOCUS

                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                        if(pVal.ItemUID == "6" && pVal.Action_Success == true)
                        {
                            ((SAPbouiCOM.EditText)_Form.Items.Item("6").Specific).Value = ((SAPbouiCOM.EditText)_Form.Items.Item("6").Specific).Value.ToString().Trim().ToUpper().Replace(" ", ""); 
                        }
                        if (pVal.ItemUID == "10" && pVal.Action_Success == true)
                        {
                            ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value = ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString().Trim().ToUpper().Replace(" ", "");
                        }
                        
                        break;
                        

                    #endregion

                    #region et_KEY_DOWN 

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("41").Specific;
                        if (pVal.ColUID == "V_1" && pVal.CharPressed == 9 && ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific).Value == "Tare Weight")
                        {
                            ((SAPbouiCOM.EditText)_Form.Items.Item("32").Specific).Value = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_1").Cells.Item(pVal.Row).Specific).Value;
                            ((SAPbouiCOM.EditText)_Form.Items.Item("34").Specific).Value = (Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("47").Specific).Value) - Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("32").Specific).Value)).ToString();
                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific).Value = Utilities.Application.Company.UserName;
                            _Matrix.Columns.Item("V_5").Editable = true;
                            _Matrix.Columns.Item("V_4").Editable = true;
                            
                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_5").Cells.Item(pVal.Row).Specific).Value = DateTime.Now.ToString("yyyyMMdd");
                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific).Value = DateTime.Now.ToString("HH:mm tt");
                            _Matrix.Columns.Item("V_0").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            _Matrix.Columns.Item("V_5").Editable = false;
                            _Matrix.Columns.Item("V_4").Editable = false;
                        }
                        else if (pVal.ColUID == "V_1" && pVal.CharPressed == 9 && ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific).Value == "Gross Weight")
                        {
                            ((SAPbouiCOM.EditText)_Form.Items.Item("47").Specific).Value = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_1").Cells.Item(pVal.Row).Specific).Value;
                            ((SAPbouiCOM.EditText)_Form.Items.Item("34").Specific).Value = (Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("47").Specific).Value) - Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("32").Specific).Value)).ToString();
                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific).Value = Utilities.Application.Company.UserName;
                            _Matrix.Columns.Item("V_5").Editable = true;
                            _Matrix.Columns.Item("V_4").Editable = true;
                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_5").Cells.Item(pVal.Row).Specific).Value = DateTime.Now.ToString("yyyyMMdd");
                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific).Value = DateTime.Now.ToString("HH:mm tt");
                            _Matrix.Columns.Item("V_0").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            _Matrix.Columns.Item("V_5").Editable = false;
                            _Matrix.Columns.Item("V_4").Editable = false;

                        }

                        break;

                    #endregion

                    #region et_ITEM_PRESSED

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && _Form.Mode == BoFormMode.fm_OK_MODE)
                        {
                            _Form.Items.Item("22").Enabled = false;

                        }
                        if (pVal.ItemUID == "45" && (_Form.Mode == BoFormMode.fm_OK_MODE || _Form.Mode == BoFormMode.fm_UPDATE_MODE))
                        {
                            _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("1000002").Specific;
                            _ComboBox.Select(2, BoSearchKey.psk_Index);

                            ((SAPbouiCOM.EditText)_Form.Items.Item("1000004").Specific).Value = Utilities.Application.Company.UserName;

                            if (_Form.Mode == BoFormMode.fm_OK_MODE)
                                _Form.Mode = BoFormMode.fm_UPDATE_MODE;

                            _Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            _Form.Mode = BoFormMode.fm_VIEW_MODE;

                        }
                        if (pVal.ItemUID == "63" && (_Form.Mode == BoFormMode.fm_OK_MODE || _Form.Mode == BoFormMode.fm_UPDATE_MODE))
                        {
                            _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("1000002").Specific;
                            _ComboBox.Select(3, BoSearchKey.psk_Index);

                            ((SAPbouiCOM.EditText)_Form.Items.Item("1000004").Specific).Value = Utilities.Application.Company.UserName;

                            if (_Form.Mode == BoFormMode.fm_OK_MODE)
                                _Form.Mode = BoFormMode.fm_UPDATE_MODE;

                            _Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            _Form.Mode = BoFormMode.fm_VIEW_MODE;

                        }
                        if(pVal.ItemUID == "53")
                        {
                            if (((SAPbouiCOM.ComboBox)_Form.Items.Item("75").Specific).Selected.Value == "Internal Weighing")
                            {
                                PrintReport();
                            }
                            else if (((SAPbouiCOM.ComboBox)_Form.Items.Item("75").Specific).Selected.Value == "External Weighing")
                            {
                                PrintReportExternal();
                            }
                        }

                        break;

                    #endregion

                    #region Set CFL Values

                    case BoEventTypes.et_CHOOSE_FROM_LIST:
                        _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                        if (pVal.ItemUID == "52" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _MasterDBDataSource.SetValue("U_CNTOOrgn", 0, _SysCFLEvent.SelectedObjects.GetValue("Name", 0).ToString().Trim());

                        }
                        break;

                    #endregion

                    #region et_COMBO_SELECT

                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        if (pVal.ItemUID == "1000001" && pVal.Action_Success == true)
                        {
                            if (((SAPbouiCOM.ComboBox)_Form.Items.Item("1000001").Specific).Selected.Value == "Product Out")
                            {
                                ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_2").Cells.Item(1).Specific).Select(0, BoSearchKey.psk_Index);
                                ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_2").Cells.Item(2).Specific).Select(1, BoSearchKey.psk_Index);
                                _Form.Items.Item("55").Enabled = true;
                            }
                            else if (((SAPbouiCOM.ComboBox)_Form.Items.Item("1000001").Specific).Selected.Value == "Product In")
                            {
                                ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_2").Cells.Item(1).Specific).Select(1, BoSearchKey.psk_Index);
                                ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_2").Cells.Item(2).Specific).Select(0, BoSearchKey.psk_Index);
                                _Form.Items.Item("55").Enabled = false;
                            }
                        }
                        if (pVal.ItemUID == "75" && pVal.Action_Success == true)
                        {
                            if (((SAPbouiCOM.ComboBox)_Form.Items.Item("75").Specific).Selected.Value == "Internal Weighing")
                            {
                                _Form.Items.Item("76").Visible = false;
                                _Form.Items.Item("77").Visible = false;
                            }
                            else if (((SAPbouiCOM.ComboBox)_Form.Items.Item("75").Specific).Selected.Value == "External Weighing")
                            {
                                _Form.Items.Item("76").Visible = true;
                                _Form.Items.Item("77").Visible = true;
                            }
                        }

                        break;

                    #endregion

                }
            }
            #endregion
        }

        #endregion

        #region FormDefault

        public void FormDefault()
        {
            _MasterDBDataSource = Form.DataSources.DBDataSources.Item("@VHWTMGTH");
            _ChildDBDataSource = Form.DataSources.DBDataSources.Item("@VHWTMGTC");
            ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value = Utilities.getMaxColumnValueNum("@VHWTMGTH", "U_DocNum");
            _Form.DataBrowser.BrowseBy = "22";
            
            //_ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("46").Specific;
            //_ComboBox.ValidValues.Add("Truck", "Truck");
            //_ComboBox.ValidValues.Add("Mini Truck", "Mini Truck");
            //_ComboBox.ValidValues.Add("Tractor", "Tractor");
            //_ComboBox.Select(0, BoSearchKey.psk_Index);

            ((SAPbouiCOM.EditText)_Form.Items.Item("20").Specific).Value = Utilities.Application.Company.UserName;
            _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("1000002").Specific;
            _ComboBox.ValidValues.Add("Open", "Open");
            _ComboBox.ValidValues.Add("Draft", "Draft");
            _ComboBox.ValidValues.Add("Closed", "Closed");
            _ComboBox.ValidValues.Add("Cancelled", "Cancelled");
            _ComboBox.Select(0, BoSearchKey.psk_Index);

            _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("75").Specific;
            _ComboBox.ValidValues.Add("Internal Weighing", "Internal Weighing");
            _ComboBox.ValidValues.Add("External Weighing", "External Weighing");
            _ComboBox.Select(0, BoSearchKey.psk_Index);

            _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("1000001").Specific;
            _ComboBox.ValidValues.Add("-", "-");
            _ComboBox.ValidValues.Add("Product Out", "Product Out");
            _ComboBox.ValidValues.Add("Product In", "Product In");
            _ComboBox.Select(0, BoSearchKey.psk_Index);
            ((SAPbouiCOM.EditText)_Form.Items.Item("24").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
            _Form.Items.Item("22").Enabled = false;
            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("41").Specific;
         
            //_Matrix.Columns.Item("V_1").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            _Matrix.AddRow(1, -1);
            _Matrix.AddRow(1, 1);
            SAPbouiCOM.Column oColumn;

            oColumn = _Matrix.Columns.Item("V_3");
            oColumn.ValidValues.Add("In Ward", "In Ward");
            oColumn.ValidValues.Add("Out Ward", "Out Ward");
            ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_3").Cells.Item(1).Specific).Select(0, BoSearchKey.psk_Index);
            ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_3").Cells.Item(2).Specific).Select(1, BoSearchKey.psk_Index);

            oColumn = _Matrix.Columns.Item("V_2");
            oColumn.ValidValues.Add("Tare Weight", "Tare Weight");
            oColumn.ValidValues.Add("Gross Weight", "Gross Weight");

            _Form.Items.Item("18").Enabled = false;
            _Form.Items.Item("20").Enabled = false;
            _Form.Items.Item("1000004").Enabled = false;
            _Form.Items.Item("1000002").Enabled = false;
            _Form.Items.Item("32").Enabled = false;
            _Form.Items.Item("47").Enabled = false;
            _Form.Items.Item("34").Enabled = false;
            _Form.Items.Item("36").Enabled = false;
            _Form.Items.Item("38").Enabled = false;
            _Form.Items.Item("40").Enabled = false;
            _Form.Items.Item("1000006").Enabled = false;
            _Form.Items.Item("49").Enabled = false;
            _Form.Items.Item("55").Enabled = false;
            _Form.Items.Item("67").Enabled = false;
            _Form.Items.Item("69").Enabled = false;
            _Matrix.Columns.Item("V_1").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            _Matrix.Columns.Item("V_3").Editable = false;
            _Matrix.Columns.Item("V_2").Editable = false;
            _Matrix.Columns.Item("V_7").Editable = false;
            _Matrix.Columns.Item("V_5").Editable = false;
            _Matrix.Columns.Item("V_4").Editable = false;

            _Form.Items.Item("76").Visible = false;
            _Form.Items.Item("77").Visible = false;

            _Form.Update();

        }

        public void FormDefaultReload()
        {
            ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value = Utilities.getMaxColumnValueNum("@VHWTMGTH", "U_DocNum");
            _Form.DataBrowser.BrowseBy = "22";
            ((SAPbouiCOM.EditText)_Form.Items.Item("20").Specific).Value = Utilities.Application.Company.UserName;
            _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("1000002").Specific;
            _ComboBox.Select(0, BoSearchKey.psk_Index);
            _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("75").Specific;
            _ComboBox.Select(0, BoSearchKey.psk_Index);
            _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("1000001").Specific;
            _ComboBox.Select(0, BoSearchKey.psk_Index);
            ((SAPbouiCOM.EditText)_Form.Items.Item("24").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
            _Form.Items.Item("22").Enabled = false;
            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("41").Specific;
            _Matrix.AddRow(1, -1);
            _Matrix.AddRow(1, 1);
            SAPbouiCOM.Column oColumn;
            oColumn = _Matrix.Columns.Item("V_3");
            ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_3").Cells.Item(1).Specific).Select(0, BoSearchKey.psk_Index);
            ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_3").Cells.Item(2).Specific).Select(1, BoSearchKey.psk_Index);
            oColumn = _Matrix.Columns.Item("V_2");
            _Form.Items.Item("18").Enabled = false;
            _Form.Items.Item("20").Enabled = false;
            _Form.Items.Item("1000004").Enabled = false;
            _Form.Items.Item("1000002").Enabled = false;
            _Form.Items.Item("32").Enabled = false;
            _Form.Items.Item("47").Enabled = false;
            _Form.Items.Item("34").Enabled = false;
            _Form.Items.Item("36").Enabled = false;
            _Form.Items.Item("38").Enabled = false;
            _Form.Items.Item("40").Enabled = false;
            _Form.Items.Item("1000006").Enabled = false;
            _Form.Items.Item("49").Enabled = false;
            _Form.Items.Item("55").Enabled = false;
            _Form.Items.Item("67").Enabled = false;
            _Form.Items.Item("69").Enabled = false;
            _Matrix.Columns.Item("V_1").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            _Matrix.Columns.Item("V_3").Editable = false;
            _Matrix.Columns.Item("V_2").Editable = false;
            _Matrix.Columns.Item("V_7").Editable = false;
            _Matrix.Columns.Item("V_5").Editable = false;
            _Matrix.Columns.Item("V_4").Editable = false;
            _Form.Items.Item("76").Visible = false;
            _Form.Items.Item("77").Visible = false;
            _Form.Update();
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

                        //case "1284":
                        //    {
                        //        _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("1000002").Specific;
                        //        _ComboBox.Select(2, BoSearchKey.psk_Index);

                        //        ((SAPbouiCOM.EditText)_Form.Items.Item("1000004").Specific).Value = Utilities.Application.Company.UserName;

                        //        if (_Form.Mode == BoFormMode.fm_OK_MODE)
                        //            _Form.Mode = BoFormMode.fm_UPDATE_MODE;

                        //        _Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        //        _Form.Mode = BoFormMode.fm_VIEW_MODE;
                        //    }
                        //    break;

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
                                if (_Form.Mode != BoFormMode.fm_FIND_MODE)
                                {
                                    if (_Form.DataSources.DBDataSources.Item("@VHWTMGTH").GetValue("U_Status", 0) == "Closed")
                                    {
                                        ControlsDisable();
                                    }
                                    else
                                    {
                                        _Form.Mode = BoFormMode.fm_OK_MODE;
                                        ControlsEnable();
                                    }
                                }
                                _Form.Items.Item("22").Enabled = true;

                            }
                            break;
                        case Constants.System_Menus.mnu_ADD:
                            {
                                if (_Form.Mode != BoFormMode.fm_ADD_MODE)
                                {
                                    if (_Form.DataSources.DBDataSources.Item("@VHWTMGTH").GetValue("U_Status", 0) == "Closed")
                                    {
                                        ControlsDisable();
                                    }
                                    else
                                    {
                                        _Form.Mode = BoFormMode.fm_OK_MODE;
                                        ControlsEnable();
                                    }
                                  
                                    _Form.Items.Item("22").Enabled = false;
                                }
                                FormDefaultReload();
                                _Form.Items.Item("22").Enabled = false;
                            }
                            break;
                        case Constants.System_Menus.mnu_FIRST:
                            {
                                if (_Form.DataSources.DBDataSources.Item("@VHWTMGTH").GetValue("U_Status", 0) == "Closed")
                                {
                                    ControlsDisable();
                                }
                                else
                                {
                                    _Form.Mode = BoFormMode.fm_OK_MODE;
                                    ControlsEnable();
                                }

                                _Form.Items.Item("22").Enabled = false;
                            }

                            break;
                        case Constants.System_Menus.mnu_LAST:
                            {
                                if (_Form.DataSources.DBDataSources.Item("@VHWTMGTH").GetValue("U_Status", 0) == "Closed")
                                {
                                    ControlsDisable();
                                }
                                else
                                {
                                    _Form.Mode = BoFormMode.fm_OK_MODE;
                                    ControlsEnable();
                                }

                                _Form.Items.Item("22").Enabled = false;
                            }
                            break;
                        case Constants.System_Menus.mnu_NEXT:
                            {
                                if (_Form.DataSources.DBDataSources.Item("@VHWTMGTH").GetValue("U_Status", 0) == "Closed")
                                {
                                    ControlsDisable();
                                }
                                else
                                {
                                    _Form.Mode = BoFormMode.fm_OK_MODE;
                                    ControlsEnable();
                                }

                                _Form.Items.Item("22").Enabled = false;
                               

                            }
                            break;
                        case Constants.System_Menus.mnu_PREVIOUS:
                            {
                                if (_Form.DataSources.DBDataSources.Item("@VHWTMGTH").GetValue("U_Status", 0) == "Closed")
                                {
                                    ControlsDisable();
                                }
                                else
                                {
                                    _Form.Mode = BoFormMode.fm_OK_MODE;
                                    ControlsEnable();
                                }
                               
                                _Form.Items.Item("22").Enabled = false;

                            }
                            break;

                        //case "1284":
                        //    {
                        //        _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("1000002").Specific;
                        //        _ComboBox.Select(3, BoSearchKey.psk_Index);

                        //        if (_Form.Mode != BoFormMode.fm_UPDATE_MODE)
                        //            _Form.Mode = BoFormMode.fm_UPDATE_MODE;

                        //        _Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        //        _Form.Mode = BoFormMode.fm_VIEW_MODE;
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

        #region Weight

        public void Weight(int RNo)
        {
            SerialPort mySerialPort = new SerialPort("COM3");
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
            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_1").Cells.Item(RowNo).Specific).Value = strValue1.ToString();
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

        #region Check Data in table

        private bool DATAEXIST()
        {
            try
            {
                string SQL = "Select Count(*)  from \"@VHWTMGTH\" where  \"U_Status\" = 'Open' and \"U_VehNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("6").Specific).Value.ToString() + "'";
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

        #region Enable Disable Control

        public void ControlsEnable()
        {
            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("41").Specific;
            _Form.Items.Item("52").Enabled = true;
            _Form.Items.Item("46").Enabled = true;
            _Form.Items.Item("6").Enabled = true;
            _Form.Items.Item("8").Enabled = true;
            _Form.Items.Item("10").Enabled = true;
            _Form.Items.Item("12").Enabled = true;
            _Form.Items.Item("14").Enabled = true;
            _Form.Items.Item("16").Enabled = true;
            _Form.Items.Item("24").Enabled = true;
            _Form.Items.Item("1000001").Enabled = true;
            _Form.Items.Item("1").Enabled = true;
            _Form.Items.Item("44").Enabled = true;
            _Form.Items.Item("45").Enabled = true;
            _Matrix.Columns.Item("V_1").Editable = true;
            _Matrix.Columns.Item("V_0").Editable = true;
            _Form.Update();
           
        }
        public void ControlsDisable()
        {
            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("41").Specific;
            _Form.Items.Item("18").Enabled = false;
            _Form.Items.Item("20").Enabled = false;
            _Form.Items.Item("1000004").Enabled = false;
            _Form.Items.Item("1000002").Enabled = false;
            _Form.Items.Item("32").Enabled = false;
            _Form.Items.Item("47").Enabled = false;
            _Form.Items.Item("34").Enabled = false;
            _Form.Items.Item("36").Enabled = false;
            _Form.Items.Item("38").Enabled = false;
            _Form.Items.Item("40").Enabled = false;
            _Form.Items.Item("1000006").Enabled = false;
            _Form.Items.Item("49").Enabled = false;
            _Matrix.Columns.Item("V_3").Editable = false;
            _Matrix.Columns.Item("V_2").Editable = false;
            _Matrix.Columns.Item("V_7").Editable = false;
            _Matrix.Columns.Item("V_5").Editable = false;
            _Matrix.Columns.Item("V_4").Editable = false;
            _Form.Items.Item("52").Enabled = false;
            _Form.Items.Item("46").Enabled = false;
            _Form.Items.Item("6").Enabled = false;
            _Form.Items.Item("8").Enabled = false;
            _Form.Items.Item("10").Enabled = false;
            _Form.Items.Item("12").Enabled = false;
            _Form.Items.Item("14").Enabled = false;
            _Form.Items.Item("16").Enabled = false;
            _Form.Items.Item("24").Enabled = false;
            _Form.Items.Item("1000001").Enabled = false;
            _Form.Items.Item("1").Enabled = false;
            _Form.Items.Item("44").Enabled = false;
            _Form.Items.Item("45").Enabled = false;
            _Matrix.Columns.Item("V_1").Editable = false;
            _Matrix.Columns.Item("V_0").Editable = false;

        }

        #endregion

        #region PrintReport

        private void PrintReport()
        {

            DataSet ds = LoadCompanySetting();
            string DatabaseServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
            string dbUser = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
            string dbPassword = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();
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
                    ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "WeighingSlip.rpt";
                    ReportDocument rpt = new ReportDocument();
                    objCryRep.Load(ReportPath);
                    objCryRep.DataSourceConnections.Clear();
                    string strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY}", dbUser, dbPassword, DatabaseServerName, Database);
                    NameValuePairs2 logonProps2 = objCryRep.DataSourceConnections[0].LogonProperties;
                    logonProps2.Set("Provider", "B1CRHPROXY");
                    logonProps2.Set("Server Type", "B1CRHPROXY");
                    logonProps2.Set("Connection String", strConnection);
                    objCryRep.DataSourceConnections[0].SetLogonProperties(logonProps2);
                    objCryRep.DataSourceConnections[0].SetConnection(DatabaseServerName, Database, false);
                    objCryRep.DataSourceConnections[0].SetLogon(dbUser, dbPassword);
                    int n = objCryRep.Database.Tables.Count;
                    // Set the logon credentials for all tables
                    _DocEntry = _Form.DataSources.DBDataSources.Item("@VHWTMGTH").GetValue("U_DocNum", 0).ToString();
                    ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);


                    // ****************** Check for subreports
                    foreach (CrystalDecisions.CrystalReports.Engine.Section oSection in objCryRep.ReportDefinition.Sections)
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject oRptObj in oSection.ReportObjects)
                        {
                            if (oRptObj.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                            {
                                // This is a subreport so set the logon credentials for this report's tables
                                CrystalDecisions.CrystalReports.Engine.SubreportObject oSubRptObj = oRptObj as CrystalDecisions.CrystalReports.Engine.SubreportObject;
                                // Open the subreport
                                CrystalDecisions.CrystalReports.Engine.ReportDocument oSubRpt = oSubRptObj.OpenSubreport(oSubRptObj.SubreportName);

                                ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);
                            }
                        }
                    }
                    //**************************************
                    objCryRep.Refresh();

                    //objCryRep.VerifyDatabase();

                    //objCryRep.SetParameterValue("Dockey@", "52");
                    objCryRep.SetParameterValue("Dockey@", _DocEntry);
                    //objCryRep.SetParameterValue("OBJECTID@", "13");
                    objCryRep.SetParameterValue("USERNAME@", currentUser);

                    _ImportFileName = _DocEntry.ToString() + ".pdf";

                    //************ Check Invalid  Char in string  ************************
                    string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                    Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
                    _ImportFileName = r.Replace(_ImportFileName, "_");
                    //**********************************************************************

                    // Export Report To Define Application Path --User
                    //string PDFFilePath = SavePDFPath.ToString();

                    //objCryRep.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, PDFFilePath + "\\" + _ImportFileName);

                    objCryRep.PrintToPrinter(1, true, 1, 3);

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

        private void PrintReportExternal()
        {

            DataSet ds = LoadCompanySetting();
            string DatabaseServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
            string dbUser = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
            string dbPassword = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();
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
                    ReportPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\Reports\\" + "WeighingSlipExternal.rpt";
                    ReportDocument rpt = new ReportDocument();
                    objCryRep.Load(ReportPath);
                    objCryRep.DataSourceConnections.Clear();
                    string strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY}", dbUser, dbPassword, DatabaseServerName, Database);
                    NameValuePairs2 logonProps2 = objCryRep.DataSourceConnections[0].LogonProperties;
                    logonProps2.Set("Provider", "B1CRHPROXY");
                    logonProps2.Set("Server Type", "B1CRHPROXY");
                    logonProps2.Set("Connection String", strConnection);
                    objCryRep.DataSourceConnections[0].SetLogonProperties(logonProps2);
                    objCryRep.DataSourceConnections[0].SetConnection(DatabaseServerName, Database, false);
                    objCryRep.DataSourceConnections[0].SetLogon(dbUser, dbPassword);
                    int n = objCryRep.Database.Tables.Count;
                    // Set the logon credentials for all tables
                    _DocEntry = _Form.DataSources.DBDataSources.Item("@VHWTMGTH").GetValue("U_DocNum", 0).ToString();
                    ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);


                    // ****************** Check for subreports
                    foreach (CrystalDecisions.CrystalReports.Engine.Section oSection in objCryRep.ReportDefinition.Sections)
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject oRptObj in oSection.ReportObjects)
                        {
                            if (oRptObj.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                            {
                                // This is a subreport so set the logon credentials for this report's tables
                                CrystalDecisions.CrystalReports.Engine.SubreportObject oSubRptObj = oRptObj as CrystalDecisions.CrystalReports.Engine.SubreportObject;
                                // Open the subreport
                                CrystalDecisions.CrystalReports.Engine.ReportDocument oSubRpt = oSubRptObj.OpenSubreport(oSubRptObj.SubreportName);

                                ClsCRDisplay.SetCrystalTablesLogin(objConInfo, objCryRep.Database.Tables);
                            }
                        }
                    }
                    //**************************************
                    objCryRep.Refresh();

                    //objCryRep.VerifyDatabase();

                    //objCryRep.SetParameterValue("Dockey@", "52");
                    objCryRep.SetParameterValue("Dockey@", _DocEntry);
                    //objCryRep.SetParameterValue("OBJECTID@", "13");
                    objCryRep.SetParameterValue("USERNAME@", currentUser);

                    _ImportFileName = _DocEntry.ToString() + ".pdf";

                    //************ Check Invalid  Char in string  ************************
                    string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                    Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
                    _ImportFileName = r.Replace(_ImportFileName, "_");
                    //**********************************************************************

                    // Export Report To Define Application Path --User
                    //string PDFFilePath = SavePDFPath.ToString();

                    //objCryRep.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, PDFFilePath + "\\" + _ImportFileName);

                    objCryRep.PrintToPrinter(1, true, 1, 3);

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

    }
}
