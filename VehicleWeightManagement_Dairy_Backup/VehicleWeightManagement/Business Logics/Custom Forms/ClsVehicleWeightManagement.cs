using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Threading;
using CrystalDecisions.CrystalReports.ViewerObjectModel;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
//using System.Windows.Forms;
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

namespace VehicleWeightManagement
{
    class ClsVehicleWeightManagement : Base
    {

        #region Variable Declaration

        SAPbouiCOM.ComboBox _ComboBox;
        SAPbouiCOM.Matrix _Matrix, _Matrix1, _Matrix2,_Matrix3, _Matrix4;
        ChooseFromListEvent _SysCFLEvent;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource, _ChildDBDataSource, _ChildDBDataSource1, _ChildDBDataSource2, _ChildDBDataSource3, _ChildDBDataSource4;
        SAPbobsCOM.Recordset _RS = null, _Rs = null, _RSSeries = null, _RSGR = null, _RSGI = null, _RSGo = null, _RSPositive = null,_RSNegative =null;
        public static int RowNo = 0;
        public static int NoOfPeacePer = 0;
        public static double VarianceAllowed = 0;
        public static string ItemGroupName = string.Empty;
        public static int count = 1;
        public string _DBName = string.Empty;
        string _DocEntry = "";
        string ReportPath = string.Empty;
        string _ImportFileName = "";
        Conditions _Cons;
        Condition _Con;
        bool VarianceChange = false;
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
                        if (pVal.ItemUID == "41" && pVal.ColUID == "0")
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
                        if (pVal.ItemUID == "54")
                        {
                            _Form.PaneLevel = 1;
                        }
                        else if (pVal.ItemUID == "55")
                        {
                            _Form.PaneLevel = 2;
                        }
                        else if (pVal.ItemUID == "1000012")
                        {
                            _Form.PaneLevel = 3;
                        }
                        else if(pVal.ItemUID == "1000031")
                        {
                             //ButtonDisable();
                            _Form.PaneLevel = 4;
                        }
                        else if(pVal.ItemUID== "1000032")
                        {
                            //ButtonDisable();
                            _Form.PaneLevel = 5;
                        }
                        
                        if (pVal.ItemUID == "1" && _Form.Mode == BoFormMode.fm_ADD_MODE)
                        {
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
                        if (pVal.ItemUID == "2000" && VarianceChange == true)
                        {
                            Utilities.Message("Please update short and excess part", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            oBubbleEvent = false;
                            return;
                        }
                        if (pVal.ItemUID == "44")
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

                    #region et_CHOOSE_FROM_LIST Code commented

                        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                            if (pVal.ItemUID == "4300" && pVal.ColUID == "V_0")
                           {
                              this.FilteGLAccount("");
                           }
                        //    if (pVal.ItemUID == "1000014" && pVal.ColUID == "V_6")
                        //    {
                        //        this.FilteGLAccount("");
                        //    }
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
                    #region et_LOST_FOCUS

                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                        if (pVal.ItemUID == "6" && pVal.Action_Success == true)
                        {
                            ((SAPbouiCOM.EditText)_Form.Items.Item("6").Specific).Value = ((SAPbouiCOM.EditText)_Form.Items.Item("6").Specific).Value.ToString().Trim().ToUpper().Replace(" ", "");
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
                        _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("56").Specific;
                        if (pVal.ColUID == "Col_7" && pVal.CharPressed == 9)
                        {
                            _Matrix.Columns.Item("Col_15").Editable = true;
                            _Matrix.Columns.Item("Col_14").Editable = true;
                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_15").Cells.Item(pVal.Row).Specific).Value = Utilities.Application.Company.UserName;
                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_14").Cells.Item(pVal.Row).Specific).Value = DateTime.Now.ToString("HH:mm tt");
                            _Matrix.Columns.Item("Col_18").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            _Matrix.Columns.Item("Col_15").Editable = false;
                            _Matrix.Columns.Item("Col_14").Editable = false;
                        }
                        if (pVal.ColUID == "Col_6" && pVal.CharPressed == 9)
                        {
                            _Form.Freeze(true);
                            _Matrix.Columns.Item("Col_17").Editable = true;
                            _Matrix.Columns.Item("Col_16").Editable = true;
                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_17").Cells.Item(pVal.Row).Specific).Value = Utilities.Application.Company.UserName;
                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_16").Cells.Item(pVal.Row).Specific).Value = DateTime.Now.ToString("HH:mm tt");
                            _Matrix.Columns.Item("Col_18").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            _Matrix.Columns.Item("Col_17").Editable = false;
                            _Matrix.Columns.Item("Col_16").Editable = false;
                            //_Matrix.Columns.Item("Col_5").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                             string _Variance= (Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific).Value)-Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_5").Cells.Item(pVal.Row).Specific).Value)).ToString();
                             string RawMilkQty = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_01").Cells.Item(pVal.Row).Specific).Value).ToString();
                            string FinalVariance = (Convert.ToDouble(_Variance) * Convert.ToDouble(RawMilkQty)).ToString();
                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_4").Cells.Item(pVal.Row).Specific).Value = FinalVariance;
                             VarianceChange = true;
                            _Form.Freeze(false);
                        }
                        break;

                    #endregion

                    #region et_ITEM_PRESSED

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && _Form.Mode == BoFormMode.fm_OK_MODE)
                        {
                            _Form.Items.Item("22").Enabled = false;
                            ButtonDisable();
                        }

                        if (pVal.ItemUID == "1000031")
                        {
                            ButtonDisable();
                           
                        }
                        if (pVal.ItemUID == "1000032")
                        {
                            ButtonDisable();
                            
                        }
                        ////if(pVal.ItemUID== "55")
                        ////{
                        ////   ButtonDisable();
                        ////}
                        
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
                        if (pVal.ItemUID == "53")
                        {
                            PrintReport();
                        }
                        if (pVal.ItemUID == "82")
                        {
                            _Form.Freeze(true);

                            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("1000014").Specific;
                            string _Sql = "Select * from (SELECT 'Sales' as \"Type\",T0.\"DocEntry\" \"InvEntry\", T2.\"DocEntry\" \"PayEntry\", T2.\"DocTotal\" \"PayAmount\", T0.\"CardCode\",T0.\"CardName\",T0.\"DocNum\",TO_Date(T0.\"DocDate\") as \"DocDate\", T0.\"DocTotal\", " +
                                        " T2.\"DocNum\" as \"PaymentNo\",T2.\"DocDate\" as \"PaymentDate\",T0.\"NumAtCard\" FROM \"OINV\" T0 LEFT OUTER JOIN \"RCT2\" T1 ON T1.\"DocEntry\" = T0.\"DocEntry\" " +
                                        " LEFT OUTER JOIN \"ORCT\" T2 ON T2.\"DocEntry\" = T1.\"DocNum\" WHERE T0.\"U_VWMNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value + "' " +
                                        " UNION ALL " +
                                        " SELECT 'Purchase' as \"Type\",T0.\"DocEntry\" \"InvEntry\", T2.\"DocEntry\" \"PayEntry\",T2.\"DocTotal\" \"PayAmount\",T0.\"CardCode\",T0.\"CardName\",T0.\"DocNum\",TO_Date(T0.\"DocDate\") as \"DocDate\", T0.\"DocTotal\", T2.\"DocNum\" as \"PaymentNo\",T2.\"DocDate\" as \"PaymentDate\",T0.\"NumAtCard\" " +
                                        " FROM \"OPCH\" T0 LEFT OUTER JOIN \"VPM2\" T1 ON T1.\"DocEntry\" = T0.\"DocEntry\" LEFT OUTER JOIN \"OVPM\" T2 ON T2.\"DocEntry\" = T1.\"DocNum\" WHERE T0.\"U_VWMNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value + "')";
                            Utilities.ExecuteSQL(ref _RS, _Sql);
                            if (_RS.RecordCount > 0)
                            {
                                for (int i = 1; i <= _RS.RecordCount; i++)
                                {
                                    _Matrix.AddRow(1, -1);
                                    DateTime Dt = Convert.ToDateTime(_RS.Fields.Item("DocDate").Value.ToString());
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_6").Cells.Item(i).Specific).Value = _RS.Fields.Item("CardCode").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_5").Cells.Item(i).Specific).Value = _RS.Fields.Item("CardName").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_4").Cells.Item(i).Specific).Value = _RS.Fields.Item("DocNum").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(i).Specific).Value = _RS.Fields.Item("DocDate").Value.ToString("yyyyMMdd");
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_2").Cells.Item(i).Specific).Value = _RS.Fields.Item("DocTotal").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_1").Cells.Item(i).Specific).Value = _RS.Fields.Item("PaymentNo").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value = _RS.Fields.Item("PaymentDate").Value.ToString("yyyyMMdd");
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_11").Cells.Item(i).Specific).Value = _RS.Fields.Item("PayAmount").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value = _RS.Fields.Item("NumAtCard").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_7").Cells.Item(i).Specific).Value = _RS.Fields.Item("Type").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_9").Cells.Item(i).Specific).Value = _RS.Fields.Item("PayEntry").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_10").Cells.Item(i).Specific).Value = _RS.Fields.Item("InvEntry").Value.ToString();
                                    _RS.MoveNext();
                                }
                            }
                            Utilities.SerializedMartix(ref _Form, ref _Matrix);
                            _Form.Mode = BoFormMode.fm_UPDATE_MODE;
                            _Form.Freeze(false);
                        }
                        if (pVal.ItemUID == "90")
                        {
                            _Form.Freeze(true);

                            /////////////

                            string _Sql1 = "select SUM(T1.\"Quantity\") AS \"Quantity\" FROM OWTR T0 INNER JOIN WTR1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\"  " +
                                           " WHERE T1.\"U_TankerCompartment\" = 'FRONT' AND T0.\"U_VWMNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value + "' ";
                            Utilities.ExecuteSQL(ref _RS, _Sql1);
                            ((SAPbouiCOM.EditText)_Form.Items.Item("1000011").Specific).Value = _RS.Fields.Item("Quantity").Value.ToString();

                            string _Sql2 = "select SUM(T1.\"Quantity\") AS \"Quantity\" FROM OWTR T0 INNER JOIN WTR1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\"  " +
                                           " WHERE T1.\"U_TankerCompartment\" = 'MIDDLE' AND T0.\"U_VWMNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value + "' ";
                            Utilities.ExecuteSQL(ref _RS, _Sql2);
                            ((SAPbouiCOM.EditText)_Form.Items.Item("1000015").Specific).Value = _RS.Fields.Item("Quantity").Value.ToString();

                            string _Sql3 = "select SUM(T1.\"Quantity\") AS \"Quantity\" FROM OWTR T0 INNER JOIN WTR1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\"  " +
                                           " WHERE T1.\"U_TankerCompartment\" = 'REAR' AND T0.\"U_VWMNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value + "' ";
                            Utilities.ExecuteSQL(ref _RS, _Sql3);
                            ((SAPbouiCOM.EditText)_Form.Items.Item("74").Specific).Value = _RS.Fields.Item("Quantity").Value.ToString();

                            //////////////

                            string _Sql = string.Empty;
                            //DeleteBlankRow3();
                            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("56").Specific;
                            if (((SAPbouiCOM.ComboBox)_Form.Items.Item("76").Specific).Selected.Value == "1")
                            {
                                _Sql = "select T0.\"DocNum\",T0.\"DocEntry\",T0.\"CardCode\",T0.\"CardName\",T0.\"U_RawMilk\",T1.\"WhsCode\",T1.\"ItemCode\",T1.\"Dscription\",T1.\"Quantity\", " +
                                              " T1.\"Price\",T1.\"LineTotal\",T1.\"U_TankerCompartment\",T0.\"NumAtCard\",case when T1.\"ItemCode\" = 'RM00019' then T0.\"U_Fat\" " +
                                              " when T1.\"ItemCode\" = 'RM00020' then T0.\"U_Snf\" else 0.00 END as \"Standard Per\" " +
                                              " FROM OWTR T0 INNER JOIN WTR1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\"  WHERE  T0.\"BPLId\" = 1 AND T0.\"U_VWMNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value + "' and T0.\"DocEntry\" not in (select \"U_DocEntry\" from \"@VHWTMGTC2\"  where \"U_VWMNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value + "') ";
                            }
                            else
                            {
                                _Sql = "select T0.\"DocNum\",T0.\"DocEntry\",T0.\"CardCode\",T0.\"CardName\",T0.\"U_RawMilk\",T1.\"WhsCode\",T1.\"ItemCode\",T1.\"Dscription\",T1.\"Quantity\", " +
                                              " T1.\"Price\",T1.\"LineTotal\",T1.\"U_TankerCompartment\",T0.\"NumAtCard\",case when T1.\"ItemCode\" = 'BRM00019' then T0.\"U_Fat\" " +
                                              " when T1.\"ItemCode\" = 'BRM00020' then T0.\"U_Snf\" else 0.00 END as \"Standard Per\" " +
                                              " FROM OWTR T0 INNER JOIN WTR1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\"  WHERE  T0.\"BPLId\" = 2 AND T0.\"U_VWMNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value + "' and T0.\"DocEntry\" not in (select \"U_DocEntry\" from \"@VHWTMGTC2\"  where \"U_VWMNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value + "')  ";

                            }
                            Utilities.ExecuteSQL(ref _RS, _Sql);
                            if (_RS.RecordCount > 0)
                            {
                                for (int i = 1; i <= _RS.RecordCount; i++)
                                {
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_12").Cells.Item(_Matrix.RowCount).Specific).Value = _RS.Fields.Item("CardCode").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_11").Cells.Item(_Matrix.RowCount).Specific).Value = _RS.Fields.Item("CardName").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_10").Cells.Item(_Matrix.RowCount).Specific).Value = _RS.Fields.Item("NumAtCard").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_9").Cells.Item(_Matrix.RowCount).Specific).Value = _RS.Fields.Item("ItemCode").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_8").Cells.Item(_Matrix.RowCount).Specific).Value = _RS.Fields.Item("Dscription").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_5").Cells.Item(_Matrix.RowCount).Specific).Value = _RS.Fields.Item("Standard Per").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_7").Cells.Item(_Matrix.RowCount).Specific).Value = _RS.Fields.Item("Quantity").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_01").Cells.Item(_Matrix.RowCount).Specific).Value = _RS.Fields.Item("U_RawMilk").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_25").Cells.Item(_Matrix.RowCount).Specific).Value = _RS.Fields.Item("Price").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_26").Cells.Item(_Matrix.RowCount).Specific).Value = _RS.Fields.Item("LineTotal").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_29").Cells.Item(_Matrix.RowCount).Specific).Value = _RS.Fields.Item("U_TankerCompartment").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_27").Cells.Item(_Matrix.RowCount).Specific).Value = _RS.Fields.Item("DocNum").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_28").Cells.Item(_Matrix.RowCount).Specific).Value = _RS.Fields.Item("DocEntry").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_41").Cells.Item(_Matrix.RowCount).Specific).Value = _RS.Fields.Item("WhsCode").Value.ToString();
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_45").Cells.Item(_Matrix.RowCount).Specific).Value = ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value;
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_15").Cells.Item(_Matrix.RowCount).Specific).Value = Utilities.Application.Company.UserName;
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_14").Cells.Item(_Matrix.RowCount).Specific).Value = DateTime.Now.ToString("HH:mm tt");
                                    
                                    //_Matrix.Columns.Item("Col_15").Editable = true;
                                    //_Matrix.Columns.Item("Col_14").Editable = true;
                                    //((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_15").Cells.Item(i).Specific).Value = Utilities.Application.Company.UserName;
                                    //((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_14").Cells.Item(i).Specific).Value = DateTime.Now.ToString("HH:mm tt");
                                    //_Matrix.Columns.Item("Col_18").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    //_Matrix.Columns.Item("Col_15").Editable = false;
                                    //_Matrix.Columns.Item("Col_14").Editable = false;

                                    _RS.MoveNext();
                                    _Matrix.AddRow(1, -1);
                                }
                            }
                            Utilities.SerializedMartix(ref _Form, ref _Matrix);
                            _Form.Mode = BoFormMode.fm_UPDATE_MODE;

                            _Form.Freeze(false);
                        }
                        //if(pVal.ItemUID=="110")
                        //{
                        //    _Form.Freeze(true);
                        //    DeleteBlankRow4();
                        //    DeleteBlankRow5();
                        //    string sqlNeg = "select  \"U_ItemCode\",\"U_ItemName\",\"U_DocEntry\",\"U_Var\",\"U_Whse\" from \"@VHWTMGTC2\" where \"U_VWMNo\"='" + ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value + "' and \"U_Var\"< 0";
                        //    Utilities.ExecuteSQL(ref _RSNegative, sqlNeg);

                        //    string sqlPos = "select  \"U_ItemCode\",\"U_ItemName\",\"U_DocNum\",\"U_Var\",\"U_Whse\" from \"@VHWTMGTC2\" where \"U_VWMNo\"='" + ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value + "' and \"U_Var\" > 0";
                        //    Utilities.ExecuteSQL(ref _RSPositive, sqlPos);
                        //    if (_RSNegative.RecordCount > 0)
                        //    {
                        //        for (int i = 1; i <= _RSNegative.RecordCount; i++)
                        //        {
                        //            _Matrix1 = _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("4300").Specific;
                        //            _Matrix1.AddRow(1, -1);
                        //            ((SAPbouiCOM.EditText)_Matrix1.Columns.Item("V_5").Cells.Item(i).Specific).Value = _RSNegative.Fields.Item("U_ItemCode").Value.ToString();
                        //            ((SAPbouiCOM.EditText)_Matrix1.Columns.Item("V_4").Cells.Item(i).Specific).Value = _RSNegative.Fields.Item("U_ItemName").Value.ToString();
                        //            ((SAPbouiCOM.EditText)_Matrix1.Columns.Item("V_3").Cells.Item(i).Specific).Value = _RSNegative.Fields.Item("U_DocEntry").Value.ToString();
                        //            ((SAPbouiCOM.EditText)_Matrix1.Columns.Item("V_2").Cells.Item(i).Specific).Value = Convert.ToDouble(_RSNegative.Fields.Item("U_Var").Value).ToString();
                        //            ((SAPbouiCOM.EditText)_Matrix1.Columns.Item("V_1").Cells.Item(i).Specific).Value = _RSNegative.Fields.Item("U_Whse").Value.ToString();
                        //            ((SAPbouiCOM.EditText)_Matrix1.Columns.Item("V_7").Cells.Item(i).Specific).Value = "";

                        //        _RSNegative.MoveNext();
                        //        }
                        //    }
                        //    if (_RSPositive.RecordCount > 0)
                        //    {
                        //        for (int i = 1; i <= _RSPositive.RecordCount; i++)
                        //        {
                        //            _Matrix2 = _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("4200").Specific;
                        //            _Matrix2.AddRow(1, -1);
                        //            ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_5").Cells.Item(i).Specific).Value = _RSPositive.Fields.Item("U_ItemCode").Value.ToString();
                        //            ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_4").Cells.Item(i).Specific).Value = _RSPositive.Fields.Item("U_ItemName").Value.ToString();
                        //            ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_3").Cells.Item(i).Specific).Value = _RSPositive.Fields.Item("U_DocEntry").Value.ToString();
                        //            ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_2").Cells.Item(i).Specific).Value = Convert.ToDouble(_RSPositive.Fields.Item("U_Var").Value).ToString();
                        //            ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_1").Cells.Item(i).Specific).Value = _RSPositive.Fields.Item("U_Whse").Value.ToString();
                        //            ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_7").Cells.Item(i).Specific).Value = "";
                        //            _RSPositive.MoveNext();
                        //        }

                        //    }
                        //    Utilities.SerializedMartix(ref _Form, ref _Matrix1);
                        //    Utilities.SerializedMartix(ref _Form, ref _Matrix2);
                        //    _Form.Mode = BoFormMode.fm_UPDATE_MODE;
                        //    _Form.Freeze(false);
                        //    VarianceChange = false;
                        //}


                        if (pVal.ItemUID == "110")
                        {
                            _Form.Freeze(true);
                            DeleteBlankRow4();
                            DeleteBlankRow5();

                            // Instead of deleting everything, we just prepare the matrices
                            _Matrix1 = (SAPbouiCOM.Matrix)_Form.Items.Item("4300").Specific;
                            _Matrix2 = (SAPbouiCOM.Matrix)_Form.Items.Item("4200").Specific;

                            string vwmNo = ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value;

                            // --- Process Negative Variance (Matrix 4300) ---
                            string sqlNeg = "select \"U_ItemCode\",\"U_ItemName\",\"U_DocEntry\",\"U_Var\",\"U_Whse\" from \"@VHWTMGTC2\" where \"U_VWMNo\"='" + vwmNo + "' and \"U_Var\" < 0";
                            Utilities.ExecuteSQL(ref _RSNegative, sqlNeg);

                            if (_RSNegative.RecordCount > 0)
                            {
                                while (!_RSNegative.EoF)
                                {
                                    string itemCode = _RSNegative.Fields.Item("U_ItemCode").Value.ToString();
                                    string docEntry = _RSNegative.Fields.Item("U_DocEntry").Value.ToString();

                                    // Check if this Item + DocEntry already exists in the matrix
                                    if (!RecordExistsInMatrix(_Matrix1, itemCode, docEntry))
                                    {
                                        _Matrix1.AddRow(1, -1); // Adds to the end
                                        int newRow = _Matrix1.VisualRowCount;

                                        ((SAPbouiCOM.EditText)_Matrix1.Columns.Item("V_5").Cells.Item(newRow).Specific).Value = itemCode;
                                        ((SAPbouiCOM.EditText)_Matrix1.Columns.Item("V_4").Cells.Item(newRow).Specific).Value = _RSNegative.Fields.Item("U_ItemName").Value.ToString();
                                        ((SAPbouiCOM.EditText)_Matrix1.Columns.Item("V_3").Cells.Item(newRow).Specific).Value = docEntry;
                                        ((SAPbouiCOM.EditText)_Matrix1.Columns.Item("V_2").Cells.Item(newRow).Specific).Value = Convert.ToDouble(_RSNegative.Fields.Item("U_Var").Value).ToString();
                                        ((SAPbouiCOM.EditText)_Matrix1.Columns.Item("V_1").Cells.Item(newRow).Specific).Value = _RSNegative.Fields.Item("U_Whse").Value.ToString();
                                        ((SAPbouiCOM.EditText)_Matrix1.Columns.Item("V_7").Cells.Item(newRow).Specific).Value = "";
                                    }
                                    _RSNegative.MoveNext();
                                }
                            }

                            // --- Process Positive Variance (Matrix 4200) ---
                            // Note: Used U_DocNum as per your original SQL for positive, but recommend U_DocEntry for consistency
                            string sqlPos = "select \"U_ItemCode\",\"U_ItemName\",\"U_DocNum\",\"U_Var\",\"U_Whse\" from \"@VHWTMGTC2\" where \"U_VWMNo\"='" + vwmNo + "' and \"U_Var\" > 0";
                            Utilities.ExecuteSQL(ref _RSPositive, sqlPos);

                            if (_RSPositive.RecordCount > 0)
                            {
                                while (!_RSPositive.EoF)
                                {
                                    string itemCode = _RSPositive.Fields.Item("U_ItemCode").Value.ToString();
                                    string docNum = _RSPositive.Fields.Item("U_DocNum").Value.ToString();

                                    if (!RecordExistsInMatrix(_Matrix2, itemCode, docNum))
                                    {
                                        _Matrix2.AddRow(1, -1);
                                        int newRow = _Matrix2.VisualRowCount;

                                        ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_5").Cells.Item(newRow).Specific).Value = itemCode;
                                        ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_4").Cells.Item(newRow).Specific).Value = _RSPositive.Fields.Item("U_ItemName").Value.ToString();
                                        ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_3").Cells.Item(newRow).Specific).Value = docNum;
                                        ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_2").Cells.Item(newRow).Specific).Value = Convert.ToDouble(_RSPositive.Fields.Item("U_Var").Value).ToString();
                                        ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_1").Cells.Item(newRow).Specific).Value = _RSPositive.Fields.Item("U_Whse").Value.ToString();
                                        ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_7").Cells.Item(newRow).Specific).Value = "";
                                    }
                                    _RSPositive.MoveNext();
                                }
                            }

                            Utilities.SerializedMartix(ref _Form, ref _Matrix1);
                            Utilities.SerializedMartix(ref _Form, ref _Matrix2);

                            _Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            _Form.Freeze(false);
                            VarianceChange = false;
                        }

                        if (pVal.ItemUID == "2000")
                        {
                            Utilities.Message("Please wait, operation is in progress...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            bool GRPOstatus = true;
                            bool GIstatus = true;
                            // Retrieve the DocEntry value once and store it in a variable
                            string docEntry = _Form.DataSources.DBDataSources.Item("@VHWTMGTC5").GetValue("DocEntry", 0).Trim();
                            if (string.IsNullOrEmpty(docEntry))
                            {
                                docEntry = "0";
                            }
                            if (docEntry != "0")
                            {
                                // Check if the Goods Receipt (OIGN) exists with the same U_VWMEntry
                                string RSQL = $"SELECT \"U_VWMEntry\" FROM \"OIGN\" WHERE \"U_VWMEntry\" = '{docEntry}'";
                                Utilities.ExecuteSQL(ref _RSGo, RSQL);
                                InventoryGoodsReceipt();
                                //if (_RSGo.RecordCount == 0)
                                //{
                                //    // If no Goods Receipt exists, execute the InventoryGoodsReceipt process
                                //    InventoryGoodsReceipt();
                                //}
                                //else
                                //{
                                //    // Add any necessary logic if the Goods Receipt already exists, or log a message if not needed
                                //    Utilities.Message("Goods Receipt already exists for this entry.", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                //    GRPOstatus = false;
                                //}
                            }
                            else
                            {
                                GRPOstatus = false;
                            }

                            // Check if the Goods Issue (OIGE) exists with the same U_VWMEntry
                            string _docEntry = _Form.DataSources.DBDataSources.Item("@VHWTMGTC4").GetValue("DocEntry", 0).Trim();
                            if (string.IsNullOrEmpty(_docEntry))
                            {
                                _docEntry = "0";
                            }
                            if(_docEntry!="0")
                            {
                                string Sql = $"SELECT \"U_VWMEntry\" AS \"U_VWMEntry\" FROM \"OWTR\" WHERE \"U_VWMEntry\" = '{_docEntry}'";
                                Utilities.ExecuteSQL(ref _RSGI, Sql);

                                InventoryTransfer();

                                //if (_RSGI.RecordCount == 0)
                                //{
                                //    // If no Goods Issue exists, execute the InventoryGoodsIssue process
                                //    InventoryTransfer();
                                //}
                                //else
                                //{
                                //    // Add any necessary logic if the Goods Issue already exists, or log a message if not needed
                                //    Utilities.Message("Inventory Transfer already exists for this entry.", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                //    GIstatus = false;
                                //}
                            }
                            else
                            {
                                GIstatus = false;
                            }
                            //if(GIstatus ==false && GRPOstatus ==false)
                            //{
                            //    _Form.Items.Item("2000").Visible = false;
                            //    _Form.Items.Item("110").Visible = false;
                            //}
                            this.Form.Select();
                            if (Utilities.Application.SBO_Application.Menus.Item("1304").Enabled)
                            {
                                Utilities.Application.SBO_Application.ActivateMenuItem("1304");
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
                        if (pVal.ItemUID == "4200" && pVal.ColUID == "V_1" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("4200").Specific;
                            _Matrix.FlushToDataSource();
                            _ChildDBDataSource4.SetValue("U_Whse", pVal.Row-1, _SysCFLEvent.SelectedObjects.GetValue("WhsCode", 0).ToString().Trim());
                            _Matrix.LoadFromDataSource();
                        }

                        if (pVal.ItemUID == "4300" && pVal.ColUID == "V_0" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("4300").Specific;
                            _Matrix.FlushToDataSource();
                            _ChildDBDataSource3.SetValue("U_ToWhse", pVal.Row - 1, _SysCFLEvent.SelectedObjects.GetValue("WhsCode", 0).ToString().Trim());
                            _Matrix.LoadFromDataSource();
                        }

                        //if (pVal.ItemUID == "56" && pVal.ColUID == "Col_12" && _SysCFLEvent.SelectedObjects != null)
                        //{
                        //    _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("56").Specific;
                        //    _Matrix.FlushToDataSource();
                        //    _ChildDBDataSource1.SetValue("U_BPCode", pVal.Row - 1, _SysCFLEvent.SelectedObjects.GetValue("CardCode", 0).ToString().Trim());
                        //    _ChildDBDataSource1.SetValue("U_BPName", pVal.Row - 1, _SysCFLEvent.SelectedObjects.GetValue("CardName", 0).ToString().Trim());

                        //    _Matrix.LoadFromDataSource();
                        //    _Matrix.Columns.Item("Col_10").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                        //}
                        //if (pVal.ItemUID == "56" && pVal.ColUID == "Col_9" && _SysCFLEvent.SelectedObjects != null)
                        //{
                        //    _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("56").Specific;
                        //    _Matrix.FlushToDataSource();
                        //    _ChildDBDataSource1.SetValue("U_ItemCode", pVal.Row - 1 , _SysCFLEvent.SelectedObjects.GetValue("ItemCode", 0).ToString().Trim());
                        //    _ChildDBDataSource1.SetValue("U_ItemName", pVal.Row - 1, _SysCFLEvent.SelectedObjects.GetValue("ItemName", 0).ToString().Trim());
                        //    _ChildDBDataSource1.SetValue("U_UserName", pVal.Row - 1,  Utilities.Application.Company.UserName);

                        //    _Matrix.LoadFromDataSource();
                        //    if (_SysCFLEvent.SelectedObjects.Rows.Count > 1)
                        //    {
                        //        AddNewRow();
                        //    }
                        //    else
                        //        if (pVal.Row == _Matrix.VisualRowCount)
                        //            AddNewRow();

                        //    Utilities.SerializedMartix(ref _Form, ref _Matrix);

                        //    _Matrix1.Columns.Item("Col_1").Editable = true;
                        //    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific).Value = DateTime.Now.ToString("HH:mm tt");
                        //    _Matrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                        //    _Matrix1.Columns.Item("Col_1").Editable = false;
                        //    _Matrix.Columns.Item("Col_6").Cel ls.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);

                        //}

                        //if (pVal.ItemUID == "1000014" && pVal.ColUID == "V_6" && _SysCFLEvent.SelectedObjects != null)
                        //{
                        //    _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("1000014").Specific;
                        //    _Matrix.FlushToDataSource();
                        //    _ChildDBDataSource2.SetValue("U_BPCode", pVal.Row - 1, _SysCFLEvent.SelectedObjects.GetValue("CardCode", 0).ToString().Trim());
                        //    _ChildDBDataSource2.SetValue("U_BPName", pVal.Row - 1, _SysCFLEvent.SelectedObjects.GetValue("CardName", 0).ToString().Trim());
                        //    _Matrix.LoadFromDataSource();
                        //    if (_SysCFLEvent.SelectedObjects.Rows.Count > 1)
                        //    {
                        //        AddNewRow2();
                        //    }
                        //    else
                        //        if (pVal.Row == _Matrix.VisualRowCount)
                        //            AddNewRow2();

                        //    Utilities.SerializedMartix(ref _Form, ref _Matrix);
                        //}

                        break;

                    #endregion
                        
                    #region et_COMBO_SELECT

                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        if (pVal.ItemUID == "1000001" && pVal.Action_Success == true)
                        {
                            if (((SAPbouiCOM.ComboBox)_Form.Items.Item("1000001").Specific).Selected.Value == "Product Out")
                            {
                                _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("41").Specific;
                                ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_2").Cells.Item(1).Specific).Select(0, BoSearchKey.psk_Index);
                                ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_2").Cells.Item(2).Specific).Select(1, BoSearchKey.psk_Index);
                                _Form.Items.Item("85").Enabled = false;
                            }
                            else if (((SAPbouiCOM.ComboBox)_Form.Items.Item("1000001").Specific).Selected.Value == "Product In")
                            {
                                _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("41").Specific;
                                ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_2").Cells.Item(1).Specific).Select(1, BoSearchKey.psk_Index);
                                ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_2").Cells.Item(2).Specific).Select(0, BoSearchKey.psk_Index);
                                _Form.Items.Item("85").Enabled = true;
                            }
                        }
                        if (pVal.ItemUID == "85" && pVal.Action_Success == true)
                        {
                            if (((SAPbouiCOM.ComboBox)_Form.Items.Item("85").Specific).Selected.Value == "Milk")
                            {
                                _Form.Items.Item("61").Enabled = true;
                            }
                            else
                            {
                                _Form.Items.Item("61").Enabled = false;
                                _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("61").Specific;
                                _ComboBox.Select(0, BoSearchKey.psk_Index);

                            }
                        }
                        if (pVal.ItemUID == "70" && pVal.Action_Success == true)
                        {
                            if (((SAPbouiCOM.ComboBox)_Form.Items.Item("70").Specific).Selected.Value == "Internal")
                            {
                                _Form.Items.Item("77").Visible = false;
                                _Form.Items.Item("78").Visible = false;
                            }
                            else if (((SAPbouiCOM.ComboBox)_Form.Items.Item("70").Specific).Selected.Value == "External")
                            {
                                _Form.Items.Item("77").Visible = true;
                                _Form.Items.Item("78").Visible = true;
                            }
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
                                if (_Form.Mode != BoFormMode.fm_FIND_MODE)
                                {
                                    if (_Form.DataSources.DBDataSources.Item("@VHWTMGTH1").GetValue("U_Status", 0) == "Closed")
                                    {
                                        ControlsDisable();
                                        ButtonDisable();
                                    }
                                    else
                                    {
                                        _Form.Mode = BoFormMode.fm_OK_MODE;
                                        ControlsEnable();
                                        ButtonDisable();
                                    }
                                }
                                _Form.Items.Item("22").Enabled = true;

                                if (((SAPbouiCOM.ComboBox)_Form.Items.Item("70").Specific).Selected.Value == "Internal")
                                {
                                    _Form.Items.Item("77").Visible = false;
                                    _Form.Items.Item("78").Visible = false;
                                }
                                else if (((SAPbouiCOM.ComboBox)_Form.Items.Item("70").Specific).Selected.Value == "External")
                                {
                                    _Form.Items.Item("77").Visible = true;
                                    _Form.Items.Item("78").Visible = true;
                                }



                            }
                            break;
                        case Constants.System_Menus.mnu_ADD:
                            {
                                //if (_Form.Mode != BoFormMode.fm_ADD_MODE)
                                //{
                                //    if (_Form.DataSources.DBDataSources.Item("@VHWTMGTH1").GetValue("U_Status", 0) == "Closed")
                                //    {
                                //        ControlsDisable();
                                //        ButtonDisable();
                                //    }
                                //    else
                                //    {
                                //        _Form.Mode = BoFormMode.fm_OK_MODE;
                                //        ControlsEnable();
                                //        //ButtonDisable();
                                //    }

                                //    _Form.Items.Item("22").Enabled = false;
                                //}
                                //try
                                //{
                                //    if (((SAPbouiCOM.ComboBox)_Form.Items.Item("70").Specific).Selected.Value == "Internal")
                                //    {
                                //        _Form.Items.Item("77").Visible = false;
                                //        _Form.Items.Item("78").Visible = false;
                                //    }
                                //    else if (((SAPbouiCOM.ComboBox)_Form.Items.Item("70").Specific).Selected.Value == "External")
                                //    {
                                //        _Form.Items.Item("77").Visible = true;
                                //        _Form.Items.Item("78").Visible = true;
                                //    }
                                //}
                                //catch {  }

                                if (_Form == null) _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;

                                if (_Form.Mode != BoFormMode.fm_ADD_MODE)
                                {
                                    SAPbouiCOM.DBDataSource dsHeader = _Form.DataSources.DBDataSources.Item("@VHWTMGTH1");
                                    if (dsHeader.Size > 0)
                                    {
                                        if (dsHeader.GetValue("U_Status", 0).Trim() == "Closed")
                                        {
                                            ControlsDisable();
                                            ButtonDisable();
                                        }
                                        else
                                        {
                                            ControlsEnable();
                                        }
                                    }
                                    else
                                    {
                                        ControlsEnable();
                                    }

                                    if (_Form.Items.Item("22") != null)
                                        _Form.Items.Item("22").Enabled = false;
                                }

                                try
                                {
                                    SAPbouiCOM.ComboBox cmbType = (SAPbouiCOM.ComboBox)_Form.Items.Item("70").Specific;
                                    bool isExternal = cmbType.Selected != null && cmbType.Selected.Value == "External";
                                    _Form.Items.Item("77").Visible = isExternal;
                                    _Form.Items.Item("78").Visible = isExternal;
                                }
                                catch { /* Handle cases where selection is empty during mode switch */ }

                            }
                            break;
                        case Constants.System_Menus.mnu_FIRST:
                            {
                                if (_Form.DataSources.DBDataSources.Item("@VHWTMGTH1").GetValue("U_Status", 0) == "Closed")
                                {
                                    //ControlsDisable();
                                    //ButtonDisable();
                                }
                                else
                                {
                                    _Form.Mode = BoFormMode.fm_OK_MODE;
                                    //ControlsEnable();
                                    //ButtonDisable();
                                }

                                _Form.Items.Item("22").Enabled = false;

                                if (((SAPbouiCOM.ComboBox)_Form.Items.Item("70").Specific).Selected.Value == "Internal")
                                {
                                    _Form.Items.Item("77").Visible = false;
                                    _Form.Items.Item("78").Visible = false;
                                }
                                else if (((SAPbouiCOM.ComboBox)_Form.Items.Item("70").Specific).Selected.Value == "External")
                                {
                                    _Form.Items.Item("77").Visible = true;
                                    _Form.Items.Item("78").Visible = true;
                                }

                            }

                            break;
                        case Constants.System_Menus.mnu_LAST:
                            {
                                if (_Form.DataSources.DBDataSources.Item("@VHWTMGTH1").GetValue("U_Status", 0) == "Closed")
                                {
                                    //ControlsDisable();
                                    //ButtonDisable();
                                }
                                else
                                {
                                    //_Form.Mode = BoFormMode.fm_OK_MODE;
                                    //ControlsEnable();
                                    //ButtonDisable();
                                }

                                _Form.Items.Item("22").Enabled = false;

                                if (((SAPbouiCOM.ComboBox)_Form.Items.Item("70").Specific).Selected.Value == "Internal")
                                {
                                    _Form.Items.Item("77").Visible = false;
                                    _Form.Items.Item("78").Visible = false;
                                }
                                else if (((SAPbouiCOM.ComboBox)_Form.Items.Item("70").Specific).Selected.Value == "External")
                                {
                                    _Form.Items.Item("77").Visible = true;
                                    _Form.Items.Item("78").Visible = true;
                                }
                            }
                            break;
                        case Constants.System_Menus.mnu_NEXT:
                            {
                                if (_Form.DataSources.DBDataSources.Item("@VHWTMGTH1").GetValue("U_Status", 0) == "Closed")
                                {
                                    //ControlsDisable();
                                    //ButtonDisable();
                                }
                                else
                                {
                                   // _Form.Mode = BoFormMode.fm_OK_MODE;
                                    //ControlsEnable();
                                    //ButtonDisable();
                                }

                                _Form.Items.Item("22").Enabled = false;

                                //if (((SAPbouiCOM.ComboBox)_Form.Items.Item("70").Specific).Selected.Value == "Internal")
                                //{
                                //    _Form.Items.Item("77").Visible = false;
                                //    _Form.Items.Item("78").Visible = false;
                                //}
                                //else if (((SAPbouiCOM.ComboBox)_Form.Items.Item("70").Specific).Selected.Value == "External")
                                //{
                                //    _Form.Items.Item("77").Visible = true;
                                //    _Form.Items.Item("78").Visible = true;
                                //}
                            }
                            break;
                        case Constants.System_Menus.mnu_PREVIOUS:
                            {
                                if (_Form.DataSources.DBDataSources.Item("@VHWTMGTH1").GetValue("U_Status", 0) == "Closed")
                                {
                                    //ControlsDisable();
                                    //ButtonDisable();
                                }
                                else
                                {
                                    //_Form.Mode = BoFormMode.fm_OK_MODE;
                                    //ControlsEnable();
                                    //ButtonDisable();
                                }

                                _Form.Items.Item("22").Enabled = false;

                                if (((SAPbouiCOM.ComboBox)_Form.Items.Item("70").Specific).Selected.Value == "Internal")
                                {
                                    _Form.Items.Item("77").Visible = false;
                                    _Form.Items.Item("78").Visible = false;
                                }
                                else if (((SAPbouiCOM.ComboBox)_Form.Items.Item("70").Specific).Selected.Value == "External")
                                {
                                    _Form.Items.Item("77").Visible = true;
                                    _Form.Items.Item("78").Visible = true;
                                }
                            }
                            break;

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
                string SQL = "Select Count(*)  from \"@VHWTMGTH1\" where  \"U_Status\" = 'Open' and \"U_VehNo\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("6").Specific).Value.ToString() + "'";
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
            _Matrix1 = (SAPbouiCOM.Matrix)_Form.Items.Item("56").Specific;

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
            _Form.Items.Item("53").Enabled = true;
            _Matrix.Columns.Item("V_1").Editable = true;
            _Matrix.Columns.Item("V_0").Editable = true;
            _Form.Items.Item("74").Enabled = false;
            _Form.Items.Item("76").Enabled = false;
            _Form.Items.Item("1000011").Enabled = false;
            _Form.Items.Item("1000015").Enabled = false;

            //_Matrix1.Columns.Item("Col_9").Editable = true;
            //_Matrix1.Columns.Item("Col_8").Editable = true;
            //_Matrix1.Columns.Item("Col_7").Editable = true;
            //_Matrix1.Columns.Item("Col_6").Editable = true;
            //_Matrix1.Columns.Item("Col_5").Editable = true;
            //_Matrix1.Columns.Item("Col_4").Editable = true;
            //_Matrix1.Columns.Item("Col_3").Editable = true;
            //_Matrix1.Columns.Item("Col_0").Editable = true;

            _Form.Update();

        }
        public void ControlsDisable()
        {
            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("41").Specific;
            _Matrix1 = (SAPbouiCOM.Matrix)_Form.Items.Item("56").Specific;
            _Form.Items.Item("20").Enabled = false;
            _Form.Items.Item("1000004").Enabled = false;
            _Form.Items.Item("1000002").Enabled = false;
            _Form.Items.Item("32").Enabled = false;
            _Form.Items.Item("47").Enabled = false;
            _Form.Items.Item("34").Enabled = false;
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
            _Form.Items.Item("53").Enabled = false;
            _Matrix.Columns.Item("V_1").Editable = false;
            _Matrix.Columns.Item("V_0").Editable = false;
            _Form.Items.Item("74").Enabled = false;
            _Form.Items.Item("76").Enabled = false;
            _Form.Items.Item("1000011").Enabled = false;
            _Form.Items.Item("1000015").Enabled = false;

            //_Matrix1.Columns.Item("Col_9").Editable = false;
            //_Matrix1.Columns.Item("Col_8").Editable = false;
            //_Matrix1.Columns.Item("Col_7").Editable = false;
            //_Matrix1.Columns.Item("Col_6").Editable = false;
            //_Matrix1.Columns.Item("Col_5").Editable = false;
            //_Matrix1.Columns.Item("Col_4").Editable = false;
            //_Matrix1.Columns.Item("Col_3").Editable = false;
            //_Matrix1.Columns.Item("Col_0").Editable = false;

        }

        public void ButtonDisable()
        {
            bool GRPOstatus = true;
            bool GIstatus = true;
            string docEntry = _Form.DataSources.DBDataSources.Item("@VHWTMGTC4").GetValue("DocEntry", 0).Trim();
          
            if (string.IsNullOrEmpty(docEntry))
            {
                docEntry = "0";
            }
            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("56").Specific;
            int rowCount = _Matrix.RowCount;
            bool hasVariance = false;
            for (int i = 1; i <= rowCount; i++) // Matrix rows are 1-based
            {
                // Get the value from the specified column in each row
                string cellValue = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_6").Cells.Item(i).Specific).Value;

                // Check if the cell is not empty
                if (!string.IsNullOrEmpty(cellValue))
                {
                    try
                    {
                        // Convert the cell value to double
                        double cellNumber = Convert.ToDouble(cellValue);

                        // Check if the value is negative
                        if (cellNumber != 0)
                        {
                            hasVariance = true;
                            // Optionally, you can break the loop if you just need to know there's at least one negative value
                            break;
                        }
                    }
                    catch (FormatException ex)
                    {
                        
                    }
                }
            }
            //double variance =Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific).Value
            if (docEntry!="0" )
            {
                string RSQL = $"SELECT \"U_VWMEntry\" FROM \"OWTR\" WHERE \"U_VWMEntry\" = '{docEntry}'";
                Utilities.ExecuteSQL(ref _RSGo, RSQL);
                if (_RSGo.RecordCount != 0)
                {
                    GRPOstatus = false;
                }
                else
                {
                    GRPOstatus = true;
                }
            }
            else if(docEntry!="0" & hasVariance)
            {
                GRPOstatus = true;
            }
            else if(docEntry=="0" && hasVariance==false)
            {
                GRPOstatus = false;
            }
            string _docEntry = _Form.DataSources.DBDataSources.Item("@VHWTMGTC5").GetValue("DocEntry", 0).Trim();
            if (string.IsNullOrEmpty(_docEntry))
            {
                _docEntry = "0";
            }
            if(_docEntry!="0")
            {
                string Sql = $"SELECT \"U_VWMEntry\" AS \"U_VWMEntry\" FROM \"OIGN\" WHERE \"U_VWMEntry\" = '{_docEntry}'";
                Utilities.ExecuteSQL(ref _RSGI, Sql);
                if (_RSGI.RecordCount != 0)
                {
                    GIstatus = false;
                }
                else
                {
                    GIstatus = true;
                }
            }
            else if (_docEntry == "0" & hasVariance)
            {
                GIstatus = true;
            }
            else if (_docEntry == "0" && !hasVariance)
            {
                GIstatus = false;
            }
          
           
            if (GIstatus == false && GRPOstatus == false )
            {
                //_Form.Items.Item("2000").Visible = false;
                //_Form.Items.Item("110").Visible = false;
                _Matrix4.Item.Enabled = false;
                //_Matrix3.Item.Enabled = false;
                _Form.Items.Item("56").Enabled = false;
                _Form.Items.Item("90").Visible = false;
                _Form.Items.Item("88").Enabled = false;
                _Form.Items.Item("92").Enabled = false;
                _Form.Items.Item("70").Enabled = false;
                _Form.Items.Item("74").Enabled = false;
                _Form.Items.Item("76").Enabled = false;
                _Form.Items.Item("1000011").Enabled = false;
                _Form.Items.Item("1000015").Enabled = false;
                _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("41").Specific;
                _Matrix1 = (SAPbouiCOM.Matrix)_Form.Items.Item("56").Specific;
                _Form.Items.Item("20").Enabled = false;
                _Form.Items.Item("1000004").Enabled = false;
                _Form.Items.Item("1000002").Enabled = false;
                _Form.Items.Item("32").Enabled = false;
                _Form.Items.Item("47").Enabled = false;
                _Form.Items.Item("34").Enabled = false;
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
                _Form.Items.Item("90").Visible = false;
                _Form.Items.Item("1000001").Enabled = false;
                _Matrix.Columns.Item("V_1").Editable = false;
                _Matrix.Columns.Item("V_0").Editable = false;
                _Form.Items.Item("74").Enabled = false;
                _Form.Items.Item("76").Enabled = false;
                _Form.Items.Item("1000011").Enabled = false;
                _Form.Items.Item("1000015").Enabled = false;
            }
            else
            {
                _Form.Items.Item("2000").Visible = true;
                //_Form.Items.Item("110").Visible = true;
                _Matrix4.Item.Enabled = true;
                _Matrix3.Item.Enabled = true;
                _Form.Items.Item("56").Enabled = true;
                _Form.Items.Item("90").Visible = true;
                _Form.Items.Item("88").Enabled = true;
                _Form.Items.Item("92").Enabled = true;
                _Form.Items.Item("70").Enabled = true;
                _Form.Items.Item("74").Enabled = true;
                _Form.Items.Item("76").Enabled = true;
                _Form.Items.Item("1000011").Enabled = true;
                _Form.Items.Item("1000015").Enabled = true;
                _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("41").Specific;
                _Matrix1 = (SAPbouiCOM.Matrix)_Form.Items.Item("56").Specific;
                _Form.Items.Item("20").Enabled = true;
                _Form.Items.Item("1000004").Enabled = true;
                _Form.Items.Item("1000002").Enabled = true;
                _Form.Items.Item("32").Enabled = true;
                _Form.Items.Item("47").Enabled = true;
                _Form.Items.Item("34").Enabled = true;
                _Form.Items.Item("1000006").Enabled = true;
                _Form.Items.Item("49").Enabled = true;
                _Matrix.Columns.Item("V_3").Editable = true;
                _Matrix.Columns.Item("V_2").Editable = true;
                _Matrix.Columns.Item("V_7").Editable = true;
                _Matrix.Columns.Item("V_5").Editable = true;
                _Matrix.Columns.Item("V_4").Editable = true;
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
                _Matrix.Columns.Item("V_1").Editable = true;
                _Matrix.Columns.Item("V_0").Editable = true;
                _Form.Items.Item("74").Enabled = true;
                _Form.Items.Item("76").Enabled = true;
                _Form.Items.Item("1000011").Enabled = true;
                _Form.Items.Item("1000015").Enabled = true;
                //_Form.Items.Item("110").Visible = true;
            }
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
                    _DocEntry = _Form.DataSources.DBDataSources.Item("@VHWTMGTH1").GetValue("U_DocNum", 0).ToString();
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
                    objCryRep.SetParameterValue("Dockey@", _DocEntry);
                    objCryRep.SetParameterValue("USERNAME@", currentUser);
                    _ImportFileName = _DocEntry.ToString() + ".pdf";

                    //************ Check Invalid  Char in string  ************************
                    string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                    Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
                    _ImportFileName = r.Replace(_ImportFileName, "_");
                    //**********************************************************************
                    objCryRep.PrintToPrinter(1, true, 1, 25);

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

        #region AddNewRow Method
        public void AddNewRow()
        {
            if (_ChildDBDataSource1 != null && _Matrix != null)
            {
                _ChildDBDataSource1.InsertRecord(_ChildDBDataSource1.Size);
                _Matrix.LoadFromDataSource();
            }
        }
        #endregion

        #region AddNewRow Method

        public void AddNewRow2()
        {
            if (_ChildDBDataSource2 != null && _Matrix != null)
            {
                _ChildDBDataSource2.InsertRecord(_ChildDBDataSource2.Size);
                _Matrix.LoadFromDataSource();
            }
        }

        #endregion

        #region FILTER BP Type
        internal void FilteGLAccount(string CFLId)
        {

            _Form.Freeze(true);
            try
            {
                _Cons = new Conditions();


                _Con = _Cons.Add();
                _Con.BracketOpenNum = 1;
                _Con.Alias = "BPLid";
                _Con.Operation = BoConditionOperation.co_EQUAL;
                _Con.CondVal = "1";
                _Con.BracketCloseNum = 1;
                _Form.ChooseFromLists.Item("ToWhse").SetConditions(_Cons);

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

        #region Delete Blank row1

        private void DeleteBlankRow1()
        {
            try
            {
                int i = 0;
                _Form.Freeze(true);
                _Matrix.FlushToDataSource();
                for (i = 1; i <= _Matrix.VisualRowCount; i++)
                {
                    if (string.IsNullOrEmpty(_ChildDBDataSource1.GetValue("U_BPCode", i - 1).Trim().ToString()) && string.IsNullOrEmpty(_ChildDBDataSource1.GetValue("U_BPCode", i - 1).Trim().ToString()))
                    {
                        _Matrix.DeleteRow(i);
                    }
                }
                Utilities.SerializedMartix(ref _Form, ref _Matrix);
            }
            catch
            {
                Utilities.SerializedMartix(ref _Form, ref _Matrix);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }

        #endregion

        #region Delete Blank row2

        private void DeleteBlankRow2()
        {
            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("1000014").Specific;
            int i = _Matrix.VisualRowCount;
            _Matrix.FlushToDataSource();
            for (int j = 1; j <= i; j++)
            {
                _Matrix.DeleteRow(1);

            }
        }

        #endregion

        #region Delete Blank row3

        private void DeleteBlankRow3()
        {
            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("56").Specific;
            int i = _Matrix.VisualRowCount;
            _Matrix.FlushToDataSource();
            for (int j = 1; j <= i; j++)
            {
                _Matrix.DeleteRow(1);

            }
        }

        #endregion

        #region Delete Blank row4

        //private void DeleteBlankRow4()
        //{
        //    _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("4200").Specific;
        //    int i = _Matrix.VisualRowCount;
        //    _Matrix.FlushToDataSource();
        //    for (int j = 1; j <= i; j++)
        //    {
        //        _Matrix.DeleteRow(1);

        //    }
        //}

        private void DeleteBlankRow4()
        {
            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("4200").Specific;
            // Loop backwards when deleting rows to maintain index integrity
            for (int i = _Matrix.VisualRowCount; i >= 1; i--)
            {
                string itemCode = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_5").Cells.Item(i).Specific).Value;
                if (string.IsNullOrEmpty(itemCode))
                {
                    _Matrix.DeleteRow(i);
                }
            }
        }

        #endregion

        #region Delete Blank row5

        private void DeleteBlankRow5()
        {
            //_Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("4300").Specific;
            //int i = _Matrix.VisualRowCount;
            //_Matrix.FlushToDataSource();
            //for (int j = 1; j <= i; j++)
            //{
            //    _Matrix.DeleteRow(1);

            //}
            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("4300").Specific;
            // Loop backwards when deleting rows to maintain index integrity
            for (int i = _Matrix.VisualRowCount; i >= 1; i--)
            {
                string itemCode = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_5").Cells.Item(i).Specific).Value;
                if (string.IsNullOrEmpty(itemCode))
                {
                    _Matrix.DeleteRow(i);
                }
            }
        }

        #endregion

        #region FormDefault

        public void FormDefault()
        {
            _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@VHWTMGTH1");
            _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@VHWTMGTC1");
            _ChildDBDataSource1 = _Form.DataSources.DBDataSources.Item("@VHWTMGTC2");
            _ChildDBDataSource2 = _Form.DataSources.DBDataSources.Item("@VHWTMGTC3");
            _ChildDBDataSource3 = _Form.DataSources.DBDataSources.Item("@VHWTMGTC4");
            _ChildDBDataSource4 = _Form.DataSources.DBDataSources.Item("@VHWTMGTC5");
            _Form.Items.Item("54").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value = Utilities.getMaxColumnValueNum("@VHWTMGTH1", "U_DocNum");
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
            _ComboBox.Select(0, BoSearchKey.psk_Index);

            _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("70").Specific;
            //_ComboBox.ValidValues.Add("-", "-");
            _ComboBox.ValidValues.Add("Internal", "Internal");
            _ComboBox.ValidValues.Add("External", "External");
            _ComboBox.Select(0, BoSearchKey.psk_Index);

            _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("1000001").Specific;
            _ComboBox.ValidValues.Add("-", "-");
            _ComboBox.ValidValues.Add("Product Out", "Product Out");
            _ComboBox.ValidValues.Add("Product In", "Product In");
            _ComboBox.Select(0, BoSearchKey.psk_Index);

            //_ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("85").Specific;
            //_ComboBox.ValidValues.Add("-", "-");
            //_ComboBox.ValidValues.Add("Milk", "Milk");
            //_ComboBox.ValidValues.Add("Others", "Others");
            //_ComboBox.Select(0, BoSearchKey.psk_Index);

            _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("61").Specific;
            _ComboBox.ValidValues.Add("-", "-");
            _ComboBox.ValidValues.Add("Self Purchase Material", "Self Purchase Material");
            _ComboBox.ValidValues.Add("Ex Party Material", "Ex Party Material");
            _ComboBox.Select(0, BoSearchKey.psk_Index);

            _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("76").Specific;
            _ComboBox.ValidValues.Add("1", "UNIT A");
            _ComboBox.ValidValues.Add("2", "UNIT B");
            _ComboBox.Select(0, BoSearchKey.psk_Index);

            ((SAPbouiCOM.EditText)_Form.Items.Item("24").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
            _Form.Items.Item("22").Enabled = false;

            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("41").Specific;
            _Matrix1 = (SAPbouiCOM.Matrix)_Form.Items.Item("56").Specific;
            _Matrix2 = (SAPbouiCOM.Matrix)_Form.Items.Item("1000014").Specific;
            _Matrix3 = (SAPbouiCOM.Matrix)_Form.Items.Item("4300").Specific;
            _Matrix4 = (SAPbouiCOM.Matrix)_Form.Items.Item("4200").Specific;

            //  _Matrix.Columns.Item("V_1").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
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

            _Form.Items.Item("20").Enabled = false;
            _Form.Items.Item("1000004").Enabled = false;
            _Form.Items.Item("1000002").Enabled = false;
            _Form.Items.Item("32").Enabled = false;
            _Form.Items.Item("47").Enabled = false;
            _Form.Items.Item("34").Enabled = false;

            _Form.Items.Item("1000006").Enabled = false;
            _Form.Items.Item("49").Enabled = false;
            _Matrix.Columns.Item("V_1").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            _Matrix.Columns.Item("V_3").Editable = false;
            _Matrix.Columns.Item("V_2").Editable = false;
            _Matrix.Columns.Item("V_7").Editable = false;
            _Matrix.Columns.Item("V_5").Editable = false;
            _Matrix.Columns.Item("V_4").Editable = false;

            _Matrix3.Columns.Item("V_5").Editable = false;
            _Matrix3.Columns.Item("V_4").Editable = false;
            _Matrix3.Columns.Item("V_3").Editable = false;
            _Matrix3.Columns.Item("V_2").Editable = false;
            _Matrix3.Columns.Item("V_1").Editable = false;
            _Matrix3.Columns.Item("V_7").Editable = false;
            //_Matrix3.Columns.Item("V_0").Editable = true;


            _Matrix4.Columns.Item("V_5").Editable = false;
            _Matrix4.Columns.Item("V_4").Editable = false;
            _Matrix4.Columns.Item("V_3").Editable = false;
            _Matrix4.Columns.Item("V_2").Editable = false;
            _Matrix4.Columns.Item("V_1").Editable = true;
            _Matrix4.Columns.Item("V_0").Editable = true;


            _Form.Items.Item("85").Enabled = false;
            _Form.Items.Item("86").Enabled = false;
            _Form.Items.Item("72").Enabled = false;
            _Form.Items.Item("61").Enabled = false;
            _Form.Items.Item("62").Enabled = false;
            _Form.Items.Item("64").Enabled = false;
            _Form.Items.Item("74").Enabled = false;
            _Form.Items.Item("76").Enabled = false;
            _Form.Items.Item("1000011").Enabled = false;
            _Form.Items.Item("1000015").Enabled = false;

            if (_Matrix1.VisualRowCount == 0)
                _Matrix1.AddRow(1, 0);
            Utilities.SerializedMartix(ref _Form, ref _Matrix1);

            oColumn = _Matrix1.Columns.Item("Col_3");
            oColumn.ValidValues.Add("Yes", "Yes");
            oColumn.ValidValues.Add("No", "No");

            ControlsEnable();
           
            bool GRPOstatus = true;
            bool GIstatus = true;
            string RSQL = $"SELECT \"U_VWMEntry\" FROM \"OIGN\" WHERE \"U_VWMEntry\" = '{_Form.DataSources.DBDataSources.Item("@VHWTMGTC4").GetValue("DocEntry", 0).Trim()}'";
            Utilities.ExecuteSQL(ref _RSGo, RSQL);
            if (_RSGo.RecordCount != 0)
            {
                GRPOstatus = false;
            }
            string Sql = $"SELECT \"U_VWMEntry\" AS \"U_VWMEntry\" FROM \"OIGE\" WHERE \"U_VWMEntry\" = '{_Form.DataSources.DBDataSources.Item("@VHWTMGTC4").GetValue("DocEntry", 0).Trim()}'";
            Utilities.ExecuteSQL(ref _RSGI, Sql);
            if (_RSGI.RecordCount != 0)
            {
                GIstatus = false;
            }
            if (GIstatus == false && GRPOstatus == false)
            {
                //_Form.Items.Item("2000").Visible = false;
                _Form.Items.Item("110").Visible = false;
            }
            _Form.Items.Item("77").Visible = false;
            _Form.Items.Item("78").Visible = false;

            _Form.Update();
            _Form.Items.Item("52").Click(SAPbouiCOM.BoCellClickType.ct_Regular);



        }

        #endregion

        #region InventoryGoodsReceipt

        private void InventoryGoodsReceipt()
        {

            int _lineCount = 0;
            SAPbobsCOM.BoObjectTypes objType = SAPbobsCOM.BoObjectTypes.oInventoryGenEntry;
            SAPbobsCOM.Documents document;
            SAPbobsCOM.Documents _GR = (SAPbobsCOM.Documents)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
            document = (SAPbobsCOM.Documents)Utilities.Application.Company.GetBusinessObject(objType);

            string docEntry = _Form.DataSources.DBDataSources.Item("@VHWTMGTC5").GetValue("DocEntry", 0).ToString().Trim();

            // Check if DocEntry is null or empty and set it to "0" if so
            if (string.IsNullOrEmpty(docEntry))
            {
                docEntry = "0";
            }

            if(docEntry!="0")
            {
                string _QryGRData = "select T0.\"LineId\",T0.\"U_ItemCode\",T0.\"U_ItemName\",T0.\"U_GRPONo\",T0.\"U_Var\",T0.\"U_Whse\",T0.\"U_UnitPrice\",T1.\"U_Branch\", T1.\"DocEntry\",T1.\"U_DocDate\" ,T1.\"DocNum\"" +
                                  " from \"@VHWTMGTC5\"  T0 inner join \"@VHWTMGTH1\" T1 on T0.\"DocEntry\" =T1.\"DocEntry\"  where T0.\"DocEntry\" = '" + docEntry + "' and  ifnull(T0.\"U_GREntry\",'') =''  ";
                Utilities.ExecuteSQL(ref _RS, _QryGRData);
                if(_RS.RecordCount>0)
                {
                    string _QrySeries = "select \"Series\" from NNM1 where \"ObjectCode\" = '59' AND \"BPLId\" = '" + _RS.Fields.Item("U_Branch").Value.ToString() + "' and \"Indicator\" = " +
                                   " (select distinct \"Indicator\" from OFPR where '" + Convert.ToDateTime(_RS.Fields.Item("U_DocDate").Value.ToString()).ToString("yyyyMMdd") + "' between \"F_TaxDate\" and \"T_TaxDate\")";
                    Utilities.ExecuteSQL(ref _RSSeries, _QrySeries);
                    if (_RS.RecordCount > 0)
                    {
                        for (int i = 0; i < _RS.RecordCount; i++)
                        {
                            if (_lineCount == 0)
                            {
                                DateTime parsedPostingDate = DateTime.Parse(_RS.Fields.Item("U_DocDate").Value.ToString());
                                DateTime parsedPDeliveryDate = DateTime.Parse(_RS.Fields.Item("U_DocDate").Value.ToString());
                                int BranchCode = Convert.ToInt32(_RS.Fields.Item("U_Branch").Value.ToString());
                                _GR.DocDate = parsedPostingDate;
                                _GR.DocDueDate = parsedPDeliveryDate;
                                _GR.JournalMemo = "Weighing Automation";
                                _GR.Comments = "Weighing Automation";
                                _GR.BPL_IDAssignedToInvoice = Convert.ToInt16(BranchCode.ToString());
                                _GR.UserFields.Fields.Item("U_VWMNo").Value = _RS.Fields.Item("DocNum").Value.ToString();
                                _GR.UserFields.Fields.Item("U_VWMEntry").Value = _RS.Fields.Item("DocEntry").Value.ToString();
                                //_GR.UserFields.Fields.Item("U_Whse").Value = _RS.Fields.Item("U_Whse").Value.ToString();
                                _GR.Series = Convert.ToInt16(_RSSeries.Fields.Item("Series").Value.ToString());
                                _lineCount++;
                            }
                            else
                                _GR.Lines.Add();
                            _GR.Lines.ItemCode = _RS.Fields.Item("U_ItemCode").Value.ToString();
                            _GR.Lines.WarehouseCode = _RS.Fields.Item("U_Whse").Value.ToString();
                            _GR.Lines.Quantity = double.Parse(_RS.Fields.Item("U_Var").Value.ToString());
                            _GR.Lines.UnitPrice = double.Parse(_RS.Fields.Item("U_UnitPrice").Value.ToString());
                            _GR.Lines.UserFields.Fields.Item("U_VWMLineNo").Value = _RS.Fields.Item("LineId").Value.ToString();
                            _GR.Lines.UserFields.Fields.Item("U_VWMEntry").Value = _RS.Fields.Item("DocEntry").Value.ToString();
                            //_GR.Lines.AccountCode = _RS.Fields.Item("U_InterBranchGLCode").Value.ToString();
                            _RS.MoveNext();
                        }
                    }
                    
                    int RetVal = _GR.Add();
                    if (RetVal == 0)
                    {
                        string objCode = "";
                        Utilities.Application.Company.GetNewObjectCode(out objCode);
                        string StrQry1 = "select \"DocNum\" from OIGN where \"DocEntry\"  = '" + objCode + "'";
                        Utilities.ExecuteSQL(ref _RSGR, StrQry1);

                        string StrQry = "update \"@VHWTMGTC5\" set \"U_GREntry\" = '" + objCode + "' where \"DocEntry\" = '" + _Form.DataSources.DBDataSources.Item("@VHWTMGTH1").GetValue("DocEntry", 0).ToString() + "' and ifnull(\"U_GREntry\",'') ='' ";
                        Utilities.ExecuteSQL(ref _RSGR, StrQry);
                        Utilities.Message("Operation completed sucessfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                    else
                    {
                        string GetLastErrorDescription = Utilities.Application.Company.GetLastErrorDescription();
                        Utilities.Message(GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }

                }
                else
                {
                    Utilities.Message("Goods Receipt already exists.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }

            }
            
        }

        #endregion

        #region InventoryTransfer
        private void InventoryTransfer()
        {

            int _lineCount = 0;
            SAPbobsCOM.StockTransfer _ST;
            _ST = (SAPbobsCOM.StockTransfer)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);

            // Fetch the DocEntry value from the data source
            string docEntry = _Form.DataSources.DBDataSources.Item("@VHWTMGTC4").GetValue("DocEntry", 0).ToString().Trim();

            // Check if DocEntry is null or empty and set it to "0" if so
            if (string.IsNullOrEmpty(docEntry))
            {
                docEntry = "0";
            }
            if(docEntry!= "0")
            {
                // Construct the SQL query with COALESCE to handle NULL values
                string _QryGRData = "SELECT T0.\"LineId\", T0.\"U_ItemCode\", T0.\"U_ItemName\", T0.\"U_GRPONo\", T0.\"U_Var\", T0.\"U_FromWhse\", T0.\"U_ToWhse\", " +
                                    "T1.\"U_Branch\", T1.\"DocEntry\", T1.\"U_DocDate\", T1.\"DocNum\" " +
                                    "FROM \"@VHWTMGTC4\" T0 " +
                                    "INNER JOIN \"@VHWTMGTH1\" T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                                    "WHERE COALESCE(T0.\"DocEntry\", '0') = '" + docEntry + "' and  ifnull(T0.\"U_INTEntry\",'') ='' ";

                // Execute the SQL query
                Utilities.ExecuteSQL(ref _RS, _QryGRData);
                if(_RS.RecordCount > 0)
                {
                    string _QrySeries = "select \"Series\" from NNM1 where \"ObjectCode\" = '67' AND \"BPLId\" = '" + _RS.Fields.Item("U_Branch").Value.ToString() + "' and \"Indicator\" = " +
                                     " (select distinct \"Indicator\" from OFPR where '" + Convert.ToDateTime(_RS.Fields.Item("U_DocDate").Value.ToString()).ToString("yyyyMMdd") + "' between \"F_TaxDate\" and \"T_TaxDate\")";
                    Utilities.ExecuteSQL(ref _RSSeries, _QrySeries);
                    if (_RS.RecordCount > 0)
                    {
                        for (int i = 0; i < _RS.RecordCount; i++)
                        {
                            if (_lineCount == 0)
                            {
                                DateTime parsedPostingDate = DateTime.Parse(_RS.Fields.Item("U_DocDate").Value.ToString());
                                //DateTime parsedPostingDate = DateTime.Parse("1/13/2026 12:00:00 AM"); 
                                DateTime parsedPDeliveryDate = DateTime.Parse(_RS.Fields.Item("U_DocDate").Value.ToString());
                                int BranchCode = Convert.ToInt32(_RS.Fields.Item("U_Branch").Value.ToString());
                                _ST.DocDate = parsedPostingDate;
                                _ST.Comments = "Weighing Automation";
                                _ST.ToWarehouse = _RS.Fields.Item("U_ToWhse").Value.ToString();
                                _ST.FromWarehouse = _RS.Fields.Item("U_FromWhse").Value.ToString();
                                _ST.UserFields.Fields.Item("U_VWMNo").Value = _RS.Fields.Item("DocNum").Value.ToString();
                                _ST.UserFields.Fields.Item("U_VWMEntry").Value = _RS.Fields.Item("DocEntry").Value.ToString();
                                //_GR.UserFields.Fields.Item("U_Whse").Value = _RS.Fields.Item("U_Whse").Value.ToString();
                                _ST.Series = Convert.ToInt16(_RSSeries.Fields.Item("Series").Value.ToString());
                                _lineCount++;
                            }
                            else
                                _ST.Lines.Add();
                            _ST.Lines.ItemCode = _RS.Fields.Item("U_ItemCode").Value.ToString();
                            _ST.Lines.WarehouseCode = _RS.Fields.Item("U_ToWhse").Value.ToString();
                            _ST.Lines.FromWarehouseCode = _RS.Fields.Item("U_FromWhse").Value.ToString();
                            double number = double.Parse(_RS.Fields.Item("U_Var").Value.ToString());
                            double positiveQty = Math.Abs(number);
                            _ST.Lines.Quantity = positiveQty;
                            _ST.Lines.UserFields.Fields.Item("U_VWMLineNo").Value = _RS.Fields.Item("LineId").Value.ToString();
                            _ST.Lines.UserFields.Fields.Item("U_VWMEntry").Value = _RS.Fields.Item("DocEntry").Value.ToString();
                            _RS.MoveNext();
                        }
                    }
                   
                    int RetVal = _ST.Add();
                    if (RetVal == 0)
                    {
                        string objCode = "";
                        Utilities.Application.Company.GetNewObjectCode(out objCode);
                        string StrQry1 = "select \"DocNum\" from OWTR where \"DocEntry\"  = '" + objCode + "'";
                        Utilities.ExecuteSQL(ref _RSGR, StrQry1);

                        string StrQry = "update \"@VHWTMGTC4\" set \"U_INTEntry\" = '" + objCode + "' where \"DocEntry\" = '" + _Form.DataSources.DBDataSources.Item("@VHWTMGTH1").GetValue("DocEntry", 0).ToString() + "' and ifnull(\"U_INTEntry\",'') =''";
                        Utilities.ExecuteSQL(ref _RSGR, StrQry);
                        Utilities.Message("Inventory transfer created sucessfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                    else
                    {
                        string GetLastErrorDescription = Utilities.Application.Company.GetLastErrorDescription();
                        Utilities.Message(GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                else
                {
                    Utilities.Message("Inventory Transfer already exists.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }

            }
          
        }
        #endregion


        private bool RecordExistsInMatrix(SAPbouiCOM.Matrix mtx, string itemCode, string docEntry)
        {
            for (int i = 1; i <= mtx.VisualRowCount; i++)
            {
                // V_5 = Item Code, V_3 = DocEntry/DocNum
                string existingItem = ((SAPbouiCOM.EditText)mtx.Columns.Item("V_5").Cells.Item(i).Specific).Value.Trim();
                string existingDoc = ((SAPbouiCOM.EditText)mtx.Columns.Item("V_3").Cells.Item(i).Specific).Value.Trim();

                if (existingItem == itemCode.Trim() && existingDoc == docEntry.Trim())
                {
                    return true; // Match found, skip this record
                }
            }
            return false; // No match found
        }

    }
}
