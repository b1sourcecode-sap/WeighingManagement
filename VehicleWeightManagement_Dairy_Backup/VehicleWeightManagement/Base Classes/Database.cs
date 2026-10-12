using System;
using System.Data;
using System.Data.SqlClient;
using Microsoft.ApplicationBlocks.Data;

namespace VehicleWeightManagement
{
	/// <summary>
	/// Summary description for Database.
	/// </summary>
	public sealed class Database
	{
		private Database() 
		{ 
		} 

		#region INITIALISE DATABASE
		
        public static void InitializeDatabase()
		{
			SAPbouiCOM.ProgressBar oProgressBar = null;
			string[] oChildTables = new string[3];
			string[] oFindColumns = new string[3];
			string[] oColumnName  = new string[2];
			try
			{
				oProgressBar = Utilities.Application.SBO_Application.StatusBar.CreateProgressBar("Initializing database...Please wait.", 50, false);
				oProgressBar.Value = 1;

				if( !Utilities.Application.Company.InTransaction )
					Utilities.Application.Company.StartTransaction();

                #region OPKL UDF

                AddColumn("OINV", "TruckNo", "Truck No", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OINV", "VWMNo", "Vehcle Management No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                #endregion
                #region Goods Receipt/Goods Issue
                AddColumn("OIGN", "VWMEntry", "Vehicle Management Entry", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OIGN", "VWMNo", "Vehcle Management No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OIGE", "VWMEntry", "Vehicle Management Entry", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OIGE", "VWMNo", "Vehcle Management No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OWTR", "VWMEntry", "Vehicle Management Entry", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OWTR", "VWMNo", "Vehcle Management No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("IGN1", "VWMEntry", "Vehicle Management Entry", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("IGN1", "VWMLineNo", "Vehcle Management Line No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OIGE", "VWMEntry", "Vehicle Management Entry", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("IGE1", "VWMLineNo", "Vehcle Management Line No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OWTR", "VWMEntry", "Vehicle Management Entry", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("WTR1", "VWMLineNo", "Vehcle ManagementLine No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                #endregion

                #region Vehicle Management Header

                CreateTable("VHWTMGTH1", "Vehicle Management Header1", SAPbobsCOM.BoUTBTableType.bott_Document);
                AddColumn("VHWTMGTH1", "CNTOOrgn", "Country Of Origin", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "VehType", "Vehicle Type", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "VehNo", "Vehicle No", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "TranspName", "Transporter Name", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "TranspPAN", "Transporter PAN", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "DriverNm", "Driver Name", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "DriverNo", "Driver No", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "LicenseNo", "License No", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "LdContNo", "Loading Counter No", 5, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "UserName", "User Name", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "ClosedBy", "Closed By", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "DocNum", "Document No", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "DocDate", "Document Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "Status", "Status", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "VisitPrPs", "Visit Purpose", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "TareWt", "Tare Weight", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "NetWt", "Net Weight", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "WBDocNum", "Weighing Bridge Document No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "WBCounter", " Weight As Per Dispatch Counter", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "NTVarWt", "Net Vaiance in Weight", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "GrossWt", "Gross Weight", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "LodWt", "Weight to be loaded", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "PickLstNo", "Pick List No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "MatType", "Material Type", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "GRPONum", "GRPO No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "GRPOEntry", "GRPO DocEntry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "GRecptNo", "Goods Receipt No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "GRecptEnt", "Goods Receipt Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "InvTransNo", "Inventory Transfer No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "InvTransEntry", "Inventory Transfer Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "MilkType", "MilkType", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "TransType", "Transaction Type", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "TrgtDocType", "Target Document Type", 25, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "TrgtDocNum", "Target Document Number", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "TrgtDocEntry", "Target Document Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "Front", "Front", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "Middle", "Middle", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "Rear", "Rear", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "Branch", "Branch", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "AmtChrg", "Amount To Be Charged", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "ExcessEntry", "Excess Entry No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH1", "ShortEntry", "Short Entry No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region Vehicle Management Line 1

                CreateTable("VHWTMGTC1", "Vehicle Management Line", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                AddColumn("VHWTMGTC1", "Date", "Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC1", "Time", "Time", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC1", "TransTyp", "Transaction Type", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC1", "WtType", "Weight Type", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC1", "Weight", "Weight", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC1", "Remarks", "Remarks", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC1", "DocNum", "Document No", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC1", "UserName", "User Name", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC1", "SKiloMtrRdr", "Start Kilo Meter Reading", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Measurement, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC1", "EKiloMtrRdr", "End Kilo Meter Reading", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Measurement, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region Vehicle Management Line 2

                CreateTable("VHWTMGTC2", "Vehicle Management Line 1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                AddColumn("VHWTMGTC2", "ItemCode", "Item Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "ItemName", "Item Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "Quantity", "Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "Whse", "Warehouse Code",10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "Price", "Price", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "LineTot", "LineTotal", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "Comp", "Compartment", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "DocNum", "DocNum", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "DocEntry", "DocEntry", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "VWMNo", "VMNo", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "Per", "Per", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "StdPer", "StdPer", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "Var", "Variance", 20, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "RawQty", "Raw Milk Qty", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "QC", "QC Checked", 5, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "UserName", "User Name", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "Time", "Time", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "Remarks", "Remarks", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "BPCode", "BP Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "BPName", "BP Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "BPRef", "BP Ref", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "PrdUser", "Production User", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "PrdTime", "Production Time", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "QCUser", "QC User", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "QCTime", "QC Time", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "QCRemark", "QC Remarks", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "PrdRemar", "Production Remarks", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC2", "ActRemar", "Accounting Remarks", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                #endregion

                #region Vehicle Management Line 3

                CreateTable("VHWTMGTC3", "Vehicle Management Line 2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                AddColumn("VHWTMGTC3", "BPCode", "BP Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC3", "BPName", "BP Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC3", "InvNo", "Invoice No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC3", "InvEntry", "Invoice Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC3", "InvDate", "Invoice Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC3", "InvAmt", "Invoice Amount", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC3", "OPaymntN", "Outgoing Payment No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC3", "OPayEntry", "Outgoing Payment Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC3", "OPaymntD", "Outgoing Payment Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC3", "OPayAmt", "Outgoing Payment Amount", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC3", "InvRefNo", "Invoice Ref No", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC3", "Type", "Type", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                #endregion
               
                #region Vehicle Management Line 4

                CreateTable("VHWTMGTC4", "Vehicle Management Line 4", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                AddColumn("VHWTMGTC4", "ItemCode", "Item Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC4", "ItemName", "Item Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC4", "GRPONo", "GRPO No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC4", "GRPOEntry", "GRPO Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC4", "Var", "Variance Qty", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC4", "FromWhse", "From Ware House", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC4", "ToWhse", "To Ware House", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC4", "INTEntry", "Inventory Transfer Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region Vehicle Management Line 5
                CreateTable("VHWTMGTC5", "Vehicle Management Line 5", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                AddColumn("VHWTMGTC5", "ItemCode", "Item Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC5", "ItemName", "Item Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC5", "GRPONo", "GRPO No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC5", "GRPOEntry", "GRPO Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC5", "Var", "Variance Qty", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC5", "Whse", "Ware House", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC5", "UnitPrice", "Unit Price", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC5", "GREntry", "Goods Receipt Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                AddColumn("OCRG", "UsedCRLimit", "Used Credit Limit", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OCRG", "CRLimit", "Credit Limit", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OCRG", "AvailCRLimit", "Available Credit Limit", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OCRG", "RSM", "RSM", 40, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OCRG", "SalesOfficer", "Sales Officer", 40, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OCRG", "GuarntType", "Guarantee Type", 40, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OCRG", "BGChqNo", "BG/Chq No", 25, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OCRG", "IssueDate", "Issue Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OCRG", "ExpiryDate", "Expiry Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OCRG", "IssuBankName", "Issuing Bank Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OCRG", "Status", "Status", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);


                #region UDO'S

                oChildTables[0] = "VHWTMGTC1";
                oChildTables[1] = "VHWTMGTC2";
                oChildTables[2] = "VHWTMGTC3";
                //oChildTables[3] = "VHWTMGTC4";
               //oChildTables[4] = "VHWTMGTC5";
                oFindColumns[0] = "DocNum";
                oFindColumns[1] = "U_DocNum";
                CreateUDO("VHWTMGTH1", "Vehicle Mgt Header", "VHWTMGTH1", SAPbobsCOM.BoUDOObjType.boud_Document, oFindColumns, oChildTables);

                #endregion


                if ( Utilities.Application.Company.InTransaction )
					Utilities.Application.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
				
				oProgressBar.Stop();
			}
			catch(Exception ex)
			{
				oProgressBar.Stop();

				if( Utilities.Application.Company.InTransaction )
					Utilities.Application.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

				throw ex;
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressBar);
				oProgressBar = null;
				GC.Collect();
			}
		}
		
		#endregion

		#region DATABASE CREATION FUNCTIONS

		#region CREATE USER DEFINED TABLE
		
		private static void CreateTable(string oTable, string oDescription, SAPbobsCOM.BoUTBTableType oType)
		{
            GC.Collect();
            SAPbobsCOM.UserTablesMD oUserTable = null;

			try
			{
				oUserTable = (SAPbobsCOM.UserTablesMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
				
				if( !oUserTable.GetByKey(oTable) )
				{
					oUserTable.TableName        = oTable;
					oUserTable.TableDescription = oDescription;
					oUserTable.TableType        = oType;

					if( oUserTable.Add() != 0 )	
                    
						throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
																
				}
			}
			catch(Exception ex)
			{
                throw ex;
			}			
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
				oUserTable = null;
				GC.Collect();
			}
		}

        private static void UpdateTable(string oTable, SAPbobsCOM.BoUTBTableType oType)
        {
            SAPbobsCOM.UserTablesMD oUserTable = null;

            try
            {
                oUserTable = (SAPbobsCOM.UserTablesMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                if (oUserTable.GetByKey(oTable) != null)
                {
                    if (oUserTable.TableType != SAPbobsCOM.BoUTBTableType.bott_MasterData)
                    {
                        //oUserTable.TableType = SAPbobsCOM.BoUTBTableType.bott_MasterData;
                        if (oUserTable.Remove() != 0)

                            throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
                oUserTable = null;
                GC.Collect();
            }
        }

		#endregion

		#region ADD COLUMN

		internal  static void AddColumn(string oTable, string oName, string oDescription, int oSize, SAPbobsCOM.BoFieldTypes oType, SAPbobsCOM.BoFldSubTypes oSubType, SAPbobsCOM.BoYesNoEnum oMandatory, SAPbobsCOM.BoYesNoEnum oIsSystemTable)
		{
			SAPbobsCOM.UserFieldsMD oUserField = null;
            int Fieldid;
			try
			{

                if (ColumnExists(oTable, oName, oIsSystemTable,out  Fieldid) == SAPbobsCOM.BoYesNoEnum.tNO)
                {
                    oUserField = (SAPbobsCOM.UserFieldsMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

                    oUserField.TableName = oTable;
                    oUserField.Name = oName;
                    oUserField.Description = oDescription;
                    oUserField.Type = oType;
                    oUserField.SubType = oSubType;

                    if (oType != SAPbobsCOM.BoFieldTypes.db_Float)
                        oUserField.EditSize = oSize;
                    else
                        oUserField.EditSize = 0;

                    
                    oUserField.Mandatory = oMandatory;

                    //if (oTable == "OPDN" && oName == "ConType")
                    //{
                    //    oUserField.ValidValues.Value = "Vehicle";
                    //    oUserField.ValidValues.Description = "Vehicle";
                    //    oUserField.ValidValues.Add();
                    //    oUserField.ValidValues.Value = "SpareParts";
                    //    oUserField.ValidValues.Description = "SpareParts";
                    //    oUserField.ValidValues.Add();
                    //    oUserField.DefaultValue = "SpareParts";
                    //}
                     if (oTable == "OUSR" && oName == "VarAlwd")
                    {
                        oUserField.ValidValues.Value = "YES";
                        oUserField.ValidValues.Description = "YES";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "NO";
                        oUserField.ValidValues.Description = "NO";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "NO";
                    }
                     if (oTable == "VHWTMGTH1" && oName == "VehType")
                     {
                         oUserField.ValidValues.Value = "Truck";
                         oUserField.ValidValues.Description = "Truck";
                         oUserField.ValidValues.Add();
                         oUserField.ValidValues.Value = "Mini Truck";
                         oUserField.ValidValues.Description = "Mini Truck";
                         oUserField.ValidValues.Add();
                         oUserField.ValidValues.Value = "Tractor";
                         oUserField.ValidValues.Description = "Tractor";
                         oUserField.ValidValues.Add();
                         oUserField.ValidValues.Value = "Trailer";
                         oUserField.ValidValues.Description = "Trailer";
                         oUserField.ValidValues.Add();
                         oUserField.DefaultValue = "Truck";

                     }
                     if (oTable == "OCRG" && oName == "GuarntType")
                     {
                         oUserField.ValidValues.Value = "Bank Gurantee";
                         oUserField.ValidValues.Description = "Bank Gurantee";
                         oUserField.ValidValues.Add();
                         oUserField.ValidValues.Value = "Post Dated Cheque";
                         oUserField.ValidValues.Description = "Post Dated Cheque";
                         oUserField.ValidValues.Add();
                         oUserField.ValidValues.Value = "Open Dated Cheque";
                         oUserField.ValidValues.Description = "Open Dated Cheque";
                         oUserField.ValidValues.Add();
                        
                     }
                     if (oTable == "OCRG" && oName == "Status")
                     {
                         oUserField.ValidValues.Value = "Active";
                         oUserField.ValidValues.Description = "Active";
                         oUserField.ValidValues.Add();
                         oUserField.ValidValues.Value = "Inactive";
                         oUserField.ValidValues.Description = "Inactive";
                         oUserField.ValidValues.Add();
                         
                     }
                    if (oUserField.Add() != 0)
                            throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                }
                
			}
			catch( Exception ex)
			{
                throw ex;
			}
			finally
			{
                if (oUserField!=null)
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField);
				oUserField = null;
				GC.Collect();
			}
		}

        internal static void UpdateColumn(string oTable, string oName, string oDescription, int oSize, SAPbobsCOM.BoFieldTypes oType, SAPbobsCOM.BoFldSubTypes oSubType, SAPbobsCOM.BoYesNoEnum oMandatory, SAPbobsCOM.BoYesNoEnum oIsSystemTable)
        {
            SAPbobsCOM.UserFieldsMD oUserField = null;
            string oTable1=oTable ;
            int Fieldid=0;
            try
            {
                
                if (ColumnExists(oTable, oName, oIsSystemTable,out  Fieldid) == SAPbobsCOM.BoYesNoEnum.tYES)
                {
                    oUserField = (SAPbobsCOM.UserFieldsMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                    if(oTable!="OCRD") 
                    oTable1 = @"@" + oTable;

                    oUserField.GetByKey(oTable1, Fieldid);
                    if (oTable == "B1F_GNDR" && (oName == "B1FINACT"))
                    {
                        if (oUserField.DefaultValue != "N")
                        {
                            oUserField.DefaultValue = "N";
                            if (oUserField.Update() != 0)
                                throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                        }
                    }
                    else if ((oTable == "OCRD" && oName == "Season") || (oTable == "B1F_ITEMSEAS" && oName == "B1FSEASD") || (oTable == "B1F_ITEMCAPS" && (oName == "B1FCAPSD" ||oName =="B1FDTST" || oName == "B1FDTCP")))
                    {
                        if (oUserField.Remove()  != 0)
                            throw new Exception(Utilities.Application.Company.GetLastErrorDescription());

                    }
                    else if (oTable == "OCRD" && (oName == "B1FBUYINP" || oName == "B1FBUYINPN"))
                    {
                        if (oUserField.ValidValues.Value == "")
                        {
                            oUserField.ValidValues.Value = "N";
                            oUserField.ValidValues.Description = "No";
                            oUserField.ValidValues.Add();
                            oUserField.ValidValues.Value = "Y";
                            oUserField.ValidValues.Description = "Yes";
                            oUserField.ValidValues.Add();
                            oUserField.DefaultValue = "N";
                            if (oUserField.Update() != 0)
                                throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                        }

                    }
                    else
                    {
                        if((oTable == "B1F_FSPM" && oName=="B1FSKUNUM") || (oTable == "B1F_FSPM" && oName=="B1FATLEN")|| (oTable == "B1F_FSAC" && oName=="B1FSEQ")||(oTable == "B1F_FSIM" && oName=="B1FSKULN")||(oTable == "B1F_VAT1" && oName=="B1FGRPSQ")||(oTable == "B1F_VAT2" && oName=="B1FGRPSQ"))
                        {
                            if (oUserField.EditSize  < oSize)
                            {
                                oUserField.EditSize = oSize;
                                if (oUserField.Update() != 0)
                                    throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                            }
                        }
                        if (oUserField.Description != oDescription)
                        {
                            oUserField.Description = oDescription;
                            if (oUserField.Update() != 0)
                                throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                        }

                    }
                }
              

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (oUserField != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField);
                oUserField = null;
                GC.Collect();
            }
        }

        #endregion

        #region IS COLUMN EXISTS

        private static SAPbobsCOM.BoYesNoEnum ColumnExists(string oTable, string oColumn, SAPbobsCOM.BoYesNoEnum oIsSystemTable, out int FieldId)
        {
            SAPbobsCOM.Recordset oRSColumn = null;
            string oSQL = string.Empty;
            FieldId = 0;
            try
            {
                if (oIsSystemTable == SAPbobsCOM.BoYesNoEnum.tNO)
                    oTable = @"@" + oTable;

                //oSQL = "Select Count(*),FieldID From CUFD Where TableID = '" + oTable + "' And AliasID = '" + oColumn + "' group by FieldID";
                oSQL = @"Select " +
                        @"Count(*) "
                        + " , "
                        + @"""FieldID"""
                        + "From CUFD Where "
                        + @"""TableID"""
                         + "="
                        + "'" + oTable.Trim() + "'"
                        + " And "
                        + @"""AliasID"""
                        + "="
                        + "'" + oColumn.Trim() + "'"
                        + " group by "
                        + @"""FieldID""".Trim();


                Utilities.ExecuteSQL(ref oRSColumn, oSQL);

                if ((int)oRSColumn.Fields.Item(0).Value == 0)
                    return SAPbobsCOM.BoYesNoEnum.tNO;

                else
                {
                    FieldId = (int)oRSColumn.Fields.Item(1).Value;
                    return SAPbobsCOM.BoYesNoEnum.tYES;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSColumn);
                oRSColumn = null;
                GC.Collect();
            }
        }
       
        #endregion

		#region CREATE UDO

		private static void CreateUDO(string oUniqueID, string oDescription, string oTable, SAPbobsCOM.BoUDOObjType oType, string[] oFindColumns, string[] oChildTables )
		{
			SAPbobsCOM.UserObjectsMD oUserObject = null;
			int                      oCount;

			try
			{
				oUserObject = (SAPbobsCOM.UserObjectsMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

				if( !oUserObject.GetByKey(oUniqueID) )
				{
					oUserObject.Code        = oUniqueID;
					oUserObject.Name        = oDescription;
					oUserObject.ObjectType  = oType;
					oUserObject.TableName   = oTable;

					oUserObject.ManageSeries         = SAPbobsCOM.BoYesNoEnum.tNO;
					oUserObject.CanCancel            = SAPbobsCOM.BoYesNoEnum.tYES;
					oUserObject.CanClose             = SAPbobsCOM.BoYesNoEnum.tYES;
					oUserObject.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;
					oUserObject.CanDelete            = SAPbobsCOM.BoYesNoEnum.tNO;

					if( oFindColumns != null )
					{
						if( oFindColumns.Length > 0 )
						{
							oUserObject.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;

							for(oCount = 0; oCount <= oFindColumns.Length-1; oCount++)
							{
								if( oFindColumns[oCount] != null )
								{
									oUserObject.FindColumns.ColumnAlias = oFindColumns[oCount];
									oUserObject.FindColumns.Add();
								}

								oFindColumns[oCount] = null;
							}
						}
					}

					if( oChildTables != null )
					{
						if( oChildTables.Length > 0 )
						{
							for(oCount = 0; oCount <= oChildTables.Length-1; oCount++)
							{
								if( oChildTables[oCount] != null )
								{
									oUserObject.ChildTables.TableName = oChildTables[oCount];

									if( oCount != oChildTables.Length - 1 )
										oUserObject.ChildTables.Add();
								}

								oChildTables[oCount] = null;
							}
						}
					}

					if( oUserObject.Add() != 0 )
						throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
				}
			}
			catch(Exception ex)
			{
				throw ex;
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObject);
				oUserObject = null;
				GC.Collect();
			}
		}
		
        #endregion

        #region CreateUserObject

        public static void CreateUserObject(string CodeID, string Name, string TableName, string Child) //used for registration of user defined table
        {
            int lRetCode = 0;
            string sErrMsg = null;
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
            try
            {
                if (oUserObjectMD == null)
                    oUserObjectMD = ((SAPbobsCOM.UserObjectsMD)(Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)));

                if (oUserObjectMD.GetByKey(CodeID) == true)
                    return; //vijay 170708

                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;


                oUserObjectMD.Code = CodeID;
                oUserObjectMD.Name = Name;
                oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData;
                oUserObjectMD.TableName = TableName;

                if (Child != "")
                {
                    oUserObjectMD.ChildTables.TableName = Child;
                    //oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
                }
                lRetCode = oUserObjectMD.Add();

                // check for errors in the process
                if (lRetCode != 0)
                    if (lRetCode == -1)
                    { }
                    else
                    { Utilities.Application.Company.GetLastError(out lRetCode, out sErrMsg); }
                else
                { }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                oUserObjectMD = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        #endregion

        #region IS KEY EXISTS

        private static SAPbobsCOM.BoYesNoEnum KeyExists(string oTable,string oKeyName,SAPbobsCOM.BoYesNoEnum oIsSystemTable)
		{
			SAPbobsCOM.Recordset oRSColumn = null;
			string               oSQL      = string.Empty;
			
			try
			{		
				if( oIsSystemTable == SAPbobsCOM.BoYesNoEnum.tNO )
					oTable = @"@" + oTable;
				oSQL = "Select Count(*) From OUKD Where TableName = '" + oTable + "' AND KeyName = '" + oKeyName + "' ";
				Utilities.ExecuteSQL(ref oRSColumn, oSQL);
            
				if( (int)oRSColumn.Fields.Item(0).Value == 0 )
					return SAPbobsCOM.BoYesNoEnum.tNO;

				else
					return SAPbobsCOM.BoYesNoEnum.tYES;
			}
			catch(Exception ex)
			{
				throw ex;
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSColumn);
				oRSColumn = null;
				GC.Collect();
			}
		}
		
        #endregion

		#region ADD KEYS

		private static void AddKey(string oTable, string[] oColumnName,string oKeyName,SAPbobsCOM.BoYesNoEnum oIsSystemTable)
		{
			SAPbobsCOM.UserKeysMD oUserKey = null;
			int                   oCount;
			try
			{				
				oUserKey = (SAPbobsCOM.UserKeysMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys);
				if(KeyExists(oTable,oKeyName,oIsSystemTable) == SAPbobsCOM.BoYesNoEnum.tNO)
				{
					oUserKey.TableName = oTable ;
					oUserKey.KeyName   = oKeyName ;
					if( oColumnName != null )
					{
						if( oColumnName.Length > 0 )
						{						
							for(oCount = 0; oCount <= oColumnName.Length-1; oCount++)
							{
								if( oColumnName[oCount] != null )
								{
									oUserKey.Elements.ColumnAlias = oColumnName[oCount];
									oUserKey.Elements.Add();
								}

								oColumnName[oCount] = null;
							}
						}
					}							
					oUserKey.Unique = SAPbobsCOM.BoYesNoEnum.tYES;					
					if( oUserKey.Add() != 0 )
						throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
				}
			}
			catch(Exception ex)
			{
				throw ex ;
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKey);
				oUserKey = null;
				GC.Collect();
			}
		}
		
        #endregion

		#endregion

	}
}
