using System;
using System.Data;
using System.Data.SqlClient;
using Microsoft.ApplicationBlocks.Data;

namespace BatchManagement_LaxmiSteel
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

                #region WB Dispatch Header

                CreateTable("WBLOADDTLH", "WB Dispatch Header", SAPbobsCOM.BoUTBTableType.bott_Document);
                AddColumn("WBLOADDTLH", "PickLstNo", "Pick List Number", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "TrnspName", "Transporter Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "TruckNo", "Truck No", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "DriverNm", "Driver Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "ContactNo", "Contact No", 25, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "TrckInDt", "Truck In Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "WtBefrLoad", "Weight Before Loading", 2, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "Remarks", "Remarks", 250, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "DocNum", "Document No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "CardCode", "Customer Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "CardName", "Customer Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "SONO", "Sales Order No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "SOEntry", "Sales Order Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "DONO", "Sales Order No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "DOEntry", "Sales Order Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "UserName", "User Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "LoadDt", "Loading Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "LoadTm", "Loading Time", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "CloseTm", "Closing Time", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "PckLstRemrk", "Pick List Remarks", 250, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "BillTo", "BillTo Address", 250, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "ShipTo", "ShipTo Address", 250, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "Status", "Status", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLH", "VehType", "Vehicle Type", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                #endregion

                #region WB Dispatch Line

                CreateTable("WBLOADDTLC", "WB Dispatch Line", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                AddColumn("WBLOADDTLC", "ItemCode", "Item Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLC", "ItemName", "Item Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLC", "ItemGrp", "Item Group", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLC", "NOPCPB", "No Of Pieces Per Bundle", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLC", "OrderQty", "Order Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLC", "UoM", "UoM", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLC", "InStockQ", "In Stock Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLC", "DelQty", "Delivered Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLC", "QtyToDis", "Quantity To Dispatch", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLC", "NoOfBund", "Number Of Bundle Created", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLC", "TotQtyIn", "Total Quantity In Bundle", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLC", "Variance", "Variance", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLC", "TotVaria", "Total Variance", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLC", "DocNum", "Document No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLC", "ItemCode", "Item Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLC", "LineNum", "Line Number", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLC", "LotNo", "Lot Number", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBLOADDTLC", "VarAlWd", "Variance Alloed", 5, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                
                #endregion 

                #region WBDETAIL

                CreateTable("WBDETAIL", "Weighing Bridge Details", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                AddColumn("WBDETAIL", "ItemCode", "Item Code", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBDETAIL", "ItemName", "Item Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBDETAIL", "LotNo", "Lot Number", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBDETAIL", "BundleNo", "Bundle No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBDETAIL", "Weight", "Weight", 2, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBDETAIL", "NOPCPB", "No Of Pieces Per Bundle", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBDETAIL", "BatchNo", "Batch No", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBDETAIL", "PickLstN", "Pick List Number", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBDETAIL", "TruckNo", "Truck No", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBDETAIL", "LineNum", "Line Number", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBDETAIL", "ItmsGrpN", "Item Group Name", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBDETAIL", "VarAlWd", "Variance Alloed", 5, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBDETAIL", "PrntSt", "Printing Status", 5, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBDETAIL", "BLineNum", "Based Doc Line Number", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                #endregion

                #region OPKL UDF

                AddColumn("OPKL", "TruckNo", "Truck No", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OPKL", "DriverNm", "Driver Name", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OPKL", "TranspNm", "Transporter Name", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OPKL", "CounterNo", "Loading Counter No", 5, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OPKL", "ContactNo", "Contact No", 25, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OPKL", "VWMNo", "Vehcle Management No", 25, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OPKL", "ExpLoadDt", "Expected Loading Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OPKL", "ExpLoadTm", "Expected Loading Time", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OPKL", "LdCntPer", "Loading Counter Contact Person", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);

                #endregion

                #region OITB UDF

                AddColumn("OITB", "VarAlWd", "Variance Alloed", 5, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OITB", "NOPCPB", "No Of Pieces Per Bundle", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);

                #endregion

                #region UDF - OUSR

                AddColumn("OUSR", "VarAlwd", "Variance Allowed", 5, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                
                #endregion

                #region Vehicle Management Header

                CreateTable("VHWTMGTH", "Vehicle Management Header", SAPbobsCOM.BoUTBTableType.bott_Document);
                AddColumn("VHWTMGTH", "CNTOOrgn", "Country Of Origin", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "VehType", "Vehicle Type", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "VehNo", "Vehicle No", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "TranspName", "Transporter Name", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "TranspPAN", "Transporter PAN", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "DriverNm", "Driver Name", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "DriverNo", "Driver No", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "LicenseNo", "License No", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "LdContNo", "Loading Counter No", 5, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "UserName", "User Name", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "ClosedBy", "Closed By", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "DocNum", "Document No", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "DocDate", "Document Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "Status", "Status", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "VisitPrPs", "Visit Purpose", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "TareWt", "Tare Weight", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "NetWt", "Net Weight", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "WBDocNum", "Weighing Bridge Document No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "WBCounter", " Weight As Per Dispatch Counter", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "NTVarWt", "Net Vaiance in Weight", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "GrossWt", "Gross Weight", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "LodWt", "Weight to be loaded", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "PickLstNo", "Pick List No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "SO", "Sales Order", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "MaterialType", "Material Type", 25, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "TrgtDocType", "Target Document Type", 25, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "TrgtDocNum", "Target Document Number", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "TrgtDocEntry", "Target Document Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "WBType", "Weighing Type", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "WBStatus", "Weighing Status", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "TransType", "Transaction Type", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTH", "AmtToBLoad", "Amount To Be Loaded", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                #endregion

                #region Vehicle Management Line

                CreateTable("VHWTMGTC", "Vehicle Management Line", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                AddColumn("VHWTMGTC", "Date", "Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC", "Time", "Time", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC", "TransTyp", "Transaction Type", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC", "WtType", "Weight Type", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC", "Weight", "Weight", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC", "Remarks", "Remarks", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC", "DocNum", "Document No", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("VHWTMGTC", "UserName", "User Name", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                #endregion

                #region WB Dispatch Trailer Header

                CreateTable("WBTRAILERH", "WB Dispatch Trailer Header", SAPbobsCOM.BoUTBTableType.bott_Document);
                AddColumn("WBTRAILERH", "PickLstNo", "Pick List Number", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "TrnspName", "Transporter Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "TruckNo", "Truck No", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "DriverNm", "Driver Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "ContactNo", "Contact No", 25, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "TrckInDt", "Truck In Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "WtBefrLoad", "Weight Before Loading", 2, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "Remarks", "Remarks", 250, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "DocNum", "Document No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "CardCode", "Customer Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "CardName", "Customer Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "SONO", "Sales Order No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "SOEntry", "Sales Order Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "DONO", "Sales Order No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "DOEntry", "Sales Order Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "UserName", "User Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "LoadDt", "Loading Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "LoadTm", "Loading Time", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "CloseTm", "Closing Time", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "PckLstRemrk", "Pick List Remarks", 250, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "BillTo", "BillTo Address", 250, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "ShipTo", "ShipTo Address", 250, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "Status", "Status", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "VMDocNum", "Vehicle Management Number", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERH", "TareWt", "Tare Weight", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                #endregion

                #region WB Dispatch Trailer Line

                CreateTable("WBTRAILERC", "WB Dispatch Trailer Line", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                AddColumn("WBTRAILERC", "ItemCode", "Item Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "ItemName", "Item Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "ItemGrp", "Item Group", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "NOPCPB", "No Of Pieces Per Bundle", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "OrderQty", "Order Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "UoM", "UoM", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "InStockQ", "In Stock Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "DelQty", "Delivered Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "QtyToDis", "Quantity To Dispatch", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "NoOfBund", "Number Of Bundle Created", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "TotQtyIn", "Total Quantity In Bundle", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "Variance", "Variance", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "TotVaria", "Total Variance", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "DocNum", "Document No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "ItemCode", "Item Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "LineNum", "Line Number", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "LotNo", "Lot Number", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "VarAlWd", "Variance Alloed", 5, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "GrossWt", "Gross Weight", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("WBTRAILERC", "NetWt", "Net Weight", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                #endregion 

                #region UDO'S

                oChildTables[0] = "WBLOADDTLC";
                oFindColumns[0] = "DocNum";
                oFindColumns[1] = "U_DocNum";
                CreateUDO("WBLOADDTLH", "WB Dispatch", "WBLOADDTLH", SAPbobsCOM.BoUDOObjType.boud_Document, oFindColumns, oChildTables);

                oChildTables[0] = "WBTRAILERC";
                oFindColumns[0] = "DocNum";
                oFindColumns[1] = "U_DocNum";
                CreateUDO("WBTRAILERH", "WB Dispatch Trailer", "WBTRAILERH", SAPbobsCOM.BoUDOObjType.boud_Document, oFindColumns, oChildTables);

                oChildTables[0] = "VHWTMGTC";
                oFindColumns[0] = "DocNum";
                oFindColumns[1] = "U_DocNum";
                CreateUDO("VHWTMGTH", "Vehicle Mgt Header", "VHWTMGTH", SAPbobsCOM.BoUDOObjType.boud_Document, oFindColumns, oChildTables);

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
                    if (oTable == "VHWTMGTH" && oName == "VehType")
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
                    if (oTable == "VHWTMGTH" && oName == "MaterialType")
                    {
                        oUserField.ValidValues.Value = "Billet";
                        oUserField.ValidValues.Description = "Billet";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "Consumables";
                        oUserField.ValidValues.Description = "Consumables";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "Cylinder";
                        oUserField.ValidValues.Description = "Cylinder";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "Furnance Oil";
                        oUserField.ValidValues.Description = "Furnance Oil";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "Fixed Asset";
                        oUserField.ValidValues.Description = "Fixed Asset";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "Spare Part";
                        oUserField.ValidValues.Description = "Spare Part";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "TMT BAR";
                        oUserField.ValidValues.Description = "TMT BAR";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "Scrap";
                        oUserField.ValidValues.Description = "Scrap";
                        oUserField.ValidValues.Add();
                       

                    }
                    if (oTable == "WBLOADDTLH" && oName == "VehType")
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
