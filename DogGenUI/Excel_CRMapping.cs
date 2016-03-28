﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Services.Client;
using System.Linq;
using System.Net;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using DocGenerator.SDDPServiceReference;

namespace DocGenerator
	{

	/// <summary>
	/// This class handles the Client_Requirements_Mapping_Workbook
	/// </summary>
	class Client_Requirements_Mapping_Workbook:aWorkbook
		{
		//private bool _client_Requirements_Mapping_Workbook = false;
		//public bool Client_Requirements_Mapping_Wbk
		//	{
		//	get{return this._client_Requirements_Mapping_Workbook;}
		//	set{this._client_Requirements_Mapping_Workbook = value;}
		//	}


		private int? _crm_Mapping = 0;
		/// <summary>
		/// This property reference the ID value of the SharePoint Mappings entry which is used to generate the Document
		/// </summary>
		public int? CRM_Mapping
			{
			get
				{
				return this._crm_Mapping;
				}
			set
				{
				this._crm_Mapping = value;
				}
			}

		public bool Generate()
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			DateTime timeStarted = DateTime.Now;
			//string hyperlinkImageRelationshipID = "";
			string strDocumentCollection_HyperlinkURL = "";
			int intHyperlinkCounter = 9;
			string strCurrentHyperlinkViewEditURI = "";
			Cell objCell = new Cell();
			int intSharedStringIndex = 0;
			//Workbook Break processing Variables
			int intRequirementBreakID_forRisks = 0;      // the ID value of the Requirement used as a break processing variable for Risks sheet
			int intRequirementBreakID_forAssumptions = 0;	// the ID value of the Requirement used as a break processing variable for Assumptions sheet

			//Content Layering Variables
			int? intLayer0upDeliverableID;
			int? intLayer1upDeliverableID;
			int? intLayer2upDeliverableID;
			string strTextDescription = "";

			//Worksheet Row Index Variables
			UInt16 intMatrixSheet_RowIndex = 6;
			UInt16 intRisksSheet_RowIndex = 2;
			UInt16 intAssumptionsSheet_RowIndex = 2;
			Dictionary<string, string> dictionaryMatrixComments = new Dictionary<string, string>();
			string strErrorText = "";
			if(this.HyperlinkEdit)
				{
				strDocumentCollection_HyperlinkURL = Properties.AppResources.SharePointSiteURL +
					Properties.AppResources.List_DocumentCollectionLibraryURI +
					Properties.AppResources.EditFormURI + this.DocumentCollectionID;
				strCurrentHyperlinkViewEditURI = Properties.AppResources.EditFormURI;
				}
			if(this.HyperlinkView)
				{
				strDocumentCollection_HyperlinkURL = Properties.AppResources.SharePointSiteURL +
					Properties.AppResources.List_DocumentCollectionLibraryURI +
					Properties.AppResources.DisplayFormURI + this.DocumentCollectionID;
				strCurrentHyperlinkViewEditURI = Properties.AppResources.DisplayFormURI;
				}

			//Initialize the Data access to SharePoint
			DesignAndDeliveryPortfolioDataContext datacontexSDDP = new DesignAndDeliveryPortfolioDataContext(new
				Uri(Properties.AppResources.SharePointSiteURL + Properties.AppResources.SharePointRESTuri));
			datacontexSDDP.Credentials = CredentialCache.DefaultCredentials;
			datacontexSDDP.MergeOption = MergeOption.NoTracking;

			// define a new objOpenXMLworksheet
			oxmlWorkbook objOXMLworkbook = new oxmlWorkbook();
			// use CreateDocumentFromTemplate method to create a new MS Word Document based on the relevant template
			if(objOXMLworkbook.CreateDocWbkFromTemplate(
				parDocumentOrWorkbook: enumDocumentOrWorkbook.Workbook,
				parTemplateURL: this.Template, 
				parDocumentType: this.DocumentType))
				{
				Console.WriteLine("\t\t\t objOXMLdocument:\n" +
				"\t\t\t+ LocalDocumentPath: {0}\n" +
				"\t\t\t+ DocumentFileName.: {1}\n" +
				"\t\t\t+ DocumentURI......: {2}", objOXMLworkbook.LocalPath, objOXMLworkbook.Filename, objOXMLworkbook.LocalURI);
				}
			else
				{
				// if the creation failed.
				Console.WriteLine("An ERROR occurred and the new MS Word Document could not be created due to above stated ERROR conditions.");
				this.ErrorMessages.Add("Application was unable to create the document based on the template - Check the Output log.");
				return false;
				}

			if(this.CRM_Mapping == null || this.CRM_Mapping == 0)
				{
				Console.WriteLine("\t\t\t *** The user didn't specify the Client Requirements Mapping to be generated.");
				this.ErrorMessages.Add("The user didn't specify the Client Requirements Mapping to be generated.");
				return false;
				}

			// Open the MS Excel Workbook 
			try
				{
				// Open the MS Excel document in Edit mode
				// https://msdn.microsoft.com/en-us/library/office/hh298534.aspx
				SpreadsheetDocument objSpreadsheetDocument = SpreadsheetDocument.Open(path: objOXMLworkbook.LocalURI, isEditable: true);
				// Obtain the WorkBookPart from the spreadsheet.
				if(objSpreadsheetDocument.WorkbookPart == null)
					{
					throw new ArgumentException(objOXMLworkbook.LocalURI +
						" does not contain a WorkbookPart.");
					}
				WorkbookPart objWorkbookPart = objSpreadsheetDocument.WorkbookPart;

				//Obtain the SharedStringTablePart - If it doesn't exisit, create new one.
				SharedStringTablePart objSharedStringTablePart;
				if(objSpreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
					{
					objSharedStringTablePart = objSpreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
					}
				else
					{
					objSharedStringTablePart = objSpreadsheetDocument.AddNewPart<SharedStringTablePart>();
					}

				// obtain the Matrix Worksheet in the Workbook.
				Sheet objMatrixWorksheet = objWorkbookPart.Workbook.Descendants<Sheet>().Where(sht => sht.Name == Properties.AppResources.Workbook_CRM_Worksheet_Name_Matrix).FirstOrDefault();
				if(objMatrixWorksheet == null)
					{
					throw new ArgumentException("The " + Properties.AppResources.Workbook_CRM_Worksheet_Name_Matrix +
						" worksheet could not be loacated in the workbook.");
					}
				// obtain the WorksheetPart of the objMatrixWorksheet
				WorksheetPart objMatrixWorksheetPart = (WorksheetPart)(objWorkbookPart.GetPartById(objMatrixWorksheet.Id));
				
				// Copy the Formats from Row 8 into the List of Formats from where it can be applied to every Row
                    Client_Requirements_Mapping_Workbook objCRMworkbook = new Client_Requirements_Mapping_Workbook();
				List<UInt32Value> listMatrixColumnStyles = new List<UInt32Value>();
				int intLastColumn = 15;
				int intStyleSourceRow = 7;
				string strCellAddress = "";
				for(int i = 0; i < intLastColumn - 1; i++)
					{
					strCellAddress = aWorkbook.GetColumnLetter(i) + intStyleSourceRow;
					Cell objSourceCell = objMatrixWorksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == strCellAddress).FirstOrDefault();
					if(objSourceCell != null)
						{
						listMatrixColumnStyles.Add(objSourceCell.StyleIndex);
						}
					else
						listMatrixColumnStyles.Add(0U);
					} // loop

				// obtain the Risks Worksheet in the Workbook.
				Sheet objRisksWorksheet = objWorkbookPart.Workbook.Descendants<Sheet>().Where(sht => sht.Name == Properties.AppResources.Workbook_CRM_Worksheet_Name_Risks).FirstOrDefault();
				if(objRisksWorksheet == null)
					{
					throw new ArgumentException("The " + Properties.AppResources.Workbook_CRM_Worksheet_Name_Risks +
						" worksheet could not be loacated in the workbook.");
					}
				WorksheetPart objRisksWorksheetPart = (WorksheetPart)(objWorkbookPart.GetPartById(objRisksWorksheet.Id));

				// Copy the Formats from Row 3 into the List of Formats from where it can be applied to every Row at a later stage
				List<UInt32Value> listRisksColumnStyles = new List<UInt32Value>();
				intLastColumn = 8;
				intStyleSourceRow = 3;
				strCellAddress = "";
				Console.WriteLine("Risk Style Values per Column");
				for(int i = 0; i < intLastColumn - 1; i++)
					{
					strCellAddress = aWorkbook.GetColumnLetter(i) + intStyleSourceRow;
					Cell objSourceCell = objRisksWorksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == strCellAddress).FirstOrDefault();
					if(objSourceCell != null)
						{
						listRisksColumnStyles.Add(objSourceCell.StyleIndex);
                              }
					else
						{
						listRisksColumnStyles.Add(0U);
						}
					Console.WriteLine("\t + {0} - {1} = {2}", i, strCellAddress, objSourceCell.StyleIndex);
					} // loop

				// obtain the Assumptions Worksheet in the Workbook.
				Sheet objAssumptionsWorksheet = objWorkbookPart.Workbook.Descendants<Sheet>().Where(sht => sht.Name == Properties.AppResources.Workbook_CRM_Worksheet_Name_Assumptions).FirstOrDefault();
				if(objAssumptionsWorksheet == null)
					{
					throw new ArgumentException("The " + Properties.AppResources.Workbook_CRM_Worksheet_Name_Assumptions +
					   " worksheet could not be loacated in the workbook.");
					}
				WorksheetPart objAssumptionsWorksheetPart = (WorksheetPart)(objWorkbookPart.GetPartById(objAssumptionsWorksheet.Id));

				// Copy the Formats from Row 3 into the List of Formats from where it can be applied to every Row at a later stage
				List<UInt32Value> listAssumptionsColumnStyles = new List<UInt32Value>();
				intLastColumn = 5;
				intStyleSourceRow = 3;
				strCellAddress = "";
				Console.WriteLine("Assumptions Style Values per Column");
				for(int i = 0; i < intLastColumn - 1; i++)
					{
					strCellAddress = aWorkbook.GetColumnLetter(i) + intStyleSourceRow;
					Cell objSourceCell = objAssumptionsWorksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == strCellAddress).FirstOrDefault();
					if(objSourceCell != null)
						{
						listAssumptionsColumnStyles.Add(objSourceCell.StyleIndex);
						}
					else
						{
						listAssumptionsColumnStyles.Add(0U);
						}
					Console.WriteLine("\t + {0} - {1} = {2}", i, strCellAddress, objSourceCell.StyleIndex);
					} // loop


				// If Hyperlinks need to be inserted, add the 
				Hyperlinks objHyperlinks = new Hyperlinks();
				//-------------------------------------
				// Begin to process the selected Mapping
				if(this.CRM_Mapping == 0)
					{
					strErrorText = "A Client Requirements Mapping was not specified for the Document Collection.";
					Console.WriteLine("### {0} ###", strErrorText);
					// If an entry was not specified - write an error in the Worksheet and record an error in the error log.
					this.LogError(strErrorText);

					//intStringIndex = oxmlWorkbook.InsertSharedStringItem(parText2Insert: strErrorText, parShareStringPart: objSharedStringTablePart);

					objCell = oxmlWorkbook.InsertCellInWorksheet(
						parColumnName: "A",
						parRowNumber: intMatrixSheet_RowIndex,
						parWorksheetPart: objMatrixWorksheetPart);
					objCell.DataType = new EnumValue<CellValues>(CellValues.String);
					objCell.CellValue = new CellValue(strErrorText);
					goto Save_and_Close_Document;
                         }

				//=============================================================
				// Begin to process the Mapping data 
				Mapping objMapping = new Mapping();
				objMapping.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: this.CRM_Mapping);
				Console.WriteLine(" + Mapping: {0} - {1}", objMapping.ID, objMapping.Title);

				// Declare the List containing the various types of objects to be processed
				// These lists consists of the various objects.
				List<MappingServiceTower> listMappingTowers = new List<MappingServiceTower>();
				List<MappingRequirement> listMappingRequirements = new List<MappingRequirement>();
				List<MappingDeliverable> listMappingDeliverables = new List<MappingDeliverable>();
				List<MappingRisk> listMappingRisks = new List<MappingRisk>();
				List<MappingAssumption> listMappingAssumptions = new List<MappingAssumption>();
				List<MappingServiceLevel> listMappingServiceLevels = new List<MappingServiceLevel>();
				
				//-------------------------------------------------------------
				// Obtain all Mapping Service Towers for the specified Mapping
		
				listMappingTowers.Clear();
				listMappingTowers = MappingServiceTower.ObtainListOfObjects(parDatacontextSDDP: datacontexSDDP, parMappingID: objMapping.ID);
				// Check if any entries were retrieved, If not end the generation of the workbook
				if(listMappingTowers.Count == 0)
					{
					strErrorText = "No Towers of Service was found for the mapping.";
					Console.WriteLine("### {0} ###", strErrorText);
					this.LogError(strErrorText);
					goto Save_and_Close_Document;
					}

				string strColumnLetter = String.Empty;
				// --- Loop through all Service Towers for the Mapping ---
				foreach(MappingServiceTower objTower in listMappingTowers)
					{
					// Write the Mapping Service Tower to the Workbook as a String
					Console.WriteLine("\t + Tower: {0} - {1}", objTower.ID, objTower.Title);
					intMatrixSheet_RowIndex += 1;
					//--- Matrix --- Tower of Service Row --- Column A --------------------------------
					oxmlWorkbook.PopulateCell(
						parWorksheetPart: objMatrixWorksheetPart,
						parColumnLetter: "A",
						parRowNumber: intMatrixSheet_RowIndex,
						parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("A"))),
						parCellDatatype: CellValues.String,
						parCellcontents: objTower.Title);
					//--- Matrix --- Tower of Service Row --- Column B --------------------------------
					oxmlWorkbook.PopulateCell(
						parWorksheetPart: objMatrixWorksheetPart,
						parColumnLetter: "B",
						parRowNumber: intMatrixSheet_RowIndex,
						parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("B"))),
						parCellDatatype: CellValues.String);
					//--- Matrix --- Tower of Service Row --- Column C --------------------------------
					oxmlWorkbook.PopulateCell(
						parWorksheetPart: objMatrixWorksheetPart,
						parColumnLetter: "C",
						parRowNumber: intMatrixSheet_RowIndex,
						parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("C"))),
						parCellDatatype: CellValues.String);
					//--- Matrix --- Tower of Service Row --- Column D --------------------------------
					oxmlWorkbook.PopulateCell(
						parWorksheetPart: objMatrixWorksheetPart,
						parColumnLetter: "D",
						parRowNumber: intMatrixSheet_RowIndex,
						parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("D"))),
						parCellDatatype: CellValues.String);
					//--- Matrix --- Tower of Service Row --- Column E --------------------------------
					oxmlWorkbook.PopulateCell(
						parWorksheetPart: objMatrixWorksheetPart,
						parColumnLetter: "E",
						parRowNumber: intMatrixSheet_RowIndex,
						parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("E"))),
						parCellDatatype: CellValues.String);
					//--- Matrix --- Tower of Service Row --- Column F --------------------------------
					oxmlWorkbook.PopulateCell(
						parWorksheetPart: objMatrixWorksheetPart,
						parColumnLetter: "F",
						parRowNumber: intMatrixSheet_RowIndex,
						parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("F"))),
						parCellDatatype: CellValues.String);
					//--- Matrix --- Tower of Service Row --- Column G --------------------------------
					intSharedStringIndex = oxmlWorkbook.InsertSharedStringItem(
						parText2Insert: Properties.AppResources.Workbook_CRM_Matrix_RowType_TowerOfService, parShareStringPart: objSharedStringTablePart);
					oxmlWorkbook.PopulateCell(
						parWorksheetPart: objMatrixWorksheetPart,
						parColumnLetter: "G",
						parRowNumber: intMatrixSheet_RowIndex,
						parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("G"))),
						parCellDatatype: CellValues.SharedString,
						parCellcontents: intSharedStringIndex.ToString());
					//--- Matrix --- Tower of Service Row --- Column H --------------------------------
					oxmlWorkbook.PopulateCell(
						parWorksheetPart: objMatrixWorksheetPart,
						parColumnLetter: "H",
						parRowNumber: intMatrixSheet_RowIndex,
						parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("H"))),
						parCellDatatype: CellValues.String);
					//--- Matrix --- Tower of Service Row --- Column I --------------------------------
					oxmlWorkbook.PopulateCell(
						parWorksheetPart: objMatrixWorksheetPart,
						parColumnLetter: "I",
						parRowNumber: intMatrixSheet_RowIndex,
						parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("I"))),
						parCellDatatype: CellValues.String);
					//--- Matrix --- Tower of Service Row --- Column J --------------------------------
					oxmlWorkbook.PopulateCell(
						parWorksheetPart: objMatrixWorksheetPart,
						parColumnLetter: "J",
						parRowNumber: intMatrixSheet_RowIndex,
						parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("J"))),
						parCellDatatype: CellValues.String);
					//--- Matrix --- Tower of Service Row --- Column K --------------------------------
					oxmlWorkbook.PopulateCell(
						parWorksheetPart: objMatrixWorksheetPart,
						parColumnLetter: "K",
						parRowNumber: intMatrixSheet_RowIndex,
						parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("K"))),
						parCellDatatype: CellValues.String);
					//--- Matrix --- Tower of Service Row --- Column L --------------------------------
					intHyperlinkCounter += 1;
					oxmlWorkbook.PopulateCell(
						parWorksheetPart: objMatrixWorksheetPart,
						parColumnLetter: "L",
						parRowNumber: intMatrixSheet_RowIndex,
						parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("L"))),
						parCellDatatype: CellValues.Number,
						parCellcontents: objTower.ID.ToString(),
						parHyperlinkCounter: intHyperlinkCounter,
						parHyperlinkURL: Properties.AppResources.SharePointURL +
							Properties.AppResources.List_MappingServiceTowers +
							Properties.AppResources.EditFormURI + objTower.ID.ToString());
					//--- Matrix --- Tower of Service Row --- Column M --------------------------------
					oxmlWorkbook.PopulateCell(
						parWorksheetPart: objMatrixWorksheetPart,
						parColumnLetter: "M",
						parRowNumber: intMatrixSheet_RowIndex,
						parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("M"))),
						parCellDatatype: CellValues.String);
					//--- Matrix --- Tower of Service Row --- Column N --------------------------------
					oxmlWorkbook.PopulateCell(
						parWorksheetPart: objMatrixWorksheetPart,
						parColumnLetter: "N",
						parRowNumber: intMatrixSheet_RowIndex,
						parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("N"))),
						parCellDatatype: CellValues.String);
					
					//========================================================================
					// Obtain all Mapping Requirements for the specified Mapping Service Tower
					
					listMappingRequirements.Clear();
					listMappingRequirements = MappingRequirement.ObtainListOfObjects(parDatacontextSDDP: datacontexSDDP, parMappingTowerID: objTower.ID);

					if(listMappingRequirements.Count() < 1)
						continue;
					
					// Process all the Mapping requirements for the specific Service Tower
					foreach(MappingRequirement objRequirement in listMappingRequirements)
						{
						Console.WriteLine("\t\t + Requirement: {0} - {1}", objRequirement.ID, objRequirement.Title);
						intMatrixSheet_RowIndex += 1;
						//--- Matrix --- Requirement Row --- Column A --------------------------------
						oxmlWorkbook.PopulateCell(
						parWorksheetPart: objMatrixWorksheetPart,
						parColumnLetter: "A",
						parRowNumber: intMatrixSheet_RowIndex,
						parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("A"))),
						parCellDatatype: CellValues.String);
						//--- Matrix --- Requirement Row --- Column B --------------------------------
						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objMatrixWorksheetPart,
							parColumnLetter: "B",
							parRowNumber: intMatrixSheet_RowIndex,
							parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("B"))),
							parCellDatatype: CellValues.String,
							parCellcontents: objRequirement.Title);
						//--- Matrix --- Requirement Row --- Column C --------------------------------
						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objMatrixWorksheetPart,
							parColumnLetter: "C",
							parRowNumber: intMatrixSheet_RowIndex,
							parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("C"))),
							parCellDatatype: CellValues.String);
						//--- Matrix --- Requirement Row --- Column D --------------------------------
						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objMatrixWorksheetPart,
							parColumnLetter: "D",
							parRowNumber: intMatrixSheet_RowIndex,
							parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("D"))),
							parCellDatatype: CellValues.String);
						//--- Matrix --- Requirement Row --- Column E --------------------------------
						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objMatrixWorksheetPart,
							parColumnLetter: "E",
							parRowNumber: intMatrixSheet_RowIndex,
							parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("E"))),
							parCellDatatype: CellValues.String);
						//--- Matrix --- Requirement Row --- Column F --------------------------------
						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objMatrixWorksheetPart,
							parColumnLetter: "F",
							parRowNumber: intMatrixSheet_RowIndex,
							parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("F"))),
							parCellDatatype: CellValues.String);

							dictionaryMatrixComments.Add("F" + intMatrixSheet_RowIndex, objRequirement.RequirementText);

						//--- Matrix --- Requirement Row --- Column G --------------------------------
						intSharedStringIndex = oxmlWorkbook.InsertSharedStringItem(
							parText2Insert: Properties.AppResources.Workbook_CRM_Matrix_RowType_Requirement, parShareStringPart: objSharedStringTablePart);
						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objMatrixWorksheetPart,
							parColumnLetter: "G",
							parRowNumber: intMatrixSheet_RowIndex,
							parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("G"))),
							parCellDatatype: CellValues.SharedString,
							parCellcontents: intSharedStringIndex.ToString());
						//--- Matrix --- Requirement Row --- Column H --------------------------------
						intSharedStringIndex = oxmlWorkbook.InsertSharedStringItem(parText2Insert: objRequirement.ComplianceStatus, 
							parShareStringPart: objSharedStringTablePart);
						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objMatrixWorksheetPart,
							parColumnLetter: "H",
							parRowNumber: intMatrixSheet_RowIndex,
							parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("H"))),
							parCellDatatype: CellValues.SharedString,
							parCellcontents: intSharedStringIndex.ToString());
						//--- Matrix --- Requirement Row --- Column I --------------------------------
						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objMatrixWorksheetPart,
							parColumnLetter: "I",
							parRowNumber: intMatrixSheet_RowIndex,
							parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("I"))),
							parCellDatatype: CellValues.String);
						//--- Matrix --- Requirement Row --- Column J --------------------------------
						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objMatrixWorksheetPart,
							parColumnLetter: "J",
							parRowNumber: intMatrixSheet_RowIndex,
							parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("J"))),
							parCellDatatype: CellValues.String,
							parCellcontents: objRequirement.SourceReference);
						//--- Matrix --- Requirement Row --- Column K --------------------------------
						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objMatrixWorksheetPart,
							parColumnLetter: "K",
							parRowNumber: intMatrixSheet_RowIndex,
							parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("K"))),
							parCellDatatype: CellValues.String);
						//--- Matrix --- Requirement Row --- Column L --------------------------------
						intHyperlinkCounter += 1;
						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objMatrixWorksheetPart,
							parColumnLetter: "L",
							parRowNumber: intMatrixSheet_RowIndex,
							parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("L"))),
							parCellDatatype: CellValues.Number,
							parCellcontents: objRequirement.ID.ToString(),
							parHyperlinkCounter: intHyperlinkCounter,
							parHyperlinkURL: Properties.AppResources.SharePointURL +
								Properties.AppResources.List_MappingRequirements +
								Properties.AppResources.EditFormURI + objRequirement.ID.ToString());
						//--- Matrix --- Requirement Row --- Column M --------------------------------
						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objMatrixWorksheetPart,
							parColumnLetter: "M",
							parRowNumber: intMatrixSheet_RowIndex,
							parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("M"))),
							parCellDatatype: CellValues.String);
						//--- Matrix --- Requirement Row --- Column N --------------------------------
						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objMatrixWorksheetPart,
							parColumnLetter: "N",
							parRowNumber: intMatrixSheet_RowIndex,
							parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("N"))),
							parCellDatatype: CellValues.String);

						//===============================================================
						// Obtain all Mapping Risk for the specified Mapping Requirement
						//===============================================================
						
						listMappingRisks.Clear();
						listMappingRisks = MappingRisk.ObtainListOfObjects(
							parDatacontextSDDP: datacontexSDDP,
							parMappingRequirementID: objRequirement.ID);
						
						// Check if any Mapping Risks were found
						if(listMappingRisks.Count != 0)
							{
							// Process all the Mapping Risks for the specific Service Requirement
							foreach(MappingRisk objRisk in listMappingRisks)
								{
								Console.WriteLine("\t\t\t + Risk: {0} - {1}", objRisk.ID, objRisk.Title);
								intMatrixSheet_RowIndex += 1;
								//--- Matrix --- Risk Row --- Column A --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "A",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("A"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Risk Row --- Column B --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "B",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("B"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Risk Row --- Column C --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "C",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("C"))),
									parCellDatatype: CellValues.String,
									parCellcontents: objRisk.Title);
								//--- Matrix --- Risk Row --- Column D --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "D",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("D"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Risk Row --- Column E --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "E",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("E"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Risk Row --- Column F --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "F",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("F"))),
									parCellDatatype: CellValues.String);

								dictionaryMatrixComments.Add("F" + intMatrixSheet_RowIndex, objRisk.Statement);

								//--- Matrix --- Risk Row --- Column G --------------------------------
								intSharedStringIndex = oxmlWorkbook.InsertSharedStringItem(
									parText2Insert: Properties.AppResources.Workbook_CRM_Matrix_RowType_Risk, parShareStringPart: objSharedStringTablePart);
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "G",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("G"))),
									parCellDatatype: CellValues.SharedString,
									parCellcontents: intSharedStringIndex.ToString());
								//--- Matrix --- Risk Row --- Column H --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "H",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("H"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Risk Row --- Column I --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "I",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("I"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Risk Row --- Column J --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "J",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("J"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Risk Row --- Column K --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "K",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("K"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Risk Row --- Column L --------------------------------
								intHyperlinkCounter += 1;
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "L",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("L"))),
									parCellDatatype: CellValues.Number,
									parCellcontents: objRisk.ID.ToString(),
									parHyperlinkCounter: intHyperlinkCounter,
									parHyperlinkURL: Properties.AppResources.SharePointURL +
										Properties.AppResources.List_MappingRisks +
										Properties.AppResources.EditFormURI + objRisk.ID.ToString());
								//--- Matrix --- Risk Row --- Column M --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "M",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("M"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Risk Row --- Column N --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "N",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("N"))),
									parCellDatatype: CellValues.String);

								//------------------------------------------------------
								// also populate a row on the Risks worksheet
								//--- Risks Columns on a Requirement Break -------------
								// checked if the Requirment changed...
								if(intRequirementBreakID_forRisks != objRequirement.ID)
									{
									intRisksSheet_RowIndex += 1;
									// Write the Requirement in the first column
									// --- Risks (new Requirement) --- Column A ---------- 
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objRisksWorksheetPart,
										parColumnLetter: "A",
										parRowNumber: intRisksSheet_RowIndex,
										parStyleId: (UInt32Value)(listRisksColumnStyles.ElementAt(aWorkbook.GetColumnNumber("A"))),
										parCellDatatype: CellValues.String,
										parCellcontents: objRequirement.Title);
									//--- Risks (new Requirement) --- Column B -----------
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objRisksWorksheetPart,
										parColumnLetter: "B",
										parRowNumber: intRisksSheet_RowIndex,
										parStyleId: (UInt32Value)(listRisksColumnStyles.ElementAt(aWorkbook.GetColumnNumber("B"))),
										parCellDatatype: CellValues.String);
									//--- Risks (new Requirement) --- Column C -----------
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objRisksWorksheetPart,
										parColumnLetter: "C",
										parRowNumber: intRisksSheet_RowIndex,
										parStyleId: (UInt32Value)(listRisksColumnStyles.ElementAt(aWorkbook.GetColumnNumber("C"))),
										parCellDatatype: CellValues.String);
									//--- Risks (new Requirement) --- Column D -----------
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objRisksWorksheetPart,
										parColumnLetter: "D",
										parRowNumber: intRisksSheet_RowIndex,
										parStyleId: (UInt32Value)(listRisksColumnStyles.ElementAt(aWorkbook.GetColumnNumber("D"))),
										parCellDatatype: CellValues.String);
									//--- Risks (new Requirement) --- Column E -----------
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objRisksWorksheetPart,
										parColumnLetter: "E",
										parRowNumber: intRisksSheet_RowIndex,
										parStyleId: (UInt32Value)(listRisksColumnStyles.ElementAt(aWorkbook.GetColumnNumber("E"))),
										parCellDatatype: CellValues.String);
									//--- Risks (new Requirement) --- Column F -----------
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objRisksWorksheetPart,
										parColumnLetter: "F",
										parRowNumber: intRisksSheet_RowIndex,
										parStyleId: (UInt32Value)(listRisksColumnStyles.ElementAt(aWorkbook.GetColumnNumber("F"))),
										parCellDatatype: CellValues.String);
									//--- Risks (new Requirement) --- Column G -----------
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objRisksWorksheetPart,
										parColumnLetter: "G",
										parRowNumber: intRisksSheet_RowIndex,
										parStyleId: (UInt32Value)(listRisksColumnStyles.ElementAt(aWorkbook.GetColumnNumber("G"))),
										parCellDatatype: CellValues.String);

									intRequirementBreakID_forRisks = objRequirement.ID;
                                             } //if(intRequirementBreakID_forRisks != objRequirement.ID)
								// Write the Risk to the Risks Worksheet
								//--- Risks - already populated Requirement (--- Column A ---)
								intRisksSheet_RowIndex += 1;
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objRisksWorksheetPart,
									parColumnLetter: "A",
									parRowNumber: intRisksSheet_RowIndex,
									parStyleId: (UInt32Value)(listRisksColumnStyles.ElementAt(aWorkbook.GetColumnNumber("A"))),
									parCellDatatype: CellValues.String);
								//--- Risks - already populated Requirement (--- Column B ---) 
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objRisksWorksheetPart,
									parColumnLetter: "B",
									parRowNumber: intRisksSheet_RowIndex,
									parStyleId: (UInt32Value)(listRisksColumnStyles.ElementAt(aWorkbook.GetColumnNumber("B"))),
									parCellDatatype: CellValues.Number,
									parCellcontents: objRisk.ID.ToString());
								//--- Risks - already populated Requirement (--- Column C ---) 
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objRisksWorksheetPart,
									parColumnLetter: "C",
									parRowNumber: intRisksSheet_RowIndex,
									parStyleId: (UInt32Value)(listRisksColumnStyles.ElementAt(aWorkbook.GetColumnNumber("C"))),
									parCellDatatype: CellValues.String,
									parCellcontents: objRisk.Title);
								//--- Risks - already populated Requirement (--- Column D ---) 
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objRisksWorksheetPart,
									parColumnLetter: "D",
									parRowNumber: intRisksSheet_RowIndex,
									parStyleId: (UInt32Value)(listRisksColumnStyles.ElementAt(aWorkbook.GetColumnNumber("D"))),
									parCellDatatype: CellValues.String,
									parCellcontents: objRisk.Statement);
								//--- Risks - already populated Requirement (--- Column E ---) 
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objRisksWorksheetPart,
									parColumnLetter: "E",
									parRowNumber: intRisksSheet_RowIndex,
									parStyleId: (UInt32Value)(listRisksColumnStyles.ElementAt(aWorkbook.GetColumnNumber("E"))),
									parCellDatatype: CellValues.String,
									parCellcontents: objRisk.Status);
								//--- Risks - already populated Requirement (--- Column F ---) 
								intSharedStringIndex = oxmlWorkbook.InsertSharedStringItem(
									parText2Insert: objRisk.Exposure, parShareStringPart: objSharedStringTablePart);
								oxmlWorkbook.PopulateCell(
										parWorksheetPart: objRisksWorksheetPart,
										parColumnLetter: "F",
										parRowNumber: intRisksSheet_RowIndex,
										parStyleId: (UInt32Value)(listRisksColumnStyles.ElementAt(aWorkbook.GetColumnNumber("F"))),
										parCellDatatype: CellValues.SharedString,
										parCellcontents: intSharedStringIndex.ToString());
								//--- Risks - already populated Requirement (--- Column G ---) 
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objRisksWorksheetPart,
									parColumnLetter: "G",
									parRowNumber: intRisksSheet_RowIndex,
									parStyleId: (UInt32Value)(listRisksColumnStyles.ElementAt(aWorkbook.GetColumnNumber("G"))),
									parCellDatatype: CellValues.String,
									parCellcontents: objRisk.Mitigation);
								} //foreach(Mappingrisk objMappingRisk in listMappingRisks)
							} // if(listMappingRisks.Count != 0)

						//=====================================================================
						// Obtain all Mapping Assumptions for the specified Mapping Requirement
						try
							{
							listMappingAssumptions.Clear();
							listMappingAssumptions = MappingAssumption.ObtainListOfObjects(
								parDatacontextSDDP: datacontexSDDP,
								parMappingRequirementID: objRequirement.ID);
							}
						catch(DataEntryNotFoundException)
							{
							// ignore if there are no Mapping Assumptions
							}

						// Check if any Mapping Assumptions were found
						if(listMappingAssumptions.Count != 0)
							{
							// Process all the Mapping Assumptions for the specific Service Requirement
							foreach(MappingAssumption objAssumption in listMappingAssumptions)
								{
								Console.WriteLine("\t\t\t + Assumption: {0} - {1}", objAssumption.ID, objAssumption.Title);
								// Write the Mapping Assumptions to the Workbook as a String
								intMatrixSheet_RowIndex += 1;
								//--- Matrix --- Assumption Row --- Column A --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "A",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("A"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Assumption Row --- Column B --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "B",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("B"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Assumption Row --- Column C --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "C",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("C"))),
									parCellDatatype: CellValues.String,
									parCellcontents: objAssumption.Title);
								//--- Matrix --- Assumption Row --- Column D --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "D",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("D"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Assumption Row --- Column E --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "E",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("E"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Assumption Row --- Column F --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "F",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("F"))),
									parCellDatatype: CellValues.String);

								dictionaryMatrixComments.Add("F" + intMatrixSheet_RowIndex, objAssumption.Description);

								//--- Matrix --- Assumption Row --- Column G --------------------------------
								intSharedStringIndex = oxmlWorkbook.InsertSharedStringItem(
									parText2Insert: Properties.AppResources.Workbook_CRM_Matrix_RowType_Assumption, 
									parShareStringPart: objSharedStringTablePart);
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "G",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("G"))),
									parCellDatatype: CellValues.SharedString,
									parCellcontents: intSharedStringIndex.ToString());
								//--- Matrix --- Assumption Row --- Column H --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "H",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("H"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Assumption Row --- Column I --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "I",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("I"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Assumption Row --- Column J --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "J",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("J"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Assumption Row --- Column K --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "K",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("K"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Assumption Row --- Column L --------------------------------
								intHyperlinkCounter += 1;
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "L",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("L"))),
									parCellDatatype: CellValues.Number,
									parCellcontents: objAssumption.ID.ToString(),
									parHyperlinkCounter: intHyperlinkCounter,
									parHyperlinkURL: Properties.AppResources.SharePointURL +
										Properties.AppResources.List_MappingAssumptions +
										Properties.AppResources.EditFormURI + objAssumption.ID.ToString());
								//--- Matrix --- Assumption Row --- Column M --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "M",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("M"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Assumption Row --- Column N --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "N",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("N"))),
									parCellDatatype: CellValues.String);

								//------------------------------------------------------
								//--- also populate the Assumptions worksheet 
								//--- Assumptions Columns on a Requirement Break --------
								// checked if the Requirment changed...
								if(intRequirementBreakID_forAssumptions != objRequirement.ID)
									{
									intAssumptionsSheet_RowIndex += 1;
									// Write the Requirement in the first column and just copy the styles for the rest
									// --- Assumptions (new Requirement) --- Column A ---------- 
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objAssumptionsWorksheetPart,
										parColumnLetter: "A",
										parRowNumber: intAssumptionsSheet_RowIndex,
										parStyleId: (UInt32Value)(listAssumptionsColumnStyles.ElementAt(aWorkbook.GetColumnNumber("A"))),
										parCellDatatype: CellValues.String,
										parCellcontents: objRequirement.Title);
									//--- Assumptions (new Requirement) --- Column B -----------
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objAssumptionsWorksheetPart,
										parColumnLetter: "B",
										parRowNumber: intAssumptionsSheet_RowIndex,
										parStyleId: (UInt32Value)(listAssumptionsColumnStyles.ElementAt(aWorkbook.GetColumnNumber("B"))),
										parCellDatatype: CellValues.String);
									//--- Assumptions (new Requirement) --- Column C -----------
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objAssumptionsWorksheetPart,
										parColumnLetter: "C",
										parRowNumber: intAssumptionsSheet_RowIndex,
										parStyleId: (UInt32Value)(listAssumptionsColumnStyles.ElementAt(aWorkbook.GetColumnNumber("C"))),
										parCellDatatype: CellValues.String);
									//--- Assumptions (new Requirement) --- Column D -----------
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objAssumptionsWorksheetPart,
										parColumnLetter: "D",
										parRowNumber: intAssumptionsSheet_RowIndex,
										parStyleId: (UInt32Value)(listAssumptionsColumnStyles.ElementAt(aWorkbook.GetColumnNumber("D"))),
										parCellDatatype: CellValues.String);

									intRequirementBreakID_forAssumptions = objRequirement.ID;
                                             } //if(intRequirementBreakID_forAssumptions != objRequirement.ID)
								// Write the Assumption detail to the Assumptions Worksheet
									{
									//--- Assumptions - already populated Requirement (--- Column A ---)
									intAssumptionsSheet_RowIndex += 1;
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objAssumptionsWorksheetPart,
										parColumnLetter: "A",
										parRowNumber: intAssumptionsSheet_RowIndex,
										parStyleId: (UInt32Value)(listAssumptionsColumnStyles.ElementAt(aWorkbook.GetColumnNumber("A"))),
										parCellDatatype: CellValues.String);
									//--- Assumptions - already populated Requirement (--- Column B ---) 
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objAssumptionsWorksheetPart,
										parColumnLetter: "B",
										parRowNumber: intAssumptionsSheet_RowIndex,
										parStyleId: (UInt32Value)(listRisksColumnStyles.ElementAt(aWorkbook.GetColumnNumber("B"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: objAssumption.ID.ToString());
									//--- Assumptions - already populated Requirement (--- Column C ---) 
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objAssumptionsWorksheetPart,
										parColumnLetter: "C",
										parRowNumber: intAssumptionsSheet_RowIndex,
										parStyleId: (UInt32Value)(listAssumptionsColumnStyles.ElementAt(aWorkbook.GetColumnNumber("C"))),
										parCellDatatype: CellValues.String,
										parCellcontents: objAssumption.Title);
									//--- Assumptions - already populated Requirement (--- Column D ---) 
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objAssumptionsWorksheetPart,
										parColumnLetter: "D",
										parRowNumber: intAssumptionsSheet_RowIndex,
										parStyleId: (UInt32Value)(listAssumptionsColumnStyles.ElementAt(aWorkbook.GetColumnNumber("D"))),
										parCellDatatype: CellValues.String,
										parCellcontents: objAssumption.Description);
									} // No break in Requirement, write the Assumption values
								} //foreach(MappingAssumption objMappingAssumption in listMappingAssumptions)
							} // if(listMappingAssumptions.Count != 0)

						//-----------------------------------------------------------------------
						// Obtain all Mapping Deliverables for the specified Mapping Requirement
						
						listMappingDeliverables.Clear();
						listMappingDeliverables = MappingDeliverable.ObtainListOfObjects(
							parDatacontextSDDP: datacontexSDDP,
							parMappingRequirementID: objRequirement.ID);

						// Check if any Mapping Deliverables were found
						if(listMappingDeliverables.Count != 0)
							{
							// Process all the Mapping Deliverables for the specific Service Requirement
							foreach(MappingDeliverable objMappingDeliverable in listMappingDeliverables)
								{
								Console.WriteLine("\t\t\t + DRM: {0} - {1}", objMappingDeliverable.ID, objMappingDeliverable.Title);
								intMatrixSheet_RowIndex += 1;
								//--- Matrix --- Deliverable Row --- Column A --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "A",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("A"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Deliverable Row --- Column B --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "B",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("B"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Deliverable Row --- Column C --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "C",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("C"))),
									parCellDatatype: CellValues.String,
									parCellcontents: objMappingDeliverable.Title);
								//--- Matrix --- Deliverable Row --- Column D --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "D",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("D"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Deliverable Row --- Column E --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "E",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("E"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Deliverable Row --- Column F --------------------------------
								if(objMappingDeliverable.NewDeliverable)
									{
									intSharedStringIndex = oxmlWorkbook.InsertSharedStringItem(
										parText2Insert: Properties.AppResources.Workbook_CRM_Matrix_NewColumn_Text,
										parShareStringPart: objSharedStringTablePart);
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objMatrixWorksheetPart,
										parColumnLetter: "F",
										parRowNumber: intMatrixSheet_RowIndex,
										parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("F"))),
										parCellDatatype: CellValues.SharedString,
										parCellcontents: intSharedStringIndex.ToString());
									if(objMappingDeliverable.NewRequirement != null)
										{
										dictionaryMatrixComments.Add("F" + intMatrixSheet_RowIndex,
											objMappingDeliverable.NewRequirement);
										}
									}
								else // if it is an EXISTING deliverable...
									{
									//--- Matrix --- Deliverable Row --- Column F -----------------------------
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objMatrixWorksheetPart,
										parColumnLetter: "F",
										parRowNumber: intMatrixSheet_RowIndex,
										parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("F"))),
										parCellDatatype: CellValues.String);
									strTextDescription = "";
									intLayer0upDeliverableID = objMappingDeliverable.MappedDeliverable.ID;
									if(objMappingDeliverable.MappedDeliverable.ContentPredecessorDeliverableID == null)
										{
										intLayer1upDeliverableID = null;
										intLayer2upDeliverableID = null;
										}
									else
										{
										intLayer1upDeliverableID = objMappingDeliverable.MappedDeliverable.ContentPredecessorDeliverableID;
										if(objMappingDeliverable.MappedDeliverable.Layer1up.ContentPredecessorDeliverableID == null)
											{
											intLayer2upDeliverableID = null;
											}
										else
											{
											intLayer2upDeliverableID =
												objMappingDeliverable.MappedDeliverable.Layer1up.ContentPredecessorDeliverableID;
											}
										}
									if(intLayer2upDeliverableID != null)
										{
										if(objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.CSDdescription != null)
											{
											strTextDescription = HTMLdecoder.CleanHTMLstring
												(objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.CSDdescription);
											}
										}
									if(intLayer1upDeliverableID != null)
										{
										if(objMappingDeliverable.MappedDeliverable.Layer1up.CSDdescription != null)
											{
											strTextDescription = strTextDescription + HTMLdecoder.CleanHTMLstring
												(objMappingDeliverable.MappedDeliverable.Layer1up.CSDdescription);
											}
										}

									if(objMappingDeliverable.MappedDeliverable.CSDdescription != null)
										{
										strTextDescription = strTextDescription + HTMLdecoder.CleanHTMLstring
												(objMappingDeliverable.MappedDeliverable.CSDdescription);
										}
									// Insert the Deliverable CSD Description
									if(strTextDescription != "")
										{
										dictionaryMatrixComments.Add("F" + intMatrixSheet_RowIndex, strTextDescription);
										}
									}
								//--- Matrix --- Deliverable Row --- Column G --------------------------------
								intSharedStringIndex = oxmlWorkbook.InsertSharedStringItem(
									parText2Insert: Properties.AppResources.Workbook_CRM_Matrix_RowType_Deliverable,
									parShareStringPart: objSharedStringTablePart);
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "G",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("G"))),
									parCellDatatype: CellValues.SharedString,
									parCellcontents: intSharedStringIndex.ToString());
								//--- Matrix --- Deliverable Row --- Column H --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "H",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("H"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Deliverable Row --- Column I --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "I",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("I"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Deliverable Row --- Column J --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "J",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("J"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Deliverable Row --- Column K --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "K",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("K"))),
									parCellDatatype: CellValues.String);
								//--- Matrix --- Deliverable Row --- Column L --------------------------------
								intHyperlinkCounter += 1;
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "L",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("L"))),
									parCellDatatype: CellValues.Number,
									parCellcontents: objMappingDeliverable.ID.ToString(),
									parHyperlinkCounter: intHyperlinkCounter,
									parHyperlinkURL: Properties.AppResources.SharePointURL +
										Properties.AppResources.List_MappingDeliverables +
										Properties.AppResources.EditFormURI + objMappingDeliverable.ID.ToString());
								//--- Matrix --- Deliverable Row --- Column M --------------------------------
								if(objMappingDeliverable.NewDeliverable)
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objMatrixWorksheetPart,
										parColumnLetter: "M",
										parRowNumber: intMatrixSheet_RowIndex,
										parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("M"))),
										parCellDatatype: CellValues.Number);
									}
								else // an EXISTING deliverable add the Deliverable reference
									{
									//--- Matrix --- Deliverable Row --- Column M -------------------------------
									intHyperlinkCounter += 1;
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objMatrixWorksheetPart,
										parColumnLetter: "M",
										parRowNumber: intMatrixSheet_RowIndex,
										parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("M"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: objMappingDeliverable.MappedDeliverable.ID.ToString(),
										parHyperlinkCounter: intHyperlinkCounter,
									parHyperlinkURL: Properties.AppResources.SharePointURL +
										Properties.AppResources.List_DeliverablesURI +
										Properties.AppResources.EditFormURI + objMappingDeliverable.MappedDeliverable.ID.ToString()
										);
									}
								//--- Matrix --- Deliverable Row --- Column N --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "N",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("N"))),
									parCellDatatype: CellValues.String);

								
								//====================================================================
								// Obtain all Service Levels for the specified Deliverable Requirement
								try
									{
									listMappingServiceLevels.Clear();
									listMappingServiceLevels = MappingServiceLevel.ObtainListOfObjects(
										parDatacontextSDDP: datacontexSDDP,
										parMappingDeliverableID: objMappingDeliverable.ID);
									}
								catch(DataEntryNotFoundException)
									{
									// ignore if there are no Mapping Deliverables
									}
								// Check if any Mapping Service Levels were found
								if(listMappingServiceLevels.Count != 0)
									{
									// Process all the Mapping Deliverables for the specific Service Requirement
									foreach(MappingServiceLevel objMappingServiceLevel in listMappingServiceLevels)
										{
										Console.WriteLine("\t\t\t\t + ServiceLevel: {0} - {1}", objMappingServiceLevel.ID, objMappingServiceLevel.Title);
										// Write the Mapping Service Level to the Workbook as a String
										intMatrixSheet_RowIndex += 1;
										// Insert the Service Level 
										//--- Matrix --- Service Level Row --- Column A --------------------------------
										oxmlWorkbook.PopulateCell(
											parWorksheetPart: objMatrixWorksheetPart,
											parColumnLetter: "A",
											parRowNumber: intMatrixSheet_RowIndex,
											parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("A"))),
											parCellDatatype: CellValues.String);
										//--- Matrix --- Service Level Row --- Column B --------------------------------
										oxmlWorkbook.PopulateCell(
											parWorksheetPart: objMatrixWorksheetPart,
											parColumnLetter: "B",
											parRowNumber: intMatrixSheet_RowIndex,
											parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("B"))),
											parCellDatatype: CellValues.String);
										//--- Matrix --- Service Level Row --- Column C --------------------------------
										oxmlWorkbook.PopulateCell(
											parWorksheetPart: objMatrixWorksheetPart,
											parColumnLetter: "C",
											parRowNumber: intMatrixSheet_RowIndex,
											parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("C"))),
											parCellDatatype: CellValues.String);
										//--- Matrix --- Service Level Row --- Column D --------------------------------
										oxmlWorkbook.PopulateCell(
											parWorksheetPart: objMatrixWorksheetPart,
											parColumnLetter: "D",
											parRowNumber: intMatrixSheet_RowIndex,
											parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("D"))),
											parCellDatatype: CellValues.String,
											parCellcontents: objMappingServiceLevel.Title);
										//--- Matrix --- Service Level Row --- Column E --------------------------------
										oxmlWorkbook.PopulateCell(
											parWorksheetPart: objMatrixWorksheetPart,
											parColumnLetter: "E",
											parRowNumber: intMatrixSheet_RowIndex,
											parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("E"))),
											parCellDatatype: CellValues.String);
										//--- Matrix --- Service Level Row --- Column F --------------------------------
										if(objMappingServiceLevel.NewServiceLevel)
											{
											intSharedStringIndex = oxmlWorkbook.InsertSharedStringItem(
												parText2Insert: Properties.AppResources.Workbook_CRM_Matrix_NewColumn_Text,
												parShareStringPart: objSharedStringTablePart);
											oxmlWorkbook.PopulateCell(
												parWorksheetPart: objMatrixWorksheetPart,
												parColumnLetter: "F",
												parRowNumber: intMatrixSheet_RowIndex,
												parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("F"))),
												parCellDatatype: CellValues.SharedString,
												parCellcontents: intSharedStringIndex.ToString());

											if(objMappingServiceLevel.RequirementText != null)
												{
												dictionaryMatrixComments.Add("F" + intMatrixSheet_RowIndex,
													objMappingServiceLevel.RequirementText);
												}
											}
										else // if it is an EXISTING ServiceLevel...
											{
											// --- Matrix --- Service Level Row --- Column F ---------------------------
											oxmlWorkbook.PopulateCell(
												parWorksheetPart: objMatrixWorksheetPart,
												parColumnLetter: "F",
												parRowNumber: intMatrixSheet_RowIndex,
												parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("F"))),
												parCellDatatype: CellValues.String);
											
											if(objMappingServiceLevel.MappedServiceLevel.CSDdescription != null)
												{
												dictionaryMatrixComments.Add("F" + intMatrixSheet_RowIndex, 
													objMappingServiceLevel.MappedServiceLevel.CSDdescription);
												}
											}
										//--- Matrix --- Service Level Row --- Column G --------------------------------
										intSharedStringIndex = oxmlWorkbook.InsertSharedStringItem(
											parText2Insert: Properties.AppResources.Workbook_CRM_Matrix_RowType_ServiceLevel,
											parShareStringPart: objSharedStringTablePart);
										oxmlWorkbook.PopulateCell(
											parWorksheetPart: objMatrixWorksheetPart,
											parColumnLetter: "G",
											parRowNumber: intMatrixSheet_RowIndex,
											parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("G"))),
											parCellDatatype: CellValues.SharedString,
											parCellcontents: intSharedStringIndex.ToString());
										//--- Matrix --- Service Level Row --- Column H --------------------------------
										oxmlWorkbook.PopulateCell(
											parWorksheetPart: objMatrixWorksheetPart,
											parColumnLetter: "H",
											parRowNumber: intMatrixSheet_RowIndex,
											parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("H"))),
											parCellDatatype: CellValues.String);
										//--- Matrix --- Service Level Row --- Column I --------------------------------
										oxmlWorkbook.PopulateCell(
											parWorksheetPart: objMatrixWorksheetPart,
											parColumnLetter: "I",
											parRowNumber: intMatrixSheet_RowIndex,
											parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("I"))),
											parCellDatatype: CellValues.String);
										//--- Matrix --- Service Level Row --- Column J --------------------------------
										oxmlWorkbook.PopulateCell(
											parWorksheetPart: objMatrixWorksheetPart,
											parColumnLetter: "J",
											parRowNumber: intMatrixSheet_RowIndex,
											parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("J"))),
											parCellDatatype: CellValues.String);
										//--- Matrix --- Service Level Row --- Column K --------------------------------
										oxmlWorkbook.PopulateCell(
											parWorksheetPart: objMatrixWorksheetPart,
											parColumnLetter: "K",
											parRowNumber: intMatrixSheet_RowIndex,
											parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("K"))),
											parCellDatatype: CellValues.String);
										//--- Matrix --- Service Level Row --- Column L --------------------------------
										intHyperlinkCounter += 1;
										oxmlWorkbook.PopulateCell(
											parWorksheetPart: objMatrixWorksheetPart,
											parColumnLetter: "L",
											parRowNumber: intMatrixSheet_RowIndex,
											parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("L"))),
											parCellDatatype: CellValues.Number,
											parCellcontents: objMappingServiceLevel.ID.ToString(),
											parHyperlinkCounter: intHyperlinkCounter,
											parHyperlinkURL: Properties.AppResources.SharePointURL +
												Properties.AppResources.List_MappingServiceLevels +
												Properties.AppResources.EditFormURI + objMappingServiceLevel.ID.ToString());
										//--- Matrix --- Service Level Row --- Column M --------------------------------
										oxmlWorkbook.PopulateCell(
											parWorksheetPart: objMatrixWorksheetPart,
											parColumnLetter: "M",
											parRowNumber: intMatrixSheet_RowIndex,
											parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("M"))),
											parCellDatatype: CellValues.String);

										//--- Matrix --- Service Level Row --- Column N --------------------------------
										if(objMappingServiceLevel.NewServiceLevel)
											{
											oxmlWorkbook.PopulateCell(
												parWorksheetPart: objMatrixWorksheetPart,
												parColumnLetter: "N",
												parRowNumber: intMatrixSheet_RowIndex,
												parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("N"))),
												parCellDatatype: CellValues.Number);
											}
										else // an EXISTING Service Level add the Deliverable reference
											{
											//--- Matrix --- Service Level Row --- Column N -------------------------
											intHyperlinkCounter += 1;
											oxmlWorkbook.PopulateCell(
												parWorksheetPart: objMatrixWorksheetPart,
												parColumnLetter: "N",
												parRowNumber: intMatrixSheet_RowIndex,
												parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("N"))),
												parCellDatatype: CellValues.Number,
												parCellcontents: objMappingServiceLevel.MappedServiceLevel.ID.ToString(),
												parHyperlinkCounter: intHyperlinkCounter,
											parHyperlinkURL: Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceLevelsURI +
												Properties.AppResources.EditFormURI + objMappingServiceLevel.MappedServiceLevel.ID.ToString()
												);
											}
										} // foreach(MappingServiceLevel objMappingServiceLevel in listMappingServiceLevels)
									} // if(listMappingServiceLevels.Count != 0)
								} // foreach(MappingDeliverable objMappingDeliverable in listMappingDeliverables)
							} // if(listMappingDeliverables.Count != 0)
						} // foreach(MappingRequirement objRequirement in listMappingRequirements)
					} //foreach(MappingServiceTower objTower in listMappingTowers)

Save_and_Close_Document:
//===============================================================
				if(dictionaryMatrixComments.Count() > 0)
					{
					// Now insert all the Comments
					//aWorkbook.InsertWorksheetComments(
					//	parWorksheetPart: objMatrixWorksheetPart,
					//	parDictionaryOfComments: dictionaryMatrixComments);
					}
				
				//Validate the document with OpenXML validator
				OpenXmlValidator objOXMLvalidator = new OpenXmlValidator(fileFormat: FileFormatVersions.Office2010);
				int errorCount = 0;
				Console.WriteLine("\n\rValidating document....");
				foreach(ValidationErrorInfo validationError in objOXMLvalidator.Validate(objSpreadsheetDocument))
					{
					errorCount += 1;
					Console.WriteLine("------------- # {0} -------------", errorCount);
					Console.WriteLine("Error ID...........: {0}", validationError.Id);
					Console.WriteLine("Description........: {0}", validationError.Description);
					Console.WriteLine("Error Type.........: {0}", validationError.ErrorType);
					Console.WriteLine("Error Part.........: {0}", validationError.Part.Uri);
					Console.WriteLine("Error Related Part.: {0}", validationError.RelatedPart);
					Console.WriteLine("Error Path.........: {0}", validationError.Path.XPath);
					Console.WriteLine("Error Path PartUri.: {0}", validationError.Path.PartUri);
					Console.WriteLine("Error Node.........: {0}", validationError.Node);
					Console.WriteLine("Error Related Node.: {0}", validationError.RelatedNode);
					Console.WriteLine("Node Local Name....: {0}", validationError.Node.LocalName);
					}

				Console.WriteLine("Workbook generation completed, saving and closing the document.");
				// Save and close the Document
				objSpreadsheetDocument.Close();

				Console.WriteLine(
					"Generation started...: {0} \nGeneration completed: {1} \n Durarion..........: {2}",
					timeStarted, DateTime.Now, (DateTime.Now - timeStarted));

				} // end Try
			catch(ArgumentException exc)
				{
				Console.WriteLine("Exception: {0} - {1}", exc.HResult, exc.Message);
				return false;
				//TODO: raise the error
				}
			catch(Exception exc)
				{
				Console.WriteLine("Exception: {0} - {1}", exc.HResult, exc.Message);
				return false;
				}
			
			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}
		}
	}