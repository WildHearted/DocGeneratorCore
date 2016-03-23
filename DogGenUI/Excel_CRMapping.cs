using System;
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
	class Client_Requirements_Mapping_Workbook:Workbook
		{
		private bool _client_Requirements_Mapping_Workbook = false;
		public bool Client_Requirements_Mapping_Wbk
			{
			get{return this._client_Requirements_Mapping_Workbook;}
			set{this._client_Requirements_Mapping_Workbook = value;}
			}


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
			Console.WriteLine("\t Begin to generate {0}", this.DocumentType);
			DateTime timeStarted = DateTime.Now;
			//string hyperlinkImageRelationshipID = "";
			string documentCollection_HyperlinkURL = "";
			int intHyperlinkCounter = 9;
			string currentListURI = "";
			string currentHyperlinkViewEditURI = "";
			string currentContentLayer = "None";
			Cell objCell = new Cell();
			int intStringIndex = 0;
			//Column Value Variables
			string strColumnTowerOfServices = "A";
			string strColumnRequirement = "B";
			string strColumnDelRiskAss = "C";
			string strColumnServiceLevel = "D";
			string strColumnE = "E";
			string strColumnNew = "F";
			string strColumnRowType = "G";
			string strColumnComliance = "H";
			string strColumnI = "I";
			string strColumnSourceReference = "J";
			string strColumnK = "K";
			string strColumnMappingReference = "L";
			string strColumnDeliverableReference = "M";
			string strColumnServiceLevelReference = "N";

			//Content Layering Variables
			int? layer0upDeliverableID;
			int? layer1upDeliverableID;
			int? layer2upDeliverableID;
			string strTextDescription = "";

			//Worksheet Row Index Variables
			UInt16 intMatrixSheet_RowIndex = 6;
			UInt16 intRisksSheet_RowIndex = 2;
			UInt16 intAssumptionsSheet_RowIndex = 2;
			//int hyperlinkCounter = 4;
			string strErrorText = "";
			if(this.HyperlinkEdit)
				{
				documentCollection_HyperlinkURL = Properties.AppResources.SharePointSiteURL +
					Properties.AppResources.List_DocumentCollectionLibraryURI +
					Properties.AppResources.EditFormURI + this.DocumentCollectionID;
				currentHyperlinkViewEditURI = Properties.AppResources.EditFormURI;
				}
			if(this.HyperlinkView)
				{
				documentCollection_HyperlinkURL = Properties.AppResources.SharePointSiteURL +
					Properties.AppResources.List_DocumentCollectionLibraryURI +
					Properties.AppResources.DisplayFormURI + this.DocumentCollectionID;
				currentHyperlinkViewEditURI = Properties.AppResources.DisplayFormURI;
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

				// construct the CommentsList to which Comments objects can be appended
				CommentList objMatrixCommentList = new CommentList();

				//WorksheetCommentsPart objMatrixCommentsPart = (WorksheetCommentsPart)(objMatrixWorksheetPart.GetPartsOfType<WorksheetCommentsPart>().First());

				// obtain the Risks Worksheet in the Workbook.
				Sheet objRisksWorksheet = objWorkbookPart.Workbook.Descendants<Sheet>().Where(sht => sht.Name == Properties.AppResources.Workbook_CRM_Worksheet_Name_Risks).FirstOrDefault();
				if(objRisksWorksheet == null)
					{
					throw new ArgumentException("The " + Properties.AppResources.Workbook_CRM_Worksheet_Name_Risks +
						" worksheet could not be loacated in the workbook.");
					}
				WorksheetPart objRisksWorksheetPart = (WorksheetPart)(objWorkbookPart.GetPartById(objRisksWorksheet.Id));

				// obtain the Assumptions Worksheet in the Workbook.
				Sheet objAssumptionsWorksheet = objWorkbookPart.Workbook.Descendants<Sheet>().Where(sht => sht.Name == Properties.AppResources.Workbook_CRM_Worksheet_Name_Assumptions).FirstOrDefault();
				if(objAssumptionsWorksheet == null)
					{
					throw new ArgumentException("The " + Properties.AppResources.Workbook_CRM_Worksheet_Name_Assumptions +
					   " worksheet could not be loacated in the workbook.");
					}
				WorksheetPart objAssumptionsWorksheetPart = (WorksheetPart)(objWorkbookPart.GetPartById(objAssumptionsWorksheet.Id));

				// If Hyperlinks need to be inserted, add the 
				Hyperlinks objHyperlinks = new Hyperlinks();
				//-------------------------------------
				// Begin to process the selects Mapping
				if(this.CRM_Mapping == 0)
					{
					strErrorText = "A Client Requirements Mapping was not specified for the Document Collection.";
					Console.WriteLine("### {0} ###", strErrorText);
					// If an entry was not specified - write an error in the Worksheet and record an error in the error log.
					this.LogError(strErrorText);

					//intStringIndex = oxmlWorkbook.InsertSharedStringItem(parText2Insert: strErrorText, parShareStringPart: objSharedStringTablePart);

					objCell = oxmlWorkbook.InsertCellInWorksheet(
						parColumnName: strColumnTowerOfServices,
						parRowIndex: intMatrixSheet_RowIndex,
						parWorksheetPart: objMatrixWorksheetPart);
					objCell.DataType = new EnumValue<CellValues>(CellValues.String);
					objCell.CellValue = new CellValue(strErrorText);
					goto Save_and_Close_Document;
                         }

				//---------------------------------------
				// Begin to process the Mapping data 
				Mapping objMapping = new Mapping();
				objMapping.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: this.CRM_Mapping);
				Console.WriteLine(" + Mapping: {0} - {1}", objMapping.ID, objMapping.Title);

				// Declare the List containing the various types of objects to be processed
				List<MappingServiceTower> listMappingTowers = new List<MappingServiceTower>();
				List<MappingRequirement> listMappingRequirements = new List<MappingRequirement>();
				List<MappingDeliverable> listMappingDeliverables = new List<MappingDeliverable>();
				List<MappingRisk> listMappingRisks = new List<MappingRisk>();
				List<MappingAssumption> listMappingAssumptions = new List<MappingAssumption>();
				List<MappingServiceLevel> listMappingServiceLevels = new List<MappingServiceLevel>();
				// Obtain all Mapping Service Towers for the specified Mapping
				try
					{
					listMappingTowers.Clear();
					listMappingTowers = MappingServiceTower.ObtainListOfObjects(parDatacontextSDDP: datacontexSDDP, parMappingID: objMapping.ID);
					}
				catch(DataEntryNotFoundException exc)
					{
					strErrorText = exc.Message;
					Console.WriteLine("### {0} ###", strErrorText);
					// If the no Service Tower (s) was not found - record an error in the error log.
					this.LogError(strErrorText);
					goto Save_and_Close_Document;
					}

				// Check if any entries were retrieved
				if(listMappingTowers.Count == 0)
					goto Save_and_Close_Document;

				//--------------------------------------------------------
				// --- Loop through all Service Towers for the Mapping ---
				foreach(MappingServiceTower objTower in listMappingTowers)
					{
					// Write the Mapping Service Tower to the Workbook as a String
					Console.WriteLine("\t + Tower: {0} - {1}", objTower.ID, objTower.Title);
					intMatrixSheet_RowIndex += 1;
					objCell = oxmlWorkbook.InsertCellInWorksheet(
						parColumnName: strColumnTowerOfServices,
						parRowIndex: intMatrixSheet_RowIndex,
						parWorksheetPart: objMatrixWorksheetPart);
					objCell.DataType = new EnumValue<CellValues>(CellValues.String);
					objCell.CellValue = new CellValue(objTower.Title);

					// Write the ROW TYPE as a Shared String value
					intStringIndex = oxmlWorkbook.InsertSharedStringItem(
						parText2Insert: Properties.AppResources.Workbook_CRM_Matrix_RowType_TowerOfService, parShareStringPart: objSharedStringTablePart);
					// now write the text to the Worksheet.
					objCell = oxmlWorkbook.InsertCellInWorksheet(
						parColumnName: strColumnRowType,
						parRowIndex: intMatrixSheet_RowIndex,
						parWorksheetPart: objMatrixWorksheetPart);
					objCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
					objCell.CellValue = new CellValue(intStringIndex.ToString());

					// Write the MAPPING Reference as a numeric value
					objCell = oxmlWorkbook.InsertCellInWorksheet(
						parWorksheetPart: objMatrixWorksheetPart,
                              parColumnName: strColumnMappingReference,
						parRowIndex: intMatrixSheet_RowIndex);
					objCell.DataType = new EnumValue<CellValues>(CellValues.Number);
					objCell.CellValue = new CellValue(objTower.ID.ToString());

					// check if Hyperlinks must be inserted
					if(documentCollection_HyperlinkURL != "")
						{
						intHyperlinkCounter += 1;
						oxmlWorkbook.InsertHyperlink(
							parWorksheetPart: objMatrixWorksheetPart,
							parCellReference: strColumnMappingReference + intMatrixSheet_RowIndex,
							parHyperLinkID: "Hyp" + intHyperlinkCounter,
							parHyperlinkURL: Properties.AppResources.SharePointURL +
								Properties.AppResources.List_MappingServiceTowers +
								currentHyperlinkViewEditURI + objTower.ID);
						}

					// Obtain all Mapping Requirements for the specified Mapping Service Tower
					try
						{
						listMappingRequirements.Clear();
						listMappingRequirements = MappingRequirement.ObtainListOfObjects(parDatacontextSDDP: datacontexSDDP, parMappingTowerID: objTower.ID);
						}
					catch(DataEntryNotFoundException)
						{
						continue; // No entries were found process the next Mapping Service Tower
						}

					// Process all the Mapping requirements for the specific Service Tower
					foreach(MappingRequirement objRequirement in listMappingRequirements)
						{
						Console.WriteLine("\t\t + Requirement: {0} - {1}", objRequirement.ID, objRequirement.Title);
						// Insert the Requirement Title cell
						intMatrixSheet_RowIndex += 1;
						objCell = oxmlWorkbook.InsertCellInWorksheet(
						parColumnName: strColumnRequirement,
						parRowIndex: intMatrixSheet_RowIndex,
						parWorksheetPart: objMatrixWorksheetPart);
						objCell.DataType = new EnumValue<CellValues>(CellValues.String);
						objCell.CellValue = new CellValue(objRequirement.Title);

						// Insert the Requirement Description as a Comment in the New Column
						if(objRequirement.RequirementText != null)
							{
							Comment objComment = oxmlWorkbook.InsertComment(
								parCellReference: strColumnNew + intMatrixSheet_RowIndex,
								parText2Add: objRequirement.RequirementText);
							objMatrixCommentList.Append(objComment);
							}

						// Write the ROW TYPE as a Shared String value
						intStringIndex = oxmlWorkbook.InsertSharedStringItem(
							parText2Insert: Properties.AppResources.Workbook_CRM_Matrix_RowType_Requirement, 
							parShareStringPart: objSharedStringTablePart);
						// now write the text to the Worksheet.
						objCell = oxmlWorkbook.InsertCellInWorksheet(
							parColumnName: strColumnRowType,
							parRowIndex: intMatrixSheet_RowIndex,
							parWorksheetPart: objMatrixWorksheetPart);
						objCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
						objCell.CellValue = new CellValue(intStringIndex.ToString());

						// Write the COMPLIANCE as a Shared String value
						intStringIndex = oxmlWorkbook.InsertSharedStringItem(
							parText2Insert: objRequirement.ComplianceStatus,
							parShareStringPart: objSharedStringTablePart);
						// now write the text to the Worksheet.
						objCell = oxmlWorkbook.InsertCellInWorksheet(
							parColumnName: strColumnComliance,
							parRowIndex: intMatrixSheet_RowIndex,
							parWorksheetPart: objMatrixWorksheetPart);
						objCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
						objCell.CellValue = new CellValue(intStringIndex.ToString());

						// Insert the SOURCE REFERENCE cell
						objCell = oxmlWorkbook.InsertCellInWorksheet(
						parColumnName: strColumnSourceReference,
						parRowIndex: intMatrixSheet_RowIndex,
						parWorksheetPart: objMatrixWorksheetPart);
						objCell.DataType = new EnumValue<CellValues>(CellValues.String);
						objCell.CellValue = new CellValue(objRequirement.SourceReference);

						// Write the MAPPING Reference as a numeric value
						objCell = oxmlWorkbook.InsertCellInWorksheet(
							parWorksheetPart: objMatrixWorksheetPart,
							parColumnName: strColumnMappingReference,
							parRowIndex: intMatrixSheet_RowIndex);
						objCell.DataType = new EnumValue<CellValues>(CellValues.Number);
						objCell.CellValue = new CellValue(objRequirement.ID.ToString());
						// Check if a hyperlink must be inserted
						if(documentCollection_HyperlinkURL != "")
							{
							intHyperlinkCounter += 1;
							oxmlWorkbook.InsertHyperlink(
								parWorksheetPart: objMatrixWorksheetPart,
								parCellReference: strColumnMappingReference + intMatrixSheet_RowIndex,
								parHyperLinkID: "Hyp" + intHyperlinkCounter,
								parHyperlinkURL: Properties.AppResources.SharePointURL +
									Properties.AppResources.List_MappingRequirements +
									currentHyperlinkViewEditURI + objRequirement.ID);
							}

						//--------------------------------------------------------------
						// Obtain all Mapping Risk for the specified Mapping Requirement
						try
							{
							listMappingRisks.Clear();
							listMappingRisks = MappingRisk.ObtainListOfObjects(
								parDatacontextSDDP: datacontexSDDP,
								parMappingRequirementID: objRequirement.ID);
							}
						catch(DataEntryNotFoundException)
							{
							// Ignore if there is no Risks recorded
							}

						// Check if any Mapping Risks were found
						if(listMappingRisks.Count != 0)
							{
							// Process all the Mapping Risks for the specific Service Requirement
							foreach(MappingRisk objRisk in listMappingRisks)
								{
								Console.WriteLine("\t\t\t + Risk: {0} - {1}", objRisk.ID, objRisk.Title);
								// Write the Mapping Risk to the Workbook as a String
								intMatrixSheet_RowIndex += 1;
								objCell = oxmlWorkbook.InsertCellInWorksheet(
									parColumnName: strColumnDelRiskAss,
									parRowIndex: intMatrixSheet_RowIndex,
									parWorksheetPart: objMatrixWorksheetPart);
								objCell.DataType = new EnumValue<CellValues>(CellValues.String);
								objCell.CellValue = new CellValue(objRisk.Title);

								// Insert the Risk Description as a Comment in the New Column
								if(objRisk.Statement != null)
									{
									Comment objComment = oxmlWorkbook.InsertComment(
										parCellReference: strColumnNew + intMatrixSheet_RowIndex,
										parText2Add: objRisk.Statement);
									objMatrixCommentList.Append(objComment);
									}

								// Write the ROW TYPE as a Shared String value
								intStringIndex = oxmlWorkbook.InsertSharedStringItem(
									parText2Insert: Properties.AppResources.Workbook_CRM_Matrix_RowType_Risk, parShareStringPart: objSharedStringTablePart);
								// now write the text to the Worksheet.
								objCell = oxmlWorkbook.InsertCellInWorksheet(
									parColumnName: strColumnRowType,
									parRowIndex: intMatrixSheet_RowIndex,
									parWorksheetPart: objMatrixWorksheetPart);
								objCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
								objCell.CellValue = new CellValue(intStringIndex.ToString());

								// Write the MAPPING Reference as a numeric value
								objCell = oxmlWorkbook.InsertCellInWorksheet(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnName: strColumnMappingReference,
									parRowIndex: intMatrixSheet_RowIndex);
								objCell.DataType = new EnumValue<CellValues>(CellValues.Number);
								objCell.CellValue = new CellValue(objRisk.ID.ToString());
								// Check if a hyperlink must be inserted
								if(documentCollection_HyperlinkURL != "")
									{
									intHyperlinkCounter += 1;
									oxmlWorkbook.InsertHyperlink(
										parWorksheetPart: objMatrixWorksheetPart,
										parCellReference: strColumnMappingReference + intMatrixSheet_RowIndex,
										parHyperLinkID: "Hyp" + intHyperlinkCounter,
										parHyperlinkURL: Properties.AppResources.SharePointURL +
											Properties.AppResources.List_MappingRisks +
											currentHyperlinkViewEditURI + objRequirement.ID);
									}
								} //foreach(Mappingrisk objMappingRisk in listMappingRisks)
							} // if(listMappingRisks.Count != 0)

						//----------------------------------------------------------------------
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
								objCell = oxmlWorkbook.InsertCellInWorksheet(
									parColumnName: strColumnDelRiskAss,
									parRowIndex: intMatrixSheet_RowIndex,
									parWorksheetPart: objMatrixWorksheetPart);
								objCell.DataType = new EnumValue<CellValues>(CellValues.String);
								objCell.CellValue = new CellValue(objAssumption.Title);

								// Insert the Assumption Description as a Comment in the New Column
								if(objAssumption.Description != null)
									{
									Comment objComment = oxmlWorkbook.InsertComment(
										parCellReference: strColumnNew + intMatrixSheet_RowIndex,
										parText2Add: objAssumption.Description);
									objMatrixCommentList.Append(objComment);
									}

								// Write the ROW TYPE as a Shared String value
								intStringIndex = oxmlWorkbook.InsertSharedStringItem(
									parText2Insert: Properties.AppResources.Workbook_CRM_Matrix_RowType_Assumption, 
									parShareStringPart: objSharedStringTablePart);
								// now write the text to the Worksheet.
								objCell = oxmlWorkbook.InsertCellInWorksheet(
									parColumnName: strColumnRowType,
									parRowIndex: intMatrixSheet_RowIndex,
									parWorksheetPart: objMatrixWorksheetPart);
								objCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
								objCell.CellValue = new CellValue(intStringIndex.ToString());

								// Write the MAPPING Reference as a numeric value
								objCell = oxmlWorkbook.InsertCellInWorksheet(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnName: strColumnMappingReference,
									parRowIndex: intMatrixSheet_RowIndex);
								objCell.DataType = new EnumValue<CellValues>(CellValues.Number);
								objCell.CellValue = new CellValue(objAssumption.ID.ToString());
								// Check if a hyperlink must be inserted
								if(documentCollection_HyperlinkURL != "")
									{
									intHyperlinkCounter += 1;
									oxmlWorkbook.InsertHyperlink(
										parWorksheetPart: objMatrixWorksheetPart,
										parCellReference: strColumnMappingReference + intMatrixSheet_RowIndex,
										parHyperLinkID: "Hyp" + intHyperlinkCounter,
										parHyperlinkURL: Properties.AppResources.SharePointURL +
											Properties.AppResources.List_MappingAssumptions +
											currentHyperlinkViewEditURI + objAssumption.ID);
									}
								} //foreach(MappingAssumption objMappingAssumption in listMappingAssumptions)
							} // if(listMappingAssumptions.Count != 0)

						//------------------------------------------
						// Obtain all Mapping Deliverables for the specified Mapping Requirement
						try
							{
							listMappingDeliverables.Clear();
							listMappingDeliverables = MappingDeliverable.ObtainListOfObjects(
								parDatacontextSDDP: datacontexSDDP,
								parMappingRequirementID: objRequirement.ID);
							}
						catch(DataEntryNotFoundException)
							{
							// ignore if there are no Mapping Deliverables
							}

						// Check if any Mapping Deliverables were found
						if(listMappingDeliverables.Count != 0)
							{
							// Process all the Mapping Deliverables for the specific Service Requirement
							foreach(MappingDeliverable objMappingDeliverable in listMappingDeliverables)
								{
								Console.WriteLine("\t\t\t + DRM: {0} - {1}", objMappingDeliverable.ID, objMappingDeliverable.Title);
								intMatrixSheet_RowIndex += 1;

								// Insert the Deliverable Description as a Comment in the New Column
								// This one is a bit different because it depends whether it is a new or existing deliverable.
								if(objMappingDeliverable.NewDeliverable)
									{
									objCell = oxmlWorkbook.InsertCellInWorksheet(
										parColumnName: strColumnDelRiskAss,
										parRowIndex: intMatrixSheet_RowIndex,
										parWorksheetPart: objMatrixWorksheetPart);
									objCell.DataType = new EnumValue<CellValues>(CellValues.String);
									objCell.CellValue = new CellValue(objMappingDeliverable.Title);

									// Insert the "New" value in the cell under the New column as a SharedString
									intStringIndex = oxmlWorkbook.InsertSharedStringItem(
										parText2Insert: Properties.AppResources.Workbook_CRM_Matrix_NewColumn_Text,
										parShareStringPart: objSharedStringTablePart);
									// now write the Shared Text Reference to the cell.
									objCell = oxmlWorkbook.InsertCellInWorksheet(
										parColumnName: strColumnNew,
										parRowIndex: intMatrixSheet_RowIndex,
										parWorksheetPart: objMatrixWorksheetPart);
									objCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
									objCell.CellValue = new CellValue(intStringIndex.ToString());

									if(objMappingDeliverable.NewRequirement != null)
										{
										Comment objComment = oxmlWorkbook.InsertComment(
											parCellReference: strColumnNew + intMatrixSheet_RowIndex,
											parText2Add: objMappingDeliverable.NewRequirement);
										objMatrixCommentList.Append(objComment);
										}
									}
								else // if it is an EXISTING deliverable
									{
									objCell = oxmlWorkbook.InsertCellInWorksheet(
										parColumnName: strColumnDelRiskAss,
										parRowIndex: intMatrixSheet_RowIndex,
										parWorksheetPart: objMatrixWorksheetPart);
									objCell.DataType = new EnumValue<CellValues>(CellValues.String);
									objCell.CellValue = new CellValue(objMappingDeliverable.MappedDeliverable.CSDheading);

									strTextDescription = "";
									layer0upDeliverableID = objMappingDeliverable.MappedDeliverable.ID;
									if(objMappingDeliverable.MappedDeliverable.ContentPredecessorDeliverableID == null)
										{
										layer1upDeliverableID = null;
										layer2upDeliverableID = null;
										}
									else
										{
										layer1upDeliverableID = objMappingDeliverable.MappedDeliverable.ContentPredecessorDeliverableID;
										if(objMappingDeliverable.MappedDeliverable.Layer1up.ContentPredecessorDeliverableID == null)
											{
											layer2upDeliverableID = null;
											}
										else
											{
											layer2upDeliverableID =
												objMappingDeliverable.MappedDeliverable.Layer1up.ContentPredecessorDeliverableID;
                                                       }
										}
									if(layer2upDeliverableID != null)
										{
										if(objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.CSDdescription != null)
											{
											strTextDescription = HTMLdecoder.CleanHTMLstring
												(objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.CSDdescription);
											}
										}
									if(layer1upDeliverableID != null)
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
										Comment objComment = oxmlWorkbook.InsertComment(
											parCellReference: strColumnNew + intMatrixSheet_RowIndex,
											parText2Add: objMappingDeliverable.NewRequirement);
										objMatrixCommentList.Append(objComment);
										}
									} // end if EXISITING Deliverable

								// Insert the Mapping Reference as a numeric value
								objCell = oxmlWorkbook.InsertCellInWorksheet(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnName: strColumnMappingReference,
									parRowIndex: intMatrixSheet_RowIndex);
								objCell.DataType = new EnumValue<CellValues>(CellValues.Number);
								objCell.CellValue = new CellValue(objMappingDeliverable.ID.ToString());
								// Check if a hyperlink must be inserted
								if(documentCollection_HyperlinkURL != "")
									{
									intHyperlinkCounter += 1;
									oxmlWorkbook.InsertHyperlink(
										parWorksheetPart: objMatrixWorksheetPart,
										parCellReference: strColumnMappingReference + intMatrixSheet_RowIndex,
										parHyperLinkID: "Hyp" + intHyperlinkCounter,
										parHyperlinkURL: Properties.AppResources.SharePointURL +
											Properties.AppResources.List_MappingDeliverables +
											currentHyperlinkViewEditURI + objMappingDeliverable.ID);
									}

								if(!objMappingDeliverable.NewDeliverable) // if it is an EXISTING deliverable
									{
									// Insert the Deliverable Reference as a numeric value under the Deliverable Reference column
									objCell = oxmlWorkbook.InsertCellInWorksheet(
										parWorksheetPart: objMatrixWorksheetPart,
										parColumnName: strColumnDeliverableReference,
										parRowIndex: intMatrixSheet_RowIndex);
									objCell.DataType = new EnumValue<CellValues>(CellValues.Number);
									objCell.CellValue = new CellValue(objMappingDeliverable.MappedDeliverable.ID.ToString());
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										intHyperlinkCounter += 1;
										oxmlWorkbook.InsertHyperlink(
											parWorksheetPart: objMatrixWorksheetPart,
											parCellReference: strColumnDeliverableReference + intMatrixSheet_RowIndex,
											parHyperLinkID: "Hyp" + intHyperlinkCounter,
											parHyperlinkURL: Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI + objMappingDeliverable.MappedDeliverable.ID);
										}
									}
								//--------------------------------------------------------------------
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
										Console.WriteLine("\t\t\t\t + DRM: {0} - {1}", objMappingServiceLevel.ID, objMappingServiceLevel.Title);
										// Write the Mapping Service Level to the Workbook as a String
										intMatrixSheet_RowIndex += 1;
										// Insert the Service Level Title as string
										// This one is a bit different because it depends whether it is a new or existing service level.
										if(objMappingServiceLevel.NewServiceLevel)
											{
											// Insert the Service Level Title in a cell under the Service Level Column
											objCell = oxmlWorkbook.InsertCellInWorksheet(
												parColumnName: strColumnServiceLevel,
												parRowIndex: intMatrixSheet_RowIndex,
												parWorksheetPart: objMatrixWorksheetPart);
											objCell.DataType = new EnumValue<CellValues>(CellValues.String);
											objCell.CellValue = new CellValue(objMappingServiceLevel.Title);
											// Insert the "New" value in the cell under the New column as 
											intStringIndex = oxmlWorkbook.InsertSharedStringItem(
												parText2Insert: Properties.AppResources.Workbook_CRM_Matrix_NewColumn_Text,
												parShareStringPart: objSharedStringTablePart);
											// now write the Shared Text Reference to the cell.
											objCell = oxmlWorkbook.InsertCellInWorksheet(
												parColumnName: strColumnNew,
												parRowIndex: intMatrixSheet_RowIndex,
												parWorksheetPart: objMatrixWorksheetPart);
											objCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
											objCell.CellValue = new CellValue(intStringIndex.ToString());

											// Insert the Service Level Text as a comment
											if(objMappingServiceLevel.RequirementText != null)
												{
												Comment objComment = oxmlWorkbook.InsertComment(
													parCellReference: strColumnNew + intMatrixSheet_RowIndex,
													parText2Add: objMappingServiceLevel.RequirementText);
												objMatrixCommentList.Append(objComment);
												}
											}
										else // if it is an EXISTING Service Level
											{
											// Insert the CSD Heading on cell under the Service Level column
											objCell = oxmlWorkbook.InsertCellInWorksheet(
												parColumnName: strColumnServiceLevel,
												parRowIndex: intMatrixSheet_RowIndex,
												parWorksheetPart: objMatrixWorksheetPart);
											objCell.DataType = new EnumValue<CellValues>(CellValues.String);
											objCell.CellValue = new CellValue(objMappingServiceLevel.MappedServiceLevel.CSDheading);
											// Insert the CSD Description if not null
											if(objMappingServiceLevel.MappedServiceLevel.CSDdescription != "")
												{
												strTextDescription = HTMLdecoder.CleanHTMLstring(objMappingServiceLevel.MappedServiceLevel.CSDdescription);
												Comment objComment = oxmlWorkbook.InsertComment(
													parCellReference: strColumnNew + intMatrixSheet_RowIndex,
													parText2Add: strTextDescription);
												objMatrixCommentList.Append(objComment);
												}
											} // end if EXISTING Service Level

										// Write the ROW TYPE as a Shared String value
										intStringIndex = oxmlWorkbook.InsertSharedStringItem(
											parText2Insert: Properties.AppResources.Workbook_CRM_Matrix_RowType_ServiceLevel,
											parShareStringPart: objSharedStringTablePart);
										// now write the text to the Worksheet.
										objCell = oxmlWorkbook.InsertCellInWorksheet(
											parColumnName: strColumnRowType,
											parRowIndex: intMatrixSheet_RowIndex,
											parWorksheetPart: objMatrixWorksheetPart);
										objCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
										objCell.CellValue = new CellValue(intStringIndex.ToString());

										// Insert the Mapping Reference as a numeric value
										objCell = oxmlWorkbook.InsertCellInWorksheet(
											parWorksheetPart: objMatrixWorksheetPart,
											parColumnName: strColumnMappingReference,
											parRowIndex: intMatrixSheet_RowIndex);
										objCell.DataType = new EnumValue<CellValues>(CellValues.Number);
										objCell.CellValue = new CellValue(objMappingServiceLevel.ID.ToString());
										// Check if a hyperlink must be inserted
										if(documentCollection_HyperlinkURL != "")
											{
											intHyperlinkCounter += 1;
											oxmlWorkbook.InsertHyperlink(
												parWorksheetPart: objMatrixWorksheetPart,
												parCellReference: strColumnMappingReference + intMatrixSheet_RowIndex,
												parHyperLinkID: "Hyp" + intHyperlinkCounter,
												parHyperlinkURL: Properties.AppResources.SharePointURL +
													Properties.AppResources.List_MappingServiceLevels +
													currentHyperlinkViewEditURI + objMappingServiceLevel.ID);
											}

										if(!objMappingServiceLevel.NewServiceLevel) // if it is an EXISTING Service Level
											{
											// Insert the Service Level Reference as a numeric value under the Service Level ID column
											objCell = oxmlWorkbook.InsertCellInWorksheet(
												parWorksheetPart: objMatrixWorksheetPart,
												parColumnName: strColumnServiceLevelReference,
												parRowIndex: intMatrixSheet_RowIndex);
											objCell.DataType = new EnumValue<CellValues>(CellValues.Number);
											objCell.CellValue = new CellValue(objMappingServiceLevel.MappedServiceLevel.ID.ToString());
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												intHyperlinkCounter += 1;
												oxmlWorkbook.InsertHyperlink(
													parWorksheetPart: objMatrixWorksheetPart,
													parCellReference: strColumnServiceLevelReference + intMatrixSheet_RowIndex,
													parHyperLinkID: "Hyp" + intHyperlinkCounter,
													parHyperlinkURL: Properties.AppResources.SharePointURL +
														Properties.AppResources.List_ServiceLevelsURI +
														currentHyperlinkViewEditURI + objMappingServiceLevel.MappedServiceLevel.ID);
												}
											}
										} // foreach(MappingServiceLevel objMappingServiceLevel in listMappingServiceLevels)
									} // if(listMappingServiceLevels.Count != 0)
								} // foreach(MappingDeliverable objMappingDeliverable in listMappingDeliverables)
							} // if(listMappingDeliverables.Count != 0)
						} // foreach(MappingRequirement objRequirement in listMappingRequirements)
					} //foreach(MappingServiceTower objTower in listMappingTowers)


Save_and_Close_Document:
				
				//----------------------------------------------
				//Append all the Comments to the Matrix Sheet
				// obtain the WorksheetCommentsPart - required to insert Comments in cells
				WorksheetCommentsPart objMatrixWorksheetCommentsPart;
				Comments objMatrixComments;
				Console.WriteLine("Comments recorded: {0}", objMatrixCommentList.Count());
				if(objMatrixCommentList.Count() != 0)
					{
					
					if(objMatrixWorksheetPart.WorksheetCommentsPart == null)
						{
						objMatrixWorksheetCommentsPart = objMatrixWorksheetPart.AddNewPart<WorksheetCommentsPart>(id: "mtxComments");
						}
					else
						{
						objMatrixWorksheetCommentsPart = objMatrixWorksheetPart.WorksheetCommentsPart;
						}


					if(objMatrixWorksheetCommentsPart.Comments == null)
						{
						objMatrixComments = new Comments();
						}
					else
						{
						objMatrixComments = objMatrixWorksheetCommentsPart.Comments;
	                         }
					CommentList objCommentsList = objMatrixComments.CommentList;
					// Add all the comments to the CommentsList.
					foreach(Comment itemComment in objMatrixCommentList)
						{
						objCommentsList.Append(itemComment);
						}
					objMatrixWorksheetCommentsPart.Comments.Save();

					} //if(objMatrixCommentList.Count() != 0)

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
