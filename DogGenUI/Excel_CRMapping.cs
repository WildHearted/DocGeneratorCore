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
				Console.WriteLine("\t\t objOXMLdocument:\n" +
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
				WorksheetPart objMatrixWorksheetPart = (WorksheetPart)(objWorkbookPart.GetPartById(objMatrixWorksheet.Id));
				WorksheetCommentsPart objMatrixCommentsPart = (WorksheetCommentsPart)(objMatrixWorksheetPart.GetPartsOfType<WorksheetCommentsPart>().First());

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
				
				if(this.HyperlinkEdit || this.HyperlinkEdit)
					{
					// Check if the Worksheet contains Hyperlinks 
					Hyperlinks objHyperlinks = new Hyperlinks();
					Hyperlinks objExistingHyperlinks = objMatrixWorksheetPart.Worksheet.Descendants<Hyperlinks>().First();
					if(objExistingHyperlinks == null) // Hyperlinks Doesn't exisit yet
						{
						// Get the PageMargins, inorder to insert the Hyperlink BEFORE the PageMargins
						PageMargins objPageMargins = objMatrixWorksheetPart.Worksheet.Descendants<PageMargins>().First();
						objMatrixWorksheetPart.Worksheet.InsertBefore<Hyperlinks>(newChild: objHyperlinks, refChild: objPageMargins);
						objMatrixWorksheetPart.Worksheet.Save();
						}
					}

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
					// First check if Hyperlinks must be inserted
					if(documentCollection_HyperlinkURL != "")
						{
						intHyperlinkCounter += 1;
						oxmlWorkbook.InsertHyperlink(
							parWorksheetPart: objMatrixWorksheetPart,
							parCellReference: strColumnMappingReference + intMatrixSheet_RowIndex,
							parHyperLinkID: intHyperlinkCounter,
							parHyperlinkURL: Properties.AppResources.SharePointURL +
								Properties.AppResources.List_MappingServiceTowers +
								currentHyperlinkViewEditURI + objTower.ID);
						}

					objCell = oxmlWorkbook.InsertCellInWorksheet(
						parWorksheetPart: objMatrixWorksheetPart,
                              parColumnName: strColumnMappingReference,
						parRowIndex: intMatrixSheet_RowIndex);
					objCell.DataType = new EnumValue<CellValues>(CellValues.Number);
					objCell.CellValue = new CellValue(objTower.ID.ToString());

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
							objMatrixCommentsPart.Comments.Append(objComment);
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

						// Check if a hyperlink must be inserted
						if(documentCollection_HyperlinkURL != "")
							{
							intHyperlinkCounter += 1;
							oxmlWorkbook.InsertHyperlink(
								parWorksheetPart: objMatrixWorksheetPart,
								parCellReference: strColumnMappingReference + intMatrixSheet_RowIndex,
								parHyperLinkID: intHyperlinkCounter,
								parHyperlinkURL: Properties.AppResources.SharePointURL +
									Properties.AppResources.List_MappingRequirements +
									currentHyperlinkViewEditURI + objRequirement.ID);
							}

						--- gaan hier aan ---
						// Check if the user specified to include the Requirement Service Level
						if(this.Requirement_Service_Level)
							{
							if(objRequirement.RequirementServiceLevel != null)
								{
								objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 3);
								objRun1 = oxmlDocument.Construct_RunText(parText2Write: objRequirement.RequirementServiceLevel);
								objParagraph.Append(objRun1);
								objBody.Append(objParagraph);
								}
							}

						// Insert the Requirement Compliance:
						objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 3);
						objRun1 = oxmlDocument.Construct_RunText(
							parText2Write: Properties.AppResources.Document_RequirementsMapping_ComplianceStatusTitle,
							parBold: true);
						objParagraph.Append(objRun1);
						if(objRequirement.ComplianceStatus != null)
							{
							objRun2 = oxmlDocument.Construct_RunText(
							parText2Write: "No Reponse");
							}
						else
							{
							objRun2 = oxmlDocument.Construct_RunText(parText2Write: objRequirement.ComplianceStatus);
							objParagraph.Append(objRun2);
							objBody.Append(objParagraph);
							}

						//------------------------------------
						// User selected to include the Risks
						if(this.Risks)
							{
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
								// Ignore if there are none
								}

							// Check if any Mapping Risks were found
							if(listMappingRisks.Count != 0)
								{
								// Insert the Risks Heading:
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
								objRun1 = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_RequirementsMapping_RisksHeading);
								objParagraph.Append(objRun1);
								objBody.Append(objParagraph);

								// Process all the Mapping Risks for the specific Service Requirement
								foreach(MappingRisk objRisk in listMappingRisks)
									{
									Console.WriteLine("\t\t\t + Risk: {0} - {1}", objRisk.ID, objRisk.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
									objRun1 = oxmlDocument.Construct_RunText(parText2Write: objRisk.Title);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objWorkbooktPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
											Properties.AppResources.List_MappingRisks +
											currentHyperlinkViewEditURI + objRisk.ID,
											parHyperlinkID: hyperlinkCounter);
										objRun1.Append(objDrawing);
										}
									objParagraph.Append(objRun1);
									objBody.Append(objParagraph);

									// Check if the Requirement Description Table
									if(this.Risk_Description)
										{
										Table tableMappingRisk = new Table();
										tableMappingRisk = CommonProcedures.BuildRiskTable(
											parMappingRisk: objRisk,
											parWidthColumn1: Convert.ToUInt32(this.PageWith * 0.3),
											parWidthColumn2: Convert.ToUInt32(this.PageWith * 0.7));
										objBody.Append(tableMappingRisk);
										}
									} //foreach(Mappingrisk objMappingRisk in listMappingRisks)
								} // if(listMappingRisks.Count != 0)
							} // if(this.Risks)

						//----------------------------------------------
						// The user selected to include the Assumptions
						if(this.Assumptions)
							{
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
								// Insert the Risks Heading:
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
								objRun1 = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_RequirementMapping_AssumptionsHeading);
								objParagraph.Append(objRun1);
								objBody.Append(objParagraph);

								// Process all the Mapping Assumptions for the specific Service Requirement
								foreach(MappingAssumption objAssumption in listMappingAssumptions)
									{
									Console.WriteLine("\t\t\t + Assumption: {0} - {1}", objAssumption.ID, objAssumption.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
									objRun1 = oxmlDocument.Construct_RunText(
										parText2Write: objAssumption.Title);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objWorkbooktPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
											Properties.AppResources.List_MappingAssumptions +
											currentHyperlinkViewEditURI + objAssumption.ID,
											parHyperlinkID: hyperlinkCounter);
										objRun1.Append(objDrawing);
										}
									objParagraph.Append(objRun1);
									objBody.Append(objParagraph);

									// Check if the Requirement Description Table
									if(this.Risk_Description)
										{
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
										objRun1 = oxmlDocument.Construct_RunText(parText2Write: objAssumption.Description);
										objParagraph.Append(objRun1);
										objBody.Append(objParagraph);
										}
									} //foreach(MappingAssumption objMappingAssumption in listMappingAssumptions)
								} // if(listMappingAssumptions.Count != 0)
							} //if(this.Assumptions)

						//------------------------------------------
						// The user selected to include the DRMs
						if(this.Deliverable_Reports_and_Meetings)
							{
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
								// Insert the Deliverable Heading:
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
								objRun1 = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_RequirementsMapping_DeliverableReportMeetingsHeading);
								objParagraph.Append(objRun1);
								objBody.Append(objParagraph);

								// Process all the Mapping Deliverables for the specific Service Requirement
								foreach(MappingDeliverable objMappingDeliverable in listMappingDeliverables)
									{
									Console.WriteLine("\t\t\t + DRM: {0} - {1}", objMappingDeliverable.ID, objMappingDeliverable.Title);
									// Insert the MappingDeliverable Title
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
									// If it is a new deliverable, use the MappingDeliverable's Title else use the actual
									// Mapped_Deliverable's CSD Description
									if(objMappingDeliverable.NewDeliverable)
										{
										objRun1 = oxmlDocument.Construct_RunText(parText2Write: objMappingDeliverable.Title);
										}
									else
										{
										objRun1 = oxmlDocument.Construct_RunText(parText2Write: objMappingDeliverable.MappedDeliverable.CSDheading);
										}
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objWorkbooktPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
											Properties.AppResources.List_MappingDeliverables +
											currentHyperlinkViewEditURI + objMappingDeliverable.ID,
											parHyperlinkID: hyperlinkCounter);
										objRun1.Append(objDrawing);
										}
									objParagraph.Append(objRun1);
									objBody.Append(objParagraph);

									// Insert the Description
									// If it a New deliverable, use the NewRequirement, ELSE process the Mapped_Deliverable's content
									if(objMappingDeliverable.NewDeliverable)
										{
										// Check if the Mapping Deliverable,Report,Meeting Description was selected
										if(this.DRM_Description)
											{

											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
											objRun1 = oxmlDocument.Construct_RunText(parText2Write: objMappingDeliverable.NewRequirement);
											objParagraph.Append(objRun1);
											objBody.Append(objParagraph);
											}
										}
									else // if(objMappingDeliverable.NewDeliverable == false)
										{
										// Check if the Mapping Deliverable,Report,Meeting Description was selected
										if(this.DRM_Description)
											{
											//Check if the Mapped_Deliverable Layer0up has Content Layers and Content Predecessors
											Console.WriteLine("\t\t\t\t + Deliverable Layer 0..: {0} - {1}",
												objMappingDeliverable.MappedDeliverable.ID, objMappingDeliverable.MappedDeliverable.Title);
											if(objMappingDeliverable.MappedDeliverable.ContentPredecessorDeliverableID == null)
												{
												layer1upDeliverableID = null;
												layer2upDeliverableID = null;
												}
											else
												{
												Console.WriteLine("\t\t\t\t + Deliverable Layer 1up: {0} - {1}",
														objMappingDeliverable.MappedDeliverable.Layer1up.ID,
														objMappingDeliverable.MappedDeliverable.Layer1up.Title);
												layer1upDeliverableID = objMappingDeliverable.MappedDeliverable.ContentPredecessorDeliverableID;
												if(objMappingDeliverable.MappedDeliverable.Layer1up.ContentPredecessorDeliverableID == null)
													{
													layer2upDeliverableID = null;
													}
												else
													{
													Console.WriteLine("\t\t\t\t + Deliverable Layer 2up: {0} - {1}",
														objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.ID,
														objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.Title);
													layer2upDeliverableID =
														objMappingDeliverable.MappedDeliverable.Layer1up.ContentPredecessorDeliverableID;
													}
												}
											// Insert Layer 2up if present and not null
											if(layer2upDeliverableID != null)
												{
												if(objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.CSDdescription != null)
													{
													// Check for Colour coding Layers and add if necessary
													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer1";
													else
														currentContentLayer = "None";

													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.ID;
														}
													else
														currentListURI = "";

													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objWorkbooktPart,
														parDocumentLevel: 5,
														parHTML2Decode: objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.CSDdescription,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref tableCaptionCounter,
														parImageCaptionCounter: ref imageCaptionCounter,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													} // if(objDeliverable.Layer1up.Layer1up.CSDdescription != null)
												} // if(layer2upDeliverableID != null)

											// Insert Layer 1up if present and not null
											if(layer1upDeliverableID != null)
												{
												if(objMappingDeliverable.MappedDeliverable.Layer1up.CSDdescription != null)
													{
													// Check for Colour coding Layers and add if necessary
													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer2";
													else
														currentContentLayer = "None";

													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objMappingDeliverable.MappedDeliverable.Layer1up.ID;
														}
													else
														currentListURI = "";

													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objWorkbooktPart,
														parDocumentLevel: 5,
														parHTML2Decode: objMappingDeliverable.MappedDeliverable.Layer1up.CSDdescription,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref tableCaptionCounter,
														parImageCaptionCounter: ref imageCaptionCounter,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													}// if(objDeliverable.Layer1up.Layer1up.CSDdescription != null)
												} // if(layer2upDeliverableID != null)

											// Insert Layer 0up if present and not null
											if(objMappingDeliverable.MappedDeliverable.CSDdescription != null)
												{
												// Check for Colour coding Layers and add if necessary
												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer3";
												else
													currentContentLayer = "None";

												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objMappingDeliverable.MappedDeliverable.ID;
													}
												else
													currentListURI = "";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objWorkbooktPart,
													parDocumentLevel: 5,
													parHTML2Decode: objMappingDeliverable.MappedDeliverable.CSDdescription,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												} // if(objDeliverable.CSDdescription != null)
											} // if (this.DRM_Description)

										//-----------------------------------------------------------------------
										// Check if the user specified to include the Deliverable DD's Obligations
										if(this.DDs_DRM_Obligations)
											{
											if(objMappingDeliverable.MappedDeliverable.DDobligations != null
											|| (layer1upDeliverableID != null && objMappingDeliverable.MappedDeliverable.Layer1up.DDobligations != null)
											|| (layer2upDeliverableID != null && objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.DDobligations != null))
												{
												// Insert the Heading
												objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
												objRun1 = oxmlDocument.Construct_RunText(
													parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
												objParagraph.Append(objRun1);
												objBody.Append(objParagraph);

												// Insert Layer 2up if present and not null
												if(layer2upDeliverableID != null)
													{
													if(objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.DDobligations != null)
														{
														if(documentCollection_HyperlinkURL != "")
															{
															hyperlinkCounter += 1;
															currentListURI = Properties.AppResources.SharePointURL +
																Properties.AppResources.List_DeliverablesURI +
																currentHyperlinkViewEditURI +
																objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.ID;
															}
														else
															currentListURI = "";

														if(this.ColorCodingLayer1)
															currentContentLayer = "Layer1";
														else
															currentContentLayer = "None";

														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objWorkbooktPart,
															parDocumentLevel: 6,
															parHTML2Decode: objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.DDobligations,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref tableCaptionCounter,
															parImageCaptionCounter: ref imageCaptionCounter,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														} //if(objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.DDobligations != null)
													} // if(layer2upDeliverableID != null)

												// Insert Layer 1up if present and not null
												if(layer1upDeliverableID != null)
													{
													if(objMappingDeliverable.MappedDeliverable.Layer1up.DDobligations != null)
														{
														// Check if a hyperlink must be inserted
														if(documentCollection_HyperlinkURL != "")
															{
															hyperlinkCounter += 1;
															currentListURI = Properties.AppResources.SharePointURL +
																Properties.AppResources.List_DeliverablesURI +
																currentHyperlinkViewEditURI +
																objMappingDeliverable.MappedDeliverable.Layer1up.ID;
															}
														else
															currentListURI = "";

														if(this.ColorCodingLayer1)
															currentContentLayer = "Layer2";
														else
															currentContentLayer = "None";

														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objWorkbooktPart,
															parDocumentLevel: 6,
															parHTML2Decode: objMappingDeliverable.MappedDeliverable.Layer1up.DDobligations,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref tableCaptionCounter,
															parImageCaptionCounter: ref imageCaptionCounter,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														} // if(objMappingDeliverable.MappedDeliverable.Layer1up.DDobligations != null)
													} // if(layer2upDeliverableID != null)

												// Insert Layer0up if not null
												if(objMappingDeliverable.MappedDeliverable.DDobligations != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objMappingDeliverable.MappedDeliverable.ID;
														}
													else
														currentListURI = "";

													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer3";
													else
														currentContentLayer = "None";

													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objWorkbooktPart,
														parDocumentLevel: 6,
														parHTML2Decode: objMappingDeliverable.MappedDeliverable.DDobligations,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref tableCaptionCounter,
														parImageCaptionCounter: ref imageCaptionCounter,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													} // if(objDeliverable.DDobligations != null)
												} //if(objDeliverable.DDoblidations != null &&)
											} // if(this.DDs_DRM_Objigations
											  //-------------------------------------------------------------------
											  // Check if the user specified to include the Client Responsibilities
										if(this.Clients_DRM_Responsibiities)
											{
											if(objMappingDeliverable.MappedDeliverable.ClientResponsibilities != null
											|| (layer1upDeliverableID != null && objMappingDeliverable.MappedDeliverable.Layer1up.ClientResponsibilities != null)
											|| (layer2upDeliverableID != null && objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.ClientResponsibilities != null))
												{
												// Insert the Heading
												objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
												objRun1 = oxmlDocument.Construct_RunText(
													parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
												objParagraph.Append(objRun1);
												objBody.Append(objParagraph);

												// Insert Layer 2up if present and not null
												if(layer2upDeliverableID != null)
													{
													if(objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.ClientResponsibilities != null)
														{
														// Check if a hyperlink must be inserted
														if(documentCollection_HyperlinkURL != "")
															{
															hyperlinkCounter += 1;
															currentListURI = Properties.AppResources.SharePointURL +
																Properties.AppResources.List_DeliverablesURI +
																currentHyperlinkViewEditURI +
																objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.ID;
															}
														else
															currentListURI = "";

														if(this.ColorCodingLayer1)
															currentContentLayer = "Layer1";
														else
															currentContentLayer = "None";

														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objWorkbooktPart,
															parDocumentLevel: 6,
															parHTML2Decode: objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.ClientResponsibilities,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref tableCaptionCounter,
															parImageCaptionCounter: ref imageCaptionCounter,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														} //if(objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.ClientResponsibilities != null)
													} // if(layer2upDeliverableID != null)

												// Insert Layer 1up if present and not null
												if(layer1upDeliverableID != null)
													{
													if(objMappingDeliverable.MappedDeliverable.Layer1up.ClientResponsibilities != null)
														{
														// Check if a hyperlink must be inserted
														if(documentCollection_HyperlinkURL != "")
															{
															hyperlinkCounter += 1;
															currentListURI = Properties.AppResources.SharePointURL +
																Properties.AppResources.List_DeliverablesURI +
																currentHyperlinkViewEditURI +
																objMappingDeliverable.MappedDeliverable.Layer1up.ID;
															}
														else
															currentListURI = "";

														if(this.ColorCodingLayer1)
															currentContentLayer = "Layer2";
														else
															currentContentLayer = "None";

														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objWorkbooktPart,
															parDocumentLevel: 6,
															parHTML2Decode: objMappingDeliverable.MappedDeliverable.Layer1up.ClientResponsibilities,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref tableCaptionCounter,
															parImageCaptionCounter: ref imageCaptionCounter,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														} // if(objMappingDeliverable.MappedDeliverable.Layer1up.ClientResponsibilities != null)
													} // if(layer2upDeliverableID != null)

												// Insert Layer0up if not null
												if(objMappingDeliverable.MappedDeliverable.ClientResponsibilities != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objMappingDeliverable.MappedDeliverable.ID;
														}
													else
														currentListURI = "";

													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer3";
													else
														currentContentLayer = "None";

													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objWorkbooktPart,
														parDocumentLevel: 6,
														parHTML2Decode: objMappingDeliverable.MappedDeliverable.ClientResponsibilities,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref tableCaptionCounter,
														parImageCaptionCounter: ref imageCaptionCounter,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													} // if(objMappingDeliverable.MappedDeliverable.ClientResponsibilities != null)
												} // if(objMappingDeliverable.MappedDeliverable.ClientResponsibilities != null &&)
											} //if(this.Clients_DRM_Responsibilities)

										//------------------------------------------------------------------
										// Check if the user specified to include the Deliverable Exclusions
										if(this.DRM_Exclusions)
											{
											if(objMappingDeliverable.MappedDeliverable.Exclusions != null
											|| (layer1upDeliverableID != null && objMappingDeliverable.MappedDeliverable.Layer1up.Exclusions != null)
											|| (layer2upDeliverableID != null && objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.Exclusions != null))
												{
												// Insert the Heading
												objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
												objRun1 = oxmlDocument.Construct_RunText(
													parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
												objParagraph.Append(objRun1);
												objBody.Append(objParagraph);

												// Insert Layer 2up if present and not null
												if(layer2upDeliverableID != null)
													{
													if(objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.Exclusions != null)
														{
														// Check if a hyperlink must be inserted
														if(documentCollection_HyperlinkURL != "")
															{
															hyperlinkCounter += 1;
															currentListURI = Properties.AppResources.SharePointURL +
																Properties.AppResources.List_DeliverablesURI +
																currentHyperlinkViewEditURI +
																objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.ID;
															}
														else
															currentListURI = "";

														if(this.ColorCodingLayer1)
															currentContentLayer = "Layer1";
														else
															currentContentLayer = "None";

														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objWorkbooktPart,
															parDocumentLevel: 6,
															parHTML2Decode: objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.Exclusions,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref tableCaptionCounter,
															parImageCaptionCounter: ref imageCaptionCounter,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														} //if(objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.Exclusions != null)
													} // if(layer2upDeliverableID != null)

												// Insert Layer 1up if present and not null
												if(layer1upDeliverableID != null)
													{
													if(objMappingDeliverable.MappedDeliverable.Layer1up.Exclusions != null)
														{
														// Check if a hyperlink must be inserted
														if(documentCollection_HyperlinkURL != "")
															{
															hyperlinkCounter += 1;
															currentListURI = Properties.AppResources.SharePointURL +
																Properties.AppResources.List_DeliverablesURI +
																currentHyperlinkViewEditURI +
																objMappingDeliverable.MappedDeliverable.Layer1up.ID;
															}
														else
															currentListURI = "";

														if(this.ColorCodingLayer1)
															currentContentLayer = "Layer2";
														else
															currentContentLayer = "None";

														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objWorkbooktPart,
															parDocumentLevel: 6,
															parHTML2Decode: objMappingDeliverable.MappedDeliverable.Layer1up.Exclusions,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref tableCaptionCounter,
															parImageCaptionCounter: ref imageCaptionCounter,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														} // if(objMappingDeliverable.MappedDeliverable.Layer1up.Exclusions != null)
													} // if(layer2upDeliverableID != null)

												// Insert Layer0up if not null
												if(objMappingDeliverable.MappedDeliverable.ClientResponsibilities != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objMappingDeliverable.MappedDeliverable.ID;
														}
													else
														currentListURI = "";

													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer3";
													else
														currentContentLayer = "None";

													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objWorkbooktPart,
														parDocumentLevel: 6,
														parHTML2Decode: objMappingDeliverable.MappedDeliverable.Exclusions,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref tableCaptionCounter,
														parImageCaptionCounter: ref imageCaptionCounter,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													} // if(objMappingDeliverable.MappedDeliverable.Exclusions != null)
												} // if(objMappingDeliverable.MappedDeliverable.Exclusions != null &&)	
											} //if(this.DRMe_Exclusions)

										//---------------------------------------------------------------
										// Check if the user specified to include the Governance Controls
										if(this.DRM_Governance_Controls)
											{
											if(objMappingDeliverable.MappedDeliverable.GovernanceControls != null
											|| (layer1upDeliverableID != null && objMappingDeliverable.MappedDeliverable.Layer1up.GovernanceControls != null)
											|| (layer2upDeliverableID != null && objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.GovernanceControls != null))
												{
												// Insert the Heading
												objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
												objRun1 = oxmlDocument.Construct_RunText(
													parText2Write: Properties.AppResources.Document_DeliverableGovernanceControls_Heading_Text);
												objParagraph.Append(objRun1);
												objBody.Append(objParagraph);

												// Insert Layer 2up if present and not null
												if(layer2upDeliverableID != null)
													{
													if(objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.GovernanceControls != null)
														{
														// Check if a hyperlink must be inserted
														if(documentCollection_HyperlinkURL != "")
															{
															hyperlinkCounter += 1;
															currentListURI = Properties.AppResources.SharePointURL +
																Properties.AppResources.List_DeliverablesURI +
																currentHyperlinkViewEditURI +
																objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.ID;
															}
														else
															currentListURI = "";

														if(this.ColorCodingLayer1)
															currentContentLayer = "Layer1";
														else
															currentContentLayer = "None";

														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objWorkbooktPart,
															parDocumentLevel: 6,
															parHTML2Decode: objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.GovernanceControls,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref tableCaptionCounter,
															parImageCaptionCounter: ref imageCaptionCounter,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														} //if(objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.GovernanceControls != null)
													} // if(layer2upDeliverableID != null)

												// Insert Layer 1up if present and not null
												if(layer1upDeliverableID != null)
													{
													if(objMappingDeliverable.MappedDeliverable.Layer1up.GovernanceControls != null)
														{
														// Check if a hyperlink must be inserted
														if(documentCollection_HyperlinkURL != "")
															{
															hyperlinkCounter += 1;
															currentListURI = Properties.AppResources.SharePointURL +
																Properties.AppResources.List_DeliverablesURI +
																currentHyperlinkViewEditURI +
																objMappingDeliverable.MappedDeliverable.Layer1up.ID;
															}
														else
															currentListURI = "";

														if(this.ColorCodingLayer1)
															currentContentLayer = "Layer2";
														else
															currentContentLayer = "None";

														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objWorkbooktPart,
															parDocumentLevel: 6,
															parHTML2Decode: objMappingDeliverable.MappedDeliverable.Layer1up.GovernanceControls,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref tableCaptionCounter,
															parImageCaptionCounter: ref imageCaptionCounter,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														} // if(objMappingDeliverable.MappedDeliverable.Layer1up.GovernanceControls != null)
													} // if(layer2upDeliverableID != null)

												// Insert Layer0up if not null
												if(objMappingDeliverable.MappedDeliverable.GovernanceControls != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objMappingDeliverable.MappedDeliverable.ID;
														}
													else
														currentListURI = "";

													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer3";
													else
														currentContentLayer = "None";

													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objWorkbooktPart,
														parDocumentLevel: 6,
														parHTML2Decode: objMappingDeliverable.MappedDeliverable.GovernanceControls,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref tableCaptionCounter,
														parImageCaptionCounter: ref imageCaptionCounter,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													} // if(objMappingDeliverable.MappedDeliverable.GovernanceControls != null)
												} // if(objMappingDeliverable.MappedDeliverable.GovernanceControls != null &&)	
											} //if(this.DRM_GovernanceControls)

										//---------------------------------------------------
										// Check if there are any Glossary Terms or Acronyms associated with the Deliverable(s).
										if(this.Acronyms_Glossary_of_Terms_Section)
											{
											// if there are GlossaryAndAcronyms to add from layer0up
											if(objMappingDeliverable.MappedDeliverable.GlossaryAndAcronyms.Count > 0)
												{
												foreach(var entry in objMappingDeliverable.MappedDeliverable.GlossaryAndAcronyms)
													{
													if(this.DictionaryGlossaryAndAcronyms.ContainsKey(entry.Key) != true)
														DictionaryGlossaryAndAcronyms.Add(entry.Key, entry.Value);
													}
												}
											// if there are GlossaryAndAcronyms to add from layer1up
											if(layer1upDeliverableID != null && objMappingDeliverable.MappedDeliverable.Layer1up.GlossaryAndAcronyms.Count > 0)
												{
												foreach(var entry in objMappingDeliverable.MappedDeliverable.Layer1up.GlossaryAndAcronyms)
													{
													if(this.DictionaryGlossaryAndAcronyms.ContainsKey(entry.Key) != true)
														DictionaryGlossaryAndAcronyms.Add(entry.Key, entry.Value);
													}
												}
											// if there are GlossaryAndAcronyms to add from layer2up
											if(layer2upDeliverableID != null && objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.GlossaryAndAcronyms.Count > 0)
												{
												foreach(var entry in objMappingDeliverable.MappedDeliverable.Layer1up.Layer1up.GlossaryAndAcronyms)
													{
													if(this.DictionaryGlossaryAndAcronyms.ContainsKey(entry.Key) != true)
														DictionaryGlossaryAndAcronyms.Add(entry.Key, entry.Value);
													}
												}
											} // if(this.Acronyms_Glossary_of_Terms_Section)
										} // if(objMappingDeliverable.NewDeliverable == false)
										  //------------------------------------------------
										  // If the user selected to include Service Levels
									if(this.Service_Level_Heading)
										{
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
											// Insert the Service Levels Heading:
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
											objRun1 = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_RequirementsMapping_ServiceLevelsHeading);
											objParagraph.Append(objRun1);
											objBody.Append(objParagraph);

											// Process all the Mapping Deliverables for the specific Service Requirement
											foreach(MappingServiceLevel objMappingServiceLevel in listMappingServiceLevels)
												{
												Console.WriteLine("\t\t\t\t + DRM: {0} - {1}", objMappingServiceLevel.ID, objMappingServiceLevel.Title);
												// Insert the MappingServiceLevel Title
												objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
												// If it is a new Mapping Service level, use the MappingService Levels's Title else use the actual
												// Mapped_ServiceLevel's CSD Description
												if(objMappingServiceLevel.NewServiceLevel)
													{
													objRun1 = oxmlDocument.Construct_RunText(parText2Write: objMappingServiceLevel.Title);
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
															parMainDocumentPart: ref objWorkbooktPart,
															parImageRelationshipId: hyperlinkImageRelationshipID,
															parClickLinkURL: Properties.AppResources.SharePointURL +
															Properties.AppResources.List_MappingServiceLevels +
															currentHyperlinkViewEditURI + objMappingServiceLevel.ID,
															parHyperlinkID: hyperlinkCounter);
														objRun1.Append(objDrawing);
														}
													}
												else
													{
													objRun1 = oxmlDocument.Construct_RunText(
														parText2Write: objMappingServiceLevel.MappedServiceLevel.CSDheading);
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
															parMainDocumentPart: ref objWorkbooktPart,
															parImageRelationshipId: hyperlinkImageRelationshipID,
															parClickLinkURL: Properties.AppResources.SharePointURL +
															Properties.AppResources.List_ServiceLevelsURI +
															currentHyperlinkViewEditURI + objMappingServiceLevel.MappedServiceLevel.ID,
															parHyperlinkID: hyperlinkCounter);
														objRun1.Append(objDrawing);
														}
													}
												objParagraph.Append(objRun1);
												objBody.Append(objParagraph);

												// Check if the user specified to include the Service Level Description
												if(this.Service_Level_Commitments_Table)
													{
													if(objMappingServiceLevel.NewServiceLevel)
														{
														objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
														objRun1 = oxmlDocument.Construct_RunText(parText2Write: objMappingServiceLevel.RequirementText);
														objParagraph.Append(objRun1);
														objBody.Append(objParagraph);
														}
													else
														{
														try
															{
															// Prepare the data which to insert into the Service Level Table
															List<string> listErrorMessagesParameter = this.ErrorMessages;
															// Populate the Service Level Table
															objServiceLevelTable = CommonProcedures.BuildSLAtable(
																parServiceLevelID: objMappingServiceLevel.MappedServiceLevel.ID,
																parWidthColumn1: Convert.ToUInt32(this.PageWith * 0.20),
																parWidthColumn2: Convert.ToUInt32(this.PageWith * 0.80),
																parMeasurement: objMappingServiceLevel.MappedServiceLevel.Measurement,
																parMeasureMentInterval: objMappingServiceLevel.MappedServiceLevel.MeasurementInterval,
																parReportingInterval: objMappingServiceLevel.MappedServiceLevel.ReportingInterval,
																parServiceHours: objMappingServiceLevel.MappedServiceLevel.ServiceHours,
																parCalculationMethod: objMappingServiceLevel.MappedServiceLevel.CalcualtionMethod,
																parCalculationFormula: objMappingServiceLevel.MappedServiceLevel.CalculationFormula,
																parThresholds: objMappingServiceLevel.MappedServiceLevel.PerfomanceThresholds,
																parTargets: objMappingServiceLevel.MappedServiceLevel.PerformanceTargets,
																parBasicServiceLevelConditions: objMappingServiceLevel.MappedServiceLevel.BasicConditions,
																parAdditionalServiceLevelConditions: "",
																parErrorMessages: ref listErrorMessagesParameter);

															if(listErrorMessagesParameter.Count != this.ErrorMessages.Count)
																this.ErrorMessages = listErrorMessagesParameter;

															objBody.Append(objServiceLevelTable);
															} // try
														catch(DataServiceClientException)
															{
															// If the entry is not found - write an error in the document and 
															// record an error in the error log.
															this.LogError("Error: The MappingServiceLevel ID " + objMappingServiceLevel.ID
																+ " doesn't exist in SharePoint and it couldn't be retrieved.");
															objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
															objRun1 = oxmlDocument.Construct_RunText(
																parText2Write: "Error: MappingServiceLevel: " + objMappingServiceLevel.ID + " is missing.",
																parIsNewSection: false,
																parIsError: true);
															objParagraph.Append(objRun1);
															objBody.Append(objParagraph);
															break;
															}
														catch(Exception exc)
															{
															Console.WriteLine("Exception occurred: {0} - {1}", exc.HResult, exc.Message);
															}
														} //else (objMappingServiceLevel.NewServiceLevel)
													} // if(this.Service_Level_Commitments_Table)
												} // foreach(MappingServiceLevel objMappingServiceLevel in listMappingServiceLevels)
											} // if(listMappingServiceLevels.Count != 0)
										} // if(this.Service_Level_Heading)
									} // foreach(MappingDeliverable objMappingDeliverable in listMappingDeliverables)
								} // if(listMappingDeliverables.Count != 0)
							} // if(this.Deliverable_Reports_and_Meetings)
						} // foreach(MappingRequirement objRequirement in listMappingRequirements)
					} //foreach(MappingServiceTower objTower in listMappingTowers)




Save_and_Close_Document:
				//----------------------------------------------
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
				//TODO: raise the error
				}
			catch(Exception exc)
				{
				Console.WriteLine("Exception: {0} - {1}", exc.HResult, exc.Message);
				}
			
			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}
		}
	}
