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
	/// This class handles the Content Status Workbook
	/// </summary>
	class Content_Status_Workbook:aWorkbook
		{
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
			int intRequirementBreakID_forAssumptions = 0;     // the ID value of the Requirement used as a break processing variable for Assumptions sheet

			//WorkString
			string strText = "";
			//Status Stats
			int intStatusNew = 0;
			int intStatusWIP = 0;
			int intStatusQA = 0;
			int intStatusDone = 0;
			int intDeliverables = 0;
			int intReports = 0;
			int intMeetings = 0;
			int intServiceLevels = 0;
			int intActivities = 0;
			int intEffortDrivers = 0;

			//Worksheet Row Index Variables
			UInt16 intStatusSheet_RowIndex = 6;
			UInt16 intColumnCounter;

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
				strErrorText = "An ERROR occurred and the new MS Excel Workbook could not be created due to above stated ERROR conditions.";
				Console.WriteLine(strErrorText);
				this.ErrorMessages.Add(strErrorText);
				return false;
				}

			if(this.SelectedNodes.Count < 1)
				{
				strErrorText = "The user didn't select any Nodes to populate the Workbook.";
				Console.WriteLine("\t\t\t ***" + strErrorText);
                    this.ErrorMessages.Add(strErrorText);
				return false;
				}

			// Open the MS Excel Workbook 
			try
				{
				// Open the MS Excel document in Edit mode
				SpreadsheetDocument objSpreadsheetDocument = SpreadsheetDocument.Open(path: objOXMLworkbook.LocalURI, isEditable: true);
				// Obtain the WorkBookPart from the spreadsheet.
				if(objSpreadsheetDocument.WorkbookPart == null)
					{
					throw new ArgumentException(objOXMLworkbook.LocalURI + " does not contain a WorkbookPart. There is a problem with the template file.");
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

				// obtain the Content Status Worksheet in the Workbook.
				Sheet objStatusWorksheet = objWorkbookPart.Workbook.Descendants<Sheet>().Where(sht => sht.Name == Properties.AppResources.Workbook_ContentStatus_WorksheetName).FirstOrDefault();
				if(objStatusWorksheet == null)
					{
					throw new ArgumentException("The " + Properties.AppResources.Workbook_ContentStatus_WorksheetName +
						" worksheet could not be loacated in the workbook.");
					}
				// obtain the WorksheetPart of the objMatrixWorksheet
				WorksheetPart objStatusWorksheetPart = (WorksheetPart)(objWorkbookPart.GetPartById(objStatusWorksheet.Id));

				// Copy the Formats from Row 7 into the List of Formats from where it can be applied to every Row
				Content_Status_Workbook objSatusWorkbook = new Content_Status_Workbook();
				List<UInt32Value> listColumnStyles = new List<UInt32Value>();
				int intLastColumn = 27;
				int intStyleSourceRow = 7;
				string strCellAddress = "";
				for(int i = 0; i < intLastColumn - 1; i++)
					{
					strCellAddress = aWorkbook.GetColumnLetter(i) + intStyleSourceRow;
					Cell objSourceCell = objStatusWorksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == strCellAddress).FirstOrDefault();
					if(objSourceCell != null)
						{
						listColumnStyles.Add(objSourceCell.StyleIndex);
						}
					else
						listColumnStyles.Add(0U);
					} // loop

				// If Hyperlinks need to be inserted, add the 
				Hyperlinks objHyperlinks = new Hyperlinks();

				// Decalre all the object to be used during the processing
				ServicePortfolio objServicePortfolio = new ServicePortfolio();
				ServiceFamily objServiceFamily = new ServiceFamily();
				ServiceProduct objServiceProduct = new ServiceProduct();
				//ServiceElement objServiceElement = new ServiceElement();
				List<ServiceElement> listServiceElements = new List<ServiceElement>();
				//ServiceFeature objServiceFeature = new ServiceFeature();
                    List<ServiceFeature> listServiceFeatures = new List<ServiceFeature>();
				List<ElementDeliverable> listElementDeliverables = new List<ElementDeliverable>();

				List<ServiceLevel> listServiceLevels = new List<ServiceLevel>();
				

				//-------------------------------------
				// Begin to process the selected Nodes

				foreach(Hierarchy itemHierarchy in this.SelectedNodes)
					{
					switch(itemHierarchy.NodeType)
						{
						case (enumNodeTypes.POR):
						case (enumNodeTypes.FRA):
							{
							objServicePortfolio.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID);
							if(objServicePortfolio.ID == 0) // the entry could not be found
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
								strErrorText = "Error: The Service Portfolio ID " + itemHierarchy.NodeID +
									" doesn't exist in SharePoint and couldn't be retrieved.";
								this.LogError(strErrorText);
								strErrorText = "Error: Service Portfolio " + itemHierarchy.NodeID + " is missing.";
								strText = strErrorText;
								}
							else
								{strText = objServicePortfolio.Title;}

							//--- Status --- Service Portfolio Row --- Column A -----
							// Write the Portfolio or Frameworkto the Workbook as a String
							Console.WriteLine("\t + Portfolio: {0} - {1}", objServicePortfolio.ID, objServicePortfolio.Title);
							intStatusSheet_RowIndex += 1;
							
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objStatusWorksheetPart,
								parColumnLetter: "A",
								parRowNumber: intStatusSheet_RowIndex,
								parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("A"))),
								parCellDatatype: CellValues.String,
								parCellcontents: strText);
							
							//--- Status --- Populate the styles for column B to Z ---
							for(int i = 1; i < intLastColumn; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objStatusWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parCellReference: i.ToString()),
									parRowNumber: intStatusSheet_RowIndex,
									parStyleId: (UInt32Value)(listColumnStyles.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}
							break;
							}
						case (enumNodeTypes.FAM):
							{
							objServiceFamily.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID);
							if(objServiceFamily.ID == 0) // the entry could not be found
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
								strErrorText = "Error: The Service Family ID " + itemHierarchy.NodeID +
									" doesn't exist in SharePoint and couldn't be retrieved.";
								this.LogError(strErrorText);
								strErrorText = "Error: Service Portfolio " + itemHierarchy.NodeID + " is missing.";
								strText = strErrorText;
								}
							else
								{
								strText = objServiceFamily.Title;
								}

							Console.WriteLine("\t\t + Family: {0} - {1}", objServicePortfolio.ID, objServicePortfolio.Title);
							intStatusSheet_RowIndex += 1;
							//--- Status --- Service Portfolio Row --- Column A -----
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objStatusWorksheetPart,
								parColumnLetter: "A",
								parRowNumber: intStatusSheet_RowIndex,
								parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("A"))),
								parCellDatatype: CellValues.String);
							// Write the Portfolio or Frameworkto the Workbook as a String
							//--- Status --- Service Family Row --- Column B -----

							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objStatusWorksheetPart,
								parColumnLetter: "B",
								parRowNumber: intStatusSheet_RowIndex,
								parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("B"))),
								parCellDatatype: CellValues.String,
								parCellcontents: strText);

							//--- Status --- Populate the styles for column B to Z ---
							for(int i = 2; i < intLastColumn; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objStatusWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parCellReference: i.ToString()),
									parRowNumber: intStatusSheet_RowIndex,
									parStyleId: (UInt32Value)(listColumnStyles.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}
							break;
							}
						case (enumNodeTypes.PRO):
							{
							objServiceProduct.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID);
							if(objServiceProduct.ID == 0) // the entry could not be found
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
								strErrorText = "Error: The Service Product ID " + itemHierarchy.NodeID +
									" doesn't exist in SharePoint and couldn't be retrieved.";
								this.LogError(strErrorText);
								strErrorText = "Error: Service Product " + itemHierarchy.NodeID + " is missing.";
								strText = strErrorText;
								}
							else
								{
								strText = objServiceProduct.Title;
								}

							intStatusNew = 0;
							intStatusWIP = 0;
							intStatusQA = 0;
							intStatusDone = 0;
							intDeliverables = 0;
							intReports = 0;
							intMeetings = 0;
							intServiceLevels = 0;
							intActivities = 0;
							intEffortDrivers = 0;
							Console.WriteLine("\t\t\t + Prodcut: {0} - {1}", objServicePortfolio.ID, objServicePortfolio.Title);
							intStatusSheet_RowIndex += 1;
							//--- Status --- Service Product Row --- Column A -----
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objStatusWorksheetPart,
								parColumnLetter: "A",
								parRowNumber: intStatusSheet_RowIndex,
								parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("A"))),
								parCellDatatype: CellValues.String);
							
							//--- Status --- Service Product Row --- Column B -----
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objStatusWorksheetPart,
								parColumnLetter: "B",
								parRowNumber: intStatusSheet_RowIndex,
								parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("B"))),
								parCellDatatype: CellValues.String);

							// Write the Product to the Workbook as a String
							//--- Status --- Service Product Row --- Column C -----
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objStatusWorksheetPart,
								parColumnLetter: "C",
								parRowNumber: intStatusSheet_RowIndex,
								parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("C"))),
								parCellDatatype: CellValues.String,
								parCellcontents: strText);

							//--- Status --- Service Product Row --- Column D -----
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objStatusWorksheetPart,
								parColumnLetter: "D",
								parRowNumber: intStatusSheet_RowIndex,
								parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("D"))),
								parCellDatatype: CellValues.String);

							//--- Status --- Service Product Row --- Column E --- Elements Planned ---
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objStatusWorksheetPart,
								parColumnLetter: "E",
								parRowNumber: intStatusSheet_RowIndex,
								parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("E"))),
								parCellDatatype: CellValues.Number,
								parCellcontents: objServiceProduct.PlannedElements.ToString());

							//--- Status --- Service Product Row --- Column F --- Elements Actual ---
							// get the actual Element values
							listServiceElements.Clear();
							listServiceElements = ServiceElement.ObtainListOfObjects(
								parDatacontextSDDP: datacontexSDDP,
								parServiceProductID: objServiceProduct.ID,
								parGetContentLayers: false);
							if(listServiceElements.Count > 0)
								{
								strText = listServiceElements.Count.ToString();
								foreach(var elementEntry in listServiceElements)
									{
									if(elementEntry.ContentStatus != null)
										{
										if(elementEntry.ContentStatus.Contains("New"))
											intStatusNew += 1;
										else if(elementEntry.ContentStatus.Contains("WIP"))
											intStatusWIP += 1;
										else if(elementEntry.ContentStatus.Contains("QA"))
											intStatusQA += 1;
										else if(elementEntry.ContentStatus.Contains("Done"))
											intStatusDone += 1;
										}
									}
								}
							else
								strText = "0";

							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objStatusWorksheetPart,
								parColumnLetter: "F",
								parRowNumber: intStatusSheet_RowIndex,
								parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("F"))),
								parCellDatatype: CellValues.Number,
								parCellcontents: strText.ToString());

							//--- Status --- Service Product Row --- Column G --- Features Planned ---
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objStatusWorksheetPart,
								parColumnLetter: "G",
								parRowNumber: intStatusSheet_RowIndex,
								parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("G"))),
								parCellDatatype: CellValues.Number,
								parCellcontents: objServiceProduct.PlannedFeatures.ToString());

							//--- Status --- Service Product Row --- Column H --- Features Actual ---
							// get the actual Features
							listServiceFeatures.Clear();
							listServiceFeatures = ServiceFeature.ObtainListOfObjects(
								parDatacontextSDDP: datacontexSDDP,
								parServiceProductID: objServiceProduct.ID,
								parGetContentLayers: false);
							if(listServiceFeatures.Count > 0)
								{
								strText = listServiceFeatures.Count.ToString();
								foreach(var featureEntry in listServiceElements)
									{
									if(featureEntry.ContentStatus != null)
										{
										if(featureEntry.ContentStatus.Contains("New"))
											intStatusNew += 1;
										else if(featureEntry.ContentStatus.Contains("WIP"))
											intStatusWIP += 1;
										else if(featureEntry.ContentStatus.Contains("QA"))
											intStatusQA += 1;
										else if(featureEntry.ContentStatus.Contains("Done"))
											intStatusDone += 1;
										}
									}
								}
							else
								strText = "0";

							// Update the Status Stats
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objStatusWorksheetPart,
								parColumnLetter: "H",
								parRowNumber: intStatusSheet_RowIndex,
								parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("H"))),
								parCellDatatype: CellValues.Number,
								parCellcontents: strText.ToString());

							//--- Status --- Service Product Row --- Column I --- Deliverables Planned ---
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objStatusWorksheetPart,
								parColumnLetter: "I",
								parRowNumber: intStatusSheet_RowIndex,
								parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("I"))),
								parCellDatatype: CellValues.Number,
								parCellcontents: objServiceProduct.PlannedDeliverables.ToString());

							//--- Status --- Service Family Row --- Column J ---Deliverables Actual ---
							// get the actual Deliverables
							listElementDeliverables.Clear();
							foreach(ServiceElement item in listServiceElements)
								{
								---- gaan hier aan ---
								}

							listDeliverables = Deliverable.ObtainListOfObjects(
								parDatacontextSDDP: datacontexSDDP,
								parServiceProductID: objServiceProduct.ID,
								parGetContentLayers: false);
							if(listServiceFeatures.Count > 0)
								{
								strText = listServiceFeatures.Count.ToString();
								foreach(var featureEntry in listServiceElements)
									{
									if(featureEntry.ContentStatus != null)
										{
										if(featureEntry.ContentStatus.Contains("New"))
											intStatusNew += 1;
										else if(featureEntry.ContentStatus.Contains("WIP"))
											intStatusWIP += 1;
										else if(featureEntry.ContentStatus.Contains("QA"))
											intStatusQA += 1;
										else if(featureEntry.ContentStatus.Contains("Done"))
											intStatusDone += 1;
										}
									}
								}
							else
								strText = "0";

							// Update the Status Stats
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objStatusWorksheetPart,
								parColumnLetter: "H",
								parRowNumber: intStatusSheet_RowIndex,
								parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("H"))),
								parCellDatatype: CellValues.Number,
								parCellcontents: strText.ToString());




							//--- Status --- Populate the styles for column B to Z ---
							for(int i = 2; i < intLastColumn; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objStatusWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parCellReference: i.ToString()),
									parRowNumber: intStatusSheet_RowIndex,
									parStyleId: (UInt32Value)(listColumnStyles.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}

							break;
							}
						default:
							{
							// just ignore any other NodeType
							break;
							}
						}// end switch(itemHierarchy.NodeType)


					}


				objCell = oxmlWorkbook.InsertCellInWorksheet(
					parColumnName: "A",
					parRowNumber: intMatrixSheet_RowIndex,
					parWorksheetPart: objMatrixWorksheetPart);
				objCell.DataType = new EnumValue<CellValues>(CellValues.String);
				objCell.CellValue = new CellValue(strErrorText);
				goto Save_and_Close_Document;



				Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}
		}
	}
