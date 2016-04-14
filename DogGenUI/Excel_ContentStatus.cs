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
		public bool Generate(ref CompleteDataSet parDataSet)
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			DateTime timeStarted = DateTime.Now;
			//string hyperlinkImageRelationshipID = "";
			string strDocumentCollection_HyperlinkURL = "";
			int intHyperlinkCounter = 9;
			string strCurrentHyperlinkViewEditURI = "";
			Cell objCell = new Cell();
			//Text Workstrings
			string strText = "";
			string strErrorText = "";
			//Status Stats variables
			int intStatusNew = 0;
			int intStatusWIP = 0;
			int intStatusQA = 0;
			int intStatusDone = 0;
			int intTotalStatus = 0;
			double dblStatusPercentage = 0;
			int intActualElements = 0;
			int intActualElementDeliverables = 0;
			int intActualElementReports = 0;
			int intActualElementMeetings = 0;
			int intActualServiceLevels = 0;
			int intActualActivities = 0;
			int intActualFeatures = 0;
			int intActualFeatureDeliverables = 0;
			int intActualFeatureReports = 0;
			int intActualFeatureMeetings = 0;
			int intTotalPlanned = 0;
			int intTotalActuals = 0;
			double dblPercentage_PlannedActuals = 0;
			
			//Worksheet Row Index Variables
			UInt16 intStatusSheet_RowIndex = 5;

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
				int intLastColumn = 29;
				int intStyleSourceRow = 6;
				string strCellAddress = "";
				for(int i = 0; i <= intLastColumn; i++)
					{
					strCellAddress = aWorkbook.GetColumnLetter(i) + intStyleSourceRow;
					Cell objSourceCell = objStatusWorksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == strCellAddress).FirstOrDefault();
					if(objSourceCell != null)
						{
						listColumnStyles.Add(objSourceCell.StyleIndex);
						//Console.WriteLine("\t\t\t\t + {0} - {1}", i, objSourceCell.StyleIndex);
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
				Deliverable objDeliverable = new Deliverable();
				List<ServiceElement> listServiceElements = new List<ServiceElement>();
				List<Deliverable> listElementDeliverables = new List<Deliverable>();
                    List<ServiceFeature> listServiceFeatures = new List<ServiceFeature>();
				List<Deliverable> listFeatureDeliverables = new List<Deliverable>();
				List<Activity> listDeliverableActivities = new List<Activity>();
				List<ServiceLevel> listDeliverableServiceLevels = new List<ServiceLevel>();

				//-------------------------------------
				// Begin to process the selected Nodes

				foreach(Hierarchy itemHierarchy in this.SelectedNodes)
					{
					switch(itemHierarchy.NodeType)
						{
					case (enumNodeTypes.POR):
					case (enumNodeTypes.FRA):
							{
							//objServicePortfolio.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID);
							objServicePortfolio = parDataSet.dsPortfolios.Where(p => p.Key == itemHierarchy.NodeID).FirstOrDefault().Value;

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
								{ strText = objServicePortfolio.Title; }

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
							for(int i = 1; i <= intLastColumn; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objStatusWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intStatusSheet_RowIndex,
									parStyleId: (UInt32Value)(listColumnStyles.ElementAt(i)),
									parCellDatatype: CellValues.String);
								//Console.WriteLine("\t\t\t\t + Column: {0} of {1}", i, intLastColumn);
								}
							break;
							}
					case (enumNodeTypes.FAM):
							{
							//objServiceFamily.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID);
							objServiceFamily = parDataSet.dsFamilies.Where(f => f.Key == itemHierarchy.NodeID).FirstOrDefault().Value;
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

							Console.WriteLine("\t\t + Family: {0} - {1}", objServiceFamily.ID, objServiceFamily.Title);
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
							for(int i = 2; i <= intLastColumn; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objStatusWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intStatusSheet_RowIndex,
									parStyleId: (UInt32Value)(listColumnStyles.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}
							break;
							}
					case (enumNodeTypes.PRO):
							{
							//objServiceProduct.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID);
							objServiceProduct = parDataSet.dsProducts.Where(p => p.Key == itemHierarchy.NodeID).FirstOrDefault().Value;
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
							intTotalStatus = 0;
							intActualElements = 0;
							intActualElementDeliverables = 0;
							intActualElementReports = 0;
							intActualElementMeetings = 0;
							intActualFeatures = 0;
							intActualFeatureDeliverables = 0;
							intActualFeatureReports = 0;
							intActualFeatureMeetings = 0;
							intActualServiceLevels = 0;
							intActualActivities = 0;
							intActualFeatureDeliverables = 0;
							intTotalPlanned = 0;
							intTotalActuals = 0;
							dblPercentage_PlannedActuals = 0;
							Console.WriteLine("\t\t\t + Prodcut: {0} - {1}", objServiceProduct.ID, objServiceProduct.Title);
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

							//-------------------------------------------
							// get the actual Element values
							intActualElements = 0;
							foreach(var elementEntry in parDataSet.dsElements
								.Where(e => e.Value.ServiceProductID == objServiceProduct.ID)
								.OrderBy(e => e.Value.SortOrder)
								.ThenBy(e => e.Value.Title))
								{
								intActualElements += 1;
                                        Console.WriteLine("\t\t\t\t + Element: {0} - {1}", elementEntry.Key, elementEntry.Value.Title);
								if(elementEntry.Value.ContentStatus != null)
									{
									if(elementEntry.Value.ContentStatus.Contains("New"))
										intStatusNew += 1;
									else if(elementEntry.Value.ContentStatus.Contains("WIP"))
										intStatusWIP += 1;
									else if(elementEntry.Value.ContentStatus.Contains("QA"))
										intStatusQA += 1;
									else if(elementEntry.Value.ContentStatus.Contains("Done"))
										intStatusDone += 1;
									}
								// Retrieve all the related ElementDeliverables
								foreach(var deliverableEntry in parDataSet.dsElementDeliverables
									.Where(ed => ed.Value.AssociatedElementID == elementEntry.Key))
									{
									Console.WriteLine("\t\t\t\t\t\t + ElementDeliverable: {0} - {1}",
										deliverableEntry.Key, deliverableEntry.Value.Title);

									objDeliverable = parDataSet.dsDeliverables
										.Where(d => d.Key == deliverableEntry.Value.AssociatedDeliverableID).First().Value;

									if(objDeliverable != null)
										{
										if(objDeliverable.DeliverableType.Contains("Deliverable"))
											intActualElementDeliverables += 1;
										else if(objDeliverable.DeliverableType.Contains("Report"))
											intActualElementReports += 1;
										else if(objDeliverable.DeliverableType.Contains("Meeting"))
											intActualElementMeetings += 1;

										if(objDeliverable.ContentStatus != null)
											{
											if(objDeliverable.ContentStatus.Contains("New"))
												intStatusNew += 1;
											else if(objDeliverable.ContentStatus.Contains("WIP"))
												intStatusWIP += 1;
											else if(objDeliverable.ContentStatus.Contains("QA"))
												intStatusQA += 1;
											else if(objDeliverable.ContentStatus.Contains("Done"))
												intStatusDone += 1;
											}

										// Retrieve all the Service Levels for each Deliverable and count the values
										foreach(var servicelevelEntry in parDataSet.dsDeliverableServiceLevels
											.Where(ds => ds.Value.AssociatedDeliverableID == deliverableEntry.Value.AssociatedDeliverableID
											&& ds.Value.AssociatedServiceProductID == objServiceProduct.ID))
											{
											Console.WriteLine("\t\t\t\t\t\t\t + DeliverableServiceLevel: {0} - {1}",
												servicelevelEntry.Value.AssociatedServiceLevelID,
												servicelevelEntry.Value.AssociatedServiceLevel.Title);
											intActualServiceLevels += 1;
											if(servicelevelEntry.Value.ContentStatus != null)
												{
												if(servicelevelEntry.Value.ContentStatus.Contains("New"))
													intStatusNew += 1;
												else if(servicelevelEntry.Value.ContentStatus.Contains("WIP"))
													intStatusWIP += 1;
												else if(servicelevelEntry.Value.ContentStatus.Contains("QA"))
													intStatusQA += 1;
												else if(servicelevelEntry.Value.ContentStatus.Contains("Done"))
													intStatusDone += 1;
												}
											} //foreach(ServiceLevel servicelevelEntry in ...)

										// Retrieve all the Activities for each Deliverable and count the values
										foreach(var activityEntry in parDataSet.dsDeliverableActivities
											.Where(da => da.Value.AssociatedDeliverableID == deliverableEntry.Value.AssociatedDeliverableID))
											{
											Console.WriteLine("\t\t\t\t\t\t + Activity: {0} - {1}",
												activityEntry.Value.AssociatedActivityID, activityEntry.Value.AssociatedActivity.Title);
											intActualActivities += 1;
											if(activityEntry.Value.AssociatedActivity.ContentStatus != null)
												{
												if(activityEntry.Value.AssociatedActivity.ContentStatus.Contains("New"))
													intStatusNew += 1;
												else if(activityEntry.Value.AssociatedActivity.ContentStatus.Contains("WIP"))
													intStatusWIP += 1;
												else if(activityEntry.Value.AssociatedActivity.ContentStatus.Contains("QA"))
													intStatusQA += 1;
												else if(activityEntry.Value.AssociatedActivity.ContentStatus.Contains("Done"))
													intStatusDone += 1;
												}
											} //foreach(Activity activityEntry in listDeliverableActivities)
										} //foreach(Deliverable deliverableEntry in listElementDeliverables)
									} // if(objDeliverable != null)
								} // end loop foreach(var elementEntry in ...)

							// Populate the rest of the columns
							//--- Status --- Service Product Row --- Column E --- Elements Planned ---
							if(objServiceProduct.PlannedElements > 0)
								{
								intTotalPlanned += Convert.ToInt16(objServiceProduct.PlannedElements);
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objStatusWorksheetPart,
									parColumnLetter: "E",
									parRowNumber: intStatusSheet_RowIndex,
									parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("E"))),
									parCellDatatype: CellValues.Number,
									parCellcontents: objServiceProduct.PlannedElements.ToString());
								}
							else
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objStatusWorksheetPart,
									parColumnLetter: "F",
									parRowNumber: intStatusSheet_RowIndex,
									parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("F"))),
									parCellDatatype: CellValues.String);
								}

								//--- Status --- Service Product Row --- Column F --- Element Deliverables Planned ---
								if(intActualElements > 0)
									{
									intTotalActuals += intActualElements;
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "F",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("F"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: intActualElements.ToString());
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "F",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("F"))),
										parCellDatatype: CellValues.String);
									}
								//--- Status --- Service Product Row --- Column G --- Element Deliverables Planned ---
								if(objServiceProduct.PlannedDeliverables > 0)
									{
									intTotalPlanned += Convert.ToInt16(objServiceProduct.PlannedDeliverables);
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "G",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("G"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: objServiceProduct.PlannedDeliverables.ToString());
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "G",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("G"))),
										parCellDatatype: CellValues.String);
									}
								//--- Status --- Service Product Row --- Column H --- Element Deliverables Actual ---
								if(intActualElementDeliverables > 0)
									{
									intTotalActuals += intActualElementDeliverables;
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "H",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("H"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: intActualElementDeliverables.ToString());
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "H",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("H"))),
										parCellDatatype: CellValues.String);
									}

								//--- Status --- Service Product Row --- Column I --- Element Reports Planned ---
								if(objServiceProduct.PlannedReports > 0)
									{
									intTotalPlanned += Convert.ToInt16(objServiceProduct.PlannedReports);
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "I",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("I"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: objServiceProduct.PlannedReports.ToString());
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "I",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("I"))),
										parCellDatatype: CellValues.String);
									}
								//--- Status --- Service Product Row --- Column J --- Element Reports Actual ---
								if(intActualElementReports > 0)
									{
									intTotalActuals += intActualElementReports;
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "J",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("J"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: intActualElementReports.ToString());
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "J",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("J"))),
										parCellDatatype: CellValues.String);
									}
								//--- Status --- Service Product Row --- Column K --- Element Meetings Planned ---
								if(objServiceProduct.PlannedMeetings > 0)
									{
									intTotalPlanned += Convert.ToInt16(objServiceProduct.PlannedMeetings);
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "K",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("K"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: objServiceProduct.PlannedMeetings.ToString());
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "K",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("K"))),
										parCellDatatype: CellValues.String);
									}
								//--- Status --- Service Product Row --- Column L --- Element Meetings Actual ---
								if(intActualElementMeetings > 0)
									{
									intTotalActuals += intActualElementMeetings;
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "L",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("L"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: intActualElementMeetings.ToString());
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "L",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("L"))),
										parCellDatatype: CellValues.String);
									}

								//-----------------------------------------------
								// Obtain the stats for the Features
								// get the actual Feature values
								intActualFeatures = 0;
								foreach(var featureEntry in parDataSet.dsFeatures
								.Where(f => f.Value.ServiceProductID == objServiceProduct.ID))
									{
									Console.WriteLine("\t\t\t\t + Feature: {0} - {1}", featureEntry.Key, featureEntry.Value.Title);
									intActualFeatures += 1;
									if(featureEntry.Value.ContentStatus != null)
										{
										if(featureEntry.Value.ContentStatus.Contains("New"))
											intStatusNew += 1;
										else if(featureEntry.Value.ContentStatus.Contains("WIP"))
											intStatusWIP += 1;
										else if(featureEntry.Value.ContentStatus.Contains("QA"))
											intStatusQA += 1;
										else if(featureEntry.Value.ContentStatus.Contains("Done"))
											intStatusDone += 1;
										}
									// Retrieve all the related FeatureDeliverables
									foreach(var featureDeliverableEntry in parDataSet.dsFeatureDeliverables
										.Where(fd => fd.Value.AssociatedFeatureID == featureEntry.Key))
										{
										objDeliverable = parDataSet.dsDeliverables
											.Where(d => d.Key == featureDeliverableEntry.Value.AssociatedDeliverableID).First().Value;										
										Console.WriteLine("\t\t\t\t\t\t + FeatureDeliverable: {0} - {1} ({2})", 
											featureDeliverableEntry.Key, 
											featureDeliverableEntry.Value.Title,
											featureDeliverableEntry.Value.AssociatedDeliverableID);
										if(objDeliverable.DeliverableType.Contains("Deliverable"))
											intActualFeatureDeliverables += 1;
										else if(objDeliverable.DeliverableType.Contains("Report"))
											intActualFeatureReports += 1;
										else if(objDeliverable.DeliverableType.Contains("Meeting"))
											intActualFeatureMeetings += 1;

										if(objDeliverable.ContentStatus != null)
											{
											if(objDeliverable.ContentStatus.Contains("New"))
												intStatusNew += 1;
											else if(objDeliverable.ContentStatus.Contains("WIP"))
												intStatusWIP += 1;
											else if(objDeliverable.ContentStatus.Contains("QA"))
												intStatusQA += 1;
											else if(objDeliverable.ContentStatus.Contains("Done"))
												intStatusDone += 1;
											}

										// Retrieve all the Service Levels for each Deliverable and count the values
										foreach(var servicelevelEntry in parDataSet.dsDeliverableServiceLevels
										.Where(ds => ds.Value.AssociatedDeliverableID == objDeliverable.ID
												&& ds.Value.AssociatedServiceProductID == objServiceProduct.ID))
											{
											Console.WriteLine("\t\t\t\t\t\t\t + DeliverableServiceLevel: {0} - {1}",
												servicelevelEntry.Value.AssociatedServiceLevelID, 
												servicelevelEntry.Value.AssociatedServiceLevel.Title);
											intActualServiceLevels += 1;
											if(servicelevelEntry.Value.AssociatedServiceLevel.ContentStatus != null)
												{
												if(servicelevelEntry.Value.AssociatedServiceLevel.ContentStatus.Contains("New"))
													intStatusNew += 1;
												else if(servicelevelEntry.Value.AssociatedServiceLevel.ContentStatus.Contains("WIP"))
													intStatusWIP += 1;
												else if(servicelevelEntry.Value.AssociatedServiceLevel.ContentStatus.Contains("QA"))
													intStatusQA += 1;
												else if(servicelevelEntry.Value.AssociatedServiceLevel.ContentStatus.Contains("Done"))
													intStatusDone += 1;
												}
											} //foreach(ServiceLevel servicelevelEntry in ...)

										// Retrieve all the Activities for each Deliverable and count the values
										foreach(var activityEntry in parDataSet.dsDeliverableActivities
										.Where(da => da.Value.AssociatedDeliverableID == objDeliverable.ID))
											{
											Console.WriteLine("\t\t\t\t\t\t + Activity: {0} - {1}",
												activityEntry.Value.AssociatedActivityID, activityEntry.Value.AssociatedActivity.Title);
											intActualActivities += 1;
											if(activityEntry.Value.AssociatedActivity.ContentStatus != null)
												{
												if(activityEntry.Value.AssociatedActivity.ContentStatus.Contains("New"))
													intStatusNew += 1;
												else if(activityEntry.Value.AssociatedActivity.ContentStatus.Contains("WIP"))
													intStatusWIP += 1;
												else if(activityEntry.Value.AssociatedActivity.ContentStatus.Contains("QA"))
													intStatusQA += 1;
												else if(activityEntry.Value.AssociatedActivity.ContentStatus.Contains("Done"))
													intStatusDone += 1;
												}
											} //foreach(Activity activityEntry in listDeliverableActivities)
										} //foreach(Deliverable deliverableEntry in listFeatureDeliverables)
									} // end loop foreach(var featureEntry in listServiceFeatures)
								//--- Status --- Service Product Row --- Column M --- Features Quantities Planned ---
								if(objServiceProduct.PlannedFeatures > 0)
									{
									intTotalPlanned += Convert.ToInt16(objServiceProduct.PlannedFeatures);
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "M",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("M"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: objServiceProduct.PlannedFeatures.ToString());
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "M",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("M"))),
										parCellDatatype: CellValues.String);
									}
								//--- Status --- Service Product Row --- Column N --- Feature Quantities Actual ---
								if(intActualFeatures > 0)
									{
									intTotalActuals += intActualFeatures;
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "N",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("N"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: intActualFeatures.ToString());
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "N",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("N"))),
										parCellDatatype: CellValues.String);
									}
								//--- Status --- Service Product Row --- Column O --- Features Deliverables Planned ---
								if(objServiceProduct.PlannedDeliverables > 0)
									{
									intTotalPlanned += Convert.ToInt16(objServiceProduct.PlannedDeliverables);
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "O",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("O"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: objServiceProduct.PlannedDeliverables.ToString());
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "O",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("O"))),
										parCellDatatype: CellValues.String);
									}
								//--- Status --- Service Product Row --- Column P --- Feature Deliverables Actual ---
								if(intActualFeatureDeliverables > 0)
									{
									intActualFeatureDeliverables += intActualFeatureDeliverables;
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "P",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("P"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: intActualFeatureDeliverables.ToString());
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "P",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("P"))),
										parCellDatatype: CellValues.String);
									}

								//--- Status --- Service Product Row --- Column Q --- Features Reports Planned ---
								if(objServiceProduct.PlannedReports > 0)
									{
									intTotalPlanned += Convert.ToInt16(objServiceProduct.PlannedReports);
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "Q",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("Q"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: objServiceProduct.PlannedReports.ToString());
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "Q",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("Q"))),
										parCellDatatype: CellValues.String);
									}
								//--- Status --- Service Product Row --- Column R --- Feature Reports Actual ---
								if(intActualFeatureReports > 0)
									{
									intTotalActuals += intActualFeatureReports;
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "R",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("R"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: intActualFeatureReports.ToString());
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "R",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("R"))),
										parCellDatatype: CellValues.String);
									}

								//--- Status --- Service Product Row --- Column S --- Features Meetings Planned ---
								if(objServiceProduct.PlannedMeetings > 0)
									{
									intTotalPlanned += Convert.ToInt16(objServiceProduct.PlannedMeetings);
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "S",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("S"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: objServiceProduct.PlannedMeetings.ToString());
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "S",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("S"))),
										parCellDatatype: CellValues.String);
									}
								//--- Status --- Service Product Row --- Column T --- Feature Meetings Actual ---
								if(intActualFeatureMeetings > 0)
									{
									intTotalActuals += intActualFeatureMeetings;
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "T",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("T"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: intActualFeatureMeetings.ToString());
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "T",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("T"))),
										parCellDatatype: CellValues.String);
									}

								//--- Status --- Service Product Row --- Column U --- Service Levels Planned ---
								if(objServiceProduct.PlannedServiceLevels > 0)
									{
									intTotalPlanned += Convert.ToInt16(objServiceProduct.PlannedServiceLevels);
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "U",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("U"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: objServiceProduct.PlannedServiceLevels.ToString());
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "U",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("U"))),
										parCellDatatype: CellValues.String);
									}

								//--- Status --- Service Product Row --- Column V --- Service Levels Actual ---
								if(intActualServiceLevels > 0)
									{
									intTotalActuals += intActualServiceLevels;
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "V",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("V"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: (intActualServiceLevels / 2).ToString());
									// Divide the Actual Service Levels by 2 beacuse the same Service Levels are counted twice because the same deliverables
									// are suppose to be associated with the Service Elements and Service Features.
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "V",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("V"))),
										parCellDatatype: CellValues.String);
									}

								//--- Status --- Service Product Row --- Column W --- Activities Planned ---
								if(objServiceProduct.PlannedActivities > 0)
									{
									intTotalPlanned += Convert.ToInt16(objServiceProduct.PlannedActivities);
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "W",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("W"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: (objServiceProduct.PlannedActivities / 2).ToString());
									// Divide the Actual Activities by 2 because the same Activities are counted twice because the same deliverables
									// are suppose to be associated with the Service Elements and Service Features.
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "W",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("W"))),
										parCellDatatype: CellValues.String);
									}

								//--- Status --- Service Product Row --- Column X --- Activities Actual ---
								if(intActualActivities > 0)
									{
									intTotalActuals += intActualActivities;
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "X",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("X"))),
										parCellDatatype: CellValues.Number,
										parCellcontents: intActualActivities.ToString());
									}
								else
									{
									oxmlWorkbook.PopulateCell(
										parWorksheetPart: objStatusWorksheetPart,
										parColumnLetter: "X",
										parRowNumber: intStatusSheet_RowIndex,
										parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("X"))),
										parCellDatatype: CellValues.String);
									}

							//--- Status --- Service Product Row --- Column Y --- % Planned vs Actuals ---
							if(intTotalActuals > 0 && intTotalPlanned > 0)
								{
								if(intTotalActuals > intTotalPlanned)
									dblPercentage_PlannedActuals = 1;
								else
									dblPercentage_PlannedActuals = intTotalActuals / intTotalPlanned;

								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objStatusWorksheetPart,
									parColumnLetter: "Y",
									parRowNumber: intStatusSheet_RowIndex,
									parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("Y"))),
									parCellDatatype: CellValues.Number,
									parCellcontents: dblPercentage_PlannedActuals.ToString());
								}
							else
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objStatusWorksheetPart,
									parColumnLetter: "Y",
									parRowNumber: intStatusSheet_RowIndex,
									parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("Y"))),
									parCellDatatype: CellValues.Number,
									parCellcontents: "0");
								}

							//--- Status --- Service Product Row --- Column Z --- blank column ---
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objStatusWorksheetPart,
								parColumnLetter: "Z",
								parRowNumber: intStatusSheet_RowIndex,
								parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("Z"))),
								parCellDatatype: CellValues.String);

							//--- Status --- Service Product Row --- Column AA --- % New Status ---
							intTotalStatus = intStatusNew + intStatusWIP + intStatusQA + intStatusDone;
							if(intStatusNew > 0 && intTotalStatus > 0)
								{
								if(intStatusNew > intTotalStatus)
									dblStatusPercentage = 1;
								else
									dblStatusPercentage = intStatusNew / intTotalStatus;

								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objStatusWorksheetPart,
									parColumnLetter: "AA",
									parRowNumber: intStatusSheet_RowIndex,
									parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("AA"))),
									parCellDatatype: CellValues.Number,
									parCellcontents: dblStatusPercentage.ToString());
								}
							else
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objStatusWorksheetPart,
									parColumnLetter: "AA",
									parRowNumber: intStatusSheet_RowIndex,
									parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("AA"))),
									parCellDatatype: CellValues.Number,
									parCellcontents: "0");
								}

							//--- Status --- Service Product Row --- Column AB --- % WIP Status ---

							if(intStatusWIP > 0 && intTotalStatus > 0)
								{
								if(intStatusWIP > intTotalStatus)
									dblStatusPercentage = 1;
								else
									dblStatusPercentage = intStatusWIP / intTotalStatus;

								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objStatusWorksheetPart,
									parColumnLetter: "AB",
									parRowNumber: intStatusSheet_RowIndex,
									parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("AB"))),
									parCellDatatype: CellValues.Number,
									parCellcontents: dblStatusPercentage.ToString());
								}
							else
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objStatusWorksheetPart,
									parColumnLetter: "AB",
									parRowNumber: intStatusSheet_RowIndex,
									parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("AB"))),
									parCellDatatype: CellValues.Number,
									parCellcontents: "0");
								}

							//--- Status --- Service Product Row --- Column AC --- % QA Status ---
							if(intStatusQA > 0 && intTotalStatus > 0)
								{
								if(intStatusQA > intTotalStatus)
									dblStatusPercentage = 1;
								else
									dblStatusPercentage = intStatusQA / intTotalStatus;

								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objStatusWorksheetPart,
									parColumnLetter: "AC",
									parRowNumber: intStatusSheet_RowIndex,
									parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("AC"))),
									parCellDatatype: CellValues.Number,
									parCellcontents: dblStatusPercentage.ToString());
								}
							else
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objStatusWorksheetPart,
									parColumnLetter: "AC",
									parRowNumber: intStatusSheet_RowIndex,
									parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("AC"))),
									parCellDatatype: CellValues.Number,
									parCellcontents: "0");
								}

							//--- Status --- Service Product Row --- Column AD --- % Done Status ---

							if(intStatusDone > 0 && intTotalStatus > 0)
								{
								if(intStatusDone > intTotalStatus)
									dblStatusPercentage = 1;
								else
									dblStatusPercentage = intStatusDone / intTotalStatus;

								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objStatusWorksheetPart,
									parColumnLetter: "AD",
									parRowNumber: intStatusSheet_RowIndex,
									parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("AD"))),
									parCellDatatype: CellValues.Number,
									parCellcontents: dblStatusPercentage.ToString());
								}
							else
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objStatusWorksheetPart,
									parColumnLetter: "AD",
									parRowNumber: intStatusSheet_RowIndex,
									parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("AD"))),
									parCellDatatype: CellValues.Number,
									parCellcontents: "0");
								}

							break;
							}
					default:
							{
							// just ignore any other NodeType
							break;
							}
						}// end switch(itemHierarchy.NodeType)


					} // end of foreach (Hierarchy itemHierarchy in this.SelectedNodes

Save_and_Close_Document:
				//===============================================================

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
