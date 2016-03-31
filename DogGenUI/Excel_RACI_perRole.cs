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
	/// This class handles the RACI Workbook per Role
	/// </summary>
	class RACI_Workbook_per_Role:aWorkbook
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

			//Text Workstrings
			string strCatalogueText = "";
			string strPortfolio = "";
			string strFamily = "";
			string strProduct = "";
			string strElement = "";
			string strDeliverable = "";
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
				Sheet objWorksheet = objWorkbookPart.Workbook.Descendants<Sheet>().Where(sht => sht.Name == Properties.AppResources.Workbook_RACI_perRole_WorksheetName).FirstOrDefault();
				if(objWorksheet == null)
					{
					throw new ArgumentException("The " + Properties.AppResources.Workbook_RACI_perRole_WorksheetName +
						" worksheet could not be loacated in the workbook.");
					}
				// obtain the WorksheetPart of the objMatrixWorksheet
				WorksheetPart objWorksheetPart = (WorksheetPart)(objWorkbookPart.GetPartById(objWorksheet.Id));

				// Copy the Cell Formats as StyleIDs into a list for later use. 
				// --- Style for Column A3 to D7
				List<UInt32Value> listColumnStylesA3D3 = new List<UInt32Value>();
				int intLastColumn = 4;
				int intStyleSourceRow = 3;
				string strCellAddress = "";
				for(int i = 0; i <= intLastColumn; i++)
					{
					strCellAddress = aWorkbook.GetColumnLetter(i) + intStyleSourceRow;
					Cell objSourceCell = objWorksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == strCellAddress).FirstOrDefault();
					if(objSourceCell != null)
						{
						listColumnStylesA3D3.Add(objSourceCell.StyleIndex);
						Console.WriteLine("\t\t\t\t + {0} - {1}", i, objSourceCell.StyleIndex);
						}
					else
						listColumnStylesA3D3.Add(0U);
					} // loop

				// If Hyperlinks need to be inserted, add the 
				Hyperlinks objHyperlinks = new Hyperlinks();

				// Decalre all the object to be used during processing
				ServicePortfolio objServicePortfolio = new ServicePortfolio();
				ServiceFamily objServiceFamily = new ServiceFamily();
				ServiceProduct objServiceProduct = new ServiceProduct();
				ServiceElement objServiceElement = new ServiceElement();
				Deliverable objDeliverable = new Deliverable();
				JobRole objJobRole = new JobRole();
				// Define the Dictionaries that will be represent the matrix
				// This dictionary will contain the the JobRole ID as the KEY and the VALUE will contain the JobRole Title
				Dictionary<int, JobRole> dictOfJobRoles = new Dictionary<int, JobRole>();
				// This dictionary contains all the Service Catalogue Srtuctures that need to be populated in the worksheet.
				// Key = intCatalogueIndex Value = Concatenated Service Catalogue Structure Text
				Dictionary<int, String> dictStructure = new Dictionary<int, string>();
				int intCatalogueIndex = 0; // This integer is used as the Key 
				// Each of the following dictionaries will contain the Matrix in which Key = intCatalogueIndex and the VALUE = JobRoleID.
				Dictionary<int, int> dictAccountableMarix = new Dictionary<int, int>();
				Dictionary<int, int> dictResponsibleMarix = new Dictionary<int, int>();
				Dictionary<int, int> dictConsultedMarix = new Dictionary<int, int>();
				Dictionary<int, int> dictInformedMarix = new Dictionary<int, int>();

				foreach(Hierarchy itemHierarchy in this.SelectedNodes)
					{
					switch(itemHierarchy.NodeType)
						{
						//-----------------------
						case (enumNodeTypes.POR):
						case (enumNodeTypes.FRA):
						//-----------------------
							{
							objServicePortfolio.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID);
							if(objServicePortfolio.ID == 0) // the entry could not be found
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
								strErrorText = "Error: The Service Portfolio ID " + itemHierarchy.NodeID +
									" doesn't exist in SharePoint and couldn't be retrieved.";
								this.LogError(strErrorText);
								strErrorText = "Error: Service Portfolio " + itemHierarchy.NodeID + " is missing.";
								strPortfolio = strErrorText;
								}
							else
								{
								strPortfolio = objServicePortfolio.Title;
								}
							Console.WriteLine("\t + Portfolio: {0} - {1}", objServicePortfolio.ID, strPortfolio);
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
								strErrorText = "Error: Service Family " + itemHierarchy.NodeID + " is missing.";
								strFamily = strErrorText;
								}
							else
								{
								strFamily = objServiceFamily.Title;
								}

							Console.WriteLine("\t\t + Family: {0} - {1}", objServiceFamily.ID, strFamily);
							break;
							}
						//-----------------------
						case (enumNodeTypes.PRO):
						//-----------------------
							{
							//--- Status --- Populate the styles for column A to B ---

							objServiceProduct.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID);
							if(objServiceProduct.ID == 0) // the entry could not be found
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
								strErrorText = "Error: The Service Product ID " + itemHierarchy.NodeID +
									" doesn't exist in SharePoint and couldn't be retrieved.";
								this.LogError(strErrorText);
								strErrorText = "Error: Service Product " + itemHierarchy.NodeID + " is missing.";
								strProduct = strErrorText;
								}
							else
								{
								strProduct = objServiceProduct.Title;
								}
							Console.WriteLine("\t\t\t + Product: {0} - {1}", objServiceProduct.ID, strProduct);
							break;
							}
						//-----------------------
						case (enumNodeTypes.ELE):
						//-----------------------
							{
							objServiceElement.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID);
							if(objServiceElement.ID == 0) // the entry could not be found
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
								strErrorText = "Error: The Service Element ID " + itemHierarchy.NodeID +
									" doesn't exist in SharePoint and couldn't be retrieved.";
								this.LogError(strErrorText);
								strErrorText = "Error: Service Element " + itemHierarchy.NodeID + " is missing.";
								strElement = strErrorText;
								}
							else
								{
								strElement = objServiceElement.Title;
								}
							Console.WriteLine("\t\t\t\t + Element: {0} - {1}", objServiceElement.ID, strElement);
							break;
							}

						//-----------------------
						case (enumNodeTypes.ELD):
						case (enumNodeTypes.ELR):
						case (enumNodeTypes.ELM):
						//-----------------------
							{
							// obtain the Deliverable object
							objDeliverable.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID);
							if(objDeliverable.ID == 0) // the entry could not be found
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
								strErrorText = "Error: The Deliverable ID " + itemHierarchy.NodeID +
									" doesn't exist in SharePoint and couldn't be retrieved.";
								this.LogError(strErrorText);
								strErrorText = "Error: Deliverable " + itemHierarchy.NodeID + " is missing.";
								strDeliverable = strErrorText;
								}
							else
								{
								strDeliverable = objDeliverable.Title;
								}
							
							// --- Add an entry to the dictCatalogue
							intCatalogueIndex += 1;
							Console.WriteLine("\t\t\t\t\t + Key: {2} \t Deliverable: {0} - {1}", objDeliverable.ID, strDeliverable, intCatalogueIndex);
							strCatalogueText = strDeliverable + " \u25C4 " + strElement + " \u25C4 " + strProduct 
								+ " \u25C4 " + strFamily + " \u25C4 " + strPortfolio;
							dictStructure.Add(intCatalogueIndex, strCatalogueText);

							// --- Process the Accountable Job Roles associated with the Deliverable
							if(objDeliverable.RACIaccountables != null
							&& objDeliverable.RACIaccountables.Count > 0)
								{
								foreach(var entry in objDeliverable.RACIaccountables)
									{
									if(!dictOfJobRoles.TryGetValue(key: entry.Key, value: out objJobRole))
										dictOfJobRoles.Add(entry.Key, entry.Value);
									// regardless whether the entry already exist in dictJobRoles add the dictCalalogue reference
									// which is intCatalogueIndex to the relevant Matrix Dictionary
									dictAccountableMarix.Add(intCatalogueIndex, entry.Key);
									}
								}

							// --- Process the Responsible Job Roles associated with the Deliverable
							if(objDeliverable.RACIresponsibles != null
							&& objDeliverable.RACIresponsibles.Count > 0)
								{
								foreach(var entry in objDeliverable.RACIresponsibles)
									{
									if(!dictOfJobRoles.TryGetValue(key: entry.Key, value: out objJobRole))
										dictOfJobRoles.Add(entry.Key, entry.Value);
									// regardless whether the entry already exist in dictJobRoles add the dictCalalogue reference
									// which is intCatalogueIndex to the relevant Matrix Dictionary
									dictResponsibleMarix.Add(intCatalogueIndex, entry.Key);
									}
								}

							// --- Process the Consulted Job Roles associated with the Deliverable
							if(objDeliverable.RACIconsulteds != null
							&& objDeliverable.RACIconsulteds.Count > 0)
								{
								foreach(var entry in objDeliverable.RACIconsulteds)
									{
									if(!dictOfJobRoles.TryGetValue(key: entry.Key, value: out objJobRole))
										dictOfJobRoles.Add(entry.Key, entry.Value);
									// regardless whether the entry already exist in dictJobRoles add the dictCalalogue reference
									// which is intCatalogueIndex to the relevant Matrix Dictionary
									dictConsultedMarix.Add(intCatalogueIndex, entry.Key);
									}
								}

							// --- Process the Informed Job Roles associated with the Deliverable
							if(objDeliverable.RACIinformeds != null
							&& objDeliverable.RACIinformeds.Count > 0)
								{
								foreach(var entry in objDeliverable.RACIinformeds)
									{
									if(!dictOfJobRoles.TryGetValue(key: entry.Key, value: out objJobRole))
										dictOfJobRoles.Add(entry.Key, entry.Value);
									// regardless whether the entry already exist in dictJobRoles add the dictCalalogue reference
									// which is intCatalogueIndex to the relevant Matrix Dictionary
									dictInformedMarix.Add(intCatalogueIndex, entry.Key);
									}
								}
								break;
							}
						} // end of Switch(itemHierarchy.NodeType)
					} // end of foreach(Hierarchy itemHierarchy in this.SelectedNodes)

				// Now we can populate the Worksheet for all the JobRoles
				Console.WriteLine("\r\n Polulating the  Worksheet...");

				// Process the JobRoles in the dictJobRoles according to the objRole.Delivery Domain and then objJobRole.Title
				int intColumnsStartNumber = 5; // Column F  - because columns use a 0 based reference
				int intColumnNumber = intColumnsStartNumber;
				string strCatalogueStructureText = "";
				//string strColumnLetter;
				string strBreak_ofDeliveryDomain = string.Empty;
				string strBreak_ofJobRole = string.Empty;
				UInt16 intRowIndex = 2;
				bool boolRACIcolumnPopulated = false;
				foreach(var entryJobRole in dictOfJobRoles.OrderBy(jr => jr.Value.DeliveryDomain).ThenBy(jt => jt.Value.Title))
					{
					// Break processing for DeliveryDomain
					if(entryJobRole.Value.DeliveryDomain != strBreak_ofDeliveryDomain)
						{
						intRowIndex += 1;
						strBreak_ofDeliveryDomain = entryJobRole.Value.DeliveryDomain;
						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objWorksheetPart,
							parColumnLetter: "A",
							parRowNumber: intRowIndex,
							parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(aWorkbook.GetColumnNumber("A"))),
							parCellDatatype: CellValues.String,
							parCellcontents: strBreak_ofDeliveryDomain);
						Console.WriteLine("+ Delivery Domain: {0}", strBreak_ofDeliveryDomain);

						for(ushort columnNo = 1; columnNo < 4; columnNo++)
							{
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: aWorkbook.GetColumnLetter(columnNo),
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(columnNo)),
								parCellDatatype: CellValues.String);
							}
						}
	
					// Break processing of JobRole
					if(entryJobRole.Value.Title != strBreak_ofJobRole)
						{
						intRowIndex += 1;
						strBreak_ofJobRole = entryJobRole.Value.Title;
						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objWorksheetPart,
							parColumnLetter: "A",
							parRowNumber: intRowIndex,
							parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(0)),
							parCellDatatype: CellValues.String);
							
						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objWorksheetPart,
							parColumnLetter: "B",
							parRowNumber: intRowIndex,
							parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(aWorkbook.GetColumnNumber("B"))),
							parCellDatatype: CellValues.String,
							parCellcontents: strBreak_ofJobRole);
						Console.WriteLine("\t + Job Role: {0}", strBreak_ofJobRole);

						for(ushort columnNo = 2; columnNo < 4; columnNo++)
							{
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: aWorkbook.GetColumnLetter(columnNo),
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(columnNo)),
								parCellDatatype: CellValues.String);
							}
						}
						
					// Determine if there is any entry in the dictAccountableMatrix with a Value == Key of the JobRole entry being processed
					boolRACIcolumnPopulated = false;
					foreach(var matrixItem in dictAccountableMarix.Where(am => am.Value == entryJobRole.Key))
						{
						intRowIndex += 1;
						//Populate the Columns A and B on the row						
						for(ushort columnNo = 0; columnNo < 2; columnNo++)
							{
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: aWorkbook.GetColumnLetter(columnNo),
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(columnNo)),
								parCellDatatype: CellValues.String);
							}
						if(boolRACIcolumnPopulated)
							{
							// Populate Column C
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "C",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(aWorkbook.GetColumnNumber("C"))),
								parCellDatatype: CellValues.String);
							}
						else
							{
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "C",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(aWorkbook.GetColumnNumber("C"))),
								parCellDatatype: CellValues.String,
								parCellcontents: "Accountable:");
							Console.WriteLine("\t\t\t + Accountable");
							boolRACIcolumnPopulated = true;
							}

						//Populate Column D
						//intRowIndex += 1;
						strCatalogueStructureText = null;
						// Obtain the Catalogue Structure VALUE (desription) from dictStructure with a Key == Value in the dict...Matrix entry
						if(!dictStructure.TryGetValue(key: matrixItem.Key, value: out strCatalogueStructureText))
							strCatalogueStructureText = "DocGenerator application Error occured...";
							
						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objWorksheetPart,
							parColumnLetter: "D",
							parRowNumber: intRowIndex,
							parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(aWorkbook.GetColumnNumber("D"))),
							parCellDatatype: CellValues.String,
							parCellcontents: strCatalogueStructureText);
						Console.WriteLine("\t\t\t\t + Deliverable: {0}", strCatalogueStructureText);

						} // end if: foreach(var item in dictAccountableMarix.Where(am => am.Value == entryJobRole.Key))

					// Determine if there is any entry in the dictResponsibleMatrix with a Key == Key of the JobRole entry being processed
					boolRACIcolumnPopulated = false;
					foreach(var matrixItem in dictResponsibleMarix.Where(m => m.Value == entryJobRole.Key).Distinct())
						{
						intRowIndex += 1;
						//Populate the Columns A and B on the row
						for(ushort columnNo = 0; columnNo < 2; columnNo++)
							{
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: aWorkbook.GetColumnLetter(columnNo),
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(columnNo)),
								parCellDatatype: CellValues.String);
							}
						if(boolRACIcolumnPopulated)
							{
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "C",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(aWorkbook.GetColumnNumber("C"))),
								parCellDatatype: CellValues.String);
							}
						else
							{
							// Populate Column C
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "C",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(aWorkbook.GetColumnNumber("C"))),
								parCellDatatype: CellValues.String,
								parCellcontents: "Responsible:");
							Console.WriteLine("\t\t\t + Responsible");
							boolRACIcolumnPopulated = true;
							}
							//Populate Columns D
							//intRowIndex += 1;
							strCatalogueStructureText = null;
							// Obtain the Catalogue Structure VALUE (desription) from dictStructure with a Key == Value in the dict...Matrix entry
							if(!dictStructure.TryGetValue(key: matrixItem.Key, value: out strCatalogueStructureText))
								strCatalogueStructureText = "DocGenerator application Error occured...";

							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "D",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(aWorkbook.GetColumnNumber("D"))),
								parCellDatatype: CellValues.String,
								parCellcontents: strCatalogueStructureText);
							Console.WriteLine("\t\t\t\t + Deliverable: {0}", strCatalogueStructureText);
						} // end loop: foreach(var matrixItem in dictResponsibleMarix.Where(m => m.Key == entryJobRole.Key))

					// Determine if there is any entry in the dictConsultedMatrix with a Key == Key of the JobRole entry being processed
					//if(dictConsultedMarix.TryGetValue(key: entryJobRole.Key, value: out intCatalogueDictionaryID))
					boolRACIcolumnPopulated = false;
					foreach(var matrixItem in dictConsultedMarix.Where(m => m.Value == entryJobRole.Key))
						{
						intRowIndex += 1;
						//Populate the Columns A and B on the row
						for(ushort columnNo = 0; columnNo < 2; columnNo++)
							{
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: aWorkbook.GetColumnLetter(columnNo),
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(columnNo)),
								parCellDatatype: CellValues.String);
							}
						// Populate Column C
						if(boolRACIcolumnPopulated)
							{
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "C",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(aWorkbook.GetColumnNumber("C"))),
								parCellDatatype: CellValues.String);
							}
						else
							{
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "C",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(aWorkbook.GetColumnNumber("C"))),
								parCellDatatype: CellValues.String,
								parCellcontents: "Consulted:");
							Console.WriteLine("\t\t\t + Consulted");
							boolRACIcolumnPopulated = true;
							}
						// Process all the Consulted Structures associated with the current JobRole.
						//Populate Columns D
						//intRowIndex += 1;
						strCatalogueStructureText = null;
						// Obtain the Catalogue Structure VALUE (desription) from dictStructure with a Key == Value in the dict...Matrix entry
						if(!dictStructure.TryGetValue(key: matrixItem.Key, value: out strCatalogueStructureText))
							strCatalogueStructureText = "DocGenerator application Error occured...";

						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objWorksheetPart,
							parColumnLetter: "D",
							parRowNumber: intRowIndex,
							parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(aWorkbook.GetColumnNumber("D"))),
							parCellDatatype: CellValues.String,
							parCellcontents: strCatalogueStructureText);
						Console.WriteLine("\t\t\t\t + Deliverable: {0}", strCatalogueStructureText);
						} // end loop: foreach(var matrixItem in dictConsultedMarix.Where(m => m.Key == entryJobRole.Key))



					// Determine if there is any entry in the dictInformedMatrix with a Key == Key of the JobRole entry being processed
					// if(dictInformedMarix.TryGetValue(key: entryJobRole.Key, value: out intCatalogueDictionaryID))
					boolRACIcolumnPopulated = false;
					foreach(var matrixItem in dictInformedMarix.Where(m => m.Value == entryJobRole.Key))
						{
						intRowIndex += 1;
						//Populate the Columns A and B on the row
						for(ushort columnNo = 0; columnNo < 2; columnNo++)
							{
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: aWorkbook.GetColumnLetter(columnNo),
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(columnNo)),
								parCellDatatype: CellValues.String);
							}
						// Populate Column C
						if(boolRACIcolumnPopulated)
							{
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "C",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(aWorkbook.GetColumnNumber("C"))),
								parCellDatatype: CellValues.String);
							}
						else
							{
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "C",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(aWorkbook.GetColumnNumber("C"))),
								parCellDatatype: CellValues.String,
								parCellcontents: "Informed:");
							Console.WriteLine("\t\t\t + Informed");
							boolRACIcolumnPopulated = true;
							}
						// Process all the Informed Structures associated with the current JobRole.
						//Populate Columns D
						//intRowIndex += 1;
						strCatalogueStructureText = null;
						// Obtain the Catalogue Structure VALUE (description) from dictStructure with a Key == Value in the dict...Matrix entry
						if(!dictStructure.TryGetValue(key: matrixItem.Key, value: out strCatalogueStructureText))
							strCatalogueStructureText = "DocGenerator application Error occured...";

						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objWorksheetPart,
							parColumnLetter: "D",
							parRowNumber: intRowIndex,
							parStyleId: (UInt32Value)(listColumnStylesA3D3.ElementAt(aWorkbook.GetColumnNumber("D"))),
							parCellDatatype: CellValues.String,
							parCellcontents: strCatalogueStructureText);
						Console.WriteLine("\t\t\t\t + Deliverable: {0}", strCatalogueStructureText);
						} // end loop: foreach(var matrixItem in dictInformedMarix.Where(m => m.Key == entryJobRole.Key))
					} // foreach(var entryJobRole in dictOfJobRoles.OrderBy(jr => jr.Value.DeliveryDomain).ThenBy(jt => jt.Value.Title))

				Console.WriteLine("Done");

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
