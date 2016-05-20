using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace DocGeneratorCore
	{

	/// <summary>
	/// This class handles the RACI Workbook per Role
	/// </summary>
	class RACI_Workbook_per_Role:aWorkbook
		{
		public void Generate(CompleteDataSet parDataSet,
			int? parRequestingUserID)
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			this.UnhandledError = false;
			DateTime timeStarted = DateTime.Now;
			//string hyperlinkImageRelationshipID = "";
			string strDocumentCollection_HyperlinkURL = "";
			//int intHyperlinkCounter = 9;
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



			try
				{
				if(this.HyperlinkEdit)
					{
					strDocumentCollection_HyperlinkURL = Properties.AppResources.SharePointURL +
						Properties.AppResources.List_DocumentCollectionLibraryURI +
						Properties.AppResources.EditFormURI + this.DocumentCollectionID;
					strCurrentHyperlinkViewEditURI = Properties.AppResources.EditFormURI;
					}

				if(this.HyperlinkView)
					{
					strDocumentCollection_HyperlinkURL = Properties.AppResources.SharePointURL +
						Properties.AppResources.List_DocumentCollectionLibraryURI +
						Properties.AppResources.DisplayFormURI + this.DocumentCollectionID;
					strCurrentHyperlinkViewEditURI = Properties.AppResources.DisplayFormURI;
					}

				//- Validate if the user selected any content to be generated
				if(this.SelectedNodes == null || this.SelectedNodes.Count < 1)
					{//- if nothing selected thow exception and exit
					throw new NoContentSpecifiedException("No content was specified/selected, therefore the document will be blank. "
						+ "Please specify/select content before submitting the document collection for generation.");
					}

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
					//- if the file creation failed.
					throw new DocumentUploadException(message: "DocGenerator was unable to create the document based on the template.");
					}

				this.LocalDocumentURI = objOXMLworkbook.LocalURI;
				this.FileName = objOXMLworkbook.Filename;

				// Open the MS Excel Workbook
				this.DocumentStatus = enumDocumentStatusses.Creating;
				// Open the MS Excel document in Edit mode
				SpreadsheetDocument objSpreadsheetDocument = SpreadsheetDocument.Open(path: objOXMLworkbook.LocalURI, isEditable: true);
				// Obtain the WorkBookPart from the spreadsheet.
				if(objSpreadsheetDocument.WorkbookPart == null)
					{
					this.DocumentStatus = enumDocumentStatusses.Failed;
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
					this.DocumentStatus = enumDocumentStatusses.Failed;
					throw new ArgumentException("The " + Properties.AppResources.Workbook_RACI_perRole_WorksheetName +
						" worksheet could not be located in the workbook.");
					}
				// obtain the WorksheetPart of the objMatrixWorksheet
				WorksheetPart objWorksheetPart = (WorksheetPart)(objWorkbookPart.GetPartById(objWorksheet.Id));

				this.DocumentStatus = enumDocumentStatusses.Building;

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
							//objServicePortfolio.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID);
							objServicePortfolio = parDataSet.dsPortfolios.Where(p => p.Key == itemHierarchy.NodeID).FirstOrDefault().Value;
							if(objServicePortfolio == null) // the entry could not be found
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
							Console.WriteLine("\t + Portfolio: {0} - {1}", itemHierarchy.NodeID, strPortfolio);
							break;
							}
					case (enumNodeTypes.FAM):
							{
							//objServiceFamily.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID);
							objServiceFamily = parDataSet.dsFamilies.Where(f => f.Key == itemHierarchy.NodeID).FirstOrDefault().Value;
							if(objServiceFamily == null || objServiceFamily.ID == 0) // the entry could not be found
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

							Console.WriteLine("\t\t + Family: {0} - {1}", itemHierarchy.NodeID, strFamily);
							break;
							}
					//-----------------------
					case (enumNodeTypes.PRO):
						//-----------------------
							{
							//--- Status --- Populate the styles for column A to B ---

							//objServiceProduct.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID);
							objServiceProduct = parDataSet.dsProducts.Where(p => p.Key == itemHierarchy.NodeID).FirstOrDefault().Value;
							if(objServiceProduct == null) // the entry could not be found
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
							Console.WriteLine("\t\t\t + Product: {0} - {1}", itemHierarchy.NodeID, strProduct);
							break;
							}
					//-----------------------
					case (enumNodeTypes.ELE):
						//-----------------------
							{
							//objServiceElement.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID);
							objServiceElement = parDataSet.dsElements.Where(e => e.Key == itemHierarchy.NodeID).FirstOrDefault().Value;
							if(objServiceElement == null || objServiceElement.ID == 0) // the entry could not be found
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
							Console.WriteLine("\t\t\t\t + Element: {0} - {1}", itemHierarchy.NodeID, strElement);
							break;
							}

					//-----------------------
					case (enumNodeTypes.ELD):
					case (enumNodeTypes.ELR):
					case (enumNodeTypes.ELM):
						//-----------------------
							{
							// obtain the Deliverable object
							// objDeliverable.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID, parGetRACI: true);
							objDeliverable = parDataSet.dsDeliverables.Where(d => d.Key == itemHierarchy.NodeID).FirstOrDefault().Value;
							if(objDeliverable == null) // the entry could not be found
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
							Console.WriteLine("\t\t\t\t\t + Key: {2} \t Deliverable: {0} - {1}", itemHierarchy.NodeID, strDeliverable, intCatalogueIndex);
							strCatalogueText = strDeliverable + " \u25C4 " + strElement + " \u25C4 " + strProduct
								+ " \u25C4 " + strFamily + " \u25C4 " + strPortfolio;
							dictStructure.Add(intCatalogueIndex, strCatalogueText);

							// --- Process the Accountable Job Roles associated with the Deliverable
							if(objDeliverable.RACIaccountables != null
							&& objDeliverable.RACIaccountables.Count > 0)
								{
								foreach(var entry in objDeliverable.RACIaccountables)
									{
									if(!dictOfJobRoles.TryGetValue(key: Convert.ToInt16(entry), value: out objJobRole))
										dictOfJobRoles.Add(Convert.ToInt16(entry),
											parDataSet.dsJobroles.Where(j => j.Key == entry).FirstOrDefault().Value);
									// regardless whether the entry already exist in dictJobRoles add the dictCalalogue reference
									// which is intCatalogueIndex to the relevant Matrix Dictionary
									dictAccountableMarix.Add(intCatalogueIndex, Convert.ToInt16(entry));
									}
								}

							// --- Process the Responsible Job Roles associated with the Deliverable
							if(objDeliverable.RACIresponsibles != null
							&& objDeliverable.RACIresponsibles.Count > 0)
								{
								foreach(var entry in objDeliverable.RACIresponsibles)
									{
									if(!dictOfJobRoles.TryGetValue(key: Convert.ToInt16(entry), value: out objJobRole))
										dictOfJobRoles.Add(Convert.ToInt16(entry),
											parDataSet.dsJobroles.Where(j => j.Key == entry).FirstOrDefault().Value);
									// regardless whether the entry already exist in dictJobRoles add the dictCalalogue reference
									// which is intCatalogueIndex to the relevant Matrix Dictionary
									dictResponsibleMarix.Add(intCatalogueIndex, Convert.ToInt16(entry));
									}
								}

							// --- Process the Consulted Job Roles associated with the Deliverable
							if(objDeliverable.RACIconsulteds != null
							&& objDeliverable.RACIconsulteds.Count > 0)
								{
								foreach(var entry in objDeliverable.RACIconsulteds)
									{
									if(!dictOfJobRoles.TryGetValue(key: Convert.ToInt16(entry), value: out objJobRole))
										dictOfJobRoles.Add(Convert.ToInt16(entry), parDataSet.dsJobroles.Where(j => j.Key == entry).FirstOrDefault().Value);
									// regardless whether the entry already exist in dictJobRoles add the dictCalalogue reference
									// which is intCatalogueIndex to the relevant Matrix Dictionary
									dictConsultedMarix.Add(intCatalogueIndex, Convert.ToInt16(entry));
									}
								}

							// --- Process the Informed Job Roles associated with the Deliverable
							if(objDeliverable.RACIinformeds != null
							&& objDeliverable.RACIinformeds.Count > 0)
								{
								foreach(var entry in objDeliverable.RACIinformeds)
									{
									if(!dictOfJobRoles.TryGetValue(key: Convert.ToInt16(entry), value: out objJobRole))
										dictOfJobRoles.Add(Convert.ToInt16(entry),
											parDataSet.dsJobroles.Where(j => j.Key == entry).FirstOrDefault().Value);
									// regardless whether the entry already exist in dictJobRoles add the dictCalalogue reference
									// which is intCatalogueIndex to the relevant Matrix Dictionary
									dictInformedMarix.Add(intCatalogueIndex, Convert.ToInt16(entry));
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

				Console.WriteLine("Document generation completed, saving and closing the document.");
				// Save and close the Document
				objSpreadsheetDocument.Close();

				this.DocumentStatus = enumDocumentStatusses.Completed;

				Console.WriteLine(
					"Generation started...: {0} \nGeneration completed: {1} \n Durarion..........: {2}",
					timeStarted, DateTime.Now, (DateTime.Now - timeStarted));

				//+ Upload the document to SharePoint
				this.DocumentStatus = enumDocumentStatusses.Uploading;
				Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");
				//- Upload the document to the Generated Documents Library and check if the upload succeeded....
				if(this.UploadDoc(parRequestingUserID: parRequestingUserID))
					{ //- Upload Succeeded
					Console.WriteLine("+ {0}, was Successfully Uploaded.", this.DocumentType);
					this.DocumentStatus = enumDocumentStatusses.Uploaded;
					}
				else
					{ //- Upload failed Failed
					Console.WriteLine("*** Uploading of {0} FAILED.", this.DocumentType);
					throw new DocumentUploadException("Error: DocGenerator was unable to upload the document to SharePoint");
					}

				//+ Done
				this.DocumentStatus = enumDocumentStatusses.Done;
				} // end Try

			//++ -------------------
			//++ Handle Exceptions
			//++ -------------------
			//+ NoContentspecified Exception
			catch(NoContentSpecifiedException exc)
				{
				this.ErrorMessages.Add(exc.Message);
				this.DocumentStatus = enumDocumentStatusses.Error;
				return; //- exit the method because there is no files to cleanup
				}

			//+ UnableToCreateDocument Exception
			catch(UnableToCreateDocumentException exc)
				{
				this.ErrorMessages.Add(exc.Message);
				this.DocumentStatus = enumDocumentStatusses.FatalError;
				return; //- exit the method because there is no files to cleanup
				}

			//+ DocumentUpload Exception
			catch(DocumentUploadException exc)
				{
				this.ErrorMessages.Add(exc.Message);
				this.DocumentStatus = enumDocumentStatusses.FatalError;
				}

			//+ OpenXMLPackage Exception
			catch(OpenXmlPackageException exc)
				{
				this.ErrorMessages.Add("Unfortunately, an unexpected error occurred during document generation and the document could not be produced. ["
					+ "[OpenXMLPackageException: " + exc.HResult + "Detail: " + exc.Message + "]");
				this.DocumentStatus = enumDocumentStatusses.FatalError;
				this.UnhandledError = true;
				}

			//+ ArgumentNull Exception
			catch(ArgumentNullException exc)
				{
				this.ErrorMessages.Add("Unfortunately, an unexpected error occurred during  ocument generation and the document could not be produced. ["
					+ "[ArgumentNullException: " + exc.HResult + "Detail: " + exc.Message + "]");
				this.DocumentStatus = enumDocumentStatusses.FatalError;
				this.UnhandledError = true;
				}

			//+ Exception (any not specified Exception)
			catch(Exception exc)
				{
				this.ErrorMessages.Add("An unexpected error occurred during the document generation and the document could not be produced. ["
					+ "[Exception: " + exc.HResult + "Detail: " + exc.Message + "]");
				this.DocumentStatus = enumDocumentStatusses.FatalError;
				this.UnhandledError = true;
				;
				}

			Console.WriteLine("\t\t End of the generation of {0}", this.DocumentType);
			//- Delete the file from the Documents Directory
			if(File.Exists(path: this.LocalDocumentURI))
				File.Delete(path: this.LocalDocumentURI);

			}
		}
	}
