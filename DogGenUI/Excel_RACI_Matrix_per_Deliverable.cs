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
	/// This class handles the RACI Matrix Workbook per Deliverable
	/// </summary>
	class RACI_Matrix_Workbook_per_Deliverable:aWorkbook
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
			JobRole objJobRole;
	
			//Text Workstrings
			string strText = "";
			string strErrorText = "";

			//Worksheet Row Index Variables
			UInt16 intRowIndex = 6;

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
					throw new ArgumentException(objOXMLworkbook.LocalURI + " does not contain a WorkbookPart. "
						+ "There is a problem with the template file.");
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
				Sheet objWorksheet = objWorkbookPart.Workbook.Descendants<Sheet>().Where(sht => sht.Name == Properties.AppResources.Workbook_RACI_Matrix_WorksheetName).FirstOrDefault();
				if(objWorksheet == null)
					{
					throw new ArgumentException("The " + Properties.AppResources.Workbook_ContentStatus_WorksheetName +
						" worksheet could not be loacated in the workbook.");
					}
				// obtain the WorksheetPart of the objMatrixWorksheet
				WorksheetPart objWorksheetPart = (WorksheetPart)(objWorkbookPart.GetPartById(objWorksheet.Id));

				// Copy the Cell Formats 
				// --- StyleId for Column G1:G6
				List<UInt32Value> listColumnStylesG1_G6 = new List<UInt32Value>();
				string strCellAddress = "";

				for(int r = 1; r < 7; r++)
					{
					strCellAddress = "G" + r;
					Cell objSourceCell = objWorksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == strCellAddress).FirstOrDefault();
					if(objSourceCell != null)
						{
						listColumnStylesG1_G6.Add(objSourceCell.StyleIndex);
						Console.WriteLine("\t\t\t\t + {0} - {1}", r, objSourceCell.StyleIndex);
						}
					else
						listColumnStylesG1_G6.Add(0U);
					} // loop

				// --- Style for Column A7 to G7
				List<UInt32Value> listColumnStylesA7_G9 = new List<UInt32Value>();
				int intLastColumn = 6;
				int intStyleSourceRow = 7;

				for(int i = 0; i <= intLastColumn; i++)
					{
					strCellAddress = aWorkbook.GetColumnLetter(i) + intStyleSourceRow;
					Cell objSourceCell = objWorksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == strCellAddress).FirstOrDefault();
					if(objSourceCell != null)
						{
						listColumnStylesA7_G9.Add(objSourceCell.StyleIndex);
						Console.WriteLine("\t\t\t\t + {0} - {1}", i, objSourceCell.StyleIndex);
						}
					else
						listColumnStylesA7_G9.Add(0U);
					} // loop

				UInt32Value uintMatrixColumnStyleID = listColumnStylesA7_G9.ElementAt(intLastColumn);

				// If Hyperlinks need to be inserted, add the 
				Hyperlinks objHyperlinks = new Hyperlinks();

				// Decalre all the object to be used during processing
				ServicePortfolio objServicePortfolio = new ServicePortfolio();
				ServiceFamily objServiceFamily = new ServiceFamily();
				ServiceProduct objServiceProduct = new ServiceProduct();
				ServiceElement objServiceElement = new ServiceElement();
				Deliverable objDeliverable = new Deliverable();
				// Define the Dictionaries that will be represent the matrix
				// This dictionary will contain the the JobRole ID as the KEY and the VALUE will contain an JobRole Object
				Dictionary<int, JobRole> dictOfJobRoles = new Dictionary<int, JobRole>();
				// Each of the following dictionaries will contain the Matrix in which Key = Row Number and the VALUE = JobRoleID.
				Dictionary<int, int> dictAccountableMarix = new Dictionary<int, int>();
				Dictionary<int, int> dictResponsibleMarix = new Dictionary<int, int>();
				Dictionary<int, int> dictConsultedMarix = new Dictionary<int, int>();
				Dictionary<int, int> dictInformedMarix = new Dictionary<int, int>();

				foreach(Hierarchy itemHierarchy in this.SelectedNodes)
					{
					switch(itemHierarchy.NodeType)
						{
					case (enumNodeTypes.POR):
					case (enumNodeTypes.FRA):
							{
							intRowIndex += 1;
							//objServicePortfolio.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID);
							objServicePortfolio = parDataSet.dsPortfolios.Where(p => p.Key == itemHierarchy.NodeID).FirstOrDefault().Value;
							if(objServicePortfolio == null) // the entry could not be found
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
								strErrorText = "Error: The Service Portfolio ID " + itemHierarchy.NodeID +
									" doesn't exist in SharePoint and couldn't be retrieved.";
								this.LogError(strErrorText);
								strErrorText = "Error: Service Portfolio " + itemHierarchy.NodeID + " is missing.";
								strText = strErrorText;
								}
							else
								{
								strText = objServicePortfolio.Title;
								}

							//--- Status --- Service Portfolio Row --- Column A -----
							// Write the Portfolio or Framework to the Workbook as a String
							Console.WriteLine("\t + Portfolio: {0} - {1}", objServicePortfolio.ID, objServicePortfolio.Title);
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "A",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA7_G9.ElementAt(aWorkbook.GetColumnNumber("A"))),
								parCellDatatype: CellValues.String,
								parCellcontents: strText);

							//--- Status --- Populate the styles for column B to G ---
							for(int i = 1; i <= intLastColumn; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_G9.ElementAt(i)),
									parCellDatatype: CellValues.String);
								//Console.WriteLine("\t\t\t\t + Column: {0} of {1}", i, intLastColumn);
								}
							break;
							}
					case (enumNodeTypes.FAM):
							{
							intRowIndex += 1;
							//objServiceFamily.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID);
							objServiceFamily = parDataSet.dsFamilies.Where(f => f.Key == itemHierarchy.NodeID).FirstOrDefault().Value;
							if(objServiceFamily == null) // the entry could not be found
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
								strErrorText = "Error: The Service Family ID " + itemHierarchy.NodeID +
									" doesn't exist in SharePoint and couldn't be retrieved.";
								this.LogError(strErrorText);
								strErrorText = "Error: Service Family " + itemHierarchy.NodeID + " is missing.";
								strText = strErrorText;
								}
							else
								{
								strText = objServiceFamily.Title;
								}

							Console.WriteLine("\t\t + Family: {0} - {1}", objServiceFamily.ID, objServiceFamily.Title);
							//--- Status --- Service Portfolio Row --- Column A -----
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "A",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA7_G9.ElementAt(aWorkbook.GetColumnNumber("A"))),
								parCellDatatype: CellValues.String);
							// Write the Family to the Workbook as a String
							//--- Status --- Service Family Row --- Column B -----

							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "B",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA7_G9.ElementAt(aWorkbook.GetColumnNumber("B"))),
								parCellDatatype: CellValues.String,
								parCellcontents: strText);

							//--- Status --- Populate the styles for column B to G ---
							for(int i = 2; i <= intLastColumn; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_G9.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}
							break;
							}
					//-----------------------
					case (enumNodeTypes.PRO):
						//-----------------------
							{
							//--- Status --- Populate the styles for column A to B ---
							intRowIndex += 1;
							for(int i = 0; i <= 1; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_G9.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}

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
							Console.WriteLine("\t\t\t + Product: {0} - {1}", objServiceProduct.ID, strText);
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "C",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA7_G9.ElementAt(aWorkbook.GetColumnNumber("C"))),
								parCellDatatype: CellValues.String,
								parCellcontents: strText);

							//--- Status --- Populate the styles for column F to G ---
							for(int i = 3; i <= intLastColumn; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_G9.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}
							break;
							}
					//-----------------------
					case (enumNodeTypes.ELE):
						//-----------------------
							{
							//--- Status --- Populate the styles for column A to C ---
							intRowIndex += 1;
							for(int i = 0; i <= 2; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_G9.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}

							//objServiceElement.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID);
							objServiceElement = parDataSet.dsElements.Where(e => e.Key == itemHierarchy.NodeID).FirstOrDefault().Value;
							if(objServiceElement.ID == 0) // the entry could not be found
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
								strErrorText = "Error: The Service Element ID " + itemHierarchy.NodeID +
									" doesn't exist in SharePoint and couldn't be retrieved.";
								this.LogError(strErrorText);
								strErrorText = "Error: Service Element " + itemHierarchy.NodeID + " is missing.";
								strText = strErrorText;
								}
							else
								{
								strText = objServiceElement.Title;
								}
							Console.WriteLine("\t\t\t\t + Element: {0} - {1}", objServiceElement.ID, strText);
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "D",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA7_G9.ElementAt(aWorkbook.GetColumnNumber("D"))),
								parCellDatatype: CellValues.String,
								parCellcontents: strText);

							//--- Status --- Populate the styles for column F to G ---
							for(int i = 4; i <= intLastColumn; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_G9.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}
							break;
							}

					//-----------------------
					case (enumNodeTypes.ELD):
					case (enumNodeTypes.ELR):
					case (enumNodeTypes.ELM):
						//-----------------------
							{
							//--- Status --- Populate the styles for column A to C ---
							intRowIndex += 1;
							for(int i = 0; i <= 3; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_G9.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}

							//objDeliverable.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: itemHierarchy.NodeID, parGetRACI: true);
							objDeliverable = parDataSet.dsDeliverables.Where(d => d.Key == itemHierarchy.NodeID).FirstOrDefault().Value;
							if(objDeliverable== null) // the entry could not be found
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
								strErrorText = "Error: The Deliverable ID " + itemHierarchy.NodeID +
									" doesn't exist in SharePoint and couldn't be retrieved.";
								this.LogError(strErrorText);
								strErrorText = "Error: Deliverable " + itemHierarchy.NodeID + " is missing.";
								strText = strErrorText;
								}
							else
								{
								strText = objDeliverable.Title;
								}
							Console.WriteLine("\t\t\t\t\t + Deliverable: {0} - {1}", objDeliverable.ID, strText);
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "E",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA7_G9.ElementAt(aWorkbook.GetColumnNumber("E"))),
								parCellDatatype: CellValues.String,
								parCellcontents: strText);

							// --- Process the Accountable Job Roles associated with the Deliverable
							if(objDeliverable.RACIaccountables != null
							&& objDeliverable.RACIaccountables.Count > 0)
								{
								foreach(var entryJobRole in objDeliverable.RACIaccountables)
									{
									if(!dictOfJobRoles.TryGetValue(key: Convert.ToInt16(entryJobRole), value: out objJobRole))
										dictOfJobRoles.Add(Convert.ToInt16(entryJobRole), 
											parDataSet.dsJobroles.Where(j => j.Key == entryJobRole).FirstOrDefault().Value);
									// regardless whether the entry already exist in dictJobRoles add a reference to the relevant Matrix Dictionary
									dictAccountableMarix.Add(intRowIndex, Convert.ToInt16(entryJobRole));
									}
								}

							// --- Process the Responsible Job Roles associated with the Deliverable
							if(objDeliverable.RACIresponsibles != null
							&& objDeliverable.RACIresponsibles.Count > 0)
								{
								foreach(var entryJobRole in objDeliverable.RACIresponsibles)
									{
									if(!dictOfJobRoles.TryGetValue(key: Convert.ToInt16(entryJobRole), value: out objJobRole))
										dictOfJobRoles.Add(Convert.ToInt16(entryJobRole),
											parDataSet.dsJobroles.Where(j => j.Key == entryJobRole).FirstOrDefault().Value);
									// regardless whether the entry already exist in dictJobRoles add a reference to the relevant Matrix Dictionary
									dictResponsibleMarix.Add(intRowIndex, Convert.ToInt16(entryJobRole));
									}
								}

							// --- Process the Consulted Job Roles associated with the Deliverable
							if(objDeliverable.RACIconsulteds != null
							&& objDeliverable.RACIconsulteds.Count > 0)
								{
								foreach(var entryJobRole in objDeliverable.RACIconsulteds)
									{
									if(!dictOfJobRoles.TryGetValue(key: Convert.ToInt16(entryJobRole), value: out objJobRole))
										dictOfJobRoles.Add(Convert.ToInt16(entryJobRole),
											parDataSet.dsJobroles.Where(j => j.Key == entryJobRole).FirstOrDefault().Value);
									// regardless whether the entry already exist in dictJobRoles add a reference to the relevant Matrix Dictionary
									dictConsultedMarix.Add(intRowIndex, Convert.ToInt16(entryJobRole));
									}
								}

							// --- Process the Informed Job Roles associated with the Deliverable
							if(objDeliverable.RACIinformeds != null
							&& objDeliverable.RACIinformeds.Count > 0)
								{
								foreach(var entryJobRole in objDeliverable.RACIinformeds)
									{
									if(!dictOfJobRoles.TryGetValue(key: Convert.ToInt16(entryJobRole), value: out objJobRole))
										dictOfJobRoles.Add(Convert.ToInt16(entryJobRole),
											parDataSet.dsJobroles.Where(j => j.Key == entryJobRole).FirstOrDefault().Value);
									// regardless whether the entry already exist in dictJobRoles add a reference to the relevant Matrix Dictionary
									dictInformedMarix.Add(intRowIndex, Convert.ToInt16(entryJobRole));
									}
								}

							//--- Status --- Populate the styles for column F to G ---
							for(int i = 5; i <= intLastColumn; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_G9.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}
							break;
							}
						} // end of Switch(itemHierarchy.NodeType)
					} // end of foreach(Hierarchy itemHierarchy in this.SelectedNodes)

				// Now Populate the Columns from Column G until the point where they JobRoles end.
				Console.WriteLine("\r\n Polulating the Matrix in the Worksheet...");
				// First sort the JobRoles in the dictJobRoles dictionary according to the Values.
				int intColumnsStartNumber = 5; // Column F  - because columns use a 0 based reference
				int intMatrixLookupJobID;
				int intColumnNumber = intColumnsStartNumber;
				string strMatricCellValue = "";
				string strColumnLetter;
				foreach(var entryJobRole in dictOfJobRoles.OrderBy(so => so.Value.Title))
					{
					intColumnNumber += 1;
					strColumnLetter = aWorkbook.GetColumnLetter(intColumnNumber);

					//Console.Write("\n Column {2}: {0} \t Id: {1}", entryJobRole.Value.Title, entryJobRole.Key, strColumnLetter);
					// Iterate through the rows for each column
					for(ushort row = 1; row < intRowIndex + 1; row++)
						{
						//Console.Write("\n\t + Row {0} - ", row);
						if(row < 7) // exception of the first 6 Rows which doesn't contain any data only a style.
							{
							// Row 2 need to be poulated with the JobRole title
							if (row ==1 && strColumnLetter == "G")
								{
								//Console.Write(" + Skip {0}{1}", strColumnLetter, row);
								}
							else if(row == 2)
								{
								oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: strColumnLetter,
								parRowNumber: row,
								parStyleId: listColumnStylesG1_G6.ElementAt(row - 1),
								parCellDatatype: CellValues.String,
								parCellcontents: entryJobRole.Value.Title);
								//Console.Write(" + styleID: [{0}] + Column Heading: {1}", listColumnStylesG1_G6.ElementAt(row - 1), entryJobRole.Value.Title);
								}
							else
								{
								oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: strColumnLetter,
								parRowNumber: row,
								parStyleId: listColumnStylesG1_G6.ElementAt(row - 1),
								parCellDatatype: CellValues.String);
								//Console.Write(" + styleID: [{0}]", listColumnStylesG1_G6.ElementAt(row - 1));
								}

							if(row == 6) //// Merge Rows 2-6 for the current column
								{
								oxmlWorkbook.MergeCell(
									parWorksheetPart: objWorksheetPart,
									parTopLeftCell: strColumnLetter + 2,
									parBottomRightCell: strColumnLetter + 6);
								//Console.Write(" + merged: [{0}]", strColumnLetter + 2 + ":" + strColumnLetter + 6);
								}
							} // end if(row < 7)
						else // if(row > 6)
							{
							strMatricCellValue = null;
							// Determine if there is a Row and Role Key match in dictResponsibleMatrix 
							if(dictResponsibleMarix.TryGetValue(key: row, value: out intMatrixLookupJobID))
								{
								foreach(var matrixItem in dictResponsibleMarix.Where(m => m.Key == row))
									{
									if(matrixItem.Value == entryJobRole.Key)
										strMatricCellValue += "R";
									}
								}
							// Determine if there is a Row and Role Key match in dictAccountableMatrix 
							if(dictAccountableMarix.TryGetValue(key: row, value: out intMatrixLookupJobID))
								{
								foreach(var matrixItem in dictAccountableMarix.Where(m => m.Key == row))
									{
									if(matrixItem.Value == entryJobRole.Key)
										strMatricCellValue += " A";
									}
								}

							// Determine if there is a Row and Role Key match in dictConsultedMatrix 
							if(dictConsultedMarix.TryGetValue(key: row, value: out intMatrixLookupJobID))
								{
								foreach(var matrixItem in dictConsultedMarix.Where(m => m.Key == row))
									{
									if(matrixItem.Value == entryJobRole.Key)
										strMatricCellValue += " C";
									}
								}

							// Determine if there is a Row and Role Key match in dictInformedMatrix 
							if(dictInformedMarix.TryGetValue(key: row, value: out intMatrixLookupJobID))
								{
								foreach(var matrixItem in dictInformedMarix.Where(m => m.Key == row))
									{
									if(matrixItem.Value == entryJobRole.Key)
										strMatricCellValue += " A";
									}
								}

							//check if the strMatricCellCalue was populate, and then populate the cell
							if(strMatricCellValue == null) // no matches were found in all the matrixes
								{
								// Insert the cell with the correct style and no value
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: strColumnLetter,
									parRowNumber: row,
									parStyleId: uintMatrixColumnStyleID,
									parCellDatatype: CellValues.String);
								//Console.Write(" + StyleID: [{0}]", uintMatrixColumnStyleID);
								}
							else // a value was found
								{
								// Insert the cell and populate the value and style
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: strColumnLetter,
									parRowNumber: row,
									parStyleId: uintMatrixColumnStyleID,
									parCellDatatype: CellValues.String,
									parCellcontents: strMatricCellValue);
								//Console.Write(" + StyleID: [{0}] + Values: [{1}]", uintMatrixColumnStyleID, strMatricCellValue);
								}
							} //else // if(row > 6)
						} // loop foreach(ushort row = 1; row < intRowIndex.....

					} // foreach(var entryJobRole in dictOfJobRoles.OrderBy(so => so.Value))

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
				Console.WriteLine("\n\nException: {0} - {1}", exc.HResult, exc.Message);
				return false;
				//TODO: raise the error
				}
			catch(Exception exc)
				{
				Console.WriteLine("\n\nException: {0} - {1}", exc.HResult, exc.Message);
				return false;
				}
			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}
		}
	}
