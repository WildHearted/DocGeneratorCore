using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using DocGeneratorCore.Database.Classes;
using DocGeneratorCore.SDDPServiceReference;

namespace DocGeneratorCore
	{
	/// <summary>
	/// This class handles the RACI Matrix Workbook per Deliverable
	/// </summary>
	class RACI_Matrix_Workbook_per_Deliverable : aWorkbook
		{
		public void Generate(
			DesignAndDeliveryPortfolioDataContext parSDDPdatacontext,
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
			JobRole objJobRole;
	
			//Text Workstrings
			string strText = "";
			string strErrorText = "";
			List<int?> listOfJobRoles = new List<int?>();

			//Worksheet Row Index Variables
			UInt16 intRowIndex = 6;

			try
				{
				if(this.HyperlinkEdit)
					{
					strDocumentCollection_HyperlinkURL = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
						Properties.AppResources.List_DocumentCollectionLibraryURI +
						Properties.AppResources.EditFormURI + this.DocumentCollectionID;
					strCurrentHyperlinkViewEditURI = Properties.AppResources.EditFormURI;
					}

				if(this.HyperlinkView)
					{
					strDocumentCollection_HyperlinkURL = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
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
					parDocumentType: this.DocumentType,
					parSDDPdataContext: parSDDPdatacontext))
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
					this.DocumentStatus = enumDocumentStatusses.FatalError;
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
					this.DocumentStatus = enumDocumentStatusses.FatalError;
					throw new ArgumentException("The " + Properties.AppResources.Workbook_ContentStatus_WorksheetName +
						" worksheet could not be loacated in the workbook.");
					}
				// obtain the WorksheetPart of the objMatrixWorksheet
				WorksheetPart objWorksheetPart = (WorksheetPart)(objWorkbookPart.GetPartById(objWorksheet.Id));

				this.DocumentStatus = enumDocumentStatusses.Building;
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
				ServicePortfolio objPortfolio = new ServicePortfolio();
				ServiceFamily objFamily = new ServiceFamily();
				ServiceProduct objProduct = new ServiceProduct();
				ServiceElement objElement = new ServiceElement();
				Deliverable objDeliverable = new Deliverable();
				// Define the Dictionaries that will be represent the matrix
				// This dictionary will contain the the JobRole ID as the KEY and the VALUE will contain an JobRole Object
				Dictionary<int, JobRole> dictOfJobRoles = new Dictionary<int, JobRole>();
				// Each of the following dictionaries will contain the Matrix in which Key = Row Number and the VALUE = JobRoleID.
				Dictionary<int, List<int?>> dictAccountableMarix = new Dictionary<int, List<int?>>();
				Dictionary<int, List<int?>> dictResponsibleMarix = new Dictionary<int, List<int?>>();
				Dictionary<int, List<int?>> dictConsultedMarix = new Dictionary<int, List<int?>>();
				Dictionary<int, List<int?>> dictInformedMarix = new Dictionary<int, List<int?>>();

				foreach(Hierarchy node in this.SelectedNodes)
					{
					switch(node.NodeType)
						{
					//+Portfolio & Framework
					case (enumNodeTypes.POR):
					case (enumNodeTypes.FRA):
						intRowIndex += 1;
						objPortfolio = ServicePortfolio.Read(parIDsp: node.NodeID);
						if(objPortfolio == null) // the entry could not be found
							{
							//-| If the entry is not found - write an error in the document and record an error in the error log.
							strErrorText = "Error: The Service Portfolio ID " + node.NodeID + " doesn't exist in SharePoint and couldn't be retrieved.";
							this.LogError(strErrorText);
							strErrorText = "Error: Service Portfolio " + node.NodeID + " is missing.";
							strText = strErrorText;
							}
						else
							{
							strText = objPortfolio.ISDheading;
							}

						//--- Status --- Service Portfolio Row --- Column A -----
						// Write the Portfolio or Framework to the Workbook as a String
						Console.WriteLine("\t + Portfolio: {0} - {1}", objPortfolio.IDsp, objPortfolio.Title);
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
					
					//+Family
					case (enumNodeTypes.FAM):
						intRowIndex += 1;
						objFamily = ServiceFamily.Read(parIDsp: node.NodeID);
						if(objFamily == null) //-| the entry could not be found
							{
							//-| If the entry is not found - write an error in the document and record an error in the error log.
							strErrorText = "Error: The Service Family ID " + node.NodeID + " doesn't exist in SharePoint and couldn't be retrieved.";
							this.LogError(strErrorText);
							strErrorText = "Error: Service Family " + node.NodeID + " is missing.";
							strText = strErrorText;
							}
						else
							{
							strText = objFamily.ISDheading;
							}

						Console.WriteLine("\t\t + Family: {0} - {1}", node.NodeID, strText);
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

					//+Product
					case (enumNodeTypes.PRO):
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

						objProduct = ServiceProduct.Read(parIDsp: node.NodeID);
						if(objProduct == null) //-| the entry could not be found in Database
							{
							// If the entry is not found - write an error in the document and record an error in the error log.
							strErrorText = "Error: The Service Product ID " + node.NodeID +
								" doesn't exist in SharePoint and couldn't be retrieved.";
							this.LogError(strErrorText);
							strErrorText = "Error: Service Product " + node.NodeID + " is missing.";
							strText = strErrorText;
							}
						else
							{
							strText = objProduct.ISDheading;
							}
						Console.WriteLine("\t\t\t + Product: {0} - {1}", node.NodeID, strText);
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
					
					//+Element
					case (enumNodeTypes.ELE):
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

						objElement = ServiceElement.Read(parIDsp: node.NodeID);
						if(objElement == null) //-| the entry could not be found
							{
							//-| If the entry is not found - write an error in the document and record an error in the error log.
							strErrorText = "Error: The Service Element ID " + node.NodeID + " doesn't exist in SharePoint and couldn't be retrieved.";
							this.LogError(strErrorText);
							strErrorText = "Error: Service Element " + node.NodeID + " is missing.";
							strText = strErrorText;
							}
						else
							{
							strText = objElement.ISDheading;
							}
						Console.WriteLine("\t\t\t\t + Element: {0} - {1}", node.NodeID, strText);
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

					//+Deliverable, Report, Meeting
					case (enumNodeTypes.ELD):
					case (enumNodeTypes.ELR):
					case (enumNodeTypes.ELM):
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
							
						objDeliverable = Deliverable.Read(parIDsp: node.NodeID);
						if(objDeliverable == null) //-| the entry could not be found
							{
							//-| If the entry is not found - write an error in the document and record an error in the error log.
							strErrorText = "Error: The Deliverable ID " + node.NodeID + " doesn't exist in SharePoint and couldn't be retrieved.";
							this.LogError(strErrorText);
							strErrorText = "Error: Deliverable " + node.NodeID + " is missing.";
							strText = strErrorText;
							}
						else
							{
							strText = objDeliverable.ISDheading;
							}
						Console.WriteLine("\t\t\t\t\t + Deliverable: {0} - {1}", node.NodeID, strText);
						oxmlWorkbook.PopulateCell(
							parWorksheetPart: objWorksheetPart,
							parColumnLetter: "E",
							parRowNumber: intRowIndex,
							parStyleId: (UInt32Value)(listColumnStylesA7_G9.ElementAt(aWorkbook.GetColumnNumber("E"))),
							parCellDatatype: CellValues.String,
							parCellcontents: strText);

						//-| Process the Accountable Job Roles associated with the Deliverable
						if(objDeliverable.RACIaccountables != null)
							{
							listOfJobRoles.Clear();
							foreach(var entryJobRole in objDeliverable.RACIaccountables)
								{
								if (!dictOfJobRoles.TryGetValue(key: Convert.ToInt16(entryJobRole), value: out objJobRole))
									{
									JobRole accountableJobRole = new JobRole();
									accountableJobRole = JobRole.Read(parIDsp: Convert.ToInt16(entryJobRole));
									if (accountableJobRole != null)
										dictOfJobRoles.Add(key: Convert.ToInt16(entryJobRole), value: accountableJobRole);
									}

								//-| regardless of whether the entry already exist in dictJobRoles add a reference to the relevant Matrix Dictionary
								if(!dictAccountableMarix.TryGetValue(key: intRowIndex, value: out listOfJobRoles))
									{//- An entry for the row doesn't exist yet...
									listOfJobRoles = new List<int?>();
									listOfJobRoles.Add(entryJobRole);
									dictAccountableMarix.Add(intRowIndex, listOfJobRoles);
									}
								else
									{//- An entry for the row already exist...
									//-- add the new JobRole Entry to the retrieved listOfJobRoles
									listOfJobRoles.Add(entryJobRole);
									//-- Remove the existing entry from the dictionaty - in order to add it back with the new JobRole added to the Value...
									dictAccountableMarix.Remove(key: intRowIndex);
									//-- Insert/Add the emtry back to the dictionary...
									dictAccountableMarix.Add(key: intRowIndex, value: listOfJobRoles);
									}
								}
							}

						//-| Process the Responsible Job Roles associated with the Deliverable
						if(objDeliverable.RACIresponsibles != null)
							{
							listOfJobRoles.Clear();
							foreach(var entryJobRole in objDeliverable.RACIresponsibles)
								{
								if (!dictOfJobRoles.TryGetValue(key: Convert.ToInt16(entryJobRole), value: out objJobRole))
									{
									JobRole responsibleJobRole = new JobRole();
									responsibleJobRole = JobRole.Read(parIDsp: Convert.ToInt16(entryJobRole));
									if (responsibleJobRole != null)
										{
										dictOfJobRoles.Add(key: Convert.ToInt16(entryJobRole), value: responsibleJobRole);
										}
									}
								//-| Regardless whether the entry already exist in dictJobRoles add a reference to the relevant Matrix Dictionary
								if(!dictResponsibleMarix.TryGetValue(key: intRowIndex, value: out listOfJobRoles))
									{//- An entry for the row doesn't exist yet...
									listOfJobRoles = new List<int?>();
									listOfJobRoles.Add(entryJobRole);
									dictResponsibleMarix.Add(key: intRowIndex, value: listOfJobRoles);
									}
								else
									{//- An entry for the row already exist... add the new JobRole Entry to the retrieved listOfJobRoles
									listOfJobRoles.Add(entryJobRole);
									//-- Remove the existing entry from the dictionaty - in order to add it back with the new JobRole added to the Value...
									dictResponsibleMarix.Remove(key: intRowIndex);
									//-- Insert/Add the emtry back to the dictionary...
									dictResponsibleMarix.Add(key: intRowIndex, value: listOfJobRoles);
									}
								}
							}

						//-| Process the Consulted Job Roles associated with the Deliverable
						if(objDeliverable.RACIconsulteds != null)
							{
							listOfJobRoles.Clear();
							foreach(var entryJobRole in objDeliverable.RACIconsulteds)
								{
								if (!dictOfJobRoles.TryGetValue(key: Convert.ToInt16(entryJobRole), value: out objJobRole))
									{
									JobRole consultedJobRole = new JobRole();
									consultedJobRole = JobRole.Read(parIDsp: Convert.ToInt16(entryJobRole));
									if (consultedJobRole != null)
										dictOfJobRoles.Add(key: Convert.ToInt16(entryJobRole), value: consultedJobRole);
									}
								//-| regardless whether the entry already exist in dictJobRoles add a reference to the relevant Matrix Dictionary
								if(!dictConsultedMarix.TryGetValue(key: intRowIndex, value: out listOfJobRoles))
									{//- An entry for the row doesn't exist yet...
									listOfJobRoles = new List<int?>();
									listOfJobRoles.Add(entryJobRole);
									dictConsultedMarix.Add(intRowIndex, listOfJobRoles);
									}
								else
									{//- An entry for the roe already exist...
										//-- add the new JobRole Entry to the retrieved listOfJobRoles
									listOfJobRoles.Add(entryJobRole);
									//-- Remove the existing entry from the dictionaty - in order to add it back with the new JobRole added to the Value...
									dictConsultedMarix.Remove(key: intRowIndex);
									//-- Insert/Add the emtry back to the dictionary...
									dictConsultedMarix.Add(key: intRowIndex, value: listOfJobRoles);
									}
								}
							}

						//-| Process the Informed Job Roles associated with the Deliverable
						if(objDeliverable.RACIinformeds != null)
							{
							foreach(var entryJobRole in objDeliverable.RACIinformeds)
								{
								if (!dictOfJobRoles.TryGetValue(key: Convert.ToInt16(entryJobRole), value: out objJobRole))
									{
									JobRole informedJobRole = new JobRole();
									informedJobRole = JobRole.Read(parIDsp: Convert.ToInt16(entryJobRole));
									if (informedJobRole != null)
										dictOfJobRoles.Add(key: Convert.ToInt16(entryJobRole), value: informedJobRole);
									}
								//-| regardless whether the entry already exist in dictJobRoles add a reference to the relevant Matrix Dictionary
								if(!dictInformedMarix.TryGetValue(key: intRowIndex, value: out listOfJobRoles))
									{//- An entry for the row doesn't exist yet...
									listOfJobRoles = new List<int?>();
									listOfJobRoles.Add(entryJobRole);
									dictInformedMarix.Add(intRowIndex, listOfJobRoles);
									}
								else
									{//- An entry for the roe already exist...
										//-- add the new JobRole Entry to the retrieved listOfJobRoles
									listOfJobRoles.Add(entryJobRole);
									//-- Remove the existing entry from the dictionaty - in order to add it back with the new JobRole added to the Value...
									dictInformedMarix.Remove(key: intRowIndex);
									//-- Insert/Add the emtry back to the dictionary...
									dictInformedMarix.Add(key: intRowIndex, value: listOfJobRoles);
									}
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
						} // end of Switch(itemHierarchy.NodeType)
					} // end of foreach(Hierarchy itemHierarchy in this.SelectedNodes)

				// Now Populate the Columns from Column G until the point where they JobRoles end.
				Console.WriteLine("\r\n Polulating the Matrix in the Worksheet...");
				// First sort the JobRoles in the dictJobRoles dictionary according to the Values.
				int intColumnsStartNumber = 5; // Column F  - because columns use a 0 based reference
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
							// Row 2 is poulated with the JobRole title
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
							if(dictResponsibleMarix.TryGetValue(key: row, value: out listOfJobRoles))
								{//- if an enttry for the row was found...
								foreach(var matrixItem in listOfJobRoles.Where(m => m.Value == entryJobRole.Key))
									{//- procesess the entry
									if(matrixItem.Value == entryJobRole.Key)
										strMatricCellValue += "R";
									}
								}
							// Determine if there is a Row and Role Key match in dictAccountableMatrix 
							if(dictAccountableMarix.TryGetValue(key: row, value: out listOfJobRoles))
								{
								foreach(var matrixItem in listOfJobRoles.Where(m => m.Value == entryJobRole.Key))
									{//- procesess the entry
									if(matrixItem.Value == entryJobRole.Key)
										strMatricCellValue += "A";
									}
								}

							// Determine if there is a Row and Role Key match in dictConsultedMatrix 
							if(dictConsultedMarix.TryGetValue(key: row, value: out listOfJobRoles))
								{
								foreach(var matrixItem in listOfJobRoles.Where(m => m.Value == entryJobRole.Key))
									{//- procesess the entry
									if(matrixItem.Value == entryJobRole.Key)
										strMatricCellValue += "C";
									}
								}

							// Determine if there is a Row and Role Key match in dictInformedMatrix 
							if(dictInformedMarix.TryGetValue(key: row, value: out listOfJobRoles))
								{
								foreach(var matrixItem in listOfJobRoles.Where(m => m.Value == entryJobRole.Key))
									{//- procesess the entry
									if(matrixItem.Value == entryJobRole.Key)
										strMatricCellValue += "I";
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
							} 
						} 
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
				if(this.UploadDoc(parSDDPdatacontext: parSDDPdatacontext, parRequestingUserID: parRequestingUserID))
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
