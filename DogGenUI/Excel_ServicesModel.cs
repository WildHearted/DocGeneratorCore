using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xl2010 = DocumentFormat.OpenXml.Office2010.Excel;
using Excel = DocumentFormat.OpenXml.Office.Excel;
using DocumentFormat.OpenXml.Validation;

namespace DocGeneratorCore
	{
	/// <summary>
	/// This class handles the Internal Services Mapping Workbook
	/// </summary>
	class Internal_Services_Model_Workbook : aWorkbook
		{
		public void Generate(CompleteDataSet parDataSet, int? parRequestingUserID)
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			this.UnhandledError = false;
			DateTime timeStarted = DateTime.Now;
			//string hyperlinkImageRelationshipID = "";
			string strDocumentCollection_HyperlinkURL = "";
			//int intHyperlinkCounter = 9;
			string strCurrentHyperlinkViewEditURI = "";
			Cell objCell = new Cell();
			string strCheckDuplicate;
			//Text Workstrings
			string strText = "";
			string strErrorText = "";

			//Worksheet Row Index Variables (one Row less than the First row that needs to be populated
			UInt16 intRowIndex = 6;

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
			try
				{
				// Decalre all the objects to be used during processing
				ServicePortfolio objPortfolio = new ServicePortfolio();
				ServiceFamily objFamily = new ServiceFamily();
				ServiceProduct objProduct = new ServiceProduct();
				ServiceElement objElement = new ServiceElement();
				Deliverable objDeliverable = new Deliverable();
				Activity objActivity = new Activity();
				JobRole objJobRole = new JobRole();

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

				//- Create and Open the MS Excel Workbook 
				this.DocumentStatus = enumDocumentStatusses.Creating;
				// Open the MS Excel document in Edit mode
				SpreadsheetDocument objSpreadsheetDocument = SpreadsheetDocument.Open(path: objOXMLworkbook.LocalURI, isEditable: true);
				// Obtain the WorkBookPart from the spreadsheet.
				if(objSpreadsheetDocument.WorkbookPart == null)
					{
					this.DocumentStatus = enumDocumentStatusses.FatalError;
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

				// obtain the **ServiceModel** Worksheet in the Workbook.
				Sheet objWorksheet = objWorkbookPart.Workbook.Descendants<Sheet>().
					Where(sht => sht.Name == Properties.AppResources.Workbook_ServicesModel_WorksheetName).FirstOrDefault();
				if(objWorksheet == null)
					{
					this.DocumentStatus = enumDocumentStatusses.FatalError;
					throw new ArgumentException("The " + Properties.AppResources.Workbook_ServicesModel_WorksheetName +
						" worksheet could not be loacated in the workbook.");
					}
				// obtain the WorksheetPart of the objMatrixWorksheet
				WorksheetPart objWorksheetPart = (WorksheetPart)(objWorkbookPart.GetPartById(objWorksheet.Id));

				//+ Store the Cell Formats  for the Rows
				// --- Style for Column A7 to H7
				List<UInt32Value> listColumnStylesA7_H7 = new List<UInt32Value>();
				int intLastColumnNo = 8;
				int intFormatRowNo = 7;
				string strCellAddress = "";

				for(int i = 0; i < 8; i++)
					{
					strCellAddress = aWorkbook.GetColumnLetter(i) + intFormatRowNo;
					Cell objSourceCell = objWorksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == strCellAddress).FirstOrDefault();
					if(objSourceCell != null)
						{
						listColumnStylesA7_H7.Add(objSourceCell.StyleIndex);
						Console.WriteLine("\t\t\t\t + {0} - {1}", i, objSourceCell.StyleIndex);
						}
					else
						listColumnStylesA7_H7.Add(0U);
					} // loop

				// --- StyleId for Column D1:D3
				List<UInt32Value> listColumnStylesD1_D3 = new List<UInt32Value>();

				// If Hyperlinks need to be inserted, add the 
				Hyperlinks objHyperlinks = new Hyperlinks();

				this.DocumentStatus = enumDocumentStatusses.Building;

				foreach(Hierarchy itemHierarchy in this.SelectedNodes)
					{
					switch(itemHierarchy.NodeType)
						{
						//+ Service Portfolio + Services Framework
						case (enumNodeTypes.POR):
						case (enumNodeTypes.FRA):
							{
							intRowIndex += 1;
							objPortfolio = parDataSet.dsPortfolios.Where(p => p.Key == itemHierarchy.NodeID).FirstOrDefault().Value;
							if(objPortfolio == null) // the entry could not be found
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
								strErrorText = "Error: The Service Portfolio ID " + itemHierarchy.NodeID +
									" doesn't exist in SharePoint and couldn't be retrieved.";
								this.LogError(strErrorText);
								strErrorText = "Error: Service Portfolio " + itemHierarchy.NodeID + " was not found.";
								strText = strErrorText;
								}
							else
								{
								strText = objPortfolio.ISDheading;
								}
							Console.WriteLine("\t\t\t + Product: {0} - {1}", itemHierarchy.NodeID, strText);
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "A",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(aWorkbook.GetColumnNumber("A"))),
								parCellDatatype: CellValues.String,
								parCellcontents: strText);

							//-- Populate the styles for column B to H ---
							for(int i = 1; i <= intLastColumnNo; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}
							break;
							}
						//+ Service Family
						case (enumNodeTypes.FAM):
							{
							intRowIndex += 1;
							//-- Populate the styles for column A
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "A",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(0)),
								parCellDatatype: CellValues.String);

							objFamily = parDataSet.dsFamilies.Where(f => f.Key == itemHierarchy.NodeID).FirstOrDefault().Value;
							if(objFamily == null) // the entry could not be found
								{
								//- If the entry is not found - write an error in the document and record an error in the error log.
								strErrorText = "Error: The Service Family ID " + itemHierarchy.NodeID +
									" doesn't exist in SharePoint and couldn't be retrieved.";
								this.LogError(strErrorText);
								strErrorText = "Error: Service Family " + itemHierarchy.NodeID + " was not found.";
								strText = strErrorText;
								}
							else
								{
								strText = objFamily.ISDheading;
								}
							Console.WriteLine("\t\t\t + Product: {0} - {1}", itemHierarchy.NodeID, strText);
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "B",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(aWorkbook.GetColumnNumber("B"))),
								parCellDatatype: CellValues.String,
								parCellcontents: strText);

							//-- Populate the styles for column C to H
							for(int i = 2; i <= intLastColumnNo; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}
							break;
							}

						//+ Service Product
						case (enumNodeTypes.PRO):
							{
							intRowIndex += 1;
							//-- Populate the styles for column A to B
							for(int i = 0; i < 3; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}
							objProduct = parDataSet.dsProducts.Where(p => p.Key == itemHierarchy.NodeID).FirstOrDefault().Value;
							if(objProduct == null) // the entry could not be found
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
								strErrorText = "Error: The Service Product ID " + itemHierarchy.NodeID +
									" doesn't exist in SharePoint and couldn't be retrieved.";
								this.LogError(strErrorText);
								strErrorText = "Error: Service Product " + itemHierarchy.NodeID + " was not found.";
								strText = strErrorText;
								}
							else
								{
								strText = objProduct.ISDheading;
								}
							Console.WriteLine("\t\t\t + Product: {0} - {1}", itemHierarchy.NodeID, strText);
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "C",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(aWorkbook.GetColumnNumber("C"))),
								parCellDatatype: CellValues.String,
								parCellcontents: strText);

							//-- Populate the styles for column D to H
							for(int i = 3; i <= intLastColumnNo; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}
							break;
							}

						//+ Service Element
						case (enumNodeTypes.ELE):
							{
							//-- Populate the styles for column A to C
							intRowIndex += 1;

							for(int i = 0; i < 4; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}

							objElement = parDataSet.dsElements.Where(e => e.Key == itemHierarchy.NodeID).FirstOrDefault().Value;
							if(objElement == null) // the entry could not be found
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
								strErrorText = "Error: The Service Element ID " + itemHierarchy.NodeID +
									" doesn't exist in SharePoint and couldn't be retrieved.";
								this.LogError(strErrorText);
								strErrorText = "Error: Service Element " + itemHierarchy.NodeID + " was not found.";
								strText = strErrorText;
								}
							else
								{
								strText = objElement.ISDheading;
								}
							Console.WriteLine("\t\t\t\t + Element: {0} - {1}", itemHierarchy.NodeID, strText);
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "D",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(aWorkbook.GetColumnNumber("D"))),
								parCellDatatype: CellValues.String,
								parCellcontents: strText);

							//-- Populate the styles for column C to D ---
							for(int i = 4; i < intLastColumnNo; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}
							break;
							}

						//+ Deliverable
						case (enumNodeTypes.ELD):
						case (enumNodeTypes.ELR):
						case (enumNodeTypes.ELM):
							{
							//-- Populate the styles for column A to D
							intRowIndex += 1;
							for(int i = 0; i < 4; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}

							objDeliverable = parDataSet.dsDeliverables.Where(d => d.Key == itemHierarchy.NodeID).FirstOrDefault().Value;
							if(objDeliverable == null) // the entry could not be found
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
								strErrorText = "Error: The Deliverable ID " + itemHierarchy.NodeID +
									" doesn't exist in SharePoint and couldn't be retrieved.";
								this.LogError(strErrorText);
								strErrorText = "Error: Deliverable " + itemHierarchy.NodeID + " was not found.";
								strText = strErrorText;
								}
							else
								{
								strText = objDeliverable.ISDheading;
								}
							Console.WriteLine("\t\t\t\t\t + Deliverable: {0} - {1}", objDeliverable.ID, strText);
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "E",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(aWorkbook.GetColumnNumber("E"))),
								parCellDatatype: CellValues.String,
								parCellcontents: strText);

							//-- Populate the styles for column F to H ---
							for(int i = 5; i < intLastColumnNo; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}
							break;
							}
							//+ Activity
						case (enumNodeTypes.EAC):
							{
							//-- Populate the styles for column A to E
							intRowIndex += 1;
							for(int i = 0; i < 5; i++)
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: aWorkbook.GetColumnLetter(parColumnNo: i),
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(i)),
									parCellDatatype: CellValues.String);
								}

							objActivity = parDataSet.dsActivities.Where(a => a.Key == itemHierarchy.NodeID).FirstOrDefault().Value;
							if(objActivity == null) // the entry could not be found
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
								strErrorText = "Error: The Activity ID " + itemHierarchy.NodeID +
									" doesn't exist in SharePoint and couldn't be retrieved.";
								this.LogError(strErrorText);
								strErrorText = "Error: Activity " + itemHierarchy.NodeID + " was not found.";
								strText = strErrorText;
								}
							else
								{
								strText = objActivity.ISDheading;
								}
							Console.WriteLine("\t\t\t\t\t\t + Activity: {0} - {1}", objActivity.ID, objActivity.Title);
							oxmlWorkbook.PopulateCell(
								parWorksheetPart: objWorksheetPart,
								parColumnLetter: "F",
								parRowNumber: intRowIndex,
								parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(aWorkbook.GetColumnNumber("F"))),
								parCellDatatype: CellValues.String,
								parCellcontents: strText);

							//+ Populate the Accountable Role
							if(objActivity.RACI_AccountableID == null)
								{//- just add the StyleID
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: "G",
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(aWorkbook.GetColumnNumber("G"))),
									parCellDatatype: CellValues.String);
								}
							else
								{//- a value exist
								if(objActivity.RACI_AccountableID.Count > 0)
									{
									foreach(int? entryAccountableJobRoleID in objActivity.RACI_AccountableID)
										{
										//- Lookup the Role from the JobRoles
										objJobRole = parDataSet.dsJobroles.Where(j => j.Key == entryAccountableJobRoleID).FirstOrDefault().Value;
										if(objJobRole == null) // the entry could not be found
											{
											// If the entry is not found - write an error in the document and record an error in the error log.
											strErrorText = "Error: The Job Role ID " + entryAccountableJobRoleID +
												" doesn't exist in SharePoint and couldn't be retrieved.";
											this.LogError(strErrorText);
											strErrorText = "Error: Job role: " + entryAccountableJobRoleID + " was not found.";
											strText = strErrorText;
											}
										else
											{
											strText = objJobRole.Title;
											}

										oxmlWorkbook.PopulateCell(
											parWorksheetPart: objWorksheetPart,
											parColumnLetter: "G",
											parRowNumber: intRowIndex,
											parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(aWorkbook.GetColumnNumber("G"))),
											parCellDatatype: CellValues.String,
											parCellcontents: strText);
										//- exit the loop after the first entry
										break;
										}
									}
								else
									{ //- No entries existed
										{//- just add the StyleID
										oxmlWorkbook.PopulateCell(
											parWorksheetPart: objWorksheetPart,
											parColumnLetter: "G",
											parRowNumber: intRowIndex,
											parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(aWorkbook.GetColumnNumber("G"))),
											parCellDatatype: CellValues.String);
										}
									}
								}

							//+ Populate the Owning entity
							if(objActivity.OwningEntity == string.Empty)
								{ //- No entries existed - just write the Style
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: "H",
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(aWorkbook.GetColumnNumber("H"))),
									parCellDatatype: CellValues.String);
								}
							else
								{
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objWorksheetPart,
									parColumnLetter: "H",
									parRowNumber: intRowIndex,
									parStyleId: (UInt32Value)(listColumnStylesA7_H7.ElementAt(aWorkbook.GetColumnNumber("H"))),
									parCellDatatype: CellValues.String,
									parCellcontents: objActivity.OwningEntity);
								}
							break;

							}
						} // end of Switch(itemHierarchy.NodeType)
					} // end
				
				Console.WriteLine("\n\rWorksheet populated....");

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
				return; //- exit the method because there is no file to cleanup
				}

			//+ UnableToCreateDocument Exception
			catch(UnableToCreateDocumentException exc)
				{
				this.ErrorMessages.Add(exc.Message);
				this.DocumentStatus = enumDocumentStatusses.FatalError;
				return; //- exit the method because there is no file to cleanup
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
				}

			Console.WriteLine("\t\t End of the generation of {0}", this.DocumentType);
			//- Delete the file from the Documents Directory
			if(File.Exists(path: this.LocalDocumentURI))
				File.Delete(path: this.LocalDocumentURI);
			}
		}
	}
