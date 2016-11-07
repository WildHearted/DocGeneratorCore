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

namespace DocGeneratorCore
	{
	//++Client_Requirements_Mapping_Workbook
	/// <summary>
	/// This class handles the Client_Requirements_Mapping_Workbook
	/// </summary>
	class Client_Requirements_Mapping_Workbook:aWorkbook
		{
		public int? CRM_Mapping {get; set;}
		//+Properties
		/// <summary>
		/// This Method generates the Client Requirements Mapping Workbook
		/// </summary>
		/// <param name="parDataSet"></param>
		/// <returns></returns>
		/// 

		//++ Methods

		//+Generate method
		public void Generate(
			ref CompleteDataSet parDataSet,
			int? parRequestingUserID)
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			this.UnhandledError = false;
			DateTime timeStarted = DateTime.Now;
			//string hyperlinkImageRelationshipID = "";
			string strDocumentCollection_HyperlinkURL = "";
			int intHyperlinkCounter = 9;
			string strCurrentHyperlinkViewEditURI = "";
			Cell objCell = new Cell();
			int intSharedStringIndex = 0;
			//- Workbook Break processing Variables
			int intRequirementBreakID_forRisks = 0;			//- the ID value of the Requirement used as a break processing variable for Risks sheet
			int intRequirementBreakID_forAssumptions = 0;     //- the ID value of the Requirement used as a break processing variable for Assumptions sheet
			string errorText = "";
			//-Content Layering Variables
			int? layer0upDeliverableID;
			int? layer1upDeliverableID;
			int? layer2upDeliverableID;
			string strTextDescription = "";

			//-Worksheet Row Index Variables
			UInt16 intMatrixSheet_RowIndex = 6;
			UInt16 intRisksSheet_RowIndex = 2;
			UInt16 intAssumptionsSheet_RowIndex = 2;
			Dictionary<string, string> dictionaryMatrixComments = new Dictionary<string, string>();
			string strErrorText = "";

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
				if(this.CRM_Mapping == null || this.CRM_Mapping == 0)
					{//- if nothing selected thow exception and exit
					throw new NoContentSpecifiedException("A Client Requirement Mapping was not specified/selected, therefore the document will be blank. "
						+ "Please specify/select a Client Requirement Mapping before submitting the document collection for generation.");
					}

				// define a new objOpenXMLworksheet
				oxmlWorkbook objOXMLworkbook = new oxmlWorkbook();
				// use CreateDocumentFromTemplate method to create a new MS Word Document based on the relevant template
				if(objOXMLworkbook.CreateDocWbkFromTemplate(
					parDocumentOrWorkbook: enumDocumentOrWorkbook.Workbook,
					parTemplateURL: this.Template,
					parDocumentType: this.DocumentType,
					parDataSet: ref parDataSet))
					{
					Console.WriteLine("\t\t\t objOXMLdocument:\n" 
						+ "\t\t\t+ LocalDocumentPath: " + objOXMLworkbook.LocalPath
						+ "\n\t\t\t+ DocumentFileName.: " + objOXMLworkbook.Filename
						+ "\n\t\t\t+ DocumentURI......: " + objOXMLworkbook.LocalURI);
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

				this.DocumentStatus = enumDocumentStatusses.Building;
				// If Hyperlinks need to be inserted, add the 
				Hyperlinks objHyperlinks = new Hyperlinks();
				//-------------------------------------
				// Begin to process the selected Mapping
				if(this.CRM_Mapping == null || this.CRM_Mapping == 0)
					{
					strErrorText = "A Client Requirements Mapping was not specified for the Document Collection.";
					Console.WriteLine("### {0} ###", strErrorText);
					// If an entry was not specified - write an error in the Worksheet and record an error in the error log.
					this.LogError(strErrorText);

					this.DocumentStatus = enumDocumentStatusses.FatalError;

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

				Deliverable objDeliverable = new Deliverable();
				Deliverable objDeliverableLayer1up = new Deliverable();
				Deliverable objDeliverableLayer2up = new Deliverable();
				DeliverableServiceLevel objDeliverableServiceLevel = new DeliverableServiceLevel();
				ServiceLevel objServiceLevel = new ServiceLevel();
				Mapping objMapping = new Mapping();
				MappingServiceTower objMappingServiceTower = new MappingServiceTower();
				MappingRequirement objMappingRequirement = new MappingRequirement();
				MappingAssumption objMappingAssumption = new MappingAssumption();
				MappingRisk objMappingRisk = new MappingRisk();

				Console.WriteLine("Retrieving the Mapping Data...");
				bool bRetrievedCRM = false;
				if(this.CRM_Mapping != null)
					{
					if(parDataSet.dsMappings != null
					&& parDataSet.dsMappings.TryGetValue(key: this.CRM_Mapping, value: out objMapping))
						{
						Console.Write("\n\t Mapping data already loaded in the Complete DataSet - no need to fetch it again");
						}
					else
						{
						// Load the Mappings data into the Complete Data Set.
						Console.Write("\n\t Mapping data NOT present in the Complete DataSet - Let's retrive it...");
						bRetrievedCRM = parDataSet.PopulateMappingDataset(parDatacontexSDDP: parDataSet.SDDPdatacontext, parMapping: this.CRM_Mapping);
						if(!bRetrievedCRM) // There was an error retriving the Mapping
							{
							errorText = "Error: Unable to retrieve the Client Requirements Mapping data for Mapping ID: " + this.CRM_Mapping
								+ ". Please check if the entry still exist in the Mappings List in SharePoint and that the DocGenerator can access SharePoint).";
							this.LogError(errorText);
							goto Save_and_Close_Document;
							}
						}
					}

				// Obtain the Mapping data 
				if(parDataSet.dsMappings.TryGetValue(key: this.CRM_Mapping, value: out objMapping))
					{
					Console.Write("\n\t + {0} - {1}", objMapping.IDsp, objMapping.Title);
					}
				else
					{
					// If the entry is not found - write an error in the document and record an error in the error log.
					errorText = "Error: The Mapping ID: " + this.CRM_Mapping
						+ " doesn't exist in SharePoint and couldn't be retrieved.";
					this.LogError(errorText);
					Console.Write("\n\t + {0} - {1}", objMapping.IDsp, errorText);

					}

				// Check if any Mapping Service Tower entries were retrieved
				if(parDataSet.dsMappingServiceTowers == null
				|| parDataSet.dsMappingServiceTowers.Count == 0)
					{
					strErrorText = "No Towers of Service was found for the mapping.";
					Console.WriteLine("### {0} ###", strErrorText);
					this.LogError(strErrorText);
					goto Save_and_Close_Document;
					}

				// Process each of the Mapping Service Towers
				// --- Loop through all Service Towers for the Mapping ---
				foreach(MappingServiceTower objTower in parDataSet.dsMappingServiceTowers.Values.OrderBy(t => t.Title))
					{
					// Write the Mapping Service Tower to the Workbook as a String
					Console.WriteLine("\n\t + Tower: {0} - {1}", objTower.IDsp, objTower.Title);
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
						parCellcontents: objTower.IDsp.ToString(),
						parHyperlinkCounter: intHyperlinkCounter,
						parHyperlinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
							Properties.AppResources.List_MappingServiceTowers +
							Properties.AppResources.EditFormURI + objTower.IDsp.ToString());
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
					// Process all the Mapping requirements for the specific Service Tower
					foreach(MappingRequirement objRequirement in parDataSet.dsMappingRequirements.Values
						.Where(r => r.MappingServiceTowerIDsp == objTower.IDsp))
						{
						Console.Write("\n\t\t + Requirement: {0} - {1}", objRequirement.IDsp, objRequirement.Title);
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

						dictionaryMatrixComments.Add("F|" + intMatrixSheet_RowIndex, objRequirement.RequirementText);

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
							parCellcontents: objRequirement.IDsp.ToString(),
							parHyperlinkCounter: intHyperlinkCounter,
							parHyperlinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
								Properties.AppResources.List_MappingRequirements +
								Properties.AppResources.EditFormURI + objRequirement.IDsp.ToString());
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
						// Process all the Mapping Risks for the specific Service Requirement
						foreach(MappingRisk objRisk in parDataSet.dsMappingRisks.Values
							.Where(r => r.MappingRequirementIDsp == objRequirement.IDsp))
							{
							Console.WriteLine("\t\t\t + Risk: {0} - {1}", objRisk.IDsp, objRisk.Title);
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

							dictionaryMatrixComments.Add("F|" + intMatrixSheet_RowIndex, objRisk.Statement);

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
								parCellcontents: objRisk.IDsp.ToString(),
								parHyperlinkCounter: intHyperlinkCounter,
								parHyperlinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
									Properties.AppResources.List_MappingRisks +
									Properties.AppResources.EditFormURI + objRisk.IDsp.ToString());
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
							if(intRequirementBreakID_forRisks != objRequirement.IDsp)
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

								intRequirementBreakID_forRisks = objRequirement.IDsp;
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
								parCellcontents: objRisk.IDsp.ToString());
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
								parCellcontents: objRisk.Mittigation);
							} //foreach(Mappingrisk objMappingRisk in listMappingRisks)

						//=====================================================================
						//  Process all Mapping Assumptions for the specified Mapping Requirement
						foreach(MappingAssumption objAssumption in parDataSet.dsMappingAssumptions.Values
							.Where(a => a.MappingRequirementIDsp == objRequirement.IDsp))
							{
							Console.WriteLine("\t\t\t + Assumption: {0} - {1}", objAssumption.IDsp, objAssumption.Title);
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

							dictionaryMatrixComments.Add("F|" + intMatrixSheet_RowIndex, objAssumption.Description);

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
								parCellcontents: objAssumption.IDsp.ToString(),
								parHyperlinkCounter: intHyperlinkCounter,
								parHyperlinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
									Properties.AppResources.List_MappingAssumptions +
									Properties.AppResources.EditFormURI + objAssumption.IDsp.ToString());
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
							if(intRequirementBreakID_forAssumptions != objRequirement.IDsp)
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

								intRequirementBreakID_forAssumptions = objRequirement.IDsp;
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
									parCellcontents: objAssumption.IDsp.ToString());
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

						//-----------------------------------------------------------------------
						// Obtain all Mapping Deliverables for the specified Mapping Requirement
						// Process all the Mapping Deliverables for the specific Service Requirement
						foreach(var objMappingDeliverable in parDataSet.dsMappingDeliverables.Values
							.Where(d => d.MappingRequirementIDsp == objMappingRequirement.IDsp).OrderBy(d => d.Title))
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
									dictionaryMatrixComments.Add("F|" + intMatrixSheet_RowIndex,
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
								layer0upDeliverableID = objMappingDeliverable.MappedDeliverableID;
								// Get the entry from the DataSet
								if(parDataSet.dsDeliverables.TryGetValue(
									key: Convert.ToInt16(objMappingDeliverable.MappedDeliverableID),
									value: out objDeliverable))
									{
									//Check if the Mapped_Deliverable Layer0up has Content Layers and Content Predecessors
									Console.WriteLine("\n\t\t + Deliverable Layer 0..: {0} - {1}", objDeliverable.IDsp, objDeliverable.Title);
									if(objDeliverable.ContentPredecessorDeliverableIDsp == null)
										{
										layer1upDeliverableID = null;
										layer2upDeliverableID = null;
										}
									else
										{
										layer1upDeliverableID = objDeliverable.ContentPredecessorDeliverableIDsp;
										// Get the entry from the DataSet
										if(parDataSet.dsDeliverables.TryGetValue(
											key: Convert.ToInt16(layer1upDeliverableID),
											value: out objDeliverableLayer1up))
											{
											if(objDeliverableLayer1up.ContentPredecessorDeliverableIDsp == null)
												{
												layer2upDeliverableID = null;
												}
											else
												{
												layer2upDeliverableID = objDeliverableLayer1up.ContentPredecessorDeliverableIDsp;
												// Get the entry from the DataSet
												if(parDataSet.dsDeliverables.TryGetValue(
													key: Convert.ToInt16(layer2upDeliverableID),
													value: out objDeliverableLayer2up))
													{
													layer2upDeliverableID = objDeliverableLayer2up.ContentPredecessorDeliverableIDsp;
													}
												else
													{
													layer2upDeliverableID = null;
													}
												}
											}
										else
											{
											layer2upDeliverableID = null;
											}
										}

									if(objDeliverable.CSDdescription != null)
										{
										strTextDescription = strTextDescription + HTMLdecoder.CleanText
												(objDeliverable.CSDdescription, parClientName: "the Client");
										}
									// Insert the Deliverable CSD Description
									if(strTextDescription != "")
										{
										dictionaryMatrixComments.Add("F|" + intMatrixSheet_RowIndex, strTextDescription);
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
									parHyperlinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
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
										parCellcontents: objDeliverable.IDsp.ToString(),
										parHyperlinkCounter: intHyperlinkCounter,
									parHyperlinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
										Properties.AppResources.List_DeliverablesURI +
										Properties.AppResources.EditFormURI + objDeliverable.IDsp.ToString()
										);
									}
								//--- Matrix --- Deliverable Row --- Column N --------------------------------
								oxmlWorkbook.PopulateCell(
									parWorksheetPart: objMatrixWorksheetPart,
									parColumnLetter: "N",
									parRowNumber: intMatrixSheet_RowIndex,
									parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("N"))),
									parCellDatatype: CellValues.String);
								}
							
								//====================================================================
								// Obtain all Service Levels for the specified Deliverable Requirement
								// Process the Mapping Service Levels 
								foreach(MappingServiceLevel objMappingServiceLevel in parDataSet.dsMappingServiceLevels.Values
									.Where(sl => sl.MappingDeliverableIDsp == objMappingDeliverable.IDsp))
									{
									Console.WriteLine("\t\t\t\t + ServiceLevel: {0} - {1}", objMappingServiceLevel.IDsp, objMappingServiceLevel.Title);
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
									if(objMappingServiceLevel.NewServiceLevel != null && objMappingServiceLevel.NewServiceLevel == true)
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
											dictionaryMatrixComments.Add("F|" + intMatrixSheet_RowIndex,
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

										if(parDataSet.dsServiceLevels.TryGetValue(
												key: Convert.ToInt16(objMappingServiceLevel.MappedServiceLevelIDsp),
												value: out objServiceLevel))
											{
											if(objServiceLevel.CSDdescription != null)
												{
												dictionaryMatrixComments.Add("F|" + intMatrixSheet_RowIndex,
													objServiceLevel.CSDdescription);
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
											parCellcontents: objMappingServiceLevel.IDsp.ToString(),
											parHyperlinkCounter: intHyperlinkCounter,
											parHyperlinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_MappingServiceLevels +
												Properties.AppResources.EditFormURI + objMappingServiceLevel.IDsp.ToString());
										//--- Matrix --- Service Level Row --- Column M --------------------------------
										oxmlWorkbook.PopulateCell(
											parWorksheetPart: objMatrixWorksheetPart,
											parColumnLetter: "M",
											parRowNumber: intMatrixSheet_RowIndex,
											parStyleId: (UInt32Value)(listMatrixColumnStyles.ElementAt(aWorkbook.GetColumnNumber("M"))),
											parCellDatatype: CellValues.String);

										//--- Matrix --- Service Level Row --- Column N --------------------------------
										if(objMappingServiceLevel.NewServiceLevel != null && objMappingServiceLevel.NewServiceLevel == true)
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
												parCellcontents: objServiceLevel.IDsp.ToString(),
												parHyperlinkCounter: intHyperlinkCounter,
											parHyperlinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_ServiceLevelsURI +
												Properties.AppResources.EditFormURI + objServiceLevel.IDsp.ToString()
												);
											}
										}
									} // foreach(MappingServiceLevel objMappingServiceLevel in listMappingServiceLevels)
								} // foreach(MappingDeliverable objMappingDeliverable in ...)
							} // foreach(MappingRequirement objRequirement in ....
					} //foreach(MappingServiceTower objTower in listMappingTowers)

Save_and_Close_Document:
//===============================================================
				if(dictionaryMatrixComments.Count() > 0)
					{
					// Now insert all the Comments
					aWorkbook.InsertWorksheetComments(
						parWorksheetPart: objMatrixWorksheetPart,
						parDictionaryOfComments: dictionaryMatrixComments);
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
				if(this.UploadDoc(parCompleteDataSet: ref parDataSet, parRequestingUserID: parRequestingUserID))
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
