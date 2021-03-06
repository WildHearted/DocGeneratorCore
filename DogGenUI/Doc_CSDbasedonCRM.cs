﻿using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Validation;
using DocGeneratorCore.Database.Classes;
using DocGeneratorCore.Database.Functions;
using DocGeneratorCore.SDDPServiceReference;


namespace DocGeneratorCore
	{
	/// <summary>
	/// This class is used to set all the properties for a
	/// CLient Service Description (CSD) based on a Client Requirements Mapping (CRM) Document.
	/// It inherits from the Document class.
	/// </summary>
	class CSD_based_on_ClientRequirementsMapping:aDocument
		{
		private bool _csd_Doc_based_on_CRM = false;
		public bool CSD_Doc_based_on_CRM
			{
			get{return this._csd_Doc_based_on_CRM;}
			set{this._csd_Doc_based_on_CRM = value;}
			}
		private int? _crm_Mapping = 0;
		/// <summary>
		/// This property reference the ID value of the SharePoint Mappings entry which is used to generate the Document
		/// </summary>
		public int? CRM_Mapping
			{
			get{return this._crm_Mapping;}
			set{this._crm_Mapping = value;}
			}
		private bool _requirements_Section = false;
		public bool Requirements_Section
			{
			get{return this._requirements_Section;}
			set{this._requirements_Section = value;}
			}
		private bool _tower_of_Service_Heading = false;
		public bool Tower_of_Service_Heading
			{
			get{return _tower_of_Service_Heading;}
			set{this._tower_of_Service_Heading = value;}
			}
		private bool _requirement_Heading = false;
		public bool Requirement_Heading
			{
			get{return this._requirement_Heading;}
			set{this._requirement_Heading = value;}
			}
		private bool _requirement_Reference = false;
		public bool Requirement_Reference
			{
			get{return this._requirement_Reference;}
			set{this._requirement_Reference = value;}
			}
		private bool _requirement_Text = false;
		public bool Requirement_Text
			{
			get{return this._requirement_Text;}
			set{this._requirement_Text = value;}
			}
		private bool _requirement_Service_Level = false;
		public bool Requirement_Service_Level
			{
			get{return this._requirement_Service_Level;}
			set{this._requirement_Service_Level = value;}
			}
		private bool _risks = false;
		public bool Risks
			{
			get{return this._risks;}
			set{this._risks = value;}
			}
		private bool _risk_Heading = false;
		public bool Risk_Heading
			{
			get{return this._risk_Heading;}
			set{this._risk_Heading = value;}
			}
		private bool _risk_Description = false;
		public bool Risk_Description
			{
			get{return this._risk_Description;}
			set{this._risk_Description = value;}
			}
		private bool _assumptions = false;
		public bool Assumptions
			{
			get{return this._assumptions;}
			set{this._assumptions = value;}
			}
		private bool _assumption_Heading = false;
		public bool Assumption_Heading
			{
			get{return this._assumption_Heading;}
			set{this._assumption_Heading = value;}
			}
		private bool _assumption_Description = false;
		public bool Assumption_Description
			{
			get{return this._assumption_Description;}
			set{this._assumption_Description = value;}
			}
		private bool _deliverables_Reports_and_Meetings = false;
		public bool Deliverable_Reports_and_Meetings
			{
			get{return this._deliverables_Reports_and_Meetings;}
			set{this._deliverables_Reports_and_Meetings = value;}
			}
		private bool _drm_Heading = false;
		public bool DRM_Heading
			{
			get{return this._drm_Heading;}
			set{this._drm_Heading = value;}
			}
		private bool _drm_Description = false;
		public bool DRM_Description
			{
			get{return this._drm_Description;}
			set{this._drm_Description = value;}
			}
		private bool _dds_DRM_Obligations = false;
		public bool DDs_DRM_Obligations
			{
			get{return this._dds_DRM_Obligations;}
			set{this._dds_DRM_Obligations = value;}
			}
		private bool _clients_DRM_Responsibilities = false;
		public bool Clients_DRM_Responsibilities
			{
			get{return this._clients_DRM_Responsibilities;}
			set{this._clients_DRM_Responsibilities = value;}
			}
		private bool _drm_Exclusions = false;
		public bool DRM_Exclusions
			{
			get{return this._drm_Exclusions;}
			set{this._drm_Exclusions = value;}
			}
		private bool _drm_Governance_Controls = false;
		public bool DRM_Governance_Controls
			{
			get{return this._drm_Governance_Controls;}
			set{this._drm_Governance_Controls = value;}
			}
		private bool _service_Levels = false;
		public bool Service_Levels
			{
			get{return this._service_Levels;}
			set{this._service_Levels = value;}
			}
		private bool _service_Level_Heading = false;
		public bool Service_Level_Heading
			{
			get{return this._service_Level_Heading;}
			set{this._service_Level_Heading = value;}
			}
		private bool _service_Level_Commitments_Table = false;
		public bool Service_Level_Commitments_Table
			{
			get{return this._service_Level_Commitments_Table;}
			set{this._service_Level_Commitments_Table = value;}
			}

		/// <summary>
		/// this option takes the values passed into the method as a list of integers
		/// which represents the options the user selected and transposing the values by
		/// setting the properties of the object.
		/// </summary>
		/// <param name="parOptions">The input must represent a List<int> object.</int></param>
		/// <returns></returns>
		public void TransposeDocumentOptions(ref List<int> parOptions)
			{
			int errors = 0;
			if(parOptions != null)
				{
				if(parOptions.Count > 0)
					{
					foreach(int option in parOptions)
						{
						switch(option)
							{
						case 168:
							this.Introductory_Section = true;
							break;
						case 169:
							this.Introduction = true;
							break;
						case 170:
							this.Executive_Summary = true;
							break;
						case 171:
							this.Requirements_Section = true;
							break;
						case 172:
							this.Tower_of_Service_Heading = true;
							break;
						case 173:
							this.Requirement_Heading = true;
							break;
						case 174:
							this.Requirement_Reference = true;
							break;
						case 175:
							this.Requirement_Text = true;
							break;
						case 176:
							this.Requirement_Service_Level = true;
							break;
						case 177:
							this.Risks = true;
							break;
						case 178:
							this.Risk_Heading = true;
							break;
						case 179:
							this.Risk_Description = true;
							break;
						case 180:
							this.Assumptions = true;
							break;
						case 181:
							this.Assumption_Heading = true;
							break;
						case 182:
							this.Deliverable_Reports_and_Meetings = true;
							break;
						case 183:
							this.DRM_Heading = true;
							break;
						case 184:
							this.DRM_Description = true;
							break;
						case 185:
							this.DDs_DRM_Obligations = true;
							break;
						case 186:
							this.Clients_DRM_Responsibilities = true;
							break;
						case 187:
							this.DRM_Exclusions = true;
							break;
						case 188:
							this.DRM_Governance_Controls = true;
							break;
						case 189:
							this.Service_Levels = true;
							break;
						case 190:
							this.Service_Level_Heading = true;
							break;
						case 191:
							this.Service_Level_Commitments_Table = true;
							break;
						case 192:
							this.Acronyms_Glossary_of_Terms_Section = true;
							break;
						case 193:
							this.Acronyms = true;
							break;
						case 194:
							this.Glossary_of_Terms = true;
							break;
						default:
							// just ignore
							break;
							}
						} // foreach(int option in parOptions)
					}
				else
					{
					this.LogError("There are no selected options - (Application Error)");
					errors += 1;
					}
				}
			else
				{
				this.LogError("The selected options are null - (Application Error)");
				errors += 1;
				}
			}
		public void Generate(
			DesignAndDeliveryPortfolioDataContext parSDDPdatacontext,
			int? parRequestingUserID,
			string parClientName)
			{
			Console.WriteLine("\t Begin to generate {0}", this.DocumentType);
			this.UnhandledError = false;
			DateTime timeStarted = DateTime.Now;
			string hyperlinkImageRelationshipID = "";
			string documentCollection_HyperlinkURL = "";
			string currentListURI = "";
			string currentHyperlinkViewEditURI = "";
			string currentContentLayer = "None";
			Table objServiceLevelTable = new Table();
			int? layer1upDeliverableID = 0;
			int tableCaptionCounter = 0;
			int imageCaptionCounter = 0;
			int numberingCounter = 49;
			int hyperlinkCounter = 9;
			int iPictureNo = 49;
			string errorText = "";
			bool bWrittenTitle = false;
			bool bWrittenServiceLevelTitle = false;

			try
				{
				if(this.HyperlinkEdit)
					{
					documentCollection_HyperlinkURL = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
						Properties.AppResources.List_DocumentCollectionLibraryURI +
						Properties.AppResources.EditFormURI + this.DocumentCollectionID;
					currentHyperlinkViewEditURI = Properties.AppResources.EditFormURI;
					}
				if(this.HyperlinkView)
					{
					documentCollection_HyperlinkURL = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
						Properties.AppResources.List_DocumentCollectionLibraryURI +
						Properties.AppResources.DisplayFormURI + this.DocumentCollectionID;
					currentHyperlinkViewEditURI = Properties.AppResources.DisplayFormURI;
					}

				//- Validate if the user selected any content to be generated
				if(this.CRM_Mapping == null || this.CRM_Mapping == 0)
					{//- if nothing selected thow exception and exit
					throw new NoContentSpecifiedException("A Client Requirement Mapping was not specified/selected, therefore the document will be blank. "
						+ "Please specify/select a Client Requirement Mapping before submitting the document collection for generation.");
					}

				// define a new objOpenXMLdocument
				oxmlDocument objOXMLdocument = new oxmlDocument();
				// use CreateDocumentFromTemplate method to create a new MS Word Document based on the relevant template
				if(objOXMLdocument.CreateDocWbkFromTemplate(
					parDocumentOrWorkbook: enumDocumentOrWorkbook.Document, 
					parTemplateURL: this.Template, 
					parDocumentType: this.DocumentType,
					parSDDPdataContext: parSDDPdatacontext))
					{
					Console.WriteLine("\t\t objOXMLdocument:\n" +
					"\t\t\t+ LocalDocumentPath: {0}\n" +
					"\t\t\t+ DocumentFileName.: {1}\n" +
					"\t\t\t+ DocumentURI......: {2}", objOXMLdocument.LocalPath, objOXMLdocument.Filename, objOXMLdocument.LocalURI);
					}
				else
					{
					//- if the file creation failed.
					throw new DocumentUploadException(message: "DocGenerator was unable to create the document based on the template.");
					}

				this.LocalDocumentURI = objOXMLdocument.LocalURI;
				this.FileName = objOXMLdocument.Filename;

				//- Create and open the new Document
				this.DocumentStatus = enumDocumentStatusses.Creating;
				//- Open the MS Word document in Edit mode
				WordprocessingDocument objWPdocument = WordprocessingDocument.Open(path: objOXMLdocument.LocalURI, isEditable: true);
				//- Define all open XML object to use for building the document
				MainDocumentPart objMainDocumentPart = objWPdocument.MainDocumentPart;
				Body objBody = objWPdocument.MainDocumentPart.Document.Body;          // Define the objBody of the document
				Paragraph objParagraph = new Paragraph();
				ParagraphProperties objParaProperties = new ParagraphProperties();
				Run objRun1 = new Run();
				Run objRun2 = new Run();
				RunProperties objRunProperties = new RunProperties();
				Text objText = new Text();
				//- Declare the HTMLdecoder object and assign the document's WordProcessing Body to the WPbody property.
				HTMLdecoder objHTMLdecoder = new HTMLdecoder();
				objHTMLdecoder.WPbody = objBody;

				//- Determine the Page Size for the current Body object.
				SectionProperties objSectionProperties = new SectionProperties();
				this.PageWith = Convert.ToUInt32(Properties.AppResources.DefaultPageWidth);
				this.PageHeight = Convert.ToUInt32(Properties.AppResources.DefaultPageHeight);

				if(objBody.GetFirstChild<SectionProperties>() != null)
					{
					objSectionProperties = objBody.GetFirstChild<SectionProperties>();
					PageSize objPageSize = objSectionProperties.GetFirstChild<PageSize>();
					PageMargin objPageMargin = objSectionProperties.GetFirstChild<PageMargin>();
					if(objPageSize != null)
						{
						this.PageWith = objPageSize.Width;
						this.PageHeight = objPageSize.Height;
						//Console.WriteLine("\t\t Page width x height: {0} x {1} twips", this.PageWith, this.PageHight);
						}
					if(objPageMargin != null)
						{
						if(objPageMargin.Left != null)
							{
							this.PageWith -= objPageMargin.Left;
							//Console.WriteLine("\t\t\t - Left Margin..: {0} twips", objPageMargin.Left);
							}
						if(objPageMargin.Right != null)
							{
							this.PageWith -= objPageMargin.Right;
							//Console.WriteLine("\t\t\t - Right Margin.: {0} twips", objPageMargin.Right);
							}
						if(objPageMargin.Top != null)
							{
							string tempTop = objPageMargin.Top.ToString();
							//Console.WriteLine("\t\t\t - Top Margin...: {0} twips", tempTop);
							this.PageHeight -= Convert.ToUInt32(tempTop);
							}
						if(objPageMargin.Bottom != null)
							{
							string tempBottom = objPageMargin.Bottom.ToString();
							//Console.WriteLine("\t\t\t - Bottom Margin: {0} twips", tempBottom);
							this.PageHeight -= Convert.ToUInt32(tempBottom);
							}
						}
					}
				//- Subtract the Table/Image Left indentation value from the Page width to ensure the table/image fits in the available space.
				this.PageWith -= Convert.ToUInt16(Properties.AppResources.Document_Table_Left_Indent);
				Console.WriteLine("\t\t Effective pageWidth x pageHeight.: {0} x {1} twips", this.PageWith, this.PageHeight);

				//- Check whether Hyperlinks need to be included and add the image to the Document Body
				if(this.HyperlinkEdit || this.HyperlinkView)
					{
					//-Insert and embed the hyperlink image in the document and keep the Image's Relationship ID in a variable for repeated use
					hyperlinkImageRelationshipID = oxmlDocument.Insert_HyperlinkImage(parMainDocumentPart: ref objMainDocumentPart,
						parSDDPdatacontext: parSDDPdatacontext);
					}

				//+ Define the objects to be used in the construction of the document
				Deliverable objDeliverable = new Deliverable();
				Deliverable objDeliverableLayer1up = new Deliverable();
				DeliverableServiceLevel objDeliverableServiceLevel = new DeliverableServiceLevel();
				ServiceLevel objServiceLevel = new ServiceLevel();
				Mapping objMapping = new Mapping();
				MappingServiceTower objMappingServiceTower = new MappingServiceTower();
				MappingRequirement objMappingRequirement = new MappingRequirement();
				MappingAssumption objMappingAssumption = new MappingAssumption();
				MappingRisk objMappingRisk = new MappingRisk();

				//+Check is Content Layering was requested and add a Ledgend for the colour coding of content
				if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 0, parNoNumberedHeading: true);
					objRun1 = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_ColourCodingLedgend_Heading,
						parBold: true);
					objParagraph.Append(objRun1);
					objBody.Append(objParagraph);

					objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 0);
					objRun1 = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_ColourCodingLedgend_Text);
					objParagraph.Append(objRun1);
					objBody.Append(objParagraph);

					objParagraph = oxmlDocument.Construct_BulletParagraph(parBulletLevel: 1);
					objRun1 = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_ColourCodingLedgend_Layer1,
						parContentLayer: "Layer1");
					objParagraph.Append(objRun1);
					objBody.Append(objParagraph);

					objParagraph = oxmlDocument.Construct_BulletParagraph(parBulletLevel: 1);
					objRun1 = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_ColourCodingLedgend_Layer2,
						parContentLayer: "Layer2");
					objParagraph.Append(objRun1);
					objBody.Append(objParagraph);

					objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 0);
					objRun1 = oxmlDocument.Construct_RunText(
						parText2Write: " ");
					objParagraph.Append(objRun1);
					objBody.Append(objParagraph);
					}

				this.DocumentStatus = enumDocumentStatusses.Building;

				//+ Insert the **Introductory Section**
				if(this.Introductory_Section)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
					objRun1 = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_IntruductorySection_HeadingText,
						parIsNewSection: true);
					objParagraph.Append(objRun1);
					objBody.Append(objParagraph);
					}

				//+ Insert the **Introduction**
				if(this.Introduction)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Introduction_HeadingText);
					// Check if a hyperlink must be inserted
					if(documentCollection_HyperlinkURL != "")
						{
						hyperlinkCounter += 1;
						Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
							parMainDocumentPart: ref objMainDocumentPart,
							parImageRelationshipId: hyperlinkImageRelationshipID,
							parClickLinkURL: documentCollection_HyperlinkURL,
							parHyperlinkID: hyperlinkCounter);
						objRun1.Append(objDrawing);
						}
					objParagraph.Append(objRun1);
					objBody.Append(objParagraph);

					if(this.IntroductionRichText != null)
						{
						try
							{
							objHTMLdecoder.DecodeHTML(parClientName: parClientName,
								parMainDocumentPart: ref objMainDocumentPart,
								parDocumentLevel: 2,
								parHTML2Decode: HTMLdecoder.CleanHTML(this.IntroductionRichText, parClientName),
								parTableCaptionCounter: ref tableCaptionCounter,
								parImageCaptionCounter: ref imageCaptionCounter,
								parNumberingCounter: ref numberingCounter,
								parPictureNo: ref iPictureNo,
								parHyperlinkID: ref hyperlinkCounter,
								parPageHeightDxa: this.PageHeight,
								parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
							}
						catch(InvalidContentFormatException exc)
							{
							Console.WriteLine("\n\nException occurred: {0}", exc.Message);
							// A Table content error occurred, record it in the error log.
							this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
								+ " contains an error in Introduction's Enhance Rich Text. Please review the content (especially tables).");
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
							objRun1 = oxmlDocument.Construct_RunText(
								parText2Write: "A content error occurred at this position and valid content could " +
								"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it. Error Detail: " 
								+ exc.Message,
								parIsNewSection: false,
								parIsError: true);
							if(documentCollection_HyperlinkURL != "")
								{
								hyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
									parMainDocumentPart: ref objMainDocumentPart,
									parImageRelationshipId: hyperlinkImageRelationshipID,
									parHyperlinkID: hyperlinkCounter,
									parClickLinkURL: documentCollection_HyperlinkURL);
								objRun1.Append(objDrawing);
								}
							objParagraph.Append(objRun1);
							objBody.Append(objParagraph);
							}
						}
					}
				
				//+ Insert the **Executive Summary**
				if(this.Executive_Summary)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_ExecutiveSummary_HeadingText);
					// Check if a hyperlink must be inserted
					if(documentCollection_HyperlinkURL != "")
						{
						hyperlinkCounter += 1;
						Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
							parMainDocumentPart: ref objMainDocumentPart,
							parImageRelationshipId: hyperlinkImageRelationshipID,
							parClickLinkURL: documentCollection_HyperlinkURL,
							parHyperlinkID: hyperlinkCounter);
						objRun1.Append(objDrawing);
						}
					objParagraph.Append(objRun1);
					objBody.Append(objParagraph);

					if(this.ExecutiveSummaryRichText != null)
						{
						try
							{
							objHTMLdecoder.DecodeHTML(parClientName: parClientName,
								parMainDocumentPart: ref objMainDocumentPart,
								parDocumentLevel: 2,
								parHTML2Decode: HTMLdecoder.CleanHTML(this.ExecutiveSummaryRichText, parClientName),
								parTableCaptionCounter: ref tableCaptionCounter,
								parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter,
								parPictureNo: ref iPictureNo,
								parHyperlinkID: ref hyperlinkCounter,
								parPageHeightDxa: this.PageHeight,
								parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
							}
						catch(InvalidContentFormatException exc)
							{
							Console.WriteLine("\n\nException occurred: {0}", exc.Message);
							// A Table content error occurred, record it in the error log.
							this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
								+ " contains an error in Excecutive Summary's Enhance Rich Text. Please review the content (especially tables).");
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
							objRun1 = oxmlDocument.Construct_RunText(
								parText2Write: "A content error occurred at this position and valid content could " +
								"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it. Error Detail: " 
								+ exc.Message,
								parIsNewSection: false,
								parIsError: true);
							if(documentCollection_HyperlinkURL != "")
								{
								hyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
									parMainDocumentPart: ref objMainDocumentPart,
									parImageRelationshipId: hyperlinkImageRelationshipID,
									parHyperlinkID: hyperlinkCounter,
									parClickLinkURL: documentCollection_HyperlinkURL);
								objRun1.Append(objDrawing);
								}
							objParagraph.Append(objRun1);
							objBody.Append(objParagraph);
							}
						}
					}

				Console.WriteLine("Retrieving the Mapping Data from SharePoint...");
				bool retrievedCRM = false;
				if (this.CRM_Mapping == null)
					{
					Console.WriteLine("The Mapping value is NOT specified.");
					errorText = "Error: The Client Requirements Mapping was not specified - ID: " + this.CRM_Mapping
							+ ". Please Please specify a Mapping before requesting the generation of the document.";
					this.LogError(errorText);
					throw new NoContentSpecifiedException(message: errorText);
					}

				retrievedCRM = UpdateLocalDatabase.UpdateMappingData(parDatacontexSDDP: parSDDPdatacontext, parMapping: Convert.ToInt16(this.CRM_Mapping));

				if (retrievedCRM == false || Properties.Settings.Default.CurrentMappingIsPopulated == false) // There was an error retriving the Mapping
					{
					errorText = "Error: Unable to retrieve the Client Requirements Mapping data for Mapping ID: " + this.CRM_Mapping
							+ ". Please check if the entry still exist in the Mappings List in SharePoint and that the DocGenerator can access SharePoint).";
					this.LogError(errorText);
					objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
					objRun1 = oxmlDocument.Construct_RunText(
						parText2Write: errorText,
						parIsNewSection: false,
						parIsError: true);
					if(documentCollection_HyperlinkURL != "")
						{
						hyperlinkCounter += 1;
						Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
							parMainDocumentPart: ref objMainDocumentPart,
							parImageRelationshipId: hyperlinkImageRelationshipID,
							parHyperlinkID: hyperlinkCounter,
							parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion
							+ Properties.AppResources.List_Mappings
							+ currentHyperlinkViewEditURI
							+ this.CRM_Mapping);
						objRun1.Append(objDrawing);
						}
					objParagraph.Append(objRun1);
					objBody.Append(objParagraph);
					goto Save_and_Close_Document;
					}

				//+| Insert the user selected Requirements content into the document
				if(this.Requirements_Section)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
					objRun1 = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_RequirementsMapping_SectionHeading,
						parIsNewSection: true);
					objParagraph.Append(objRun1);
					objBody.Append(objParagraph);

					//+| Obtain the Mapping data 
					//-| Read the **Mapping** data from the local database
					objMapping = Mapping.Read(parIDsp: Convert.ToInt16(this.CRM_Mapping));
					if (objMapping != null)
						{
						Console.Write("\n\t + {0} - {1}", objMapping.IDsp, objMapping.Title);
						}
					else
						{
						// If the entry is not found - write an error in the document and record an error in the error log.
						errorText = "Error: The Mapping ID: " + this.CRM_Mapping
							+ " doesn't exist in SharePoint and couldn't be retrieved.";
						this.LogError(errorText);
						objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
						objRun1 = oxmlDocument.Construct_RunText(
							parText2Write: errorText,
							parIsNewSection: true,
							parIsError: true);
						objParagraph.Append(objRun1);
						objBody.Append(objParagraph);
						Console.Write("\n\t + {0} - {1}", objMapping.IDsp, errorText);
						}

					//+| Process each of the Mapping **Service Towers**
					List<MappingServiceTower> mappingServiceTowers = MappingServiceTower.ReadMappingServiceTowersForMapping(parMappingIDsp: this.CRM_Mapping);
					//- **--- Loop through all Service Towers for the Mapping ---**
					foreach (MappingServiceTower objTower in mappingServiceTowers.OrderBy(t => t.Title))
						{
						//-| Write the Mapping Service Tower to the Document
						Console.Write("\n\t\t + Tower: {0} - {1}", objTower.IDsp, objTower.Title);
						objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
						objRun1 = oxmlDocument.Construct_RunText(parText2Write: objTower.Title);
						//-| Check if a hyperlink must be inserted
						if(documentCollection_HyperlinkURL != "")
							{
							hyperlinkCounter += 1;
							Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
								parMainDocumentPart: ref objMainDocumentPart,
								parImageRelationshipId: hyperlinkImageRelationshipID,
								parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
									Properties.AppResources.List_MappingServiceTowers +
									currentHyperlinkViewEditURI + objTower.IDsp,
								parHyperlinkID: hyperlinkCounter);
							objRun1.Append(objDrawing);
							}
						objParagraph.Append(objRun1);
						objBody.Append(objParagraph);

						//- Check if the user selected to generate the Requirements
						if(this.Requirement_Heading == false)
							{
							continue; //- skip the rest and process the next Service Tower entry
							}

						//+ **Mapping requirements** for the specific Service Tower
						List<MappingRequirement> mappingRequirements = MappingRequirement.ReadAllForServiceTower(objTower.IDsp);
						foreach (MappingRequirement objRequirement in mappingRequirements.OrderBy(r => r.SortOrder).ThenBy(r => r.Title))
							{
							Console.Write("\n\t\t + Requirement: {0} - {1}", objRequirement.IDsp, objRequirement.Title);
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3);
							objRun1 = oxmlDocument.Construct_RunText(
								parText2Write: objRequirement.Title);

							//- Check if a hyperlink must be inserted
							if(documentCollection_HyperlinkURL != "")
								{
								hyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
									parMainDocumentPart: ref objMainDocumentPart,
									parImageRelationshipId: hyperlinkImageRelationshipID,
									parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
									Properties.AppResources.List_MappingRequirements +
									currentHyperlinkViewEditURI + objRequirement.IDsp,
									parHyperlinkID: hyperlinkCounter);
								objRun1.Append(objDrawing);
								}
							objParagraph.Append(objRun1);
							objBody.Append(objParagraph);

							//- Check if the Requirement Reference is required
							if(this.Requirement_Reference)
								{
								objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 3);
								objRun1 = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_RequirementsMapping_ReferenceSourceTitle,
									parBold: true);
								objRun2 = oxmlDocument.Construct_RunText(parText2Write: objRequirement.SourceReference);
								objParagraph.Append(objRun1);
								objParagraph.Append(objRun2);
								objBody.Append(objParagraph);
								}

							//- Check if the user specified to include the Requirement Text
							if(this.Requirement_Text)
								{
								if(objRequirement.RequirementText != null)
									{
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 3);
									objRun1 = oxmlDocument.Construct_RunText(parText2Write: objRequirement.RequirementText);
									objParagraph.Append(objRun1);
									objBody.Append(objParagraph);
									}
								}

							//- Check if the user specified to include the Requirement Service Level
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

							//- Insert the Requirement Compliance:
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 3);
							objRun1 = oxmlDocument.Construct_RunText(
								parText2Write: Properties.AppResources.Document_RequirementsMapping_ComplianceStatusTitle,
								parBold: true);
							objParagraph.Append(objRun1);
							if(objRequirement.ComplianceStatus == null)
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

							//+include the Risks
							bWrittenTitle = false;
							if(this.Risks)
								{
								//- Process all the Mapping Risks for the specific Service Requirement
								List<MappingRisk> mappingRisks = MappingRisk.ReadAllForRequirement(parMappingRequirementIDsp: objRequirement.IDsp);
								foreach (MappingRisk objRisk in mappingRisks.OrderBy(r => r.Title))
									{
									// Insert the Risks Heading if not written yet.
									if(!bWrittenTitle)
										{
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun1 = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_RequirementsMapping_RisksHeading);
										objParagraph.Append(objRun1);
										objBody.Append(objParagraph);
										bWrittenTitle = true;
										}

									Console.Write("\n\t\t\t + Risk: {0} - {1}", objRisk.IDsp, objRisk.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
									objRun1 = oxmlDocument.Construct_RunText(parText2Write: objRisk.Title);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_MappingRisks +
											currentHyperlinkViewEditURI + objRisk.IDsp,
											parHyperlinkID: hyperlinkCounter);
										objRun1.Append(objDrawing);
										}
									objParagraph.Append(objRun1);
									objBody.Append(objParagraph);

									//- Check if the Risk Description Table is required
									if(this.Risk_Description)
										{
										Table tableMappingRisk = new Table();
										tableMappingRisk = CommonProcedures.BuildRiskTable(
											parMappingRisk: objRisk,
											parWidthColumn1: Convert.ToInt16(this.PageWith * 0.3),
											parWidthColumn2: Convert.ToInt16(this.PageWith * 0.7));
										objBody.Append(tableMappingRisk);
										}
									} //-foreach(Mappingrisk objMappingRisk ...)
								} //- if(this.Risks)

							//+ include the Assumptions
							if(this.Assumptions)
								{
								bWrittenTitle = false;
								//- Process all the Mapping Assumptions for the specific Service Requirement
								List<MappingAssumption> mappingAssumptions = MappingAssumption.ReadAllForRequirement(parMappingRequirementIDsp: objRequirement.IDsp);
								foreach (MappingAssumption objAssumption in mappingAssumptions.OrderBy(a => a.Title))
									{
									//- Insert the Assumption Heading
									if(!bWrittenTitle)
										{
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun1 = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_RequirementMapping_AssumptionsHeading);
										objParagraph.Append(objRun1);
										objBody.Append(objParagraph);
										bWrittenTitle = true;
										}

									Console.Write("\n\t\t\t + Assumption: {0} - {1}", objAssumption.IDsp, objAssumption.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
									objRun1 = oxmlDocument.Construct_RunText(
										parText2Write: objAssumption.Title);
									//- Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_MappingAssumptions +
											currentHyperlinkViewEditURI + objAssumption.IDsp,
											parHyperlinkID: hyperlinkCounter);
										objRun1.Append(objDrawing);
										}
									objParagraph.Append(objRun1);
									objBody.Append(objParagraph);

									//- Check if the Requirement Description Table
									if(this.Assumption_Description)
										{
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
										objRun1 = oxmlDocument.Construct_RunText(parText2Write: objAssumption.Description);
										objParagraph.Append(objRun1);
										objBody.Append(objParagraph);
										}
									} //-foreach(MappingAssumption objMappingAssumption in ...)
								} //-if(this.Assumptions)

							
							//+Include the **Deliverables, Reports & Meetings**
							if(this.Deliverable_Reports_and_Meetings)
								{
								bWrittenTitle = false;
								//- Process all the **Mapping Deliverables** for the specific Requirement
								List<MappingDeliverable> mappingDeliverables = MappingDeliverable.ReadAllForRequirement(parMappingRequirementIDsp: objRequirement.IDsp);
								foreach (var objMappingDeliverable in mappingDeliverables.OrderBy(d => d.Title))
									{
									if(!bWrittenTitle)
										{
										//- Insert the **Mapping Deliverable Heading**:
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun1 = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_RequirementsMapping_DeliverableReportMeetingsHeading);
										objParagraph.Append(objRun1);
										objBody.Append(objParagraph);
										bWrittenTitle = true;
										}

									Console.Write("\n\t\t\t + DRM: {0} - {1}", objMappingDeliverable.ID, objMappingDeliverable.Title);
									//- Insert the MappingDeliverable Title
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
									
									//- If it is a new deliverable, use the MappingDeliverable's Title else use the actual
									//- Mapped_Deliverable's CSD Description
									if(objMappingDeliverable.NewDeliverable)
										{
										objRun1 = oxmlDocument.Construct_RunText(parText2Write: objMappingDeliverable.Title);
										}
									else //- Existing Deliverable
										{
										//- Get the **Deliverable** entry from the Local Database
										objDeliverable = Deliverable.Read(parIDsp: Convert.ToInt16(objMappingDeliverable.MappedDeliverableID));
										if (objDeliverable != null)
											{
											Console.Write("\t + {0} - {1}", objDeliverable.IDsp, objDeliverable.Title);
											//- Insert the Deliverable CSD Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
											objRun1 = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.CSDheading);
											}
										else
											{
											Console.Write("\t + {0} - {1}", objMappingDeliverable.MappedDeliverableID, "Not found!");
											// Insert the Deliverable CSD Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
											objRun1 = oxmlDocument.Construct_RunText(parText2Write: "Deliverable coluld not be found!", parIsError: true);
											this.LogError("Error: The Deliverable ID: " + objMappingDeliverable.MappedDeliverableID
															+ " has probably been deleted since it was originally mapped. Please review the "
															+ "content and select a deliverable thatt currently exist or create a new one.");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
											objRun1 = oxmlDocument.Construct_RunText(
												parText2Write: "Error: The Deliverable ID: " + objMappingDeliverable.MappedDeliverableID
															+ " has probably been deleted since it was originally mapped. Please review the "
															+ "content and select a deliverable thatt currently exist or create a new one.",
												parIsNewSection: false,
												parIsError: true);
											continue;
											}
										}
									//- Insert hyperlink if required
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_MappingDeliverables +
											currentHyperlinkViewEditURI + objMappingDeliverable.ID,
											parHyperlinkID: hyperlinkCounter);
										objRun1.Append(objDrawing);
										}
									objParagraph.Append(objRun1);
									objBody.Append(objParagraph);

									//- Insert the **Deliverable Description**
									//- If it a **New** deliverable, use the NewRequirement, ELSE process the **Mapped_Deliverable's** content
									if(objMappingDeliverable.NewDeliverable)
										{
										//- Check if the user selected to insert **Deliverable Description** was selected
										if(this.DRM_Description)
											{
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
											objRun1 = oxmlDocument.Construct_RunText(parText2Write: objMappingDeliverable.NewRequirement);
											objParagraph.Append(objRun1);
											objBody.Append(objParagraph);
											}
										}
									else //- if(objMappingDeliverable.NewDeliverable != true :. it is an existing deliverable that must be included)
										{
										//- Check if the **Mapping Deliverable,Report,Meeting Description** was selected *AND* the Deliverable was found.
										if(this.DRM_Description && objDeliverable != null)
											{
											//-Check if the **Mapped_Deliverable** *Layer0up* has Content Layers and Content Predecessors
											Console.WriteLine("\n\t\t + Deliverable Layer 0..: {0} - {1}", objDeliverable.IDsp, objDeliverable.Title);
											if(objDeliverable.ContentPredecessorDeliverableIDsp == null)
												layer1upDeliverableID = null;
											else
												{
												layer1upDeliverableID = objDeliverable.ContentPredecessorDeliverableIDsp;
												//-| Get the entry from the Database
												objDeliverableLayer1up = Deliverable.Read(parIDsp: Convert.ToInt16(layer1upDeliverableID));
												if (objDeliverableLayer1up == null)
													layer1upDeliverableID = null;
												}
											
											// Insert Layer 1up if present and not null
											if(layer1upDeliverableID != null)
												{
												if(objDeliverableLayer1up.CSDdescription != null)
													{
													// Check for Colour coding Layers and add if necessary
													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer1";
													else
														currentContentLayer = "None";

													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer1up.IDsp;
														}
													else
														currentListURI = "";
													try
														{
														objHTMLdecoder.DecodeHTML(parClientName: parClientName,
															parMainDocumentPart: ref objMainDocumentPart,
															parDocumentLevel: 5,
															parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverableLayer1up.CSDdescription, parClientName),
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref tableCaptionCounter,
															parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter,
															parPictureNo: ref iPictureNo,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightDxa: this.PageHeight,
															parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
														}
													catch(InvalidContentFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
															+ " contains an error in CSD Description's Enhance Rich Text. Please review the "
															+ "content (especially tables).");
														objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
														objRun1 = oxmlDocument.Construct_RunText(
															parText2Write: "A content error occurred at this position and valid content could "
															+ "not be interpreted and inserted here. Please review the content in the SharePoint "
															+ "system and correct it. Error Detail: " + exc.Message,
															parIsNewSection: false,
															parIsError: true);
														if(documentCollection_HyperlinkURL != "")
															{
															hyperlinkCounter += 1;
															Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
																parMainDocumentPart: ref objMainDocumentPart,
																parImageRelationshipId: hyperlinkImageRelationshipID,
																parHyperlinkID: hyperlinkCounter,
																parClickLinkURL: currentListURI);
															objRun1.Append(objDrawing);
															}
														objParagraph.Append(objRun1);
														objBody.Append(objParagraph);
														}
													}//- if(objDeliverable.Layer1up.Layer1up.CSDdescription != null)
												} //- if(layer2upDeliverableID != null)

											//- Insert Layer0up if present and not null
											if(objDeliverable.CSDdescription != null)
												{
												// Check for Colour coding Layers and add if necessary
												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";

												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.IDsp;
													}
												else
													currentListURI = "";

												try
													{
													objHTMLdecoder.DecodeHTML(parClientName: parClientName,
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 5,
														parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverable.CSDdescription, parClientName),
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref tableCaptionCounter,
														parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter,
														parPictureNo: ref iPictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													//- A Table content error occurred, record it in the error log.
													this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
														+ " contains an error in CSD Description's Enhance Rich Text. Please review the "
														+ "content (especially tables).");
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
													objRun1 = oxmlDocument.Construct_RunText(
														parText2Write: "A content error occurred at this position and valid content could "
														+ "not be interpreted and inserted here. Please review the content in the SharePoint "
														+ "system and correct it. Error Detail: " + exc.Message,
														parIsNewSection: false,
														parIsError: true);
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
															parMainDocumentPart: ref objMainDocumentPart,
															parImageRelationshipId: hyperlinkImageRelationshipID,
															parHyperlinkID: hyperlinkCounter,
															parClickLinkURL: currentListURI);
														objRun1.Append(objDrawing);
														}
													objParagraph.Append(objRun1);
													objBody.Append(objParagraph);
													}
												} //- if(objDeliverable.CSDdescription != null)
											} //- if (this.DRM_Description)

											//-----------------------------------------------------------------------
											//- Check if the user specified to include the Deliverable DD's Obligations
											if(this.DDs_DRM_Obligations && objDeliverable != null)
												{
												if(objDeliverable.DDobligations != null
												|| (layer1upDeliverableID != null && objDeliverableLayer1up.DDobligations != null))
													{
													// Insert the Heading
													objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
													objRun1 = oxmlDocument.Construct_RunText(
														parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
													objParagraph.Append(objRun1);
													objBody.Append(objParagraph);

													// Insert Layer 1up if present and not null
													if(layer1upDeliverableID != null)
														{
														if(objDeliverableLayer1up.DDobligations != null)
															{
															// Check if a hyperlink must be inserted
															if(documentCollection_HyperlinkURL != "")
																{
																hyperlinkCounter += 1;
																currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
																	Properties.AppResources.List_DeliverablesURI +
																	currentHyperlinkViewEditURI +
																	objDeliverableLayer1up.IDsp;
																}
															else
																currentListURI = "";

															if(this.ColorCodingLayer1)
																currentContentLayer = "Layer1";
															else
																currentContentLayer = "None";
															try
																{
																objHTMLdecoder.DecodeHTML(parClientName: parClientName,
																	parMainDocumentPart: ref objMainDocumentPart,
																	parDocumentLevel: 6,
																	parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverableLayer1up.DDobligations, parClientName),
																	parContentLayer: currentContentLayer,
																	parTableCaptionCounter: ref tableCaptionCounter,
																	parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter,
																	parPictureNo: ref iPictureNo,
																	parHyperlinkID: ref hyperlinkCounter,
																	parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
																	parHyperlinkURL: currentListURI,
																	parPageHeightDxa: this.PageHeight,
																	parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
																}
															catch(InvalidContentFormatException exc)
																{
																Console.WriteLine("\n\nException occurred: {0}", exc.Message);
																//- A Table content error occurred, record it in the error log.
																this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
																	+ " contains an error in DD's Obligation's Enhance Rich Text. Please review the "
																	+ "content (especially tables).");
																objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
																objRun1 = oxmlDocument.Construct_RunText(
																	parText2Write: "A content error occurred at this position and valid content could "
																	+ "not be interpreted and inserted here. Please review the content in the SharePoint "
																	+ "system and correct it. Error Detail: " + exc.Message,
																	parIsNewSection: false,
																	parIsError: true);
																if(documentCollection_HyperlinkURL != "")
																	{
																	hyperlinkCounter += 1;
																	Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
																		parMainDocumentPart: ref objMainDocumentPart,
																		parImageRelationshipId: hyperlinkImageRelationshipID,
																		parHyperlinkID: hyperlinkCounter,
																		parClickLinkURL: currentListURI);
																	objRun1.Append(objDrawing);
																	}
																objParagraph.Append(objRun1);
																objBody.Append(objParagraph);
																}
															} //- if(objDeliverableLayer1up.DDobligations != null)
														} //- if(layer2upDeliverableID != null)

													//- Insert Layer0up if not null
													if(objDeliverable.DDobligations != null)
														{
														// Check if a hyperlink must be inserted
														if(documentCollection_HyperlinkURL != "")
															{
															hyperlinkCounter += 1;
															currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
																Properties.AppResources.List_DeliverablesURI +
																currentHyperlinkViewEditURI +
																objDeliverable.IDsp;
															}
														else
															currentListURI = "";

														if(this.ColorCodingLayer1)
															currentContentLayer = "Layer2";
														else
															currentContentLayer = "None";

														try
															{
															objHTMLdecoder.DecodeHTML(parClientName: parClientName,
																parMainDocumentPart: ref objMainDocumentPart,
																parDocumentLevel: 6,
																parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverable.DDobligations, parClientName),
																parContentLayer: currentContentLayer,
																parTableCaptionCounter: ref tableCaptionCounter,
																parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter,
																parPictureNo: ref iPictureNo,
																parHyperlinkID: ref hyperlinkCounter,
																parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
																parHyperlinkURL: currentListURI,
																parPageHeightDxa: this.PageHeight,
																parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
															}
														catch(InvalidContentFormatException exc)
															{
															Console.WriteLine("\n\nException occurred: {0}", exc.Message);
															//- A Table content error occurred, record it in the error log.
															this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
																+ " contains an error in DD's Obligation's  Enhance Rich Text. Please review the "
																+ "content (especially tables).");
															objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
															objRun1 = oxmlDocument.Construct_RunText(
																parText2Write: "A content error occurred at this position and valid content could "
																+ "not be interpreted and inserted here. Please review the content in the SharePoint "
																+ "system and correct it. Error Detail: " + exc.Message,
																parIsNewSection: false,
																parIsError: true);
															if(documentCollection_HyperlinkURL != "")
																{
																hyperlinkCounter += 1;
																Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
																	parMainDocumentPart: ref objMainDocumentPart,
																	parImageRelationshipId: hyperlinkImageRelationshipID,
																	parHyperlinkID: hyperlinkCounter,
																	parClickLinkURL: currentListURI);
																objRun1.Append(objDrawing);
																}
															objParagraph.Append(objRun1);
															objBody.Append(objParagraph);
															}
														} //- if(objDeliverable.DDobligations != null)
													} //- if(objDeliverable.DDoblidations != null &&)
												} //- if(this.DDs_DRM_Objigations
											
											//- Check if the user specified to include the Client Responsibilities
											if(this.Clients_DRM_Responsibilities && objDeliverable != null)
												{
												if(objDeliverable.ClientResponsibilities != null
												|| (layer1upDeliverableID != null && objDeliverableLayer1up.ClientResponsibilities != null))
													{
													// Insert the Heading
													objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
													objRun1 = oxmlDocument.Construct_RunText(
														parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
													objParagraph.Append(objRun1);
													objBody.Append(objParagraph);
																									
													// Insert Layer 1up if present and not null
													if(layer1upDeliverableID != null)
														{
														if(objDeliverableLayer1up.ClientResponsibilities != null)
															{
															// Check if a hyperlink must be inserted
															if(documentCollection_HyperlinkURL != "")
																{
																hyperlinkCounter += 1;
																currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
																	Properties.AppResources.List_DeliverablesURI +
																	currentHyperlinkViewEditURI +
																	objDeliverableLayer1up.IDsp;
																}
															else
																currentListURI = "";

															if(this.ColorCodingLayer1)
																currentContentLayer = "Layer1";
															else
																currentContentLayer = "None";
															try
																{
																objHTMLdecoder.DecodeHTML(parClientName: parClientName,
																	parMainDocumentPart: ref objMainDocumentPart,
																	parDocumentLevel: 6,
																	parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverableLayer1up.ClientResponsibilities, parClientName),
																	parContentLayer: currentContentLayer,
																	parTableCaptionCounter: ref tableCaptionCounter,
																	parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter,
																	parPictureNo: ref iPictureNo,
																	parHyperlinkID: ref hyperlinkCounter,
																	parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
																	parHyperlinkURL: currentListURI,
																	parPageHeightDxa: this.PageHeight,
																	parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
																}
															catch(InvalidContentFormatException exc)
																{
																Console.WriteLine("\n\nException occurred: {0}", exc.Message);
																// A Table content error occurred, record it in the error log.
																this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
																	+ " contains an error in Client's Responsibilities Enhance Rich Text. "
																	+ "Please review the content (especially tables).");
																objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
																objRun1 = oxmlDocument.Construct_RunText(
																	parText2Write: "A content error occurred at this position and valid content could "
																	+ "not be interpreted and inserted here. Please review the content in the SharePoint "
																	+ "system and correct it. Error Detail: " + exc.Message,
																	parIsNewSection: false,
																	parIsError: true);
																if(documentCollection_HyperlinkURL != "")
																	{
																	hyperlinkCounter += 1;
																	Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
																		parMainDocumentPart: ref objMainDocumentPart,
																		parImageRelationshipId: hyperlinkImageRelationshipID,
																		parHyperlinkID: hyperlinkCounter,
																		parClickLinkURL: currentListURI);
																	objRun1.Append(objDrawing);
																	}
																objParagraph.Append(objRun1);
																objBody.Append(objParagraph);
																}
															} //- if(objDeliverableLayer1up.ClientResponsibilities != null)
														} //- if(layer2upDeliverableID != null)

													//- Insert Layer0up if not null
													if(objDeliverable.ClientResponsibilities != null)
														{
														// Check if a hyperlink must be inserted
														if(documentCollection_HyperlinkURL != "")
															{
															hyperlinkCounter += 1;
															currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
																Properties.AppResources.List_DeliverablesURI +
																currentHyperlinkViewEditURI +
																objDeliverable.IDsp;
															}
														else
															currentListURI = "";

														if(this.ColorCodingLayer1)
															currentContentLayer = "Layer2";
														else
															currentContentLayer = "None";

														try
															{
															objHTMLdecoder.DecodeHTML(parClientName: parClientName,
																parMainDocumentPart: ref objMainDocumentPart,
																parDocumentLevel: 6,
																parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverable.ClientResponsibilities, parClientName),
																parContentLayer: currentContentLayer,
																parTableCaptionCounter: ref tableCaptionCounter,
																parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter,
																parPictureNo: ref iPictureNo,
																parHyperlinkID: ref hyperlinkCounter,
																parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
																parHyperlinkURL: currentListURI,
																parPageHeightDxa: this.PageHeight,
																parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
															}
														catch(InvalidContentFormatException exc)
															{
															Console.WriteLine("\n\nException occurred: {0}", exc.Message);
															// A Table content error occurred, record it in the error log.
															this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
																+ " contains an error in Client Responsibilies Enhance Rich Text. Please review the "
																+ "content (especially tables).");
															objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
															objRun1 = oxmlDocument.Construct_RunText(
																parText2Write: "A content error occurred at this position and valid content could "
																+ "not be interpreted and inserted here. Please review the content in the SharePoint "
																+ "system and correct it. Error Detail: " + exc.Message,
																parIsNewSection: false,
																parIsError: true);
															if(documentCollection_HyperlinkURL != "")
																{
																hyperlinkCounter += 1;
																Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
																	parMainDocumentPart: ref objMainDocumentPart,
																	parImageRelationshipId: hyperlinkImageRelationshipID,
																	parHyperlinkID: hyperlinkCounter,
																	parClickLinkURL: currentListURI);
																objRun1.Append(objDrawing);
																}
															objParagraph.Append(objRun1);
															objBody.Append(objParagraph);
															}
														} //- if(objDeliverable.ClientResponsibilities != null)
													} //- if(objDeliverable.ClientResponsibilities != null &&)
												} //- if(this.Clients_DRM_Responsibilities)

											//------------------------------------------------------------------
											//- Check if the user specified to include the Deliverable Exclusions
											if(this.DRM_Exclusions && objDeliverable != null)
												{
												if(objDeliverable.Exclusions != null
												|| (layer1upDeliverableID != null && objDeliverableLayer1up.Exclusions != null))
													{
													// Insert the Heading
													objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
													objRun1 = oxmlDocument.Construct_RunText(
														parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
													objParagraph.Append(objRun1);
													objBody.Append(objParagraph);
													
													//- Insert Layer 1up if present and not null
													if(layer1upDeliverableID != null)
														{
														if(objDeliverableLayer1up.Exclusions != null)
															{
															//- Check if a hyperlink must be inserted
															if(documentCollection_HyperlinkURL != "")
																{
																hyperlinkCounter += 1;
																currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
																	Properties.AppResources.List_DeliverablesURI +
																	currentHyperlinkViewEditURI +
																	objDeliverableLayer1up.IDsp;
																}
															else
																currentListURI = "";

															if(this.ColorCodingLayer1)
																currentContentLayer = "Layer1";
															else
																currentContentLayer = "None";

															try
																{
																objHTMLdecoder.DecodeHTML(parClientName: parClientName,
																	parMainDocumentPart: ref objMainDocumentPart,
																	parDocumentLevel: 6,
																	parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverableLayer1up.Exclusions, parClientName),
																	parContentLayer: currentContentLayer,
																	parTableCaptionCounter: ref tableCaptionCounter,
																	parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter,
																	parPictureNo: ref iPictureNo,
																	parHyperlinkID: ref hyperlinkCounter,
																	parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
																	parHyperlinkURL: currentListURI,
																	parPageHeightDxa: this.PageHeight,
																	parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
																}
															catch(InvalidContentFormatException exc)
																{
																Console.WriteLine("\n\nException occurred: {0}", exc.Message);
																//- A Table content error occurred, record it in the error log.
																this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
																	+ " contains an error in Exclusions Enhance Rich Text. Please review the "
																	+ "content (especially tables).");
																objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
																objRun1 = oxmlDocument.Construct_RunText(
																	parText2Write: "A content error occurred at this position and valid content could "
																	+ "not be interpreted and inserted here. Please review the content in the SharePoint "
																	+ "system and correct it. Error Detail: " + exc.Message,
																	parIsNewSection: false,
																	parIsError: true);
																if(documentCollection_HyperlinkURL != "")
																	{
																	hyperlinkCounter += 1;
																	Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
																		parMainDocumentPart: ref objMainDocumentPart,
																		parImageRelationshipId: hyperlinkImageRelationshipID,
																		parHyperlinkID: hyperlinkCounter,
																		parClickLinkURL: currentListURI);
																	objRun1.Append(objDrawing);
																	}
																objParagraph.Append(objRun1);
																objBody.Append(objParagraph);
																}
															} //- if(objDeliverableLayer1up.Exclusions != null)
														} //- if(layer2upDeliverableID != null)

													//- Insert Layer0up if not null
													if(objDeliverable.ClientResponsibilities != null)
														{
														//- Check if a hyperlink must be inserted
														if(documentCollection_HyperlinkURL != "")
															{
															hyperlinkCounter += 1;
															currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
																Properties.AppResources.List_DeliverablesURI +
																currentHyperlinkViewEditURI +
																objDeliverable.IDsp;
															}
														else
															currentListURI = "";

														if(this.ColorCodingLayer1)
															currentContentLayer = "Layer2";
														else
															currentContentLayer = "None";

														try
															{
															objHTMLdecoder.DecodeHTML(parClientName: parClientName,
																parMainDocumentPart: ref objMainDocumentPart,
																parDocumentLevel: 6,
																parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverable.Exclusions, parClientName),
																parContentLayer: currentContentLayer,
																parTableCaptionCounter: ref tableCaptionCounter,
																parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter,
																parPictureNo: ref iPictureNo,
																parHyperlinkID: ref hyperlinkCounter,
																parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
																parHyperlinkURL: currentListURI,
																parPageHeightDxa: this.PageHeight,
																parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
															}
														catch(InvalidContentFormatException exc)
															{
															Console.WriteLine("\n\nException occurred: {0}", exc.Message);
															//- A Table content error occurred, record it in the error log.
															this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
																+ " contains an error in Exclusions Enhance Rich Text. Please review the "
																+ "content (especially tables).");
															objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
															objRun1 = oxmlDocument.Construct_RunText(
																parText2Write: "A content error occurred at this position and valid content could "
																+ "not be interpreted and inserted here. Please review the content in the SharePoint "
																+ "system and correct it. Error Detail: " + exc.Message,
																parIsNewSection: false,
																parIsError: true);
															if(documentCollection_HyperlinkURL != "")
																{
																hyperlinkCounter += 1;
																Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
																	parMainDocumentPart: ref objMainDocumentPart,
																	parImageRelationshipId: hyperlinkImageRelationshipID,
																	parHyperlinkID: hyperlinkCounter,
																	parClickLinkURL: documentCollection_HyperlinkURL);
																objRun1.Append(objDrawing);
																}
															objParagraph.Append(objRun1);
															objBody.Append(objParagraph);
															}
														} //- if(objDeliverable.Exclusions != null)
													} //- if(objDeliverable.Exclusions != null &&)	
												} //- if(this.DRMe_Exclusions)

											
											//- Check if the user specified to include the **Governance Controls**
											if(this.DRM_Governance_Controls && objDeliverable != null)
												{
												if(objDeliverable.GovernanceControls != null
												|| (layer1upDeliverableID != null
												&& objDeliverableLayer1up.GovernanceControls != null))
													{
													//- Insert the Heading
													objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
													objRun1 = oxmlDocument.Construct_RunText(
														parText2Write: Properties.AppResources.Document_DeliverableGovernanceControls_Heading_Text);
													objParagraph.Append(objRun1);
													objBody.Append(objParagraph);

													//- Insert Layer 1up if present and not null
													if(layer1upDeliverableID != null)
														{
														if(objDeliverableLayer1up.GovernanceControls != null)
															{
															//- Check if a hyperlink must be inserted
															if(documentCollection_HyperlinkURL != "")
																{
																hyperlinkCounter += 1;
																currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
																	Properties.AppResources.List_DeliverablesURI +
																	currentHyperlinkViewEditURI +
																	objDeliverableLayer1up.IDsp;
																}
															else
																currentListURI = "";

															if(this.ColorCodingLayer1)
																currentContentLayer = "Layer1";
															else
																currentContentLayer = "None";

															try
																{
																objHTMLdecoder.DecodeHTML(parClientName: parClientName,
																	parMainDocumentPart: ref objMainDocumentPart,
																	parDocumentLevel: 6,
																	parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverableLayer1up.GovernanceControls, parClientName),
																	parContentLayer: currentContentLayer,
																	parTableCaptionCounter: ref tableCaptionCounter,
																	parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter,
																	parPictureNo: ref iPictureNo,
																	parHyperlinkID: ref hyperlinkCounter,
																	parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
																	parHyperlinkURL: currentListURI,
																	parPageHeightDxa: this.PageHeight,
																	parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
																}
															catch(InvalidContentFormatException exc)
																{
																Console.WriteLine("\n\nException occurred: {0}", exc.Message);
																// A Table content error occurred, record it in the error log.
																this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
																	+ " contains an error in Governance Controls Enhance Rich Text. Please review the "
																	+ "content (especially tables).");
																objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
																objRun1 = oxmlDocument.Construct_RunText(
																	parText2Write: "A content error occurred at this position and valid content could "
																	+ "not be interpreted and inserted here. Please review the content in the SharePoint "
																	+ "system and correct it. Error Detail: " + exc.Message,
																	parIsNewSection: false,
																	parIsError: true);
																if(documentCollection_HyperlinkURL != "")
																	{
																	hyperlinkCounter += 1;
																	Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
																		parMainDocumentPart: ref objMainDocumentPart,
																		parImageRelationshipId: hyperlinkImageRelationshipID,
																		parHyperlinkID: hyperlinkCounter,
																		parClickLinkURL: currentListURI);
																	objRun1.Append(objDrawing);
																	}
																objParagraph.Append(objRun1);
																objBody.Append(objParagraph);
																}
															} // if(objDeliverableLayer1up.GovernanceControls != null)
														} // if(layer2upDeliverableID != null)

													// Insert Layer0up if not null
													if(objDeliverable.GovernanceControls != null)
														{
														// Check if a hyperlink must be inserted
														if(documentCollection_HyperlinkURL != "")
															{
															hyperlinkCounter += 1;
															currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
																Properties.AppResources.List_DeliverablesURI +
																currentHyperlinkViewEditURI +
																objDeliverable.IDsp;
															}
														else
															currentListURI = "";

														if(this.ColorCodingLayer1)
															currentContentLayer = "Layer2";
														else
															currentContentLayer = "None";

														try
															{
															objHTMLdecoder.DecodeHTML(parClientName: parClientName,
																parMainDocumentPart: ref objMainDocumentPart,
																parDocumentLevel: 6,
																parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverable.GovernanceControls, parClientName),
																parContentLayer: currentContentLayer,
																parTableCaptionCounter: ref tableCaptionCounter,
																parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter,
																parPictureNo: ref iPictureNo,
																parHyperlinkID: ref hyperlinkCounter,
																parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
																parHyperlinkURL: currentListURI,
																parPageHeightDxa: this.PageHeight,
																parPageWidthDxa: this.PageWith, 
																parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint 
																	+ Properties.Settings.Default.CurrentURLSharePointSitePortion);
															}
														catch(InvalidContentFormatException exc)
															{
															Console.WriteLine("\n\nException occurred: {0}", exc.Message);
															// A Table content error occurred, record it in the error log.
															this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
																+ " contains an error in Governance Controls Enhance Rich Text. Please review the "
																+ "content (especially tables).");
															objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
															objRun1 = oxmlDocument.Construct_RunText(
																parText2Write: "A content error occurred at this position and valid content could "
																+ "not be interpreted and inserted here. Please review the content in "
																+ "the SharePoint system and correct it. Error Detail: " + exc.Message,
																parIsNewSection: false,
																parIsError: true);
															if(documentCollection_HyperlinkURL != "")
																{
																hyperlinkCounter += 1;
																Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
																	parMainDocumentPart: ref objMainDocumentPart,
																	parImageRelationshipId: hyperlinkImageRelationshipID,
																	parHyperlinkID: hyperlinkCounter,
																	parClickLinkURL: currentListURI);
																objRun1.Append(objDrawing);
																}
															objParagraph.Append(objRun1);
															objBody.Append(objParagraph);
															}
														} // if(objDeliverable.GovernanceControls != null)
													} // if(objDeliverable.GovernanceControls != null &&)	
												} //if(this.DRM_GovernanceControls)

											
											// Check if there are any **Glossary Terms or Acronyms** associated with the Deliverable(s).
											if(this.Acronyms_Glossary_of_Terms_Section && objDeliverable != null)
												{
												//- if there are GlossaryAndAcronyms to add from layer0up
												if(objDeliverable.GlossaryAndAcronyms != null)
													{
													foreach(var entry in objDeliverable.GlossaryAndAcronyms)
														{
														if(this.ListGlossaryAndAcronyms.Contains(entry) != true)
															ListGlossaryAndAcronyms.Add(entry);
														}
													}
												// if there are GlossaryAndAcronyms to add from layer1up
												if(layer1upDeliverableID != null
												&& objDeliverableLayer1up.GlossaryAndAcronyms != null)
													{
													foreach(var entry in objDeliverableLayer1up.GlossaryAndAcronyms)
														{
														if(this.ListGlossaryAndAcronyms.Contains(entry) != true)
															ListGlossaryAndAcronyms.Add(entry);
														}
													}
												} // if(this.Acronyms_Glossary_of_Terms_Section)

										//+include Service Levels
										if(this.Service_Level_Heading)
											{
											// Obtain all Service Levels for the specified Deliverable Requirement
											bWrittenServiceLevelTitle = false;
											// Process the Mapping Service Levels 

											List<MappingServiceLevel> mappingServiceLevels = MappingServiceLevel.ReadAllForMappingDeliverable(parMappingDeliverableIDsp: objMappingDeliverable.IDsp);
											foreach (MappingServiceLevel objMappingServiceLevel in mappingServiceLevels.OrderBy(s => s.Title))
												{
												if(!bWrittenServiceLevelTitle)
													{
													// Insert the Service Levels Heading:
													objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
													objRun1 = oxmlDocument.Construct_RunText(
														parText2Write: Properties.AppResources.Document_RequirementsMapping_ServiceLevelsHeading);
													objParagraph.Append(objRun1);
													objBody.Append(objParagraph);
													bWrittenServiceLevelTitle = true;
													}

												Console.Write("\n\t\t\t\t\t + DRM: {0} - {1}", objMappingServiceLevel.IDsp, objMappingServiceLevel.Title);
												//-| Insert the MappingServiceLevel Title
												objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
												//-| If it is a new Mapping Service level, use the MappingService Levels's Title else use the actual
												//-| Mapped_ServiceLevel's CSD Title
												if(objMappingServiceLevel.NewServiceLevel != null
												&& objMappingServiceLevel.NewServiceLevel == true)
													{
													objRun1 = oxmlDocument.Construct_RunText(parText2Write: objMappingServiceLevel.Title);
													//- Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
															parMainDocumentPart: ref objMainDocumentPart,
															parImageRelationshipId: hyperlinkImageRelationshipID,
															parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint 
																+ Properties.Settings.Default.CurrentURLSharePointSitePortion 
																+ Properties.AppResources.List_MappingServiceLevels 
																+ currentHyperlinkViewEditURI + objMappingServiceLevel.IDsp,
															parHyperlinkID: hyperlinkCounter);
														objRun1.Append(objDrawing);
														}
													}
												else //-| objMappingServiceLevel.NewServiceLevel != true)
													{
													objServiceLevel = ServiceLevel.Read(parIDsp: Convert.ToInt16(objMappingServiceLevel.MappedServiceLevelIDsp));
													if(objServiceLevel != null)
														{
														Console.Write("\t Existing Service Level: {0} - {1}", objServiceLevel.IDsp,
															objServiceLevel.Title);
														objRun1 = oxmlDocument.Construct_RunText(
														parText2Write: objServiceLevel.CSDheading);
														//- Check if a hyperlink must be inserted
														if(documentCollection_HyperlinkURL != "")
															{
															hyperlinkCounter += 1;
															Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
																parMainDocumentPart: ref objMainDocumentPart,
																parImageRelationshipId: hyperlinkImageRelationshipID,
																parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
																Properties.AppResources.List_ServiceLevelsURI +
																currentHyperlinkViewEditURI + objMappingServiceLevel.IDsp,
																parHyperlinkID: hyperlinkCounter);
															objRun1.Append(objDrawing);
															}
														}
													else
														{
														//- If the entry is not found - write an error in the document and record.
														errorText = "Error: The Service Level ID: " + objMappingServiceLevel.MappedServiceLevelIDsp
															+ " doesn't exist in SharePoint and it couldn't be retrieved.";
														this.LogError(errorText);
														objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
														objRun1 = oxmlDocument.Construct_RunText(
															parText2Write: errorText,
															parIsNewSection: false,
															parIsError: true);
														}
													objParagraph.Append(objRun1);
													objBody.Append(objParagraph);

													// Check if the user specified to include the Service Level Description
													if(this.Service_Level_Commitments_Table)
														{
														if(objMappingServiceLevel.NewServiceLevel != null
														&& objMappingServiceLevel.NewServiceLevel == true)
															{
															objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
															objRun1 = oxmlDocument.Construct_RunText(
																parText2Write: objMappingServiceLevel.RequirementText);
															objParagraph.Append(objRun1);
															objBody.Append(objParagraph);
															}
														else
															{
															// Prepare the data which to insert into the Service Level Table
															List<string> listErrorMessagesParameter = this.ErrorMessages;
															// Populate the Service Level Table
															objServiceLevelTable = CommonProcedures.BuildSLAtable(
																parMainDocumentPart: ref objMainDocumentPart,
																parClientName: parClientName,
																parServiceLevelID: objServiceLevel.IDsp,
																parWidthColumn1: Convert.ToInt16(this.PageWith * 0.20),
																parWidthColumn2: Convert.ToInt16(this.PageWith * 0.80),
																parMeasurement: objServiceLevel.Measurement,
																parMeasureMentInterval: objServiceLevel.MeasurementInterval,
																parReportingInterval: objServiceLevel.ReportingInterval,
																parServiceHours: objServiceLevel.ServiceHours,
																parCalculationMethod: objServiceLevel.CalculationMethod,
																parCalculationFormula: objServiceLevel.CalculationFormula,
																parThresholds: objServiceLevel.PerformanceThresholds,
																parTargets: objServiceLevel.PerformanceTargets,
																parBasicServiceLevelConditions: objServiceLevel.BasicConditions,
																parAdditionalServiceLevelConditions: "",
																parErrorMessages: ref listErrorMessagesParameter,
																parNumberingCounter: ref numberingCounter);

															if(listErrorMessagesParameter.Count != this.ErrorMessages.Count)
																this.ErrorMessages = listErrorMessagesParameter;

															objBody.Append(objServiceLevelTable);
															} //-else (objMappingServiceLevel.NewServiceLevel)
														} //- if(this.Service_Level_Commitments_Table
													} //- && objMappingServiceLevel.NewServiceLevel == true)
												} //- foreach(MappingServiceLevel objMappingServiceLevel in ....)
											} // if(this.ServiceLevelHeading)...
										} // if(objMappingDeliverable.NewDeliverable != true)
									} // foreach(MappingDeliverable objMappingDeliverable in .....)
								} // if(this.Deliverable_Reports_and_Meetings)
							} // foreach(MappingRequirement objRequirement in listMappingRequirements)
						} //foreach(MappingServiceTower objTower in listMappingTowers)
					} // if(this.RequirementSection)

				//---G
				// Insert the Glossary of Terms and Acronym Section
				if(this.Acronyms_Glossary_of_Terms_Section)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
					objRun1 = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_GlossaryAndAcronymSection_HeadingText,
						parIsNewSection: true);
					objParagraph.Append(objRun1);
					objBody.Append(objParagraph);

					// Insert a blank paragrpah
					objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
					objRun1 = oxmlDocument.Construct_RunText(
						parText2Write: " ");
					objParagraph.Append(objRun1);
					objBody.Append(objParagraph);

					List<string> listErrors = this.ErrorMessages;
					if(this.ListGlossaryAndAcronyms.Count > 0)
						{
						Table tableGlossaryAcronym = new Table();
						tableGlossaryAcronym = CommonProcedures.BuildGlossaryAcronymsTable(
							parSDDPdatacontext: parSDDPdatacontext,
							parGlossaryAcronyms: this.ListGlossaryAndAcronyms,
							parWidthColumn1: Convert.ToInt16(this.PageWith * 0.3),
							parWidthColumn2: Convert.ToInt16(this.PageWith * 0.2),
							parWidthColumn3: Convert.ToInt16(this.PageWith * 0.5),
							parErrorMessages: ref listErrors);
						objBody.Append(tableGlossaryAcronym);
						}     //if(this.TermAndAcronymList.Count > 0)
					} // if (this.Acronyms)

Save_and_Close_Document:

				if(this.ErrorMessages.Count > 0)
					{
					//--------------------------------------------------
					// Insert the Document Generation Error Section

					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
					objRun1 = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_Error_Section_Heading,
						parIsNewSection: true);
					objParagraph.Append(objRun1);
					objBody.Append(objParagraph);

					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Error_Heading);
					objParagraph.Append(objRun1);
					objBody.Append(objParagraph);

					foreach(var errorMessageEntry in this.ErrorMessages)
						{
						objParagraph = oxmlDocument.Construct_Error(errorMessageEntry);
						objBody.Append(objParagraph);
						}
					}

				//----------------------------------------------
				//Validate the document with OpenXML validator
				OpenXmlValidator objOXMLvalidator = new OpenXmlValidator(fileFormat: DocumentFormat.OpenXml.FileFormatVersions.Office2010);
				int errorCount = 0;
				Console.WriteLine("\n\rValidating document....");
				foreach(ValidationErrorInfo validationError in objOXMLvalidator.Validate(objWPdocument))
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
				objWPdocument.Close();

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
		} // end of CSD_ClientRequirementsMapping_Document class
	}
