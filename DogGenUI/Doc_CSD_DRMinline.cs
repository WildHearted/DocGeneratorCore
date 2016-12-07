using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocGeneratorCore.Database.Classes;
using DocGeneratorCore.SDDPServiceReference;

namespace DocGeneratorCore
	{
	/// <summary>
	///      This class contains all the Client Service Description (CSD) with inline DRM
	///      (Deliverable Report Meeting).
	/// </summary>
	internal class CSD_Document_DRM_Inline:External_Document
		{
		private bool _drm_Description = false;

		public bool DRM_Description
			{
			get { return this._drm_Description; }
			set { this._drm_Description = value; }
			}

		private bool _drm_Inputs = false;

		public bool DRM_Inputs
			{
			get { return this._drm_Inputs; }
			set { this._drm_Inputs = value; }
			}

		private bool _drm_Outputs = false;

		public bool DRM_Outputs
			{
			get { return this._drm_Outputs; }
			set { this._drm_Outputs = value; }
			}

		private bool _dds_DRM_Obligations = false;

		public bool DDS_DRM_Obligations
			{
			get { return this._dds_DRM_Obligations; }
			set { this._dds_DRM_Obligations = value; }
			}

		private bool _clients_DRM_Responsibilities = false;

		public bool Clients_DRM_Responsibilities
			{
			get { return this._clients_DRM_Responsibilities; }
			set { this._clients_DRM_Responsibilities = value; }
			}

		private bool _drm_Exclusions = false;

		public bool DRM_Exclusions
			{
			get { return this._drm_Exclusions; }
			set { this._drm_Exclusions = value; }
			}

		private bool _drm_Governance_Controls = false;

		public bool DRM_Governance_Controls
			{
			get { return this._drm_Governance_Controls; }
			set { this._drm_Governance_Controls = value; }
			}

		/// <summary>
		///      this option takes the values passed into the method as a list of integers which
		///      represents the options the user selected and transposing the values by setting the
		///      properties of the object.
		/// </summary>
		/// <param name="parOptions">
		///      The input must represent a List <int>object.</int>
		/// </param>
		/// <returns>
		/// </returns>
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
						case 144:
							this.Introductory_Section = true;
							break;

						case 145:
							this.Introduction = true;
							break;

						case 146:
							this.Executive_Summary = true;
							break;

						case 147:
							this.Service_Portfolio_Section = true;
							break;

						case 148:
							this.Service_Portfolio_Description = true;
							break;

						case 149:
							this.Service_Family_Heading = true;
							break;

						case 150:
							this.Service_Family_Description = true;
							break;

						case 151:
							this.Service_Product_Heading = true;
							break;

						case 152:
							this.Service_Product_Description = true;
							break;

						case 153:
							this.Service_Feature_Heading = true;
							break;

						case 154:
							this.Service_Feature_Description = true;
							break;

						case 155:
							this.Deliverables_Reports_Meetings = true;
							break;

						case 156:
							this.DRM_Heading = true;
							break;

						case 157:
							this.DRM_Description = true;
							break;

						case 158:
							this.DDS_DRM_Obligations = true;
							break;

						case 159:
							this.Clients_DRM_Responsibilities = true;
							break;

						case 160:
							this.DRM_Exclusions = true;
							break;

						case 161:
							this.DRM_Governance_Controls = true;
							break;

						case 162:
							this.Service_Levels = true;
							break;

						case 163:
							this.Service_Level_Heading = true;
							break;

						case 164:
							this.Service_Level_Commitments_Table = true;
							break;

						case 165:
							this.Acronyms_Glossary_of_Terms_Section = true;
							break;

						case 166:
							this.Acronyms = true;
							break;

						case 167:
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
			bool drmHeading = false;
			Table objActivityTable = new Table();
			Table objServiceLevelTable = new Table();
			int? layer1upFeatureID = 0;
			int? layer1upDeliverableID = 0;
			int tableCaptionCounter = 0;
			int imageCaptionCounter = 0;
			int numberingCounter = 49;
			int iPictureNo = 49;
			int hyperlinkCounter = 9;

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
				if(this.SelectedNodes == null || this.SelectedNodes.Count < 1)
					{//- if nothing selected thow exception and exit
					throw new NoContentSpecifiedException("No content was specified/selected, therefore the document will be blank. "
						+ "Please specify/select content before submitting the document collection for generation.");
					}

				// define a new objOpenXMLdocument
				oxmlDocument objOXMLdocument = new oxmlDocument();
				// use CreateDocumentFromTemplate method to create a new MS Word Document based on
				// the relevant template
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

				// Create and open the new Document
				this.DocumentStatus = enumDocumentStatusses.Creating;
				// Open the MS Word document in Edit mode
				WordprocessingDocument objWPdocument = WordprocessingDocument.Open(path: objOXMLdocument.LocalURI, isEditable: true);
				// Define all open XML object to use for building the document
				MainDocumentPart objMainDocumentPart = objWPdocument.MainDocumentPart;
				Body objBody = objWPdocument.MainDocumentPart.Document.Body;          // Define the objBody of the document
				Paragraph objParagraph = new Paragraph();
				ParagraphProperties objParaProperties = new ParagraphProperties();
				Run objRun = new Run();
				RunProperties objRunProperties = new RunProperties();
				Text objText = new Text();
				// Declare the HTMLdecoder object and assign the document's WordProcessing Body to
				// the WPbody property.
				HTMLdecoder objHTMLdecoder = new HTMLdecoder();
				objHTMLdecoder.WPbody = objBody;

				// Determine the Page Size for the current Body object.
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
				// Subtract the Table/Image Left indentation value from the Page width to ensure the
				// table/image fits in the available space.
				this.PageWith -= Convert.ToUInt16(Properties.AppResources.Document_Table_Left_Indent);
				//Console.WriteLine("\t\t Effective pageWidth x pageHeight.: {0} x {1} twips", this.PageWith, this.PageHight);

				// Check whether Hyperlinks need to be included and add the image to the Document Body
				if(this.HyperlinkEdit || this.HyperlinkView)
					{
					//Insert and embed the hyperlink image in the document and keep the Image's Relationship ID in a variable for repeated use
					hyperlinkImageRelationshipID = oxmlDocument.Insert_HyperlinkImage(
						parMainDocumentPart: ref objMainDocumentPart,
						parSDDPdatacontext: parSDDPdatacontext);
					}

				// Define the objects to be used in the construction of the document
				ServicePortfolio objPortfolio = new ServicePortfolio();
				ServiceFamily objFamily = new ServiceFamily();
				ServiceProduct objProduct = new ServiceProduct();
				ServiceFeature objFeature = new ServiceFeature();
				ServiceFeature objFeatureLayer1up = new ServiceFeature();
				Deliverable objDeliverable = new Deliverable();
				Deliverable objDeliverableLayer1up = new Deliverable();
				DeliverableActivity objDeliverableActivity = new DeliverableActivity();
				DeliverableServiceLevel objDeliverableServiceLevel = new DeliverableServiceLevel();
				Activity objActivity = new Activity();
				ServiceLevel objServiceLevel = new ServiceLevel();

				//Check is Content Layering was requested and add a Ledgend for the colour coding of content
				if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 0, parNoNumberedHeading: true);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_ColourCodingLedgend_Heading,
						parBold: true);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 0);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_ColourCodingLedgend_Text);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					objParagraph = oxmlDocument.Construct_BulletParagraph(parBulletLevel: 1);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_ColourCodingLedgend_Layer1,
						parContentLayer: "Layer1");
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					objParagraph = oxmlDocument.Construct_BulletParagraph(parBulletLevel: 1);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_ColourCodingLedgend_Layer2,
						parContentLayer: "Layer2");
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 0);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: " ");
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);
					}

				this.DocumentStatus = enumDocumentStatusses.Building;

				//++ Insert the Introductory Section
				if(this.Introductory_Section)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_IntruductorySection_HeadingText,
						parIsNewSection: true);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);
					}

				//+ Insert the Introduction
				if(this.Introduction)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Introduction_HeadingText);
					// Check if a hyperlink must be inserted
					if(this.HyperlinkEdit || this.HyperlinkView)
						{
						hyperlinkCounter += 1;
						Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
							parMainDocumentPart: ref objMainDocumentPart,
							parImageRelationshipId: hyperlinkImageRelationshipID,
							parClickLinkURL: documentCollection_HyperlinkURL,
							parHyperlinkID: hyperlinkCounter);
						objRun.Append(objDrawing);
						}
					objParagraph.Append(objRun);
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
								+ " contains an error in Introduction's Enhance Rich Text. "
								+ "Please review the content.");
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "A content error occurred at this position and valid content could " +
								"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it. Error Detail: "
								+ exc.Message,
								parIsNewSection: false,
								parIsError: true);
							if(this.HyperlinkEdit || this.HyperlinkView)
								{
								hyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
									parMainDocumentPart: ref objMainDocumentPart,
									parImageRelationshipId: hyperlinkImageRelationshipID,
									parHyperlinkID: hyperlinkCounter,
									parClickLinkURL: documentCollection_HyperlinkURL);
								objRun.Append(objDrawing);
								}
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);
							}
						}
					}

				//+ Insert the Executive Summary
				if(this.Executive_Summary)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_ExecutiveSummary_HeadingText);
					// Check if a hyperlink must be inserted
					if(this.HyperlinkEdit || this.HyperlinkView)
						{
						hyperlinkCounter += 1;
						Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
							parMainDocumentPart: ref objMainDocumentPart,
							parImageRelationshipId: hyperlinkImageRelationshipID,
							parClickLinkURL: documentCollection_HyperlinkURL,
							parHyperlinkID: hyperlinkCounter);
						objRun.Append(objDrawing);
						}
					objParagraph.Append(objRun);
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
								+ " contains an error in Executive Summary's Enhance Rich Text. "
								+ "Please review the content (especially tables).");
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "A content error occurred at this position and valid content could " +
								"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it. Error Detail: "
								+ exc.Message,
								parIsNewSection: false,
								parIsError: true);
							if(this.HyperlinkEdit || this.HyperlinkView)
								{
								hyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
									parMainDocumentPart: ref objMainDocumentPart,
									parImageRelationshipId: hyperlinkImageRelationshipID,
									parHyperlinkID: hyperlinkCounter,
									parClickLinkURL: documentCollection_HyperlinkURL);
								objRun.Append(objDrawing);
								}
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);
							}
						}
					}

				//++ Insert the user selected content
				if(this.SelectedNodes == null || this.SelectedNodes.Count <= 0)
					goto Process_Glossary_and_Acronyms;

				foreach(Hierarchy node in this.SelectedNodes)
					{
					Console.Write("\nNode: {0} - {1} {2} {3}", node.Sequence, node.Level, node.NodeType, node.NodeID);

					switch(node.NodeType)
						{
					//+ Service Framework & Service Portfolios
					case enumNodeTypes.FRA: //-| Service Framework
					case enumNodeTypes.POR: //-| Service Portfolio

						if(!this.Service_Portfolio_Section)
							break;

						objPortfolio = ServicePortfolio.Read(parIDsp: node.NodeID);
						if (objPortfolio != null)
							{
							Console.Write("\t + {0} - {1}", objPortfolio.IDsp, objPortfolio.Title);
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
							objRun = oxmlDocument.Construct_RunText(parText2Write: objPortfolio.CSDheading, parIsNewSection: true);

							// Check if a hyperlink must be inserted
							if(this.HyperlinkEdit || this.HyperlinkView)
								{
								hyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
									parMainDocumentPart: ref objMainDocumentPart,
									parImageRelationshipId: hyperlinkImageRelationshipID,
									parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
										Properties.AppResources.List_ServicePortfoliosURI +
										currentHyperlinkViewEditURI + objPortfolio.IDsp,
									parHyperlinkID: hyperlinkCounter);
								objRun.Append(objDrawing);
								}
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);

							//-| Check if the user specified to include the Service Porfolio Description
							if(this.Service_Portfolio_Description
							&& objPortfolio.CSDdescription != null)
								{
								currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
									Properties.AppResources.List_ServicePortfoliosURI +
									currentHyperlinkViewEditURI + objPortfolio.IDsp;

								try
									{
									objHTMLdecoder.DecodeHTML(parClientName: parClientName,
										parMainDocumentPart: ref objMainDocumentPart,
										parDocumentLevel: 1,
										parHTML2Decode: HTMLdecoder.CleanHTML(objPortfolio.CSDdescription, parClientName),
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
									//-| A Table content error occurred, record it in the error log.
									this.LogError("Error: The Service Portfolio ID: " + objPortfolio.IDsp
										+ " contains an error in CSD Description's Enhance Rich Text. "
										+ "Please review the content (especially tables).");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "A content error occurred at this position and valid content could "
										+ "not be interpreted and inserted here. Please review the content in the SharePoint "
										+ "system and correct it." + exc.Message,
										parIsNewSection: false,
										parIsError: true);
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parHyperlinkID: hyperlinkCounter,
											parClickLinkURL: currentListURI);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
								}
							}
						else
							{
							//-| If the entry is not found - write an error in the document and record
							//-| an error in the error log.
							this.LogError("Error: The Service Portfolio ID " + node.NodeID +
								" doesn't exist in SharePoint and couldn't be retrieved.");
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "Error: Service Portfolio " + node.NodeID + " is missing.",
								parIsNewSection: true,
								parIsError: true);
							objParagraph.Append(objRun);
							}
						break;

					//+ Service Family
					case enumNodeTypes.FAM:
						if(!this.Service_Family_Heading)
							break;

						//-| Get the entry from the Database
						objFamily = ServiceFamily.Read(parIDsp: node.NodeID);
						if (objFamily != null)
							{
							Console.WriteLine("\t + {0} - {1}", objFamily.IDsp, objFamily.Title);
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: objFamily.CSDheading,
								parIsNewSection: false);
							//- Check if a hyperlink must be inserted
							if(this.HyperlinkEdit || this.HyperlinkView)
								{
								hyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
									parMainDocumentPart: ref objMainDocumentPart,
									parImageRelationshipId: hyperlinkImageRelationshipID,
									parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
									Properties.AppResources.List_ServiceFamiliesURI +
									currentHyperlinkViewEditURI + objFamily.IDsp,
									parHyperlinkID: hyperlinkCounter);
								objRun.Append(objDrawing);
								}
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);

							//+ Insert the Service Family Description
							if(this.Service_Family_Description
							&& objFamily.CSDdescription != null)
								{
								currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
									Properties.AppResources.List_ServicePortfoliosURI +
									currentHyperlinkViewEditURI +
									objFamily.IDsp;
								try
									{
									objHTMLdecoder.DecodeHTML(parClientName: parClientName,
										parMainDocumentPart: ref objMainDocumentPart,
										parDocumentLevel: 2,
										parHTML2Decode: HTMLdecoder.CleanHTML(objFamily.CSDdescription, parClientName),
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
									this.LogError("Error: The Service Family ID: " + objFamily.IDsp
										+ " contains an error in CSD Description's Enhance Rich Text. "
										+ "Please review the content (especially tables).");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "A content error occurred at this position and valid content could "
										+ "not be interpreted and inserted here. Please review the content in the SharePoint "
										+ "system and correct it." + exc.Message,
										parIsNewSection: false,
										parIsError: true);
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parHyperlinkID: hyperlinkCounter,
											parClickLinkURL: currentListURI);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
								}
							}
						else
							{
							// If the entry is not found - write an error in the document and
							this.LogError("Error: The Service Family ID " + node.NodeID
								+ " doesn't exist in SharePoint and couldn't be retrieved.");
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "Error: Service Family " + node.NodeID + " is missing.",
								parIsNewSection: false,
								parIsError: true);
							objParagraph.Append(objRun);
							break;
							}
						break;

					//+ Service Product
					case enumNodeTypes.PRO:

						if(this.Service_Product_Heading)
							break;

						//-| Get the entry from the Database
						objProduct = ServiceProduct.Read(parIDsp: node.NodeID);
						if (objProduct != null)
							{
							Console.Write("\t + {0} - {1}", objProduct.IDsp, objProduct.Title);
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: objProduct.CSDheading,
								parIsNewSection: false);
							// Check if a hyperlink must be inserted
							if(this.HyperlinkEdit || this.HyperlinkView)
								{
								hyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
									parMainDocumentPart: ref objMainDocumentPart,
									parImageRelationshipId: hyperlinkImageRelationshipID,
									parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
									Properties.AppResources.List_ServiceProductsURI +
									currentHyperlinkViewEditURI + objProduct.IDsp,
									parHyperlinkID: hyperlinkCounter);
								objRun.Append(objDrawing);
								}
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);

							//+ Insert the Service Product Description
							if(this.Service_Product_Description
							&& objProduct.CSDdescription != null)
								{
								currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
									Properties.AppResources.List_ServiceProductsURI +
									currentHyperlinkViewEditURI +
									objProduct.IDsp;

								try
									{
									objHTMLdecoder.DecodeHTML(parClientName: parClientName,
										parMainDocumentPart: ref objMainDocumentPart,
										parDocumentLevel: 3,
										parHTML2Decode: HTMLdecoder.CleanHTML(objProduct.CSDdescription, parClientName),
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
									this.LogError("Error: The Service Product ID: " + objProduct.IDsp
										+ " contains an error in CSD Description's Enhance Rich Text. "
										+ "Please review the content (especially tables).");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 3);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "A content error occurred at this position and valid content could "
										+ "not be interpreted and inserted here. Please review the content in the SharePoint "
										+ "system and correct it." + exc.Message,
										parIsNewSection: false,
										parIsError: true);
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parHyperlinkID: hyperlinkCounter,
											parClickLinkURL: currentListURI);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
								}
							else
								{
								// If the entry is not found - write an error in the document
								this.LogError("Error: The Service Product ID " + node.NodeID
									+ " doesn't exist in SharePoint and couldn't be retrieved.");
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: "Error: Service Family " + node.NodeID + " is missing.",
									parIsNewSection: false,
									parIsError: true);
								objParagraph.Append(objRun);
								}
							} //- if(this.Service_Product_Heading)
						break;

					//+ Service Feature
					case enumNodeTypes.FEA:
						if(!this.Service_Feature_Heading)
							break;

						//-| Get the entry from the Database
						objFeature = ServiceFeature.Read(parIDsp: node.NodeID);
						if (objFeature != null)
							{
							Console.Write("\t + {0} - {1}", objFeature.IDsp, objFeature.Title);

							//-| Insert the Service Feature CSD Heading...
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
							objRun = oxmlDocument.Construct_RunText(parText2Write: objFeature.CSDheading);
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);

							//- Check if the Feature has Content Layers and Content Predecessors
							if(objFeature.ContentPredecessorFeatureIDsp == null)
								{
								layer1upFeatureID = null;
								}
							else
								{
								//- Get the entry from the Layer1up Database
								objFeatureLayer1up = ServiceFeature.Read(parIDsp: Convert.ToInt16(objFeature.ContentPredecessorFeatureIDsp));
								if (objFeatureLayer1up != null)
									{
									layer1upFeatureID = objFeature.ContentPredecessorFeatureIDsp;
									}
								else
									{
									layer1upFeatureID = null;
									}
								}

							//+ Insert the Service Feature Description
							if(this.Service_Feature_Description)
								{
								//- Insert Layer1up if present and not null
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upFeatureID != null
								&& objFeatureLayer1up.CSDdescription != null)
									{
									//- Insert a hyperlink if needed
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceFeaturesURI +
											currentHyperlinkViewEditURI +
											objFeatureLayer1up.IDsp;
										}
									else
										currentListURI = "";

									//- Check for Colour coding Layers and add if necessary
									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objFeatureLayer1up.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objFeatureLayer1up.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									try
										{
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 4,
											parHTML2Decode: HTMLdecoder.CleanHTML(objFeatureLayer1up.CSDdescription, parClientName),
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
										this.LogError("Error: The Service Feature ID: " + objFeatureLayer1up.IDsp
											+ " contains an error in CSD Description's Enhance Rich Text. "
											+ "Please review the content (especially tables).");
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content in the SharePoint "
											+ "system and correct it." + exc.Message,
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											hyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: hyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									} //- if(layer2upFeatureID != null)

								// Insert Layer0up if not null
								if(objFeature.CSDdescription != null)
									{
									// Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceFeaturesURI +
											currentHyperlinkViewEditURI +
											objFeature.IDsp;
										}
									else
										currentListURI = "";

									//- Check for Colour coding Layers and add if necessary
									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objFeature.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objFeature.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									try
										{
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 4,
											parHTML2Decode: HTMLdecoder.CleanHTML(objFeature.CSDdescription, parClientName),
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
										this.LogError("Error: The Service Feature ID: " + objFeature.IDsp
											+ " contains an error in CSD Description's Enhance Rich Text. "
											+ "Please review the content (especially tables).");
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content in the SharePoint "
											+ "system and correct it." + exc.Message,
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											hyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: hyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}
								} //- if(this.Service_Feature_Description)
							}
						else
							{
							//- If the entry is not found - write an error in the document and record an error in the error log.
							this.LogError("Error: The Service Feature ID " + node.NodeID
								+ " doesn't exist in SharePoint and couldn't be retrieved.");
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "Error: Service Feature " + node.NodeID + " is missing.",
								parIsNewSection: false,
								parIsError: true);
							objParagraph.Append(objRun);
							}
						drmHeading = false;
						break;

					//+ Deliverables, Reports, Meetings
					case enumNodeTypes.FED:  //-| Deliverable associated with Feature
					case enumNodeTypes.FER:  //-| Report deliverable associated with Feature
					case enumNodeTypes.FEM:  //-| Meeting deliverable associated with Feature

						if(!this.DRM_Heading)
							break;

						//- This that the Deliverables, Reports and Meetings Heading appear only once after a Feature.
						if(drmHeading == false)
							{
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: Properties.AppResources.Document_DeliverableReportsMeetings_Heading);
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);
							drmHeading = true;
							}

						//-| Get the Layer0up entry from the Database
						objDeliverable = Deliverable.Read(parIDsp: node.NodeID);
						if (objDeliverable != null)
							{
							Console.Write("\t + {0} - {1}", objDeliverable.IDsp, objDeliverable.Title);

							//-| Insert the Deliverable CSD Heading
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
							objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.CSDheading);
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);

							//-|Check if the Deliverable Layer0up has Content Layers and Content Predecessors
							Console.Write("\n\t\t + Deliverable Layer0up..: {0} - {1}", objDeliverable.IDsp, objDeliverable.Title);
							if(objDeliverable.ContentPredecessorDeliverableIDsp == null)
								{
								layer1upDeliverableID = null;
								}
							else
								{
								//-| Get the entry from the Database
								objDeliverableLayer1up = Deliverable.Read(parIDsp: Convert.ToInt16(objDeliverable.ContentPredecessorDeliverableIDsp));
								if (objDeliverableLayer1up != null)
									{
									Console.Write("\n\t\t + Deliverable Layer1up..: {0} - {1}", objDeliverableLayer1up.IDsp, objDeliverableLayer1up.Title);
									layer1upDeliverableID = objDeliverable.ContentPredecessorDeliverableIDsp;
									}
								else
									{
									layer1upDeliverableID = null;
									}
								}

							//+ Insert Deliverable Description
							if(this.DRM_Description)
								{
								//-| Insert Layer1up if present and not null
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upDeliverableID != null
								&& objDeliverableLayer1up.CSDdescription != null)
									{
									//- Check for Colour coding Layers and add if necessary
									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objDeliverableLayer1up.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objDeliverableLayer1up.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									if(this.HyperlinkEdit || this.HyperlinkView)
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
											parDocumentLevel: 6,
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
										this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
											+ " contains an error in CSD Description's Enhance Rich Text. "
											+ "Please review the content (especially tables).");
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content in the SharePoint "
											+ "system and correct it." + exc.Message,
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											hyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: hyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}//- if(objDeliverable.Layer1up.Layer1up.CSDdescription != null)
									} //- if(layer2upDeliverableID != null)

								// Insert Layer0up if present and not null
								if(objDeliverable.CSDdescription != null)
									{
									//- Check for Colour coding Layers and add if necessary
									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objDeliverable.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objDeliverable.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									if(this.HyperlinkEdit || this.HyperlinkView)
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
											parDocumentLevel: 6,
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
										// A Table content error occurred, record it in the error log.
										this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
											+ " contains an error in CSD Description's Enhance Rich Text. "
											+ "Please review the content (especially tables).");
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content in the SharePoint "
											+ "system and correct it." + exc.Message,
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											hyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: hyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										} // catch...
									} //- if(objDeliverable.CSDdescription != null)
								}

							//+ Insert the Deliverable Inputs
							if(this.DRM_Inputs)
								{
								if(objDeliverable.Inputs != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.Inputs != null))
									{
									//- Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									//- Insert Layer1up if present and not null
									if(layer1upDeliverableID != null
									&& this.PresentationMode == enumPresentationMode.Layered
									&& objDeliverableLayer1up.Inputs != null)
										{
										//- Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverableLayer1up.IDsp;
											}
										else
											currentListURI = "";

										//- Check for Colour coding Layers and add if necessary
										currentContentLayer = "None";
										if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
											{
											if(objDeliverableLayer1up.ContentLayer.Contains("1"))
												currentContentLayer = "Layer1";
											else if(objDeliverableLayer1up.ContentLayer.Contains("2"))
												currentContentLayer = "Layer2";

											try
												{
												objHTMLdecoder.DecodeHTML(parClientName: parClientName,
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverableLayer1up.Inputs, parClientName),
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
												// A Table content error occurred, record it in the
												// error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
													+ " contains an error in Input's Enhance Rich Text. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(this.HyperlinkEdit || this.HyperlinkView)
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: hyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: currentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												} //- catch...
											} //- if(objDeliverableLayer1up.Inputs
										} //- if(layer1upDeliverableID != null)

									// Insert Layer0up if not null
									if(objDeliverable.Inputs != null)
										{
										// Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverable.IDsp;
											}
										else
											currentListURI = "";

										//- Check for Colour coding Layers and add if necessary
										currentContentLayer = "None";
										if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
											{
											if(objDeliverable.ContentLayer.Contains("1"))
												currentContentLayer = "Layer1";
											else if(objDeliverable.ContentLayer.Contains("2"))
												currentContentLayer = "Layer2";
											}

										try
											{
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 7,
												parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverable.Inputs, parClientName),
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
											this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in Input's Enhance Rich Text. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it." + exc.Message,
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: hyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											} //- catch
										} //- if(recDeliverable.Inputs != null)
									} //- if(objDeliverable.Inputs  &&...)
								} //- if(this.DRM_Inputs)

							//+ Insert the Deliverable Outputs
							if(this.DRM_Outputs)
								{
								if(objDeliverable.Outputs != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.Outputs != null))
									{
									//- Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableOutputs_Heading_Text);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									//- Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null
									&& objDeliverableLayer1up.Outputs != null)
										{
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverableLayer1up.IDsp;
											}
										else
											currentListURI = "";

										//- Check for Colour coding Layers and add if necessary
										currentContentLayer = "None";
										if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
											{
											if(objDeliverableLayer1up.ContentLayer.Contains("1"))
												currentContentLayer = "Layer1";
											else if(objDeliverableLayer1up.ContentLayer.Contains("2"))
												currentContentLayer = "Layer2";
											}

										try
											{
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 7,
												parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverableLayer1up.Outputs, parClientName),
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
											this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
												+ " contains an error in Output's Enhance Rich Text. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it." + exc.Message,
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: hyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											} //- catch(...
										} //- if(layer2upDeliverableID != null)

									//- Insert Layer0up if not null
									if(objDeliverable.Outputs != null)
										{
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverable.IDsp;
											}
										else
											currentListURI = "";

										//- Check for Colour coding Layers and add if necessary
										currentContentLayer = "None";
										if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
											{
											if(objDeliverable.ContentLayer.Contains("1"))
												{ currentContentLayer = "Layer1"; }
											else if(objDeliverable.ContentLayer.Contains("2"))
												{ currentContentLayer = "Layer2"; }
											}

										try
											{
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 7,
												parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverable.Outputs, parClientName),
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
											this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in Output's Enhance Rich Text. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it." + exc.Message,
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: hyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											} //- catch(...
										} //- if(objDeliverable.Outputs != null)
									} //- if(objDeliverables.Outputs !== null &&)
								} //- if(this.DRM_Outputs)

							//+ Insert the Deliverable DD's Obligations
							if(this.DDS_DRM_Obligations)
								{
								if(objDeliverable.DDobligations != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.DDobligations != null))
									{
									// Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									//- Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null
									&& objDeliverableLayer1up.DDobligations != null)
										{
										//- Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverableLayer1up.IDsp;
											}
										else
											currentListURI = "";

										//- Check for Colour coding Layers and add if necessary
										currentContentLayer = "None";
										if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
											{
											if(objDeliverableLayer1up.ContentLayer.Contains("1"))
												currentContentLayer = "Layer1";
											else if(objDeliverableLayer1up.ContentLayer.Contains("2"))
												currentContentLayer = "Layer2";
											}

										try
											{
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 7,
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
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
												+ " contains an error in DD's Obligations' Enhance Rich Text. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it." + exc.Message,
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: hyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											} //- catch(...
										} //- if(layer2upDeliverableID != null)

									// Insert Layer0up if not null
									if(objDeliverable.DDobligations != null)
										{
										// Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverable.IDsp;
											}
										else
											currentListURI = "";

										//- Check for Colour coding Layers and add if necessary
										currentContentLayer = "None";
										if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
											{
											if(objDeliverable.ContentLayer.Contains("1"))
												currentContentLayer = "Layer1";
											else if(objDeliverable.ContentLayer.Contains("2"))
												currentContentLayer = "Layer2";
											}

										try
											{
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 7,
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
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in DD's Obligations' Enhance Rich Text. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it." + exc.Message,
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: hyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											} //- catch(...
										} //- if(objDeliverable.DDobligations != null)
									} //- if(objDeliverable.DDoblidations != null &&)
								} //- if(this.DDs_DRM_Obligations)

							//+ Insert the Client Responsibilities
							if(this.Clients_DRM_Responsibilities)
								{
								if(objDeliverable.ClientResponsibilities != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.ClientResponsibilities != null))
									{
									// Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									//- Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null
									&& objDeliverableLayer1up.ClientResponsibilities != null)
										{
										//- Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverableLayer1up.IDsp;
											}
										else
											currentListURI = "";

										//- Check for Colour coding Layers and add if necessary
										currentContentLayer = "None";
										if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
											{
											if(objDeliverableLayer1up.ContentLayer.Contains("1"))
												currentContentLayer = "Layer1";
											else if(objDeliverableLayer1up.ContentLayer.Contains("2"))
												currentContentLayer = "Layer2";
											}

										try
											{
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 7,
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
											this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
												+ " contains an error in Client Responsibilities Enhance Rich Text. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it." + exc.Message,
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: hyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											} //- catch
										} //- if(layer1upDeliverableID != null)

									// Insert Layer0up if not null
									if(objDeliverable.ClientResponsibilities != null)
										{
										//- Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverable.IDsp;
											}
										else
											currentListURI = "";

										//- Check for Colour coding Layers and add if necessary
										currentContentLayer = "None";
										if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
											{
											if(objDeliverable.ContentLayer.Contains("1"))
												currentContentLayer = "Layer1";
											else if(objDeliverable.ContentLayer.Contains("2"))
												currentContentLayer = "Layer2";
											}

										try
											{
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 7,
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
											this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in Client Responsibilities Enhance Rich Text. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it." + exc.Message,
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: hyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											} // catch
										} // if(objDeliverable.ClientResponsibilities != null)
									}
								} //if(this.Clients_DRM_Responsibilities)

							//+ Insert the Deliverable Exclusions
							if(this.DRM_Exclusions)
								{
								if(objDeliverable.Exclusions != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.Exclusions != null))
									{
									//- Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									//- Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null
									&& objDeliverableLayer1up.Exclusions != null)
										{
										//- Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverableLayer1up.IDsp;
											}
										else
											currentListURI = "";

										//- Check for Colour coding Layers and add if necessary
										currentContentLayer = "None";
										if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
											{
											if(objDeliverableLayer1up.ContentLayer.Contains("1"))
												currentContentLayer = "Layer1";
											else if(objDeliverableLayer1up.ContentLayer.Contains("2"))
												currentContentLayer = "Layer2";
											}

										try
											{
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 7,
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
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
												+ " contains an error in Exclusions Enhance Rich Text. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it." + exc.Message,
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: hyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											} //- catch...
										} //- if(layer2upDeliverableID != null)

									// Insert Layer0up if not null
									if(objDeliverable.Exclusions != null)
										{
										// Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverable.IDsp;
											}
										else
											currentListURI = "";

										//- Check for Colour coding Layers and add if necessary
										currentContentLayer = "None";
										if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
											{
											if(objDeliverable.ContentLayer.Contains("1"))
												currentContentLayer = "Layer1";
											else if(objDeliverable.ContentLayer.Contains("2"))
												currentContentLayer = "Layer2";
											}

										try
											{
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 7,
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
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in Exclusions Enhance Rich Text. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it." + exc.Message,
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: hyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											} //- catch...
										} //- if(objDeliverable.Exclusions != null)
									} //- if(objDeliverable.Exclusions != null &&)
								} //- if(this.DRMe_Exclusions)

							//+ Insert the Governance Controls
							if(this.DRM_Governance_Controls)
								{
								if(objDeliverable.GovernanceControls != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.GovernanceControls != null))
									{
									//- Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableGovernanceControls_Heading_Text);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									// Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null
									&& objDeliverableLayer1up.GovernanceControls != null)
										{
										// Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverableLayer1up.IDsp;
											}
										else
											currentListURI = "";

										//- Check for Colour coding Layers and add if necessary
										currentContentLayer = "None";
										if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
											{
											if(objDeliverableLayer1up.ContentLayer.Contains("1"))
												currentContentLayer = "Layer1";
											else if(objDeliverableLayer1up.ContentLayer.Contains("2"))
												currentContentLayer = "Layer2";
											}

										try
											{
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 7,
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
											this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in Exclusions Enhance Rich Text. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it." + exc.Message,
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: hyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											} //- catch...
										} //- if(objDeliverableLayer1up.GovernanceControls != null)

									//- Insert Layer0up if not null
									if(objDeliverable.GovernanceControls != null)
										{
										//- Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverable.IDsp;
											}
										else
											currentListURI = "";

										//- Check for Colour coding Layers and add if necessary
										currentContentLayer = "None";
										if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
											{
											if(objDeliverable.ContentLayer.Contains("1"))
												currentContentLayer = "Layer1";
											else if(objDeliverable.ContentLayer.Contains("2"))
												currentContentLayer = "Layer2";
											}

										try
											{
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 7,
												parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverable.GovernanceControls, parClientName),
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
											this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in Governance Controls Enhance Rich Text. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it." + exc.Message,
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: hyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											} //- catch(InvalidContentFormatException exc)
										} //- if(objDeliverable.GovernanceControls != null
									} //- if(objDeliverable.GovernanceControls != null &&...
								} //- if(this.DRM_GovernanceControls)

							//+ Check if there are any Glossary Terms or Acronyms associated with the Deliverable(s).
							if(this.Acronyms_Glossary_of_Terms_Section)
								{
								// if there are GlossaryAndAcronyms to add from layer0up
								if(objDeliverable.GlossaryAndAcronyms != null && objDeliverable.GlossaryAndAcronyms != null)
									{
									foreach(var entry in objDeliverable.GlossaryAndAcronyms)
										{
										if(this.ListGlossaryAndAcronyms.Contains(entry) != true)
											ListGlossaryAndAcronyms.Add(entry);
										}
									}
								// if there are GlossaryAndAcronyms to add from layer1up
								if(objDeliverableLayer1up.GlossaryAndAcronyms != null && objDeliverableLayer1up.GlossaryAndAcronyms != null)
									{
									foreach(var entry in objDeliverableLayer1up.GlossaryAndAcronyms)
										{
										if(this.ListGlossaryAndAcronyms.Contains(entry) != true)
											ListGlossaryAndAcronyms.Add(entry);
										}
									}
								} // if(this.Acronyms_Glossary_of_Terms_Section)
							} //- if(parDataset.dsDeliverables.TryGetValue(....
						else
							{
							// If the entry is not found - write an error in the document and record
							// an error in the error log.
							this.LogError("Error: The Deliverable ID " + node.NodeID
								+ " doesn't exist in SharePoint and couldn't be retrieved.");
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "Error: Deliverable " + node.NodeID + " is missing.",
								parIsNewSection: false,
								parIsError: true);
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);
							}
						break;

					//+ Service Levels
					case enumNodeTypes.FSL:
						if(!this.Service_Level_Heading)
							break;

						//-| Populate the Service Level Heading
						objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
						objRun = oxmlDocument.Construct_RunText(
							parText2Write: Properties.AppResources.Document_ServiceLevels_Heading_Text);
						objParagraph.Append(objRun);
						objBody.Append(objParagraph);

						//-| Check if the user specified to include the Deliverable Description
						if(this.Service_Level_Commitments_Table)
							break;

						//-| Prepare the data to insert into the Service Level Table
						objDeliverableServiceLevel = DeliverableServiceLevel.Read(parIDsp: node.NodeID);
						if (objDeliverableServiceLevel != null)
							{
							Console.WriteLine("\t\t + Deliverable ServiceLevel: {0} - {1}", objDeliverableServiceLevel.IDsp,
								objDeliverableServiceLevel.Title);

							//-| Get the Service Level entry from the Database
							if(objDeliverableServiceLevel.AssociatedServiceLevelIDsp != null)
								{
								objServiceLevel = ServiceLevel.Read(parIDsp: Convert.ToInt16(objDeliverableServiceLevel.AssociatedServiceLevelIDsp));
								if (objServiceLevel != null)
									{
									Console.WriteLine("\t\t\t + Service Level: {0} - {1}", objServiceLevel.IDsp, objServiceLevel.Title);
									Console.WriteLine("\t\t\t + Service Hour.: {0}", objServiceLevel.ServiceHours);

									//-| Insert the Service Level ISD Description
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
									objRun = oxmlDocument.Construct_RunText(parText2Write: objServiceLevel.CSDheading);
									// Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_ServiceLevelsURI +
												currentHyperlinkViewEditURI + objServiceLevel.IDsp,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									List<string> listErrorMessagesParameter = this.ErrorMessages;
									// Populate the Service Level Table
									objServiceLevelTable = CommonProcedures.BuildSLAtable(
										parMainDocumentPart: ref objMainDocumentPart,
										parClientName: parClientName,
										parServiceLevelID: objServiceLevel.IDsp,
										parWidthColumn1: Convert.ToInt16(this.PageWith * 0.30),
										parWidthColumn2: Convert.ToInt16(this.PageWith * 0.70),
										parMeasurement: objServiceLevel.Measurement,
										parMeasureMentInterval: objServiceLevel.MeasurementInterval,
										parReportingInterval: objServiceLevel.ReportingInterval,
										parServiceHours: objServiceLevel.ServiceHours,
										parCalculationMethod: objServiceLevel.CalculationMethod,
										parCalculationFormula: objServiceLevel.CalculationFormula,
										parThresholds: objServiceLevel.PerformanceThresholds,
										parTargets: objServiceLevel.PerformanceTargets,
										parBasicServiceLevelConditions: objServiceLevel.BasicConditions,
										parAdditionalServiceLevelConditions: objDeliverableServiceLevel.AdditionalConditions,
										parErrorMessages: ref listErrorMessagesParameter,
										parNumberingCounter: ref numberingCounter);

									if(listErrorMessagesParameter.Count != this.ErrorMessages.Count)
										this.ErrorMessages = listErrorMessagesParameter;

									objBody.Append(objServiceLevelTable);
									} // if(parDatabase.dsServiceLevels.TryGetValue(
								} // if(objDeliverableServiceLevel.AssociatedServiceLevelID != null)
							else
								{
								// If the entry is not found - write an error in the document and
								// record an error in error log.
								this.LogError("Error: The DeliverableServiceLevel ID " + node.NodeID
									+ " doesn't exist in SharePoint and it couldn't be retrieved.");
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: "Error: DeliverableServiceLevel: " + node.NodeID + " is missing.",
									parIsNewSection: false,
									parIsError: true);
								if(this.HyperlinkEdit || this.HyperlinkView)
									{
									hyperlinkCounter += 1;
									Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
										parMainDocumentPart: ref objMainDocumentPart,
										parImageRelationshipId: hyperlinkImageRelationshipID,
										parHyperlinkID: hyperlinkCounter,
										parClickLinkURL: currentListURI);
									objRun.Append(objDrawing);
									}
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);
								break;
								}
							} //if(parDatabase.dsDeliverableServiceLevels.TryGetValue(
						else
							{
							// If the entry is not found - write an error in the document and record
							// an error in the error log.
							this.LogError("Error: The DeliverableServiceLevel ID " + node.NodeID
								+ " doesn't exist in SharePoint and it couldn't be retrieved.");
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "Error: DeliverableServiceLevel: " + node.NodeID + " is missing.",
								parIsNewSection: false,
								parIsError: true);
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);
							break;
							}
						break;
						} //- switch (node.NodeType)
					} //- foreach(Hierarchy node in this.SelectedNodes)
			  //++Glossary of Terms and Acronym Section
Process_Glossary_and_Acronyms:
				if(this.Acronyms_Glossary_of_Terms_Section && this.ListGlossaryAndAcronyms.Count == 0)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_GlossaryAndAcronymSection_HeadingText,
						parIsNewSection: true);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					//- Insert a blank paragrpah
					objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: " ");
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);
					//+Insert Terms and Acronyms Table
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
						} //- if(this.TermAndAcronymList.Count > 0)
					} //- if (this.Acronyms)

				//++ Insert the Document Generation Error Section
				if(this.ErrorMessages.Count > 0)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_Error_Section_Heading,
						parIsNewSection: true);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Error_Heading);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					foreach(var errorMessageEntry in this.ErrorMessages)
						{
						objParagraph = oxmlDocument.Construct_Error(errorMessageEntry);
						objBody.Append(objParagraph);
						}
					}

				//+Validate the document with OpenXML validator
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
				if(this.UploadDoc(parRequestingUserID: parRequestingUserID, parSDDPdatacontext: parSDDPdatacontext))
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

			//++ Handle Exceptions
			//+ Non-Contentspecified Exception
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
				}

			Console.WriteLine("\t\t End of the generation of {0}", this.DocumentType);
			//- Delete the file from the Documents Directory
			if(File.Exists(path: this.LocalDocumentURI))
				File.Delete(path: this.LocalDocumentURI);
			} //- end of Generate Method
		} // end of CSD_inline DRM class
	}