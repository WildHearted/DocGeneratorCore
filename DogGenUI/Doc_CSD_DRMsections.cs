using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocGeneratorCore.Database.Classes;
using DocGeneratorCore.SDDPServiceReference;

namespace DocGeneratorCore
	{
	/// <summary>
	///      This class represent the Client Service Description (CSD) with sperate DRM (Deliverable
	///      Report Meeting) sections It inherits from the DRM Sections Class.
	/// </summary>
	internal class CSD_Document_DRM_Sections:External_DRM_Sections
		{
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
							case 99:
							this.Introductory_Section = true;
							break;

							case 100:
							this.Introduction = true;
							break;

							case 101:
							this.Executive_Summary = true;
							break;

							case 102:
							this.Service_Portfolio_Section = true;
							break;

							case 103:
							this.Service_Portfolio_Description = true;
							break;

							case 104:
							this.Service_Family_Heading = true;
							break;

							case 105:
							this.Service_Family_Description = true;
							break;

							case 106:
							this.Service_Product_Heading = true;
							break;

							case 107:
							this.Service_Product_Description = true;
							break;

							case 108:
							this.Service_Feature_Heading = true;
							break;

							case 109:
							this.Service_Feature_Description = true;
							break;

							case 110:
							this.Deliverables_Reports_Meetings = true;
							break;

							case 111:
							this.DRM_Heading = true;
							break;

							case 112:
							this.DRM_Summary = true;
							break;

							case 113:
							this.Service_Levels = true;
							break;

							case 114:
							this.Service_Level_Heading = true;
							break;

							case 115:
							this.Service_Level_Commitments_Table = true;
							break;

							case 116:
							this.DRM_Section = true;
							break;

							case 117:
							this.Deliverables = true;
							break;

							case 118:
							this.Deliverable_Heading = true;
							break;

							case 119:
							this.Deliverable_Description = true;
							break;

							case 120:
							this.DDs_Deliverable_Obligations = true;
							break;

							case 121:
							this.Clients_Deliverable_Responsibilities = true;
							break;

							case 122:
							this.Deliverable_Exclusions = true;
							break;

							case 123:
							this.Deliverable_Governance_Controls = true;
							break;

							case 124:
							this.Reports = true;
							break;

							case 125:
							this.Report_Heading = true;
							break;

							case 126:
							this.Report_Description = true;
							break;

							case 127:
							this.DDs_Report_Obligations = true;
							break;

							case 128:
							this.Clients_Report_Responsibilities = true;
							break;

							case 129:
							this.Report_Exclusions = true;
							break;

							case 130:
							this.Report_Governance_Controls = true;
							break;

							case 131:
							this.Meetings = true;
							break;

							case 132:
							this.Meeting_Heading = true;
							break;

							case 133:
							this.Meeting_Description = true;
							break;

							case 134:
							this.DDs_Meeting_Obligations = true;
							break;

							case 135:
							this.Clients_Meeting_Responsibilities = true;
							break;

							case 136:
							this.Meeting_Exclusions = true;
							break;

							case 137:
							this.Meeting_Governance_Controls = true;
							break;

							case 138:
							this.Service_Level_Section = true;
							break;

							case 139:
							this.Service_Level_Heading = true;
							break;

							case 140:
							this.Service_Level_Commitments_Table = true;
							break;

							case 141:
							this.Acronyms_Glossary_of_Terms_Section = true;
							break;

							case 142:
							this.Acronyms = true;
							break;

							case 143:
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
			Dictionary<int, string> dictDeliverables = new Dictionary<int, string>();
			Dictionary<int, string> dictReports = new Dictionary<int, string>();
			Dictionary<int, string> dictMeetings = new Dictionary<int, string>();
			Dictionary<int, string> dictSLAs = new Dictionary<int, string>();
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
					documentCollection_HyperlinkURL = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion
						+ Properties.AppResources.List_DocumentCollectionLibraryURI
						+ Properties.AppResources.EditFormURI + this.DocumentCollectionID;
					currentHyperlinkViewEditURI = Properties.AppResources.EditFormURI;
					}
				if(this.HyperlinkView)
					{
					documentCollection_HyperlinkURL = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion
						+ Properties.AppResources.List_DocumentCollectionLibraryURI
						+ Properties.AppResources.DisplayFormURI + this.DocumentCollectionID;
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

				//+ Create and open the new Document
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
						Console.WriteLine("\t\t Page width x height: {0} x {1} twips", this.PageWith, this.PageHeight);
						}
					if(objPageMargin != null)
						{
						if(objPageMargin.Left != null)
							{
							this.PageWith -= objPageMargin.Left;
							Console.WriteLine("\t\t\t - Left Margin..: {0} twips", objPageMargin.Left);
							}
						if(objPageMargin.Right != null)
							{
							this.PageWith -= objPageMargin.Right;
							Console.WriteLine("\t\t\t - Right Margin.: {0} twips", objPageMargin.Right);
							}
						if(objPageMargin.Top != null)
							{
							string tempTop = objPageMargin.Top.ToString();
							Console.WriteLine("\t\t\t - Top Margin...: {0} twips", tempTop);
							this.PageHeight -= Convert.ToUInt32(tempTop);
							}
						if(objPageMargin.Bottom != null)
							{
							string tempBottom = objPageMargin.Bottom.ToString();
							Console.WriteLine("\t\t\t - Bottom Margin: {0} twips", tempBottom);
							this.PageHeight -= Convert.ToUInt32(tempBottom);
							}
						}
					}
				// Subtract the Table/Image Left indentation value from the Page width to ensure the
				// table/image fits in the available space.
				this.PageWith -= Convert.ToUInt16(Properties.AppResources.Document_Table_Left_Indent);
				Console.WriteLine("\t\t Effective pageWidth x pageHeight.: {0} x {1} twips", this.PageWith, this.PageHeight);

				// Check whether Hyperlinks need to be included and add the image to the Document Body
				if(this.HyperlinkEdit || this.HyperlinkView)
					{
					//Insert and embed the hyperlink image in the document and keep the Image's Relationship ID in a variable for repeated use
					hyperlinkImageRelationshipID = oxmlDocument.Insert_HyperlinkImage(
						parMainDocumentPart: ref objMainDocumentPart,
						parSDDPdatacontext: parSDDPdatacontext);
					}

				//+ Define the objects to be used in the construction of the document
				ServicePortfolio objPortfolio = new ServicePortfolio();
				ServiceFamily objFamily = new ServiceFamily();
				ServiceProduct objProduct = new ServiceProduct();
				ServiceFeature objFeature = new ServiceFeature();
				ServiceFeature objFeatureLayer1up = new ServiceFeature();
				Deliverable objDeliverable = new Deliverable();
				Deliverable objDeliverableLayer1up = new Deliverable();
				DeliverableActivity objDeliverableActivity = new DeliverableActivity();
				DeliverableServiceLevel objDeliverableServiceLevel = new DeliverableServiceLevel();
				ServiceLevel objServiceLevel = new ServiceLevel();
				Activity objActivity = new Activity();

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

				//++Introductory Section
				if(this.Introductory_Section)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_IntruductorySection_HeadingText,
						parIsNewSection: true);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);
					}

				//+Insert the Introduction
				if(this.Introduction)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Introduction_HeadingText);
					// Check if a hyperlink must be inserted
					if(documentCollection_HyperlinkURL != "")
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
							objHTMLdecoder.DecodeHTML(
								parClientName: parClientName,
								parMainDocumentPart: ref objMainDocumentPart,
								parDocumentLevel: 2,
								parHTML2Decode: HTMLdecoder.CleanHTML(this.IntroductionRichText, parClientName),
								parTableCaptionCounter: ref tableCaptionCounter,
								parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
								parPictureNo: ref iPictureNo,
								parHyperlinkID: ref hyperlinkCounter,
								parPageHeightDxa: this.PageHeight,
								parPageWidthDxa: this.PageWith, 
								parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
							}
						catch(InvalidContentFormatException exc)
							{
							Console.WriteLine("\n\nException occurred: {0}", exc.Message);
							// A Table content error occurred, record it in the error log.
							this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
								+ " contains an error in Introduction's Enhance Rich Text. Please review the content (especially tables).");
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "A content error occurred at this position and valid content could " +
								"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it. " + exc.Message,
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
								objRun.Append(objDrawing);
								}
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);
							}
						}
					}
				
				//+Insert the Executive Summary
				if(this.Executive_Summary)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_ExecutiveSummary_HeadingText);
					// Check if a hyperlink must be inserted
					if(documentCollection_HyperlinkURL != "")
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
								parPageWidthDxa: this.PageWith, 
								parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
							}
						catch(InvalidContentFormatException exc)
							{
							Console.WriteLine("\n\nException occurred: {0}", exc.Message);
							// A Table content error occurred, record it in the error log.
							this.LogError("Error: The Docuement Collection ID: " + this.DocumentCollectionID
								+ " contains an error in its Executive summary Enhance Rich Text column. Please review the content (especially tables).");
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "A content error occurred at this position and valid content could " +
								"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it. " + exc.Message,
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
								objRun.Append(objDrawing);
								}
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);
							}
						}
					}
				
				//++Insert the user selected content

				if(this.SelectedNodes.Count <= 0)
					goto Process_Glossary_and_Acronyms;
				foreach(Hierarchy node in this.SelectedNodes)
					{
					Console.Write("\nNode: SEQ:{0} LeveL:{1} NodeType:{2} NodeID:{3}", node.Sequence, node.Level, node.NodeType, node.NodeID);

					switch(node.NodeType)
						{
						//+Service Framework & Service Porfolio
						case enumNodeTypes.FRA:  //-| Service Framework
						case enumNodeTypes.POR:  //-| Service Portfolio
							{
							if(!this.Service_Portfolio_Section)
								{
								break;
								}

							objPortfolio = ServicePortfolio.Read(parIDsp: node.NodeID);
							if (objPortfolio != null)
								{
								Console.Write("\t + {0} - {1}", objPortfolio.IDsp, objPortfolio.Title);

								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: objPortfolio.CSDheading,
									parIsNewSection: true);
								//-| Check if a hyperlink must be inserted
								if(documentCollection_HyperlinkURL != "")
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
								if(this.Service_Portfolio_Description)
									{
									if(objPortfolio.CSDdescription != null)
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
											Console.WriteLine("\n\nException occurred: {0}\n", exc.Message);
											//-| A Table content error occurred, record it in the error log.
											this.LogError("Error: The Service Portfolio ID: " + node.NodeID
												+ " contains an error in one of its Enhance Rich Text columns. Please review "
												+ " the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ " system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}
									}
								}
							else
								{
								//-| If the entry is not found - write an error in the document and
								//-| record an error in the error log.
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
							}

						//+Service Family
						case enumNodeTypes.FAM:
							{
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
								// Check if a hyperlink must be inserted
								if(documentCollection_HyperlinkURL != "")
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
								//-|Insert the Service Family Description
								if(this.Service_Family_Description)
									{
									if(objFamily.CSDdescription != null)
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
											this.LogError("Error: The Service Family ID: " + node.NodeID
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}
									}
								}
							else
								{
								//-| If the entry is not found - write an error in the document and
								//-| record an error in the error log.
								this.LogError("Error: The Service Family ID " + node.NodeID
									+ " doesn't exist in SharePoint and couldn't be retrieved.");
								objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: "Error: Service Family " + node.NodeID + " is missing.",
									parIsNewSection: false,
									parIsError: true);
								objParagraph.Append(objRun);
								}

							break;
							}

						//+Service Products
						case enumNodeTypes.PRO:
							{
							if(!this.Service_Product_Heading)
								{ break; }

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
								if(documentCollection_HyperlinkURL != "")
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
								//+Insert the Service Product Description
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
										// A Table content error occurred, record it in the error log.
										this.LogError("Error: The Service Product ID: " + node.NodeID
											+ " contains an error in one of its Enhance Rich Text columns. "
											+ "Please review the content (especially tables).");
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content in the SharePoint "
											+ "system and correct it. " + exc.Message,
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
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}
								}
							else
								{//- If the entry is not found - write an error in the document and
								this.LogError("Error: The Service Product ID " + node.NodeID
									+ " doesn't exist in SharePoint and couldn't be retrieved.");
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: "Error: Service Family " + node.NodeID + " is missing.",
									parIsNewSection: false,
									parIsError: true);
								objParagraph.Append(objRun);
								}

							break;
							}

						//+Service Feature
						case enumNodeTypes.FEA:
							{
							if(!this.Service_Feature_Heading)
								break;

							//-| Get the entry from the Database
							objFeature = ServiceFeature.Read(parIDsp: node.NodeID);
							if (objFeature != null)
								{
								Console.Write("\t + {0} - {1}", objFeature.IDsp, objFeature.Title);

								//+Insert the Service Feature CSD Heading...
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
								objRun = oxmlDocument.Construct_RunText(parText2Write: objFeature.CSDheading);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);

								//Check if the Feature Layer0up has Content Layers and Content Predecessors
								if(objFeature.ContentPredecessorFeatureIDsp == null)
									{
									layer1upFeatureID = null;
									}
								else
									{
									//-| Get the Layer1up entry from the Database
									objFeatureLayer1up = ServiceFeature.Read(parIDsp: Convert.ToInt16(objFeature.ContentPredecessorFeatureIDsp));
									if (objFeatureLayer1up != null)
										{
										layer1upFeatureID = objFeatureLayer1up.IDsp;
										}
									else
										{
										layer1upFeatureID = null;
										}
									}

								//+Insert the Service Feature Description
								if(!this.Service_Feature_Description)
									break;

								// Insert Layer1up if present and not null
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upFeatureID != null
								&& objFeatureLayer1up.CSDdescription != null)
									{
									//- Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceFeaturesURI +
											currentHyperlinkViewEditURI +
											objFeatureLayer1up.IDsp;
										}
									else
										currentListURI = "";

									//- Set the Content Layer Colour Coding
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
											+ " contains an error in one of its Enhance Rich Text columns. "
											+ "Please review the content (especially tables).");
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid "
											+ "content could not be interpreted and inserted here. "
											+ "Please review the content in the SharePoint system and correct it. Detail of Error: "
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
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}

								// Insert Layer 0up if not null
								if(objFeature.CSDdescription != null)
									{
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceFeaturesURI +
											currentHyperlinkViewEditURI +
											objFeature.IDsp;
										}
									else
										currentListURI = "";

									//- Set the Content Layer Colour Coding
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
										this.LogError("Error: The Service Feature ID: " + node.NodeID
											+ " contains an error in one of its Enhance Rich Text columns. "
											+ "Please review the content (especially tables).");
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content in the SharePoint "
											+ "system and correct it. " + exc.Message,
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
								// record an error in the error log.
								this.LogError("Error: The Service Feature ID " + node.NodeID
									+ " doesn't exist in SharePoint and couldn't be retrieved.");
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: "Error: Service Feature " + node.NodeID + " is missing.",
									parIsNewSection: false,
									parIsError: true);
								objParagraph.Append(objRun);
								}
							drmHeading = false;
							break;
							}
						//+Deliverables, Reports & Meetings
						case enumNodeTypes.FED:  //-| Deliverable associated with Feature
						case enumNodeTypes.FER:  //-| Report deliverable associated with Feature
						case enumNodeTypes.FEM:  //-| Meeting deliverable associated with Feature
							{
							if(!this.DRM_Heading)
								break;

							//- Only need to insert this once...
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

								//- Insert the Deliverable/Report/Meeting CSD Heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
								objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.CSDheading);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);

								//+ Add the deliverable/report/meeting to the relevant Dictionary to include in DRM section
								if(node.NodeType == enumNodeTypes.FED) //+ Deliverable
									{
									if(dictDeliverables.ContainsKey(objDeliverable.IDsp) != true)
										dictDeliverables.Add(objDeliverable.IDsp, objDeliverable.CSDheading);
									}
								else if(node.NodeType == enumNodeTypes.FER) //+ Report
									{
									if(dictReports.ContainsKey(objDeliverable.IDsp) != true)
										dictReports.Add(objDeliverable.IDsp, objDeliverable.CSDheading);
									}
								else if(node.NodeType == enumNodeTypes.FEM) //+ Meeting
									{
									if(dictMeetings.ContainsKey(objDeliverable.IDsp) != true)
										dictMeetings.Add(objDeliverable.IDsp, objDeliverable.CSDheading);
									}

								//-|Check if the Deliverable Layer0up is Layered
								Console.Write("\n\t\t + Deliverable Layer 0..: {0} - {1}", objDeliverable.IDsp, objDeliverable.Title);
								if(objDeliverable.ContentPredecessorDeliverableIDsp == null)
									{
									layer1upDeliverableID = null;
									}
								else
									{
									//-| Get the Layer1up entry from the Database
									objDeliverableLayer1up = Deliverable.Read(parIDsp: Convert.ToInt16(objDeliverable.ContentPredecessorDeliverableIDsp));
									if (objDeliverableLayer1up != null)
										{
										layer1upDeliverableID = objDeliverableLayer1up.IDsp;
										}
									else
										{
										layer1upDeliverableID = null;
										}
									}

								//+ Insert the Deliverable Summary
								if(!this.DRM_Summary)
									break;

								// Insert Layer1up if present and required
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upDeliverableID != null
								&& objDeliverableLayer1up.CSDsummary != null)
									{
									//- Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objFeatureLayer1up.IDsp;
										}
									else
										currentListURI = "";

									//-Set the Content Layer Colour Coding
									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objFeatureLayer1up.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objFeatureLayer1up.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
									objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverableLayer1up.CSDsummary, parContentLayer: currentContentLayer);

									//-Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_DeliverablesURI +
											currentHyperlinkViewEditURI + objDeliverableLayer1up.IDsp;

										Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: currentListURI,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									else
										{
										currentListURI = "";
										}

									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}

								// Insert Layer0up if present and not null
								if(objDeliverable.CSDsummary != null)
									{
									//- Set the Content Layer Colour Coding
									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objDeliverable.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objDeliverable.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
									objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.CSDsummary,
										parContentLayer: currentContentLayer);

									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_DeliverablesURI +
											currentHyperlinkViewEditURI + objDeliverable.IDsp;

										Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: currentListURI,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									else
										currentListURI = "";

									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									} // if(objDeliverable.CSDsummary != null)

								//+ Insert the hyperlink to the bookmark of the Deliverable's relevant position in DRM Section.
								objParagraph = oxmlDocument.Construct_BookmarkHyperlink(
								parBodyTextLevel: 6,
								parBookmarkValue: "Deliverable_" + objDeliverable.IDsp);
								objBody.Append(objParagraph);
								}
							else
								{
								// If the entry is not found - write an error in the document and
								// record an error in the error log.
								this.LogError("Error: The Deliverable ID " + node.NodeID
									+ " doesn't exist in SharePoint and couldn't be retrieved.");
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: "Error: Deliverable " + node.NodeID + " is missing.",
									parIsNewSection: false,
									parIsError: true);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);
								}
							break;
							}

						//+Service Level
						case enumNodeTypes.FSL:  // Service Level associated with Deliverable pertaining to Service Feature
							{
							if(!this.Service_Level_Heading)
								break;

							// Populate the Service Level Heading
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: Properties.AppResources.Document_ServiceLevels_Heading_Text);
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);

							// Check if the user specified to include the Deliverable Description
							if(this.Service_Level_Commitments_Table)
								{
								// Prepare the data which to insert into the Service Level Table
								objDeliverableServiceLevel = DeliverableServiceLevel.Read(parIDsp: node.NodeID);
								if (objDeliverableServiceLevel != null)
									{
									Console.WriteLine("\t\t + Deliverable ServiceLevel: {0} - {1}", objDeliverableServiceLevel.IDsp,
										objDeliverableServiceLevel.Title);

									// Get the Service Level entry from the Database
									if(objDeliverableServiceLevel.AssociatedServiceLevelIDsp != null)
										{
										objServiceLevel = ServiceLevel.Read(parIDsp: Convert.ToInt16(objDeliverableServiceLevel.AssociatedServiceLevelIDsp));
										if (objServiceLevel != null)
											{
											Console.WriteLine("\t\t\t + Service Level: {0} - {1}", objServiceLevel.IDsp, objServiceLevel.Title);
											Console.WriteLine("\t\t\t + Service Hour.: {0}", objServiceLevel.ServiceHours);

											//-| Add the Service Level entry to the Service Level
											//-| Dictionay (list)
											if(dictSLAs.ContainsKey(objServiceLevel.IDsp) != true)
												{
												//-| NOTE: the DeliverableServiceLevel ID is used NOT the ServiceLevel ID.
												dictSLAs.Add(objDeliverableServiceLevel.IDsp, objServiceLevel.ISDheading);
												}
											// Insert the Service Level CSD Description
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
											objRun = oxmlDocument.Construct_RunText(parText2Write: objServiceLevel.CSDheading);
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
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
											}
										else
											{
											// If the entry is not found - write an error in the
											// document and record an error in error log.
											this.LogError("Error: The DeliverableServiceLevel ID " + node.NodeID
												+ " doesn't exist in SharePoint and it couldn't be retrieved.");
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "Error: DeliverableServiceLevel: " + node.NodeID + " is missing.",
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //- if(objDeliverableServiceLevel.AssociatedServiceLevelID != null)
									} //- if(parDatabase.dsDeliverableServiceLevels.TryGetValue(
								} //- if(this.Service_Level_Commitments_Table)
							break;
							} //- case enumNodeTypes.FSL:
						} //- switch (node.NodeType)
					} //- foreach(Hierarchy node in this.SelectedNodes)

				//++ Insert the Deliverable, Report, Meeting (DRM) Section
				if(this.DRM_Section)
					{
					Console.Write("\nGenerating Deliverable, Report, Meeting sections...\n");

					//-| Insert the Deliverables, Reports and Meetings Section
					if(dictDeliverables.Count > 0 || dictReports.Count > 0 || dictMeetings.Count > 0)
						{
						objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
						objRun = oxmlDocument.Construct_RunText(
							parText2Write: Properties.AppResources.Document_DRM_Section_Text,
							parIsNewSection: true);
						objParagraph.Append(objRun);
						objBody.Append(objParagraph);
						}
					else
						goto Process_ServiceLevels;

					//+Deliverables

					if(this.Deliverables && dictDeliverables.Count == 0)
						goto Process_Reports;

					Console.Write("\n\tDeliverables:\n");
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Deliverables_Heading_Text);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					string deliverableBookMark = "Deliverable_";
					//+ Insert the individual Deliverables in the section
					foreach(KeyValuePair<int, string> deliverableItem in dictDeliverables.OrderBy(k => k.Value))
						{
						if(this.Deliverable_Heading)
							{
							objDeliverable = Deliverable.Read(parIDsp: deliverableItem.Key);
							if (objDeliverable != null)
								{
								Console.Write("\n\t + {0} - {1}", objDeliverable.IDsp, objDeliverable.CSDheading);

								//+ Insert the Deliverable's CSD Heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3,
									parBookMark: deliverableBookMark + objDeliverable.IDsp);
								objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.CSDheading);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);

								//Check if the Deliverable's Layer0up has Content Layers and Content Predecessors
								if(objDeliverable.ContentPredecessorDeliverableIDsp == null)
									{
									layer1upDeliverableID = null;
									}
								else
									{
									// Get the entry from the Database
									objDeliverableLayer1up = Deliverable.Read(parIDsp: Convert.ToInt16(objDeliverable.ContentPredecessorDeliverableIDsp));
									if (objDeliverableLayer1up != null)
										{
										layer1upDeliverableID = objDeliverable.IDsp;
										}
									else
										{
										layer1upDeliverableID = null;
										}
									}

								//+Insert the Deliverable CSD Description
								if(this.Deliverable_Description)
									{
									//- Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null
									&& objDeliverableLayer1up.CSDdescription != null)
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
											Console.Write("\n\t\t + Layer1up {0} - {1}", objDeliverableLayer1up.IDsp, objDeliverableLayer1up.Title);
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
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
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}

									// Insert Layer0up if not null
									if(objDeliverable.CSDdescription != null)
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
											Console.Write("\n\t\t + Layer0up{0} - {1}", objDeliverable.IDsp, objDeliverable.Title);
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
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
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}//- if(objDeliverable.CSDdescription)
									} //- if(this.Deliverable_Description)

								//+Insert the Deliverable Inputs
								if(this.Deliverable_Inputs)
									{
									if(objDeliverable.Inputs != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.Inputs != null))
										{
										//- Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null
										&& objDeliverableLayer1up.Inputs != null)
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
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";
											try
												{
												objHTMLdecoder.DecodeHTML(parClientName: parClientName,
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
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
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											}

										//- Insert Layer0up if there are Input content
										if(objDeliverable.Inputs != null)
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
													parDocumentLevel: 4,
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
												// A Table content error occurred, record it in the
												// error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} // if(objDeliverable.Inputs != null)
										} //if(this.Deliverable_Inputs)
									} //if(this.Deliverable_Inputs)

								//+ Insert the Deliverable Outputs
								if(this.Deliverable_Outputs)
									{
									if(objDeliverable.Outputs != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.Outputs != null))
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableOutputs_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null
										&& objDeliverableLayer1up.Outputs != null)
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
													parDocumentLevel: 4,
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
												// A Table content error occurred, record it in the
												// error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											}

										// Insert Layer0up if not null
										if(objDeliverable.Outputs != null)
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
													parDocumentLevel: 4,
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
												// A Table content error occurred, record it in the
												// error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											}
										} //if(recDeliverables.Outputs !== null &&)
									} //if(this.Deliverable_Outputs)

								//+ Insert the Deliverable DD's Obligations
								if(this.DDs_Deliverable_Obligations)
									{
									if(objDeliverable.DDobligations != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.DDobligations != null))
										{
										//- Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null
										&& objDeliverableLayer1up.DDobligations != null)
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
													parDocumentLevel: 4,
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
												// A Table content error occurred, record it in the
												// error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											}

										// Insert Layer0up if not null
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
													parDocumentLevel: 4,
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
												// A Table content error occurred, record it in the
												// error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} //- if(objDeliverable.DDobligations != null)
										} //- if(recDeliverable.DDoblidations != null &&)
									}

								//+ Insert the Client Responsibilities
								if(this.Clients_Deliverable_Responsibilities)
									{
									if(objDeliverable.ClientResponsibilities != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.ClientResponsibilities != null))
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer 1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
											&& layer1upDeliverableID != null
											&& objDeliverableLayer1up.ClientResponsibilities != null)
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
													parDocumentLevel: 4,
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
												// A Table content error occurred, record it in the
												// error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											}
										} // if(... layer1upDeliverableID != null)

									// Insert Layer0up if not null
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
												parDocumentLevel: 4,
												parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverable.ClientResponsibilities,parClientName),
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
											// A Table content error occurred
											this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //- if(recDeliverable.ClientResponsibilities != null)
									} //- if(recDeliverable.ClientResponsibilities != null &&)
								} //- if(this.Clients_Deliverable_Responsibilities)

							//+ Insert the Deliverable Exclusions
							if(this.Deliverable_Exclusions)
								{
								if(objDeliverable.Exclusions != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.Exclusions != null))
									{
									// Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									// Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null
										&& objDeliverableLayer1up.Exclusions != null)
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
												parDocumentLevel: 4,
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
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}

									// Insert Layer0up if not null
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
												parDocumentLevel: 4,
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
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //- if(recDeliverable.Exclusions != null)
									} //- if(recDeliverable.Exclusions != null &&)
								} //- if(this.Deliverable_Exclusions)

							//+Insert the Governance Controls
							if(this.Deliverable_Governance_Controls)
								{
								if(objDeliverable.GovernanceControls != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.GovernanceControls != null))
									{
									// Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
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
												parDocumentLevel: 4,
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
											this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}

									// Insert Layer0up if not null
									if(objDeliverable.GovernanceControls != null)
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
												parDocumentLevel: 4,
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
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} // if(objDeliverable.GovernanceControls != null)
									} // if(recDeliverable.GovernanceControls != null &&)
								} //if(this.Deliverable_GovernanceControls)

							//+ Insert Glossary Terms or Acronyms associated with the Deliverable(s).
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
								if(layer1upDeliverableID != null && objDeliverableLayer1up.GlossaryAndAcronyms != null)
									{
									foreach(var entry in objDeliverableLayer1up.GlossaryAndAcronyms)
										{
										if(this.ListGlossaryAndAcronyms.Contains(entry) != true)
											ListGlossaryAndAcronyms.Add(entry);
										}
									}
								} // if(this.Acronyms_Glossary_of_Terms_Section)
							else
								{
								// If the entry is not found - write an error in the document and
								// record an error in the error log.
								this.LogError("Error: The Deliverable ID " + deliverableItem.Key
									+ " doesn't exist in SharePoint and couldn't be retrieved.");
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: "Error: Deliverable " + deliverableItem.Key + " is missing.",
									parIsNewSection: false,
									parIsError: true);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);
								}
							} //- if(this.Deliverable_Heading)
						} //- foreach (KeyValuePair<int, String>.....

Process_Reports:   //+Reports
					if(dictReports.Count == 0 && this.Reports == false)
						goto Process_Meetings;

					Console.Write("\n Reports:");
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Reports_Heading_Text);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);
					string reportBookMark = "Report_";
					//+ Insert the individual Report in the section
					foreach(KeyValuePair<int, string> reportItem in dictReports.OrderBy(key => key.Value))
						{
						//- User selected to include Reports Headings
						if(this.Report_Heading)
							{
							//- Get the entry from the Database
							objDeliverable = Deliverable.Read(parIDsp: reportItem.Key);
							if (objDeliverable != null)
								{
								Console.Write("\t + {0} - {1}", objDeliverable.IDsp, objDeliverable.CSDheading);

								//+ Insert the Reports's CSD Heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3,
									parBookMark: reportBookMark + objDeliverable.IDsp);
								objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.CSDheading);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);

								//-Check if the Report's Layer0up has Content Layers and Content Predecessors
								if(objDeliverable.ContentPredecessorDeliverableIDsp == null)
									{
									layer1upDeliverableID = null;
									}
								else
									{
									//-| Get the Layer1up entry from the Database
									objDeliverableLayer1up = Deliverable.Read(parIDsp: Convert.ToUInt16(objDeliverable.ContentPredecessorDeliverableIDsp));
									if (objDeliverableLayer1up != null)
										{
										layer1upDeliverableID = objDeliverableLayer1up.IDsp;
										}
									else
										{
										layer1upDeliverableID = null;
										}
									}

								//+ Insert the Deliverable CSD Description
								if(this.Report_Description)
									{
									// Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null
									&& objDeliverableLayer1up.CSDdescription != null)
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
											Console.Write("\n\t\t + Layer1up {0} - {1}", objDeliverableLayer1up.IDsp, objDeliverableLayer1up.Title);
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
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
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}

									// Insert Layer0up if not null
									if(objDeliverable.CSDdescription != null)
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
											Console.Write("\n\t\t + Layer0up {0} - {1}", objDeliverable.IDsp, objDeliverable.Title);
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
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
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}
									}

								//+ Insert the Report Inputs
								if(this.Report_Inputs)
									{
									if(objDeliverable.Inputs != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.Inputs != null))
										{
										//- Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null
										&& objDeliverableLayer1up.Inputs != null)
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

											//-Check for Colour coding Layers and add if necessary
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
													parDocumentLevel: 4,
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
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											}

										// Insert Layer0up if not null
										if(objDeliverable.Inputs != null)
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

											//-Check for Colour coding Layers and add if necessary
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
													parDocumentLevel: 4,
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
												// A Table content error occurred
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} // if(recReport.Inputs != null)
										} //if(recReports.Inputs != null &&)
									} //if(this.Report_Inputs)

								//+Insert the Deliverable Outputs
								if(this.Report_Outputs)
									{
									if(objDeliverable.Outputs != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.Outputs != null))
										{
										//- Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableOutputs_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null
										&& objDeliverableLayer1up.Outputs != null)
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

											//-Check for Colour coding Layers and add if necessary
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
													parDocumentLevel: 4,
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
												// A Table content error occurred, record it in the
												// error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											}

										// Insert Layer0up if not null
										if(objDeliverable.Outputs != null)
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

											//-Check for Colour coding Layers and add if necessary
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
													parDocumentLevel: 4,
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
												// A Table content error occurred, record it in the
												// error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} // if(objReport.Outputs != null)
										} //if(recReport.Outputs !== null &&)
									} //if(this.Report_Outputs)

								//-----------------------------------------------------------------------
								//+Insert the Report DD's Obligations
								if(this.DDs_Report_Obligations)
									{
									if(objDeliverable.DDobligations != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.DDobligations != null))
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null
										&& objDeliverableLayer1up.DDobligations != null)
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

											//-Check for Colour coding Layers and add if necessary
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
													parDocumentLevel: 4,
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
												// A Table content error occurred, record it in the
												// error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} // if(recReport.Layer1up.DDobligations != null)

										// Insert Layer0up if not null
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

											//-Check for Colour coding Layers and add if necessary
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
													parDocumentLevel: 4,
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
												// A Table content error occurred, record it in the
												// error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} // if(objReport.DDobligations != null)
										} //if(recReport.DDoblidations != null &&)
									} //if(this.DDs_Report_Obligations)

								//+ Insert the Client Responsibilities
								if(this.Clients_Report_Responsibilities)
									{
									if(objDeliverable.ClientResponsibilities != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.ClientResponsibilities != null))
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null
										&& objDeliverableLayer1up.ClientResponsibilities != null)
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

											//-Check for Colour coding Layers and add if necessary
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
													parDocumentLevel: 4,
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
												// A Table content error occurred, record it in the
												// error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											}

										// Insert Layer0up if not null
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
													parDocumentLevel: 4,
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
												// A Table content error occurred, record it in the
												// error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} // if(recReport.ClientResponsibilities != null)
										} // if(recReport.ClientResponsibilities != null &&)
									} //if(this.Clients_Report_Responsibilities)

								//+Insert the Deliverable Exclusions
								if(this.Report_Exclusions)
									{
									if(objDeliverable.Exclusions != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.Exclusions != null))
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null
										&& objDeliverableLayer1up.Exclusions != null)
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

											//-Check for Colour coding Layers and add if necessary
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
													parDocumentLevel: 4,
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
												// A Table content error occurred, record it in the
												// error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											}

										// Insert Layer0up if not null
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

											//-Check for Colour coding Layers and add if necessary
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
													parDocumentLevel: 4,
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
												// A Table content error occurred, record it in the
												// error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} // if(recReport.Exclusions != null)
										} // if(recReport.Exclusions != null &&)
									} //if(this.Report_Exclusions)

								//+ Insert the Governance Controls
								if(this.Report_Governance_Controls)
									{
									if(objDeliverable.GovernanceControls != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.GovernanceControls != null))
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
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

											//-Check for Colour coding Layers and add if necessary
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
													parDocumentLevel: 4,
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
												// A Table content error occurred, record it in the
												// error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											}

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

											//-Check for Colour coding Layers and add if necessary
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
													parDocumentLevel: 4,
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
												// A Table content error occurred, record it in the
												// error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} // if(recReport.GovernanceControls != null)
										} // if(recReport.GovernanceControls != null &&)
									} //if(this.Report_GovernanceControls)

								//+ Add the Glossary Terms or Acronyms associated with the Deliverable(s).
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
									if(layer1upDeliverableID != null && objDeliverableLayer1up.GlossaryAndAcronyms != null)
										{
										foreach(var entry in objDeliverableLayer1up.GlossaryAndAcronyms)
											{
											if(this.ListGlossaryAndAcronyms.Contains(entry) != true)
												ListGlossaryAndAcronyms.Add(entry);
											}
										}
									} // if(this.Acronyms_Glossary_of_Terms_Section)
								}
							else
								{
								// If the entry is not found - write an error in the document and
								// record an error in the error log.
								this.LogError("Error: The Deliverable ID " + reportItem.Key
									+ " doesn't exist in SharePoint and couldn't be retrieved.");
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: "Error: Deliverable " + reportItem.Key + " is missing.",
									parIsNewSection: false,
									parIsError: true);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);
								}
							} // if(this.ReportHeading
						} // foreach (KeyValuePair<int, String>.....

Process_Meetings:  //+Meetings
					if(dictMeetings.Count == 0 || this.Meetings == false)
						goto Process_ServiceLevels;

					Console.Write("\n Meetings:");
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Meetings_Heading_Text);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);
					deliverableBookMark = "Meeting_";

					//+ Insert the individual Meetings in the section
					foreach(KeyValuePair<int, string> meetingItem in dictMeetings.OrderBy(key => key.Value))
						{
						// Get the entry from the Database
						objDeliverable = Deliverable.Read(parIDsp: meetingItem.Key);
						if (objDeliverable != null)
							{
							Console.Write("\t + {0} - {1}", objDeliverable.IDsp, objDeliverable.Title);

							//+ Insert the Reports's CSD Heading
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3, parBookMark: deliverableBookMark + objDeliverable.IDsp);
							objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.CSDheading);
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);

							//Check if the Report's Layer0up has Content Layers and Content Predecessors
							if(objDeliverable.ContentPredecessorDeliverableIDsp == null)
								layer1upDeliverableID = null;
								{
								//-| Get the entry from the Database
								objDeliverableLayer1up = Deliverable.Read(parIDsp: Convert.ToInt16(objDeliverable.ContentPredecessorDeliverableIDsp));
								if (objDeliverableLayer1up != null)
									{
									layer1upDeliverableID = objDeliverableLayer1up.IDsp;
									}
								else
									{
									layer1upDeliverableID = null;
									}
								}

							//+Insert the Deliverable CSD Description
							if(this.Meeting_Description)
								{
								// Insert Layer1up if present and not null
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upDeliverableID != null
								&& objDeliverableLayer1up.CSDdescription != null)
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
										Console.Write("\n\t\t + Layer1up {0} - {1}", objDeliverableLayer1up.IDsp, objDeliverableLayer1up.Title);
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 4,
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
											+ " contains an error in one of its Enhance Rich Text columns. "
											+ "Please review the content (especially tables).");
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content in the SharePoint "
											+ "system and correct it. " + exc.Message,
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
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}

								// Insert Layer0up if not null
								if(objDeliverable.CSDdescription != null)
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
										Console.Write("\n\t\t + Layer0up {0} - {1}", objDeliverable.IDsp, objDeliverable.Title);
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 4,
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
											+ " contains an error in one of its Enhance Rich Text columns. "
											+ "Please review the content (especially tables).");
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content in the SharePoint "
											+ "system and correct it. " + exc.Message,
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
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}
								} //if(this.Meeting_Description)

							//+Insert the Report Inputs
							if(this.Meeting_Inputs)
								{
								if(objDeliverable.Inputs != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.Inputs != null))
									{
									// Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									// Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null
									&& objDeliverableLayer1up.Inputs != null)
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
												parDocumentLevel: 4,
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
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}

									// Insert Layer0up if not null
									if(objDeliverable.Inputs != null)
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
												parDocumentLevel: 4,
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
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} // if(recMeeting.Inputs != null)
									} //if(recMeeting.Inputs != null &&)
								} //if(this.Meeting_Inputs)

							//+ Insert the Deliverable Outputs
							if(this.Meeting_Outputs)
								{
								if(objDeliverable.Outputs != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.Outputs != null))
									{
									// Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableOutputs_Heading_Text);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									// Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null
									&& objDeliverableLayer1up.Outputs != null)
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
												parDocumentLevel: 4,
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
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}

									// Insert Layer0up if not null
									if(objDeliverable.Outputs != null)
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
												parDocumentLevel: 4,
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
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} // if(recMeeting.Outputs != null)
									} //if(recMeeting.Outputs !== null &&)
								} //if(this.Meeting_Outputs)

							//+Insert the Report DD's Obligations
							if(this.DDs_Report_Obligations)
								{
								if(objDeliverable.DDobligations != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.DDobligations != null))
									{
									// Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									// Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null
									&& objDeliverableLayer1up.DDobligations != null)
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
												parDocumentLevel: 4,
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
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}

									// Insert Layer0up if not null
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
												parDocumentLevel: 4,
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
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} // if(recMeeting.DDobligations != null)
									} //if(recMeeting.DDoblidations != null &&)
								} //if(this.DDs_Report_Obligations)

							//+Insertthe Client Responsibilities
							if(this.Clients_Report_Responsibilities)
								{
								if(objDeliverable.ClientResponsibilities != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.ClientResponsibilities != null))
									{
									// Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									// Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null
									&& objDeliverableLayer1up.ClientResponsibilities != null)
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
												parDocumentLevel: 4,
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
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}

									// Insert Layer0up if not null
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
												parDocumentLevel: 4,
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
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} // if(recMeeting.ClientResponsibilities != null)
									} // if(recMeeting.ClientResponsibilities != null &&)
								} //if(this.Clients_Report_Responsibilities)

							//+Insert the Deliverable Exclusions
							if(this.Meeting_Exclusions)
								{
								if(objDeliverable.Exclusions != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.Exclusions != null))
									{
									// Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									// Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null
									&& objDeliverableLayer1up.Exclusions != null)
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
												parDocumentLevel: 4,
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
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}

									// Insert Layer0up if not null
									if(objDeliverable.Exclusions != null)
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
												parDocumentLevel: 4,
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
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} // if(recMeeting.Exclusions != null)
									} // if(objMeeting.Exclusions != null &&)
								} //if(this.Deliverable_Exclusions)

							//+Insert the Governance Controls
							if(this.Deliverable_Governance_Controls)
								{
								if(objDeliverable.GovernanceControls != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.GovernanceControls != null))
									{
									// Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
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
											currentContentLayer = "Layer2";
										else
											currentContentLayer = "None";

										try
											{
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
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
											this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}

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
												parDocumentLevel: 4,
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
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it. " + exc.Message,
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} // if(recMeeting.GovernanceControls != null)
									} // if(recMeeting.GovernanceControls != null &&)
								} //- if(this.Deliverable_Governance_Controls)

							//+Insert any Glossary Terms or Acronyms associated with the Deliverable(s).
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
								if(layer1upDeliverableID != null && objDeliverableLayer1up.GlossaryAndAcronyms != null)
									{
									foreach(var entry in objDeliverableLayer1up.GlossaryAndAcronyms)
										{
										if(this.ListGlossaryAndAcronyms.Contains(entry) != true)
											ListGlossaryAndAcronyms.Add(entry);
										}
									}
								} //- if(this.Acronyms_Glossary_of_Terms_Section)
							}
						else
							{
							// If the entry is not found - write an error in the document and record
							// an error in the error log.
							this.LogError("Error: The Deliverable ID " + meetingItem.Key
								+ " doesn't exist in SharePoint and couldn't be retrieved.");
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "Error: Deliverable " + meetingItem.Key + " is missing.",
								parIsNewSection: false,
								parIsError: true);
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);
							}
						} //- foreach(....

					} //- if(this.DRMSection)

Process_ServiceLevels:   //++ Insert the Service Levels Section
				if(this.Service_Level_Section == false
				|| dictSLAs.Count == 0)
					{
					goto Process_Glossary_and_Acronyms;
					}
				else
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_ServiceLevel_Section_Text,
						parIsNewSection: true);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					string servicelevelBookMark = "ServiceLevel_";
					// Insert the individual Service Levels in the section
					foreach(KeyValuePair<int, string> servicelevelItem in dictSLAs.OrderBy(sortkey => sortkey.Value))
						{
						// Prepare the data which to insert into the Service Level Table
						objDeliverableServiceLevel = DeliverableServiceLevel.Read(parIDsp: servicelevelItem.Key);
						if (objDeliverableServiceLevel != null)
							{
							Console.WriteLine("\t\t + Deliverable ServiceLevel: {0} - {1}", objDeliverableServiceLevel.IDsp,
								objDeliverableServiceLevel.Title);

							// Get the Service Level entry from the Database
							if(objDeliverableServiceLevel.AssociatedServiceLevelIDsp != null)
								{
								objServiceLevel = ServiceLevel.Read(parIDsp: Convert.ToInt16(objDeliverableServiceLevel.AssociatedServiceLevelIDsp));
								if (objServiceLevel != null)
									{
									Console.WriteLine("\t\t\t + Service Level: {0} - {1}", objServiceLevel.IDsp, objServiceLevel.Title);
									Console.WriteLine("\t\t\t + Service Hour.: {0}", objServiceLevel.ServiceHours);

									if(this.Service_Level_Commitments_Table)
										{
										// Insert the Service Level CSD Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2,
											parBookMark: servicelevelBookMark + objServiceLevel.IDsp);
										objRun = oxmlDocument.Construct_RunText(parText2Write: objServiceLevel.CSDheading);
										// Check if a hyperlink must be inserted
										if(documentCollection_HyperlinkURL != "")
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
										// Insert the Service Level Commitments Table
										if(objServiceLevel.CSDdescription != null)
											{
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_ServiceLevelsURI +
												currentHyperlinkViewEditURI +
												objServiceLevel.IDsp;
											currentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(parClientName: parClientName,
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 2,
														parHTML2Decode: HTMLdecoder.CleanHTML(objServiceLevel.CSDdescription, parClientName),
														parContentLayer: currentContentLayer,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
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
												this.LogError("Error: The Service Level ID: " + objServiceLevel.IDsp
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it. " + exc.Message,
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} // if(objServiceLevel.CSDdescription != null)

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
										} //if(this.Service_Level_Commitments_Table)
									} //if(parDatabase.dsServiceLevels.TryGetValue(
								} // if(objDeliverableServiceLevel.AssociatedServiceLevelID != null)
							} // if(parDatabase.dsDeliverableServiceLevels.TryGetValue(
						else
							{
							// If the entry is not found - write an error in the document and record an
							// error in the error log.
							this.LogError("Error: The DeliverableServiceLevel ID " + dictSLAs.Keys + " - " + dictSLAs.Values
								+ " doesn't exist in SharePoint and it couldn't be retrieved.");
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "Error: DeliverableServiceLevel " + dictSLAs.Keys + " is missing.",
								parIsNewSection: false,
								parIsError: true);
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);
							break;
							}
						} //-foreach(servicelevelItem in dictSLA....
					} //- else...


Process_Glossary_and_Acronyms: //++Glossary & Acronyms
				if(this.ListGlossaryAndAcronyms.Count == 0)
					goto Save_and_Close_Document;

				if(this.Acronyms_Glossary_of_Terms_Section)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_GlossaryAndAcronymSection_HeadingText,
						parIsNewSection: true);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					// Insert a blank paragrpah
					objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: " ");
					objParagraph.Append(objRun);
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
						}
					}

Save_and_Close_Document: //+Error Section

				if(this.ErrorMessages.Count > 0)
					{
					//+ Insert the Document Generation Error Section

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
						objParagraph = oxmlDocument.Construct_Error( errorMessageEntry);
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

				//++ Upload the document to SharePoint
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
				}

			Console.WriteLine("\t\t End of the generation of {0}", this.DocumentType);
			//- Delete the file from the Documents Directory
			if(File.Exists(path: this.LocalDocumentURI))
				File.Delete(path: this.LocalDocumentURI);

			} //- end of Generate method
		} //- end of CSD_Document_DRM_Sections class
	}