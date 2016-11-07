using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocGeneratorCore.Database.Classes;

namespace DocGeneratorCore
	{
	/// <summary>
	///      This class represent the Internal Service Definition (ISD) with inline DRM (Deliverable
	///      Report Meeting) It inherits from the Internal_DRM_Inline Class.
	/// </summary>
	internal class ISD_Document_DRM_Inline:Internal_DRM_Inline
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
						case 60:
							this.Introductory_Section = true;
							break;

						case 61:
							this.Introduction = true;
							break;

						case 62:
							this.Executive_Summary = true;
							break;

						case 63:
							this.Service_Portfolio_Section = true;
							break;

						case 64:
							this.Service_Portfolio_Description = true;
							break;

						case 65:
							this.Service_Family_Heading = true;
							break;

						case 66:
							this.Service_Family_Description = true;
							break;

						case 67:
							this.Service_Product_Heading = true;
							break;

						case 68:
							this.Service_Product_Description = true;
							break;

						case 69:
							this.Service_Product_Key_Client_Benefits = true;
							break;

						case 70:
							this.Service_Product_KeyDD_Benefits = true;
							break;

						case 71:
							this.Service_Element_Heading = true;
							break;

						case 72:
							this.Service_Element_Description = true;
							break;

						case 73:
							this.Service_Element_Objectives = true;
							break;

						case 74:
							this.Service_Element_Key_Client_Benefits = true;
							break;

						case 75:
							this.Service_Element_Key_Client_Advantages = true;
							break;

						case 76:
							this.Service_Element_Key_DD_Benefits = true;
							break;

						case 77:
							this.Service_Element_Critical_Success_Factors = true;
							break;

						case 78:
							this.Service_Element_Key_Performance_Indicators = true;
							break;

						case 79:
							this.Service_Element_High_Level_Process = true;
							break;

						case 80:
							this.Deliverables_Reports_Meetings = true;
							break;

						case 81:
							this.DRM_Heading = true;
							break;

						case 82:
							this.DRM_Description = true;
							break;

						case 83:
							this.DRM_Inputs = true;
							break;

						case 84:
							this.DRM_Outputs = true;
							break;

						case 85:
							this.DDS_DRM_Obligations = true;
							break;

						case 86:
							this.Clients_DRM_Responsibilities = true;
							break;

						case 87:
							this.DRM_Exclusions = true;
							break;

						case 88:
							this.DRM_Governance_Controls = true;
							break;

						case 89:
							this.Service_Levels = true;
							break;

						case 90:
							this.Service_Level_Heading = true;
							break;

						case 91:
							this.Service_Level_Commitments_Table = true;
							break;

						case 92:
							this.Activities = true;
							break;

						case 93:
							this.Activity_Heading = true;
							break;

						case 94:
							this.Activity_Description_Table = true;
							break;

						case 95:
							this.Acronyms_Glossary_of_Terms_Section = true;
							break;

						case 96:
							this.Acronyms = true;
							break;

						case 97:
							this.Glossary_of_Terms = true;
							break;

						case 98:
							this.Document_Acceptance_Section = true;
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
			ref CompleteDataSet parDataSet,
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
			bool layerHeadingWritten = false;
			bool drmHeading = false;
			Table objActivityTable = new Table();
			Table objServiceLevelTable = new Table();
			int? layer1upElementID = 0;
			int? layer1upDeliverableID = 0;
			int tableCaptionCounter = 0;
			int imageCaptionCounter = 0;
			int numberingCounter = 49;
			int iPictureNo = 49;
			int intHyperlinkCounter = 9;

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
					parDataSet: ref parDataSet))
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
				Console.WriteLine("\t\t Effective pageWidth x pageHeight.: {0} x {1} twips", this.PageWith, this.PageHeight);

				// Check whether Hyperlinks need to be included and add the image to the Document Body
				if(this.HyperlinkEdit || this.HyperlinkView)
					{
					//Insert and embed the hyperlink image in the document and keep the Image's Relationship ID in a variable for repeated use
					hyperlinkImageRelationshipID = oxmlDocument.Insert_HyperlinkImage(parMainDocumentPart: ref objMainDocumentPart,
						parDataSet: ref parDataSet);
					}


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

				// Define the objects to be used in the construction of the document
				ServicePortfolio objPortfolio = new ServicePortfolio();
				ServiceFamily objFamily = new ServiceFamily();
				ServiceProduct objProduct = new ServiceProduct();
				ServiceElement objElement = new ServiceElement();
				ServiceElement objElementLayer1up = new ServiceElement();
				Deliverable objDeliverable = new Deliverable();
				Deliverable objDeliverableLayer1up = new Deliverable();
				DeliverableActivity objDeliverableActivity = new DeliverableActivity();
				DeliverableServiceLevel objDeliverableServiceLevel = new DeliverableServiceLevel();
				Activity objActivity = new Activity();
				ServiceLevel objServiceLevel = new ServiceLevel();

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
						intHyperlinkCounter += 1;
						Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
							parMainDocumentPart: ref objMainDocumentPart,
							parImageRelationshipId: hyperlinkImageRelationshipID,
							parClickLinkURL: documentCollection_HyperlinkURL,
							parHyperlinkID: intHyperlinkCounter);
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
								parHyperlinkID: ref intHyperlinkCounter,
								parPageHeightDxa: this.PageHeight,
								parPageWidthDxa: this.PageWith,
								parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
							}
						catch(InvalidContentFormatException exc)
							{
							Console.WriteLine("\n\nException occurred: {0}", exc.Message);
							// A Table content error occurred, record it in the error log.
							this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
								+ " contains an error in Introduction's Enhance Rich Text. " + exc.Message);
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "A content error occurred at this position and valid content could " +
								"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it. Error Detail: "
								+ exc.Message,
								parIsNewSection: false,
								parIsError: true);
							if(this.HyperlinkEdit || this.HyperlinkView)
								{
								intHyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
									parMainDocumentPart: ref objMainDocumentPart,
									parImageRelationshipId: hyperlinkImageRelationshipID,
									parHyperlinkID: intHyperlinkCounter,
									parClickLinkURL: documentCollection_HyperlinkURL);
								objRun.Append(objDrawing);
								}
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);
							}
						}
					}

				//++ Insert the Executive Summary
				if(this.Executive_Summary)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_ExecutiveSummary_HeadingText);
					// Check if a hyperlink must be inserted
					if(this.HyperlinkEdit || this.HyperlinkView)
						{
						intHyperlinkCounter += 1;
						Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
							parMainDocumentPart: ref objMainDocumentPart,
							parImageRelationshipId: hyperlinkImageRelationshipID,
							parClickLinkURL: documentCollection_HyperlinkURL,
							parHyperlinkID: intHyperlinkCounter);
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
							parHyperlinkID: ref intHyperlinkCounter,
							parPageHeightDxa: this.PageHeight,
							parPageWidthDxa: this.PageWith, 
							parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
							}
						catch(InvalidContentFormatException exc)
							{
							Console.WriteLine("\n\nException occurred: {0}", exc.Message);
							// A Table content error occurred, record it in the error log.
							this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
								+ " contains an error in Executive Summary's Enhance Rich Text. " + exc.Message);
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "A content error occurred at this position and valid content could " +
								"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it. Error Detail: "
								 + exc.Message,
								parIsNewSection: false,
								parIsError: true);
							if(this.HyperlinkEdit || this.HyperlinkView)
								{
								intHyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
									parMainDocumentPart: ref objMainDocumentPart,
									parImageRelationshipId: hyperlinkImageRelationshipID,
									parHyperlinkID: intHyperlinkCounter,
									parClickLinkURL: documentCollection_HyperlinkURL);
								objRun.Append(objDrawing);
								}
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);
							}
						}
					}

				// Insert the user selected content
				if(this.SelectedNodes.Count <= 0)
					goto Process_Glossary_and_Acronyms;

				foreach(Hierarchy node in this.SelectedNodes)
					{
					Console.Write("\nNode: Seq:{0} LeveL:{1} Type:{2} ID:{3}", node.Sequence, node.Level, node.NodeType, node.NodeID);

					switch(node.NodeType)
						{
					//++ Service Portfolio or Service Framework
					case enumNodeTypes.FRA:  // Service Framework
					case enumNodeTypes.POR:  //Service Portfolio

						if(!this.Service_Portfolio_Section)
							break;

						if(parDataSet.dsPortfolios.TryGetValue(
							key: node.NodeID,
							value: out objPortfolio))
							{
							Console.Write("\t\t + {0} - {1}", objPortfolio.IDsp, objPortfolio.Title);
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: objPortfolio.ISDheading,
								parIsNewSection: true);

							// Check if a hyperlink must be inserted
							if(this.HyperlinkEdit || this.HyperlinkView)
								{
								intHyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
									parMainDocumentPart: ref objMainDocumentPart,
									parImageRelationshipId: hyperlinkImageRelationshipID,
									parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
										Properties.AppResources.List_ServicePortfoliosURI +
										currentHyperlinkViewEditURI + objPortfolio.IDsp,
									parHyperlinkID: intHyperlinkCounter);
								objRun.Append(objDrawing);
								}
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);

							//+ Insert the Service Porfolio Description
							if (!this.Service_Portfolio_Description
							|| string.IsNullOrWhiteSpace(objPortfolio.ISDdescription))
								break;

							try
								{
								objHTMLdecoder.DecodeHTML(parClientName: parClientName,
									parMainDocumentPart: ref objMainDocumentPart,
									parDocumentLevel: 1,
									parHTML2Decode: HTMLdecoder.CleanHTML(objPortfolio.ISDdescription, parClientName),
									parHyperlinkID: ref intHyperlinkCounter,
									parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
									parHyperlinkURL: currentListURI,
									parContentLayer: currentContentLayer,
									parTableCaptionCounter: ref tableCaptionCounter,
									parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
									parPictureNo: ref iPictureNo,
									parPageHeightDxa: this.PageHeight,
									parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
								}
							catch(InvalidContentFormatException exc)
								{
								Console.WriteLine("\n\nException occurred: {0}", exc.Message);
								//- A Table content error occurred, record it in the error log.
								this.LogError("Error: The Service Portfolio ID: " + node.NodeID
									+ " contains an error in the Enhance Rich Text column ISD Description. "
										+ exc.Message);
								objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: "A content error occurred at this position and valid content could "
									+ "not be interpreted and inserted here. Please review the content in the "
									+ "SharePoint system and correct it.",
									parIsNewSection: false,
									parIsError: true);
								if(this.HyperlinkEdit || this.HyperlinkView)
									{
									intHyperlinkCounter += 1;
									Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
										parMainDocumentPart: ref objMainDocumentPart,
										parImageRelationshipId: hyperlinkImageRelationshipID,
										parHyperlinkID: intHyperlinkCounter,
										parClickLinkURL: documentCollection_HyperlinkURL);
									objRun.Append(objDrawing);
									}
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);
								}
							} //- Try
						else
							{
							// If the entry is not found - write an error in the document and record
							// an error in the error log.
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

					//++Service Family
					case enumNodeTypes.FAM:

						if(!this.Service_Family_Heading)
							break;

						// Get the entry from the DataSet
						if(parDataSet.dsFamilies.TryGetValue(
							key: node.NodeID,
							value: out objFamily))
							{
							Console.Write("\t\t + {0} - {1}", objFamily.IDsp, objFamily.Title);
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: objFamily.ISDheading,
								parIsNewSection: false);
							// Check if a hyperlink must be inserted
							if(this.HyperlinkEdit || this.HyperlinkView)
								{
								intHyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
									parMainDocumentPart: ref objMainDocumentPart,
									parImageRelationshipId: hyperlinkImageRelationshipID,
									parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
									Properties.AppResources.List_ServiceFamiliesURI +
									currentHyperlinkViewEditURI + objFamily.IDsp,
									parHyperlinkID: intHyperlinkCounter);
								objRun.Append(objDrawing);
								}
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);

							//+ Insert the Service Family Description
							if(!this.Service_Family_Description
							|| objFamily.ISDdescription == null)
								break;

							try
								{
								objHTMLdecoder.DecodeHTML(parClientName: parClientName,
									parMainDocumentPart: ref objMainDocumentPart,
									parDocumentLevel: 2,
									parHTML2Decode: HTMLdecoder.CleanHTML(objFamily.ISDdescription, parClientName),
									parHyperlinkID: ref intHyperlinkCounter,
									parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
									parHyperlinkURL: currentListURI,
									parContentLayer: currentContentLayer,
									parTableCaptionCounter: ref tableCaptionCounter,
									parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
									parPictureNo: ref iPictureNo,
									parPageHeightDxa: this.PageHeight,
									parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
								}
							catch(InvalidContentFormatException exc)
								{
								Console.WriteLine("\n\nException occurred: {0}", exc.Message);
								// A Table content error occurred, record it in the error log.
								this.LogError("Error: The Service Family ID: " + node.NodeID
									+ " contains an error in the Enhance Rich Text column ISD Description. "
									+ exc.Message);
								objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: "A content error occurred at this position and valid content could "
									+ "not be interpreted and inserted here. Please review the "
									+ "content in the SharePoint system and correct it.",
									parIsNewSection: false,
									parIsError: true);
								if(this.HyperlinkEdit || this.HyperlinkView)
									{
									intHyperlinkCounter += 1;
									Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
										parMainDocumentPart: ref objMainDocumentPart,
										parImageRelationshipId: hyperlinkImageRelationshipID,
										parHyperlinkID: intHyperlinkCounter,
										parClickLinkURL: currentListURI);
									objRun.Append(objDrawing);
									}
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);
								}
							}
						else
							{
							//- If the entry is not found - write an error in the document and record an error in the error log.
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

					//++ Service Product
					case enumNodeTypes.PRO:

						if(!this.Service_Product_Heading)
							break;

						// Get the entry from the DataSet
						if(parDataSet.dsProducts.TryGetValue(
							key: node.NodeID,
							value: out objProduct))
							{
							Console.Write("\t\t + {0} - {1}", objProduct.IDsp, objProduct.Title);
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: objProduct.ISDheading,
								parIsNewSection: false);

							// Check if a hyperlink must be inserted
							if(this.HyperlinkEdit || this.HyperlinkView)
								{
								intHyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
									parMainDocumentPart: ref objMainDocumentPart,
									parImageRelationshipId: hyperlinkImageRelationshipID,
									parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
									Properties.AppResources.List_ServiceProductsURI +
									currentHyperlinkViewEditURI + objProduct.IDsp,
									parHyperlinkID: intHyperlinkCounter);
								objRun.Append(objDrawing);
								}
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);

							//+Insert the Service Product Description
							if(this.Service_Product_Description
							|| objProduct.ISDdescription != null)
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
										parHTML2Decode: HTMLdecoder.CleanHTML(objProduct.ISDdescription, parClientName),
										parHyperlinkID: ref intHyperlinkCounter,
										parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
										parHyperlinkURL: currentListURI,
										parContentLayer: currentContentLayer,
										parTableCaptionCounter: ref tableCaptionCounter,
										parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
										parPictureNo: ref iPictureNo,
										parPageHeightDxa: this.PageHeight,
										parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
									}
								catch(InvalidContentFormatException exc)
									{
									Console.WriteLine("\n\nException occurred: {0}", exc.Message);
									// A Table content error occurred, record it in the error log.
									this.LogError("Error: The Service Product ID: " + node.NodeID
										+ " contains an error in the Enhance Rich Text column ISD Description. "
										+ exc.Message);
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "A content error occurred at this position and valid content could "
										+ "not be interpreted and inserted here. "
										+ "Please review the content in the SharePoint system and correct it. Error Detail: " + exc.Message,
										parIsNewSection: false,
										parIsError: true);
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										intHyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parHyperlinkID: intHyperlinkCounter,
											parClickLinkURL: currentListURI);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
								}

							//+ Insert Product Key DD Benefits
							if(this.Service_Product_KeyDD_Benefits
							&& objProduct.KeyDDbenefits != null)
								{
								currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
									Properties.AppResources.List_ServiceProductsURI +
									currentHyperlinkViewEditURI +
									objProduct.IDsp;
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Product_KeyDD_Benefits,
									parIsNewSection: false);
								// Check if a hyperlink must be inserted
								if(this.HyperlinkEdit || this.HyperlinkView)
									{
									intHyperlinkCounter += 1;
									Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
										parMainDocumentPart: ref objMainDocumentPart,
										parImageRelationshipId: hyperlinkImageRelationshipID,
										parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
										Properties.AppResources.List_ServiceProductsURI +
										currentHyperlinkViewEditURI + objProduct.IDsp,
										parHyperlinkID: intHyperlinkCounter);
									objRun.Append(objDrawing);
									}
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);

								try
									{
									objHTMLdecoder.DecodeHTML(parClientName: parClientName,
										parMainDocumentPart: ref objMainDocumentPart,
										parDocumentLevel: 4,
										parHTML2Decode: HTMLdecoder.CleanHTML(objProduct.KeyDDbenefits, parClientName),
										parHyperlinkID: ref intHyperlinkCounter,
										parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
										parHyperlinkURL: currentListURI,
										parContentLayer: currentContentLayer,
										parTableCaptionCounter: ref tableCaptionCounter,
										parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
										parPictureNo: ref iPictureNo,
										parPageHeightDxa: this.PageHeight,
										parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
									}
								catch(InvalidContentFormatException exc)
									{
									Console.WriteLine("\n\nException occurred: {0}", exc.Message);
									// A Table content error occurred, record it in the error log.
									this.LogError("Error: Service Product ID: " + node.NodeID
										+ " contains an error in the Enhance Rich Text column Key DD Benifits. "
										+ exc.Message);
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "A content error occurred at this position and valid content could "
										+ "not be interpreted and inserted here. "
										+ "Please review the content in the SharePoint system and correct it. Error Detail: "
										+ exc.Message,
										parIsNewSection: false,
										parIsError: true);
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										intHyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parHyperlinkID: intHyperlinkCounter,
											parClickLinkURL: currentListURI);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
								}

							//+ Product Key Client Benefits
							if(this.Service_Product_Key_Client_Benefits
							&& objProduct.KeyClientBenefits != null)
								{
								currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
									Properties.AppResources.List_ServiceProductsURI +
									currentHyperlinkViewEditURI +
									objProduct.IDsp;

								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Product_ClientKeyBenefits,
									parIsNewSection: false);
								// Check if a hyperlink must be inserted
								if(this.HyperlinkEdit || this.HyperlinkView)
									{
									intHyperlinkCounter += 1;
									Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
										parMainDocumentPart: ref objMainDocumentPart,
										parImageRelationshipId: hyperlinkImageRelationshipID,
										parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
										Properties.AppResources.List_ServiceProductsURI +
										currentHyperlinkViewEditURI + objProduct.IDsp,
										parHyperlinkID: intHyperlinkCounter);
									objRun.Append(objDrawing);
									}
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);

								try
									{
									objHTMLdecoder.DecodeHTML(parClientName: parClientName,
										parMainDocumentPart: ref objMainDocumentPart,
										parDocumentLevel: 4,
										parHTML2Decode: HTMLdecoder.CleanHTML(objProduct.KeyClientBenefits, parClientName),
										parHyperlinkID: ref intHyperlinkCounter,
										parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
										parHyperlinkURL: currentListURI,
										parContentLayer: currentContentLayer,
										parTableCaptionCounter: ref tableCaptionCounter,
										parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
										parPictureNo: ref iPictureNo,
										parPageHeightDxa: this.PageHeight,
										parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
									}
								catch(InvalidContentFormatException exc)
									{
									Console.WriteLine("\n\nException occurred: {0}", exc.Message);
									// A Table content error occurred, record it in the error log.
									this.LogError("Error: The Service Product ID: " + node.NodeID
										+ " contains an error in the Enhance Rich Text column Key Client Benefits. "
										+ exc.Message);
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "A content error occurred at this position and valid content could "
										+ "not be interpreted and inserted here. "
										+ "Please review the content in the SharePoint system and correct it. Error Detail: "
										+ exc.Message,
										parIsNewSection: false,
										parIsError: true);
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										intHyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parHyperlinkID: intHyperlinkCounter,
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
							// If the entry is not found - write an error in the document and record
							// an error in the error log.
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

					//++Service Element
					case enumNodeTypes.ELE:  // Service Element

						if(!this.Service_Element_Heading)
							break;

						// Get the entry from the DataSet
						if(parDataSet.dsElements.TryGetValue(
							key: node.NodeID,
							value: out objElement))
							{
							Console.Write("\t\t + {0} - {1}", objElement.IDsp, objElement.Title);

							//+ Insert the Service Element ISD Heading...
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
							objRun = oxmlDocument.Construct_RunText(parText2Write: objElement.ISDheading);
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);

							//Check if the Element Layer0up has Content Layers and Content Predecessors
							if(objElement.ContentPredecessorElementIDsp == null)
								{
								layer1upElementID = null;
								}
							else
								{
								// Get the entry from the DataSet
								if(parDataSet.dsElements.TryGetValue(
									key: Convert.ToInt16(objElement.ContentPredecessorElementIDsp),
									value: out objElementLayer1up))
									{
									layer1upElementID = objElementLayer1up.IDsp;
									}
								else
									{
									layer1upElementID = null;
									}
								}

							//+ Include the Service Element Description
							if(this.Service_Element_Description)
								{
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upElementID != null
								&& objElementLayer1up.ISDdescription != null)
									{
									// Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElementLayer1up.IDsp;
										}
									else
										currentListURI = "";

									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objElementLayer1up.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objElementLayer1up.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									try
										{
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 4,
											parHTML2Decode: HTMLdecoder.CleanHTML(objElementLayer1up.ISDdescription, parClientName),
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
											parPictureNo: ref iPictureNo,
											parHyperlinkID: ref intHyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										// A Table content error occurred, record it in the error log.
										this.LogError("Error: Service Element ID: " + objElementLayer1up.IDsp
											+ " contains an error in the Enhance Rich Text column ISD Description. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content "
											+ "in the SharePoint system and correct it.",
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: intHyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									} //- if(this.PresentationMode == enumPresentationMode.Layered)

								// Insert Layer0up if not null
								if(objElement.ISDdescription != null)
									{
									// Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElement.IDsp;
										}
									else
										currentListURI = "";

									//- Check for Colour coding of Content Layers
									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objElement.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objElement.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									try
										{
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 4,
											parHTML2Decode: HTMLdecoder.CleanHTML(objElement.ISDdescription, parClientName),
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
											parPictureNo: ref iPictureNo,
											parHyperlinkID: ref intHyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										// A Table content error occurred, record it in the error log.
										this.LogError("Error: Service Element ID: " + objElement.IDsp
											+ " contains an error in the Enhance Rich Text column ISD Description. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content "
											+ "in the SharePoint system and correct it.",
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: intHyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}
								} //- if(this.Service_Element_Description)

							//+ Insert the Service Element Objectives
							if(this.Service_Element_Objectives)
								{
								//-Insert the heading
								//-Prepeare the heading paragraph to be inserted, but only insert it if required...
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Element_Objectives,
									parIsNewSection: false);
								objParagraph.Append(objRun);
								layerHeadingWritten = false;

								// Insert Layer1up if present and not null
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upElementID != null
								&& objElementLayer1up.Objectives != null)
									{
									// insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}
									// Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElementLayer1up.IDsp;
										}
									else
										currentListURI = "";

									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objElementLayer1up.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objElementLayer1up.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									try
										{
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: HTMLdecoder.CleanHTML(objElementLayer1up.Objectives, parClientName),
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
											parPictureNo: ref iPictureNo,
											parHyperlinkID: ref intHyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										// A Table content error occurred, record it in the error log.
										this.LogError("Error: Service Element ID: " + objElementLayer1up.IDsp
											+ " contains an error in the Enhance Rich Text column Objectives. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content "
											+ "in the SharePoint system and correct it.",
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: intHyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									} //- if(this.PresentationMode == enumPresentationMode.Layered)

								// Insert Layer 0up if not null
								if(objElement.Objectives != null)
									{
									// insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}
									// Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElement.IDsp;
										}
									else
										currentListURI = "";

									//- Check for Colour coding of Content Layers
									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objElement.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objElement.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									try
										{
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: HTMLdecoder.CleanHTML(objElement.Objectives, parClientName),
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
											parPictureNo: ref iPictureNo,
											parHyperlinkID: ref intHyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										// A Table content error occurred, record it in the error log.
										this.LogError("Error: Service Element ID: " + objElement.IDsp
											+ " contains an error in the Enhance Rich Text column Objectives. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content "
											+ "in the SharePoint system and correct it.",
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: intHyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}
								} //if(this.Service_Element_Objectives)

							//+ Insert the Critical Success Factors
							if(this.Service_Element_Critical_Success_Factors)
								{
								//- Prepare the heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Element_CriticalSuccessFactors,
									parIsNewSection: false);
								objParagraph.Append(objRun);
								layerHeadingWritten = false;

								//- Insert Layer1up if present and not null
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upElementID != null
								&& objElementLayer1up.CriticalSuccessFactors != null)
									{
									// insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}
									// Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElementLayer1up.IDsp;
										}
									else
										currentListURI = "";

									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objElementLayer1up.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objElementLayer1up.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									try
										{
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: HTMLdecoder.CleanHTML(objElementLayer1up.CriticalSuccessFactors, parClientName),
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
											parPictureNo: ref iPictureNo,
											parHyperlinkID: ref intHyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										// A Table content error occurred, record it in the error log.
										this.LogError("Error: Service Element ID: " + objElementLayer1up.IDsp
											+ " contains an error in the Enhance Rich Text column Critical Success Factors. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content "
											+ "in the SharePoint system and correct it.",
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: intHyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									} //- if (this.PresentationMode == Layered)

								// Insert Layer 0up if not null
								if(objElement.CriticalSuccessFactors != null)
									{
									// insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}
									// Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElement.IDsp;
										}
									else
										currentListURI = "";

									//- Check for Colour coding of Content Layers
									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objElement.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objElement.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									try
										{
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: HTMLdecoder.CleanHTML(objElement.CriticalSuccessFactors, parClientName),
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
											parPictureNo: ref iPictureNo,
											parHyperlinkID: ref intHyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										// A Table content error occurred, record it in the error log.
										this.LogError("Error: Service Element ID: " + objElement.IDsp
											+ " contains an error in the Enhance Rich Text column Critical Success Factors. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content "
											+ "in the SharePoint system and correct it.",
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: intHyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}
								} //- if(this.Service_Element_CriticalSuccessFactors)

							//+ Insert the Key Client Advantages
							if(this.Service_Element_Key_Client_Advantages)
								{
								// Insert the heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Element_ClientKeyAdvantages,
									parIsNewSection: false);
								objParagraph.Append(objRun);
								layerHeadingWritten = false;

								// Insert Layer1up if present and not null
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upElementID != null
								&& objElementLayer1up.KeyClientAdvantages != null)
									{
									// insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}
									// Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElementLayer1up.IDsp;
										}
									else
										currentListURI = "";

									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objElementLayer1up.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objElementLayer1up.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									try
										{
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: HTMLdecoder.CleanHTML(objElementLayer1up.KeyClientAdvantages, parClientName),
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
											parPictureNo: ref iPictureNo,
											parHyperlinkID: ref intHyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										// A Table content error occurred, record it in the error log.
										this.LogError("Error: Service Element ID: " + objElementLayer1up.IDsp
											+ " contains an error in the Enhance Rich Text column Key Client Advantages. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content "
											+ "in the SharePoint system and correct it.",
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: intHyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									} //- if(this.PresentationMode == Layered)

								// Insert Layer0up if not null
								if(objElement.KeyClientAdvantages != null)
									{
									// insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}
									// Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElement.IDsp;
										}
									else
										currentListURI = "";

									//- Check for Colour coding of Content Layers
									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objElement.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objElement.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									try
										{
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: HTMLdecoder.CleanHTML(objElement.KeyClientAdvantages, parClientName),
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
											parPictureNo: ref iPictureNo,
											parHyperlinkID: ref intHyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										// A Table content error occurred, record it in the error log.
										this.LogError("Error: Service Element ID: " + objElement.IDsp
											+ " contains an error in the Enhance Rich Text column Key Client Advantages. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content "
											+ "in the SharePoint system and correct it.",
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: intHyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}
								} //- if(this.Service_Element_Key Client Advantages)

							//+ Insert Key Client Benefits
							// Check if the user specified to include the Service Element Key Client Benefits
							if(this.Service_Element_Key_Client_Benefits)
								{
								// Insert the heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Element_ClientKeyBenefits,
									parIsNewSection: false);
								objParagraph.Append(objRun);
								layerHeadingWritten = false;

								// Insert Layer1up if present and not null
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upElementID != null
								&& objElementLayer1up.KeyClientBenefits != null)
									{
									// insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}
									// Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElementLayer1up.IDsp;
										}
									else
										currentListURI = "";

									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objElementLayer1up.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objElementLayer1up.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									try
										{
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: HTMLdecoder.CleanHTML(objElementLayer1up.KeyClientBenefits, parClientName),
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
											parPictureNo: ref iPictureNo,
											parHyperlinkID: ref intHyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										// A Table content error occurred, record it in the error log.
										this.LogError("Error: Service Element ID: " + objElementLayer1up.IDsp
											+ " contains an error in the Enhance Rich Text column Key Client Benefits. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content "
											+ "in the SharePoint system and correct it.",
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: intHyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									} //- if(this.PresentationMode == Layered)

								// Insert Layer 0up if not null
								if(objElement.KeyClientBenefits != null)
									{
									// insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}

									// Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElement.IDsp;
										}
									else
										currentListURI = "";

									//- Check for Colour coding of Content Layers
									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objElement.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objElement.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									try
										{
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: HTMLdecoder.CleanHTML(objElement.KeyClientBenefits, parClientName),
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
											parPictureNo: ref iPictureNo,
											parHyperlinkID: ref intHyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										// A Table content error occurred, record it in the error log.
										this.LogError("Error: Service Element ID: " + objElement.IDsp
											+ " contains an error in the Enhance Rich Text column Key Client Benefits. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content "
											+ "in the SharePoint system and correct it.",
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: intHyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}
								} //- if(this.Service_Element_KeyClientBenefits)

							//+ Insert the Key DD Benefits
							// Check if the user specified to include the Service  Element Key DD Benefits
							if(this.Service_Element_Key_DD_Benefits)
								{
								// Insert the heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Element_KeyDDBenefits,
									parIsNewSection: false);
								objParagraph.Append(objRun);
								layerHeadingWritten = false;

								// Insert Layer1up if present and not null
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upElementID != null
								&& objElementLayer1up.KeyDDbenefits != null)
									{
									// insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}
									// Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElementLayer1up.IDsp;
										}
									else
										currentListURI = "";

									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objElementLayer1up.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objElementLayer1up.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									try
										{
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: HTMLdecoder.CleanHTML(objElementLayer1up.KeyDDbenefits, parClientName),
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
											parPictureNo: ref iPictureNo,
											parHyperlinkID: ref intHyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										// A Table content error occurred, record it in the error log.
										this.LogError("Error: Service Element ID: " + objElementLayer1up.IDsp
											+ " contains an error in the Enhance Rich Text column Key DD Benefits. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content "
											+ "in the SharePoint system and correct it.",
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: intHyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									} //- if(this.PresentationMode == Layered)

								// Insert Layer0up if not null
								if(objElement.KeyDDbenefits != null)
									{
									//- insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}
									//- Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElement.IDsp;
										}
									else
										currentListURI = "";

									//- Check for Colour coding of Content Layers
									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objElement.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objElement.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									try
										{
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: HTMLdecoder.CleanHTML(objElement.KeyDDbenefits, parClientName),
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
											parPictureNo: ref iPictureNo,
											parHyperlinkID: ref intHyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										// A Table content error occurred, record it in the error log.
										this.LogError("Error: Service Element ID: " + objElement.IDsp
											+ " contains an error in the Enhance Rich Text column Key DD Benefits. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content "
											+ "in the SharePoint system and correct it.",
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: intHyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}
								} //- if(this.Service_Element_KeyDDbenefits)

							//+ Insert the Key Performance Indicators
							// Check if the user specified to include the Service Element Key Performance Indicators
							if(this.Service_Element_Key_Performance_Indicators)
								{
								// Set the heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Element_KPI,
									parIsNewSection: false);
								objParagraph.Append(objRun);
								layerHeadingWritten = false;

								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upElementID != null
								&& objElementLayer1up.KeyPerformanceIndicators != null)
									{
									// insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}
									// Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElementLayer1up.IDsp;
										}
									else
										currentListURI = "";

									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objElementLayer1up.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objElementLayer1up.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									try
										{
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: HTMLdecoder.CleanHTML(objElementLayer1up.KeyPerformanceIndicators, parClientName),
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
											parPictureNo: ref iPictureNo,
											parHyperlinkID: ref intHyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										// A Table content error occurred, record it in the error log.
										this.LogError("Error: Service Element ID: " + objElementLayer1up.IDsp
											+ " contains an error in the Enhance Rich Text column Key Performance Indicators. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content "
											+ "in the SharePoint system and correct it.",
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: intHyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									} //- if(this.PresentationMode == Layered)

								// Insert Layer0up if not null
								if(objElement.KeyPerformanceIndicators != null)
									{
									//- insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}
									//- Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElement.IDsp;
										}
									else
										currentListURI = "";

									//- Check for Colour coding of Content Layers
									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objElement.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objElement.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									try
										{
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: HTMLdecoder.CleanHTML(objElement.KeyPerformanceIndicators, parClientName),
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
											parPictureNo: ref iPictureNo,
											parHyperlinkID: ref intHyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										// A Table content error occurred, record it in the error log.
										this.LogError("Error: Service Element ID: " + objElement.IDsp
											+ " contains an error in the Enhance Rich Text column Key Performance Indicators. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content "
											+ "in the SharePoint system and correct it.",
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: intHyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}
								} //if(this.Service_Element_KeyPerformanceIndicators)

							//+ Insert the High Level Process
							// Check if the user specified to include the Service  Element High Level Process
							if(this.Service_Element_High_Level_Process)
								{
								//- Insert the heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Element_HighLevelProcess,
									parIsNewSection: false);
								objParagraph.Append(objRun);
								layerHeadingWritten = false;

								// Insert Layer1up if present and not null
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upElementID != null
								&& objElementLayer1up.ProcessLink != null)
									{
									// insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}
									// Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElementLayer1up.IDsp;
										}
									else
										currentListURI = "";

									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objElementLayer1up.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objElementLayer1up.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: objElementLayer1up.ProcessLink);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									} //- if(this.PresentationMode == Layered)

								// Insert Layer0up if not null
								if(objElement.ProcessLink != null)
									{
									//- insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}
									//- Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElement.IDsp;
										}
									else
										currentListURI = "";

									//- Check for Colour coding of Content Layers
									currentContentLayer = "None";
									if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
										{
										if(objElement.ContentLayer.Contains("1"))
											currentContentLayer = "Layer1";
										else if(objElement.ContentLayer.Contains("2"))
											currentContentLayer = "Layer2";
										}

									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: objElement.ProcessLink);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
								} //if(this.Service_Element_HighLevelProcess)
							}
						else
							{
							// If the entry is not found - write an error in the document and record
							// an error in the error log.
							this.LogError("Error: The Service Element ID " + node.NodeID
								+ " doesn't exist in SharePoint and couldn't be retrieved.");
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "Error: Service Element " + node.NodeID + " is missing.",
								parIsNewSection: false,
								parIsError: true);
							objParagraph.Append(objRun);
							}
						drmHeading = false;
						break;

					//++ Deliverable, Report, Meeting
					case enumNodeTypes.ELD:  // Deliverable associated with Element
					case enumNodeTypes.ELR:  // Report deliverable associated with Element
					case enumNodeTypes.ELM:  // Meeting deliverable associated with Element

						if(!this.DRM_Heading)
							break;

						if(drmHeading == false)
							{
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: Properties.AppResources.Document_DeliverableReportsMeetings_Heading);
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);
							drmHeading = true;
							}

						// Get the entry from the DataSet
						if(parDataSet.dsDeliverables.TryGetValue(
							key: node.NodeID,
							value: out objDeliverable))
							{
							Console.Write("\t\t + {0} - {1}", objDeliverable.IDsp, objDeliverable.Title);

							//- Insert the Deliverable ISD Heading
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
							objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.ISDheading);
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);

							//- Check if the Deliverable Layer0up has a Content Predecessors
							if(objDeliverable.ContentPredecessorDeliverableIDsp == null)
								{
								layer1upDeliverableID = null;
								}
							else
								{
								// Get the entry from the DataSet
								if(parDataSet.dsDeliverables.TryGetValue(
									key: Convert.ToInt16(objDeliverable.ContentPredecessorDeliverableIDsp),
									value: out objDeliverableLayer1up))
									{
									layer1upDeliverableID = objDeliverableLayer1up.IDsp;
									}
								else
									{
									layer1upDeliverableID = null;
									}
								}

							//+ Insert the Deliverable Description
							if(this.DRM_Description)
								{
								// Insert Layer1up if present and not null
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upDeliverableID != null
								&& objDeliverableLayer1up.ISDdescription != null)
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
										intHyperlinkCounter += 1;
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
											parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverableLayer1up.ISDdescription, parClientName),
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
											parPictureNo: ref iPictureNo,
											parHyperlinkID: ref intHyperlinkCounter,
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
											+ " contains an error in the Enhance Rich Text column ISD Description. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could " +
											"not be interpreted and inserted here. Please review the content in the SharePoint "
											+ "system and correct it. Error Detail: " + exc.Message,
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: intHyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									} //- if(this.PresentationMode == enumPresentationMode.Layered)

								// Insert Layer0up if present and not null
								if(objDeliverable.ISDdescription != null)
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
										intHyperlinkCounter += 1;
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
											parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverable.ISDdescription, parClientName),
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
											parPictureNo: ref iPictureNo,
											parHyperlinkID: ref intHyperlinkCounter,
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
											+ " contains an error in the Enhance Rich Text column ISD Description. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "A content error occurred at this position and valid content could "
											+ "not be interpreted and inserted here. Please review the content in the SharePoint "
											+ "system and correct it. Error Detail: " + exc.Message,
											parIsNewSection: false,
											parIsError: true);
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: intHyperlinkCounter,
												parClickLinkURL: currentListURI);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									} //- if(objDeliverable.ISDdescription != null)
								} //- if (this.DRM_Description)

							//+ Include the Deliverable Inputs
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
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null
									&& objDeliverableLayer1up.Inputs != null)
										{
										// Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
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
												parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverableLayer1up.Inputs, parClientName),
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
												parPictureNo: ref iPictureNo,
												parHyperlinkID: ref intHyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: Deliverable ID: " + objDeliverableLayer1up.IDsp
												+ " contains an error in the Enhance Rich Text column Inputs. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content "
												+ "in the SharePoint system and correct it.",
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												intHyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: intHyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //- if(this.PresentationMode == enumPresentationMode.Layered)

									// Insert Layer0up if not null
									if(objDeliverable.Inputs != null)
										{
										//- Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
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
												parHyperlinkID: ref intHyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in the Enhance Rich Text column Inputs. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content "
												+ "in the SharePoint system and correct it.",
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												intHyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: intHyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //- if(recDeliverable.Inputs != null)
									} //- if(objDeliverable.Inputs  &&...)
								} //- if(this.DRM_Inputs)

							//+ Include the Deliverable Outputs
							if(this.DRM_Outputs)
								{
								if(objDeliverable.Outputs != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.Outputs != null))
									{
									// Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
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
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
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
												parHyperlinkID: ref intHyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: Deliverable ID: " + objDeliverableLayer1up.IDsp
												+ " contains an error in the Enhance Rich Text column Outputs. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content "
												+ "in the SharePoint system and correct it.",
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												intHyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: intHyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //- if(this.PresentationMode == enumPresentationMode.Layered)

									// Insert Layer0up if not null
									if(objDeliverable.Outputs != null)
										{
										// Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
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
												parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverable.Outputs, parClientName),
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
												parPictureNo: ref iPictureNo,
												parHyperlinkID: ref intHyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in the Enhance Rich Text column Outputs. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content "
												+ "in the SharePoint system and correct it.",
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												intHyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: intHyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //- if(objDeliverable.Outputs != null)
									} //- if(objDeliverables.Outputs !== null &&)
								} //- if(this.DRM_Outputs)

							//+ Include the Deliverable DD's Obligations
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

									// Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null
									&& objDeliverableLayer1up.DDobligations != null)
										{
										// Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
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
												parHyperlinkID: ref intHyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: Deliverable ID: " + objDeliverableLayer1up.IDsp
												+ " contains an error in the Enhance Rich Text column DD's Obligations. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content "
												+ "in the SharePoint system and correct it.",
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												intHyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: intHyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //- if(this.PresentationMode == enumPresentationMode.Layered)

									// Insert Layer0up if not null
									if(objDeliverable.DDobligations != null)
										{
										// Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
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
												parHyperlinkID: ref intHyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in the Enhance Rich Text column DD's Obligations. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content "
												+ "in the SharePoint system and correct it.",
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												intHyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: intHyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} // if(objDeliverable.DDobligations != null)
									} //if(objDeliverable.DDoblidations != null &&)
								} //if(this.DDs_DRM_Obligations)

							//+ Include the Client Responsibilities
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

									// Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null
									&& objDeliverableLayer1up.ClientResponsibilities != null)
										{
										// Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
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
												parHyperlinkID: ref intHyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: Deliverable ID: " + objDeliverableLayer1up.IDsp
												+ " contains an error in the Enhance Rich Text column Client's Responsibilities. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content "
												+ "in the SharePoint system and correct it.",
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												intHyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: intHyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //- if(this.PresentationMode == enumPresentationMode.Layered)

									// Insert Layer0up if not null
									if(objDeliverable.ClientResponsibilities != null)
										{
										// Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
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
												parHyperlinkID: ref intHyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in the Enhance Rich Text column Client's Responsibilities. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content "
												+ "in the SharePoint system and correct it.",
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												intHyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: intHyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //- if(objDeliverable.ClientResponsibilities != null)
									} //- if(objDeliverable.ClientResponsibilities != null &&)
								} //- if(this.Clients_DRM_Responsibilities)

							//+ Insert the Deliverable Exclusions
							if(this.DRM_Exclusions)
								{
								if(objDeliverable.Exclusions != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.Exclusions != null))
									{
									// Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
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
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
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
												parHyperlinkID: ref intHyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: Deliverable ID: " + objDeliverableLayer1up.IDsp
												+ " contains an error in the Enhance Rich Text column Exclusions. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content "
												+ "in the SharePoint system and correct it.",
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												intHyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: intHyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //- if(this.PresentationMode == enumPresentationMode.Layered)

									// Insert Layer0up if not null
									if(objDeliverable.ClientResponsibilities != null)
										{
										// Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
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
												parHyperlinkID: ref intHyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in the Enhance Rich Text column Exclusions. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content "
												+ "in the SharePoint system and correct it.",
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												intHyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: intHyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //- if(objDeliverable.Exclusions != null)
									} //- if(objDeliverable.Exclusions != null &&)
								} //- if(this.DRMe_Exclusions)

							//+ Include the Governance Controls
							if(this.DRM_Governance_Controls)
								{
								if(objDeliverable.GovernanceControls != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.GovernanceControls != null))
									{
									// Insert the Heading
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
											intHyperlinkCounter += 1;
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
												parHyperlinkID: ref intHyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: Deliverable ID: " + objDeliverableLayer1up.IDsp
												+ " contains an error in the Enhance Rich Text column Governance Controls. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content "
												+ "in the SharePoint system and correct it.",
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												intHyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: intHyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //- if(this.PresentationMode == enumPresentationMode.Layered)

									// Insert Layer0up if not null
									if(objDeliverable.GovernanceControls != null)
										{
										// Check if a hyperlink must be inserted
										if(this.HyperlinkEdit || this.HyperlinkView)
											{
											intHyperlinkCounter += 1;
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
												parHyperlinkID: ref intHyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith,
												parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in the Enhance Rich Text column Governance Controls. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content "
												+ "in the SharePoint system and correct it.",
												parIsNewSection: false,
												parIsError: true);
											if(this.HyperlinkEdit || this.HyperlinkView)
												{
												intHyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: intHyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} // if(objDeliverable.GovernanceControls != null)
									} // if(objDeliverable.GovernanceControls != null &&)
								} //if(this.DRM_GovernanceControls)

							//+ Check if there are any Glossary Terms or Acronyms associated with the Deliverable(s).
							if(this.Acronyms_Glossary_of_Terms_Section)
								{
								// if there are GlossaryAndAcronyms to add from layer0up
								if(objDeliverable.GlossaryAndAcronyms != null)
									{
									foreach(var entry in objDeliverable.GlossaryAndAcronyms)
										{
										if(this.DictionaryGlossaryAndAcronyms.ContainsKey(entry.Key) != true)
											DictionaryGlossaryAndAcronyms.Add(entry.Key, entry.Value);
										}
									}
								// if there are GlossaryAndAcronyms to add from layer1up
								if(layer1upDeliverableID != null && objDeliverableLayer1up.GlossaryAndAcronyms != null)
									{
									foreach(var entry in objDeliverableLayer1up.GlossaryAndAcronyms)
										{
										if(this.DictionaryGlossaryAndAcronyms.ContainsKey(entry.Key) != true)
											DictionaryGlossaryAndAcronyms.Add(entry.Key, entry.Value);
										}
									}
								} // if(this.Acronyms_Glossary_of_Terms_Section)
							}
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

					//++ Activities
					case enumNodeTypes.EAC:  // Activity associated with Deliverable pertaining to Service Element

						if(!this.Activities)
							break;

						objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
						objRun = oxmlDocument.Construct_RunText(
							parText2Write: Properties.AppResources.Document_Activities_Heading);
						objParagraph.Append(objRun);
						objBody.Append(objParagraph);

						if(!this.Activity_Description_Table)
							break;

						// Get the entry from the DataSet
						if(parDataSet.dsActivities.TryGetValue(
							key: node.NodeID,
							value: out objActivity))
							{
							Console.WriteLine("\t\t + {0} - {1}", objActivity.IDsp, objActivity.Title);

							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
							objRun = oxmlDocument.Construct_RunText(parText2Write: objActivity.ISDheading);
							// Check if a hyperlink must be inserted
							if(this.HyperlinkEdit || this.HyperlinkView)
								{
								intHyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
									parMainDocumentPart: ref objMainDocumentPart,
									parImageRelationshipId: hyperlinkImageRelationshipID,
									parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
										Properties.AppResources.List_ActvitiesURI +
										currentHyperlinkViewEditURI + objActivity.IDsp,
									parHyperlinkID: intHyperlinkCounter);
								objRun.Append(objDrawing);
								}
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);

							// Check if the user specified to include the Deliverable Description
							if(this.Activity_Description_Table)
								{
								objActivityTable = CommonProcedures.BuildActivityTable(
									parWidthColumn1: Convert.ToInt16(this.PageWith * 0.25),
									parWidthColumn2: Convert.ToInt16(this.PageWith * 0.75),
									parActivityDesciption: objActivity.ISDdescription,
									parActivityInput: objActivity.Input,
									parActivityOutput: objActivity.Output,
									parActivityAssumptions: objActivity.Assumptions,
									parActivityOptionality: objActivity.Optionality);
								objBody.Append(objActivityTable);
								} // if (this.Activity_Description_Table)
							}
						else
							{
							// If the entry is not found - write an error in the document and record
							// an error in the error log.
							this.LogError("Error: The Activity ID " + node.NodeID
								+ " doesn't exist in SharePoint and it couldn't be retrieved.");
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "Error: Activity " + node.NodeID + " is missing.",
								parIsNewSection: false,
								parIsError: true);
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);
							break;
							}
						break;

					//++ Service Levels
					case enumNodeTypes.ESL:  // Service Level associated with Deliverable pertaining to Service Element

						if(!this.Service_Level_Heading)
							break;

						// Populate the Service Level Heading
						objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
						objRun = oxmlDocument.Construct_RunText(
							parText2Write: Properties.AppResources.Document_ServiceLevels_Heading_Text);
						objParagraph.Append(objRun);
						objBody.Append(objParagraph);

						// Check if the user specified to include the Service Level Commitments Table
						if(!this.Service_Level_Commitments_Table)
							break;

						// Prepare the data which to insert into the Service Level Table
						if(parDataSet.dsDeliverableServiceLevels.TryGetValue(
							key: node.NodeID,
							value: out objDeliverableServiceLevel))
							{
							Console.WriteLine("\t\t + Deliverable ServiceLevel: {0} - {1}", objDeliverableServiceLevel.IDsp,
								objDeliverableServiceLevel.Title);

							// Get the Service Level entry from the DataSet
							if(objDeliverableServiceLevel.AssociatedServiceLevelIDsp != null)
								{
								if(parDataSet.dsServiceLevels.TryGetValue(
									key: Convert.ToInt16(objDeliverableServiceLevel.AssociatedServiceLevelIDsp),
									value: out objServiceLevel))
									{
									Console.WriteLine("\t\t\t + Service Level: {0} - {1}", objServiceLevel.IDsp,
										objServiceLevel.Title);
									Console.WriteLine("\t\t\t + Service Hour.: {0}", objServiceLevel.ServiceHours);

									// Insert the Service Level ISD Description
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 8);
									objRun = oxmlDocument.Construct_RunText(parText2Write: objServiceLevel.ISDheading);
									// Check if a hyperlink must be inserted
									if(this.HyperlinkEdit || this.HyperlinkView)
										{
										intHyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_ServiceLevelsURI +
												currentHyperlinkViewEditURI + objServiceLevel.IDsp,
											parHyperlinkID: intHyperlinkCounter);
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
										parCalculationMethod: objServiceLevel.CalcualtionMethod,
										parCalculationFormula: objServiceLevel.CalculationFormula,
										parThresholds: objServiceLevel.PerfomanceThresholds,
										parTargets: objServiceLevel.PerformanceTargets,
										parBasicServiceLevelConditions: objServiceLevel.BasicConditions,
										parAdditionalServiceLevelConditions: objDeliverableServiceLevel.AdditionalConditions,
										parErrorMessages: ref listErrorMessagesParameter,
										parNumberingCounter: ref numberingCounter);

									if(listErrorMessagesParameter.Count != this.ErrorMessages.Count)
										this.ErrorMessages = listErrorMessagesParameter;

									objBody.Append(objServiceLevelTable);
									} //if(parDataSet.dsServiceLevels.TryGetValue(
								} // if(objDeliverableServiceLevel.AssociatedServiceLevelID != null)
							} // if(parDataSet.dsDeliverableServiceLevels.TryGetValue(
						else
							{
							// If the entry is not found - write an error in the document and record
							// an error in the error log.
							this.LogError("Error: The DeliverableServiceLevel ID " + node.NodeID
								+ " doesn't exist in SharePoint and it couldn't be retrieved.");
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 8);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "Error: DeliverableServiceLevel: " + node.NodeID + " is missing.",
								parIsNewSection: false,
								parIsError: true);
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);
							break;
							}
						break;
						} //switch (node.NodeType)
					} // foreach(Hierarchy node in this.SelectedNodes)

//++ Insert the Glossary of Terms and Acronym Section
Process_Glossary_and_Acronyms:
				if(this.DictionaryGlossaryAndAcronyms.Count == 0)
					goto Process_Document_Acceptance_Section;

				// Insert the Acronyms and Glossary of Terms scetion
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
					if(this.DictionaryGlossaryAndAcronyms.Count > 0)
						{
						Table tableGlossaryAcronym = new Table();
						tableGlossaryAcronym = CommonProcedures.BuildGlossaryAcronymsTable(
							parSDDPdatacontext: parDataSet.SDDPdatacontext,
							parDictionaryGlossaryAcronym: this.DictionaryGlossaryAndAcronyms,
							parWidthColumn1: Convert.ToInt16(this.PageWith * 0.3),
							parWidthColumn2: Convert.ToInt16(this.PageWith * 0.2),
							parWidthColumn3: Convert.ToInt16(this.PageWith * 0.5),
							parErrorMessages: ref listErrors);
						objBody.Append(tableGlossaryAcronym);
						}     //if(this.TermAndAcronymList.Count > 0)
					} // if (this.Acronyms)

//++ Generate the Document Acceptance Section if it was selected
Process_Document_Acceptance_Section:

				if(this.Document_Acceptance_Section)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_AcceptanceText,
						parIsNewSection: true);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);
					if(this.DocumentAcceptanceRichText != null)
						{
						try
							{
							objHTMLdecoder.DecodeHTML(parClientName: parClientName,
								parMainDocumentPart: ref objMainDocumentPart,
								parDocumentLevel: 1,
								parPageWidthDxa: this.PageWith,
								parPageHeightDxa: this.PageHeight,
								parHTML2Decode: HTMLdecoder.CleanHTML(this.DocumentAcceptanceRichText, parClientName),
								parTableCaptionCounter: ref tableCaptionCounter,
								parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter,
								parPictureNo: ref iPictureNo,
								parHyperlinkID: ref intHyperlinkCounter,
								parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_DocumentCollectionLibraryURI +
												currentHyperlinkViewEditURI + DocumentCollectionID);
							}
						catch (InvalidContentFormatException exc)
							{
							Console.WriteLine("\n\nException occurred: {0}", exc.Message);
							// A content error occurred, record it in the error log.
							this.LogError("Error: in Document Acceptance Section. "
								+ " The Enhance Rich Text column ISD Document Acceptance. " + exc.Message);
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "A content error occurred at this position and valid content could "
								+ "not be interpreted and inserted here. Please review the content "
								+ "in the SharePoint system and correct it. Error Detail: |" + exc.Message + "|",
								parIsNewSection: false,
								parIsError: true);
							if (this.HyperlinkEdit || this.HyperlinkView)
								{
								intHyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
									parMainDocumentPart: ref objMainDocumentPart,
									parImageRelationshipId: hyperlinkImageRelationshipID,
									parHyperlinkID: intHyperlinkCounter,
									parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_ServiceLevelsURI +
												currentHyperlinkViewEditURI + DocumentCollectionID);
								objRun.Append(objDrawing);
								}
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);
							}
						}
					}
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

				//+ Validate the document with OpenXML validator
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
				}

			Console.WriteLine("\t\t End of the generation of {0}", this.DocumentType);
			//- Delete the file from the Documents Directory
			if(File.Exists(path: this.LocalDocumentURI))
				File.Delete(path: this.LocalDocumentURI);
			}
		} // end of ISD_Document_DRM_Inline class
	}