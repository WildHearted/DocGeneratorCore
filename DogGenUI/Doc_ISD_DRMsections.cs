﻿using System;
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
	///<summary>
	///This class represent the Internal Service Definition (ISD) with sperate DRM (Deliverable
	///Report Meeting) sections It inherits from the Internal DRM Sections Class.
	///</summary>
	internal class ISD_Document_DRM_Sections : Internal_DRM_Sections
		{
		///<summary>
		///this option takes the values passed into the method as a list of integers which
		///represents the options the user selected and transposing the values by setting the
		///properties of the object.
		///</summary>
		///<param name="parOptions">
		///The input must represent a List <int>object.</int>
		///</param>
		///<returns>
		///</returns>
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
							case 1:
							this.Introductory_Section = true;
							break;

							case 2:
							this.Introduction = true;
							break;

							case 3:
							this.Executive_Summary = true;
							break;

							case 4:
							this.Service_Portfolio_Section = true;
							break;

							case 5:
							this.Service_Portfolio_Description = true;
							break;

							case 6:
							this.Service_Family_Heading = true;
							break;

							case 7:
							this.Service_Family_Description = true;
							break;

							case 8:
							this.Service_Product_Heading = true;
							break;

							case 9:
							this.Service_Product_Description = true;
							break;

							case 10:
							this.Service_Product_Key_Client_Benefits = true;
							break;

							case 11:
							this.Service_Product_KeyDD_Benefits = true;
							break;

							case 12:
							this.Service_Element_Heading = true;
							break;

							case 13:
							this.Service_Element_Description = true;
							break;

							case 14:
							this.Service_Element_Objectives = true;
							break;

							case 15:
							this.Service_Element_Key_Client_Benefits = true;
							break;

							case 16:
							this.Service_Element_Key_Client_Advantages = true;
							break;

							case 17:
							this.Service_Element_Key_DD_Benefits = true;
							break;

							case 18:
							this.Service_Element_Critical_Success_Factors = true;
							break;

							case 19:
							this.Service_Element_Key_Performance_Indicators = true;
							break;

							case 20:
							this.Service_Element_High_Level_Process = true;
							break;

							case 21:
							this.Deliverables_Reports_Meetings = true;
							break;

							case 27:
							this.DRM_Heading = true;
							break;

							case 22:
							this.DRM_Summary = true;
							break;

							case 23:
							this.Service_Levels = true;
							break;

							case 24:
							this.Service_Level_Heading = true;
							break;

							case 25:
							this.Service_Level_Commitments_Table = true;
							break;

							case 26:
							this.Activities = true;
							break;

							case 28:
							this.Activity_Heading = true;
							break;

							case 29:
							this.Activity_Description_Table = true;
							break;

							case 32:
							this.DRM_Section = true;
							break;

							case 33:
							this.Deliverables = true;
							break;

							case 34:
							this.Deliverable_Heading = true;
							break;

							case 35:
							this.Deliverable_Description = true;
							break;

							case 36:
							this.Deliverable_Inputs = true;
							break;

							case 37:
							this.Deliverable_Outputs = true;
							break;

							case 38:
							this.DDs_Deliverable_Obligations = true;
							break;

							case 39:
							this.Clients_Deliverable_Responsibilities = true;
							break;

							case 40:
							this.Deliverable_Exclusions = true;
							break;

							case 41:
							this.Deliverable_Governance_Controls = true;
							break;

							case 42:
							this.Reports = true;
							break;

							case 43:
							this.Report_Heading = true;
							break;

							case 44:
							this.Report_Description = true;
							break;

							case 45:
							this.DDs_Report_Obligations = true;
							break;

							case 46:
							this.Clients_Report_Responsibilities = true;
							break;

							case 47:
							this.Report_Exclusions = true;
							break;

							case 48:
							this.Report_Governance_Controls = true;
							break;

							case 49:
							this.Meetings = true;
							break;

							case 50:
							this.Meeting_Heading = true;
							break;

							case 51:
							this.Meeting_Description = true;
							break;

							case 52:
							this.DDs_Meeting_Obligations = true;
							break;

							case 53:
							this.Clients_Meeting_Responsibilities = true;
							break;

							case 54:
							this.Meeting_Exclusions = true;
							break;

							case 55:
							this.Meeting_Governance_Controls = true;
							break;

							case 56:
							this.Acronyms_Glossary_of_Terms_Section = true;
							break;

							case 57:
							this.Acronyms = true;
							break;

							case 58:
							this.Glossary_of_Terms = true;
							break;

							case 59:
							this.Document_Acceptance_Section = true;
							break;

							default:
							//-| just ignore
							break;
							}
						} //-| foreach(int option in parOptions)
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
			bool layerHeadingWritten = false;
			bool drmHeading = false;
			Table objActivityTable = new Table();
			Table objServiceLevelTable = new Table();
			Dictionary<int, string> dictDeliverables = new Dictionary<int, string>();
			Dictionary<int, string> dictReports = new Dictionary<int, string>();
			Dictionary<int, string> dictMeetings = new Dictionary<int, string>();
			Dictionary<int, string> dictSLAs = new Dictionary<int, string>();
			List<Deliverable> listDeliverables = new List<Deliverable>();
			int? layer1upElementID = 0;
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

				//-| define a new objOpenXMLdocument
				oxmlDocument objOXMLdocument = new oxmlDocument();
				//-| use CreateDocumentFromTemplate method to create a new MS Word Document based on
				//-| the relevant template
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

				//- Validate if the user selected any content to be generated
				if(this.SelectedNodes == null || this.SelectedNodes.Count < 1)
					{//- if nothing selected thow exception and exit
					throw new NoContentSpecifiedException("No content was specified/selected, therefore the document will be blank. "
						+ "Please specify/select content before submitting the document collection for generation.");
					}

				//-| Create and open the new Document
				this.DocumentStatus = enumDocumentStatusses.Creating;
				//-| Open the MS Word document in Edit mode
				WordprocessingDocument objWPdocument = WordprocessingDocument.Open(path: objOXMLdocument.LocalURI, isEditable: true);
				//-| Define all open XML object to use for building the document
				MainDocumentPart objMainDocumentPart = objWPdocument.MainDocumentPart;
				Body objBody = objWPdocument.MainDocumentPart.Document.Body;          //-| Define the objBody of the document
				Paragraph objParagraph = new Paragraph();
				ParagraphProperties objParaProperties = new ParagraphProperties();
				Run objRun = new Run();
				RunProperties objRunProperties = new RunProperties();
				Text objText = new Text();
				//-| Declare the HTMLdecoder object and assign the document's WordProcessing Body to
				//-| the WPbody property.
				HTMLdecoder objHTMLdecoder = new HTMLdecoder();
				objHTMLdecoder.WPbody = objBody;

				//-| Define the objects to be used in the construction of the document
				ServicePortfolio objPortfolio = new ServicePortfolio();
				ServiceFamily objFamily = new ServiceFamily();
				ServiceProduct objProduct = new ServiceProduct();
				ServiceElement objElement = new ServiceElement();
				ServiceElement objElementLayer1up = new ServiceElement();
				//ServiceElement objElementLayer2up = new ServiceElement();
				Deliverable objDeliverable = new Deliverable();
				Deliverable objDeliverableLayer1up = new Deliverable();
				Deliverable objDeliverableLayer2up = new Deliverable();
				DeliverableActivity objDeliverableActivity = new DeliverableActivity();
				DeliverableServiceLevel objDeliverableServiceLevel = new DeliverableServiceLevel();
				Activity objActivity = new Activity();
				ServiceLevel objServiceLevel = new ServiceLevel();

				//-| Determine the Page Size for the current Body object.
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
				//-| Subtract the Table/Image Left indentation value from the Page width to ensure the
				//-| table/image fits in the available space.
				this.PageWith -= Convert.ToUInt16(Properties.AppResources.Document_Table_Left_Indent);
				//Console.WriteLine("\t\t Effective pageWidth x pageHeight.: {0} x {1} twips", this.PageWith, this.PageHight);

				//-| Check whether Hyperlinks need to be included and add the image to the Document Body
				if(this.HyperlinkEdit || this.HyperlinkView)
					{
					//Insert and embed the hyperlink image in the document and keep the Image's Relationship ID in a variable for repeated use
					hyperlinkImageRelationshipID = oxmlDocument.Insert_HyperlinkImage(
						parMainDocumentPart: ref objMainDocumentPart,
						parSDDPdatacontext: parSDDPdatacontext);
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

				//-| Check if there was an error during preparation for document generation
				if(this.ErrorMessages != null && this.ErrorMessages.Count > 0)
					{
					goto Close_Document;
					}

				this.DocumentStatus = enumDocumentStatusses.Building;
				
				//+ Insert the Introductory Section
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
					//-| Check if a hyperlink must be inserted
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
							//-| A Table content error occurred, record it in the error log.
							this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
								+ " contains an error in the Enhance Rich Text column Introduction. "
								+ exc.Message);
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
							objRun = oxmlDocument.Construct_RunText(
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
					//-| Check if a hyperlink must be inserted
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
								parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
							}
						catch(InvalidContentFormatException exc)
							{
							Console.WriteLine("\n\nException occurred: {0}", exc.Message);
							//-| A Table content error occurred, record it in the error log.
							this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
								+ " contains an error in the Enhance Rich Text column Executive Summary. "
								+ exc.Message);
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
							objRun = oxmlDocument.Construct_RunText(
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
								objRun.Append(objDrawing);
								}
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);
							}
						}
					}
				//++ Insert the user selected content
				if(this.SelectedNodes.Count <= 0)
					goto Process_Glossary_and_Acronyms;

				foreach(Hierarchy node in this.SelectedNodes)
					{
					Console.Write("\nNode: SEQ:{0} - Level:{1} Type:{2} NodeID:{3}", node.Sequence, node.Level, node.NodeType, node.NodeID);

					switch(node.NodeType)
						{
						//+ServicePortfolio & ServiceFramework
						case enumNodeTypes.FRA:  //-| Service Framework
						case enumNodeTypes.POR:  //Service Portfolio
							{
							currentContentLayer = "None";
							if(this.Service_Portfolio_Section)
								{
								objPortfolio = ServicePortfolio.Read(parIDsp: node.NodeID);
								if (objPortfolio != null)
									{
									Console.Write("\t + {0} - {1}", objPortfolio.IDsp, objPortfolio.Title);

									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: objPortfolio.ISDheading,
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
										if(objPortfolio.ISDdescription != null)
											{
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_ServicePortfoliosURI +
												currentHyperlinkViewEditURI + objPortfolio.IDsp;
											try
												{
												objHTMLdecoder.DecodeHTML(parClientName: parClientName,
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 1,
													parHTML2Decode: HTMLdecoder.CleanHTML(objPortfolio.ISDdescription, parClientName),
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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Service Portfolio ID: " + node.NodeID
													+ " contains an error in the Enhance Rich Text column ISD Description. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
								} //-|if(this.Service_Portfolio_Section)
							break;
							}

						//+Service Family
						case enumNodeTypes.FAM:  //-| Service Family
							{
							currentContentLayer = "None";
							if(this.Service_Family_Heading)
								{
								//-| Get the entry from the Database
								objFamily = ServiceFamily.Read(parIDsp: node.NodeID);
								if (objFamily != null)
									{
									Console.Write("\t + {0} - {1}", objFamily.IDsp, objFamily.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: objFamily.ISDheading,
										parIsNewSection: false);
									//-| Check if a hyperlink must be inserted
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
									//-| Check if the user specified to include the Service Family Description
									if(this.Service_Family_Description)
										{
										if(objFamily.ISDdescription != null)
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
													parHTML2Decode: HTMLdecoder.CleanHTML(objFamily.ISDdescription, parClientName),
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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Service Family ID: " + node.NodeID
													+ " contains an error in the Enhance Rich Text column ISD Description. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
									break;
									}
								} //-|if(this.Service_Portfolio_Section)
							break;
							}
						//+Service Product
						case enumNodeTypes.PRO:  //-| Service Product
							{
							currentContentLayer = "None";
							if(this.Service_Product_Heading)
								{
								//-| Get the entry from the Database
								objProduct = ServiceProduct.Read(parIDsp: node.NodeID);
								if (objProduct != null)
									{
									Console.Write("\t + {0} - {1}", objProduct.IDsp, objProduct.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: objProduct.ISDheading,
										parIsNewSection: false);
									//-| Check if a hyperlink must be inserted
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
									//-| Check if the user specified to include the Service Product Description
									if(this.Service_Product_Description)
										{
										if(objProduct.ISDdescription != null)
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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Service Product ID: " + node.NodeID
													+ " contains an error in the Enhance Rich Text column ISD Description. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											}
										}
									if(this.Service_Product_KeyDD_Benefits)
										{
										if(objProduct.KeyDDbenefits != null)
											{
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_ServiceProductsURI +
												currentHyperlinkViewEditURI +
												objProduct.IDsp;
											Console.WriteLine("\t\t + {0} - {1}", objProduct.IDsp,
												Properties.AppResources.Document_Product_KeyDD_Benefits);
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Product_KeyDD_Benefits,
												parIsNewSection: false);
											//-| Check if a hyperlink must be inserted
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
											try
												{
												objHTMLdecoder.DecodeHTML(parClientName: parClientName,
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: HTMLdecoder.CleanHTML(objProduct.KeyDDbenefits, parClientName),
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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Service Product ID: " + node.NodeID
													+ " contains an error in the Enhance Rich Text column Key DD Benefits. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											}
										}

									if(this.Service_Product_Key_Client_Benefits)
										{
										if(objProduct.KeyClientBenefits != null)
											{
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_ServiceProductsURI +
												currentHyperlinkViewEditURI +
												objProduct.IDsp;

											Console.WriteLine("\t\t + {0} - {1}", objProduct.IDsp,
												Properties.AppResources.Document_Product_ClientKeyBenefits);
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Product_ClientKeyBenefits,
												parIsNewSection: false);
											//-| Check if a hyperlink must be inserted
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
											try
												{
												objHTMLdecoder.DecodeHTML(parClientName: parClientName,
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: HTMLdecoder.CleanHTML(objProduct.KeyClientBenefits, parClientName),
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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Service Product ID: " + node.NodeID
													+ " contains an error in the Enhance Rich Text column Key Client Benefits. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
									this.LogError("Error: The Service Product ID " + node.NodeID
										+ " doesn't exist in SharePoint and couldn't be retrieved.");
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Error: Service Family " + node.NodeID + " is missing.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									}
								} //if(this.Service_Product_Heading)
							break;
							}

						//+Service Element
						case enumNodeTypes.ELE:
						currentContentLayer = "None";
						if(!this.Service_Element_Heading)
							break;

						//-| Get the entry from the Database
						objElement = ServiceElement.Read(parIDsp: node.NodeID);
						if (objElement != null)
							{
							Console.Write("\t + {0} - {1}", objElement.IDsp, objElement.Title);

							//-| Insert the Service Element ISD Heading...
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
							objRun = oxmlDocument.Construct_RunText(parText2Write: objElement.ISDheading);
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);

							//-|Check if the Element Layer0up has Content Layers and Content Predecessors
							if (objElement.ContentLayer == "Layer 2")
								{
								if (objElement.ContentPredecessorElementIDsp == null)
									{
									layer1upElementID = null;
									objElementLayer1up = null;
									}
								else
									{
									//-| Set the Layer1up object from the Database
									objElementLayer1up = ServiceElement.Read(parIDsp: Convert.ToInt16(objElement.ContentPredecessorElementIDsp));
									if (objElementLayer1up != null)
										layer1upElementID = objElementLayer1up.IDsp;
									else
										{
										layer1upElementID = null;
										objElementLayer1up = null;
										}
									}
								}
							else
								{
								objElementLayer1up = null;
								layer1upElementID = null;
								}

							//+ Include the Service Element Description
							if(this.Service_Element_Description)
								//- Insert Layer1up if present and required
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upElementID != null
								&& objElementLayer1up.ISDdescription != null)
									{
									//-| Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElementLayer1up.IDsp;
										}
									else
										currentListURI = "";

									//- Set the Content Layer Colour Coding
									
									if (this.ColorCodingLayer1 && objElementLayer1up.ContentLayer == "Layer 1")
										currentContentLayer = "Layer1";
									else
										currentContentLayer = "None";


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
											parHyperlinkID: ref hyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										//-| A Table content error occurred, record it in the error log.
										this.LogError("Error: The Service Element ID: " + node.NodeID
											+ " contains an error in the Enhance Rich Text column ISD Description. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
										objRun = oxmlDocument.Construct_RunText(
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
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									} //- if(layer1upElementID != null)

							//-| Insert Layer0up if not null
							if(objElement.ISDdescription != null)
								{
								//-| Check if a hyperlink must be inserted
								if(documentCollection_HyperlinkURL != "")
									{
									hyperlinkCounter += 1;
									currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
										Properties.AppResources.List_ServiceElementsURI +
										currentHyperlinkViewEditURI +
										objElement.IDsp;
									}
								else
									currentListURI = "";

								//- Set the Content Layer Colour Coding
								if (this.ColorCodingLayer2 && objElement.ContentLayer == "Layer 2")
									currentContentLayer = "Layer2";
								else
									currentContentLayer = "None";

								try
									{
									objHTMLdecoder.DecodeHTML(parClientName: parClientName,
										parMainDocumentPart: ref objMainDocumentPart,
										parDocumentLevel: 4,
										parHTML2Decode: HTMLdecoder.CleanHTML(objElement.ISDdescription,parClientName),
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
									//-| A Table content error occurred, record it in the error log.
									this.LogError("Error: The Service Element ID: " + node.NodeID
										+ " contains an error in the Enhance Rich Text column ISD Description. "
										+ exc.Message);
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
									objRun = oxmlDocument.Construct_RunText(
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
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
								}

							//+ Insert the Service Element Objectives
							//-| Check if the user specified to include the Service Service Element Objectives
							if(this.Service_Element_Objectives)
								{
								//-| Insert the heading
								layerHeadingWritten = false;
								//Prepeare the heading paragraph to be inserted, but only insert it if required...
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Element_Objectives,
									parIsNewSection: false);
								objParagraph.Append(objRun);
								layerHeadingWritten = false;

								//-| Insert Layer1up if present and not null
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upElementID != null)
									{
									if(objElementLayer1up.Objectives != null)
										{
										//- insert the Heading if not inserted yet.
										if(!layerHeadingWritten)
											{
											objBody.Append(objParagraph);
											layerHeadingWritten = true;
											}
										//- Check if a hyperlink must be inserted
										if(documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElementLayer1up.IDsp;
											}
										else
											currentListURI = "";

										//- Set the Content Layer Colour Coding
										
										if(this.ColorCodingLayer1 && objElementLayer1up.ContentLayer == "Layer 1")
											currentContentLayer = "Layer1";
										else
											currentContentLayer = "None";

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
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											//-| A Table content error occurred, record it in the error log.
											this.LogError("Error: The Service Element ID: " + node.NodeID
												+ " contains an error in the Enhance Rich Text column Objectives. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}
									} //- if(layer-upElementID != null)

								//-| Insert Layer0up if not null
								if(objElement.Objectives != null)
									{
									//- insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}

									//- Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
										Properties.AppResources.List_ServiceElementsURI +
										currentHyperlinkViewEditURI +
										objElement.IDsp;
										}
									else
										currentListURI = "";

									//- Set the Content Layer Colour Coding
									
									if (this.ColorCodingLayer2 && objElement.ContentLayer == "Layer 2")
										currentContentLayer = "Layer2";
									else
										currentContentLayer = "None";

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
											parHyperlinkID: ref hyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										//-| A Table content error occurred, record it in the error log.
										this.LogError("Error: The Service Element ID: " + node.NodeID
											+ " contains an error in the Enhance Rich Text column Objectives. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
										objRun = oxmlDocument.Construct_RunText(
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
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}
								} //- if(this.Service_Element_Objectives)

							//+ Insert the Critical Success Factors
							//- Check if the user specified to include the Service Service Element Critical Success Factors
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
								&& layer1upElementID != null)
									{
									if(objElementLayer1up.CriticalSuccessFactors != null)
										{
										//- insert the Heading if not inserted yet.
										if(!layerHeadingWritten)
											{
											objBody.Append(objParagraph);
											layerHeadingWritten = true;
											}
										//- Check if a hyperlink must be inserted
										if(documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElementLayer1up.IDsp;
											}
										else
											currentListURI = "";

										//- Set the Content Layer Colour Coding
										
										if(this.ColorCodingLayer1 && objElementLayer1up.ContentLayer == "Layer 1")
											currentContentLayer = "Layer1";
										else
											currentContentLayer = "None";
										
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
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											//-| A Table content error occurred, record it in the error log.
											this.LogError("Error: The Service Element ID: " + node.NodeID
												+ " contains an error in the Enhance Rich Text column Critical Success Factors. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}
									} //- if(layer1upElementID != null)

								//-| Insert Layer0up if not null
								if(objElement.CriticalSuccessFactors != null)
									{
									//- insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}
									//- Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
										Properties.AppResources.List_ServiceElementsURI +
										currentHyperlinkViewEditURI +
										objElement.IDsp;
										}
									else
										currentListURI = "";

									//- Set the Content Layer Colour Coding									
									if (this.ColorCodingLayer2 && objElement.ContentLayer == "Layer 2")
										currentContentLayer = "Layer2";
									else
										currentContentLayer = "None";


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
											parHyperlinkID: ref hyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										//-| A Table content error occurred, record it in the error log.
										this.LogError("Error: The Service Element ID: " + node.NodeID
											+ " contains an error in the Enhance Rich Text column Critical Success Factors. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
										objRun = oxmlDocument.Construct_RunText(
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
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}
								} //- if(this.Service_Element_CriticalSuccessFactors)

							//+ Insert the Key Client Advantages
							//- Check if the user specified to include the Service Service Key Client Advantages
							if(this.Service_Element_Key_Client_Advantages)
								{
								//- Insert the heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Element_ClientKeyAdvantages,
									parIsNewSection: false);
								objParagraph.Append(objRun);
								layerHeadingWritten = false;

								//- Insert Layer 2up if present and not null
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upElementID != null)
									{
									if(objElementLayer1up.KeyClientAdvantages != null)
										{
										//- insert the Heading if not inserted yet.
										if(!layerHeadingWritten)
											{
											objBody.Append(objParagraph);
											layerHeadingWritten = true;
											}
										//- Check if a hyperlink must be inserted
										if(documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												objElementLayer1up.IDsp;
											}
										else
											currentListURI = "";

										//- Set the Content Layer Colour Coding
										
										if(this.ColorCodingLayer1 && objElementLayer1up.ContentLayer == "Layer 1")
											currentContentLayer = "Layer1";
										else
											currentContentLayer = "None";
										
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
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											//-| A Table content error occurred, record it in the error log.
											this.LogError("Error: The Service Element ID: " + node.NodeID
												+ " contains an error in the Enhance Rich Text column Key client Advantages. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}
									} //- if(layer1upElementID != null)

								//-| Insert Layer 0up if not null
								if(objElement.KeyClientAdvantages != null)
									{
									//- insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}
									//- Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElement.IDsp;
										}
									else
										currentListURI = "";

									//- Set the Content Layer Colour Coding
									if (this.ColorCodingLayer2 && objElement.ContentLayer == "Layer 2")
										currentContentLayer = "Layer2";
									else
										currentContentLayer = "None";

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
											parHyperlinkID: ref hyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										//-| A Table content error occurred, record it in the error log.
										this.LogError("Error: The Service Element ID: " + node.NodeID
											+ " contains an error in the Enhance Rich Text column Key Client Advantages. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
										objRun = oxmlDocument.Construct_RunText(
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
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}
								} //- if(this.Service_Element_Key Client Advantages)

							//+ Insert Key Client Benefits
							//- Check if the user specified to include the Service Element Key Client Benefits
							if(this.Service_Element_Key_Client_Benefits)
								{
								//- Insert the heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Element_ClientKeyBenefits,
									parIsNewSection: false);
								objParagraph.Append(objRun);
								layerHeadingWritten = false;

								//-| Insert Layer1up if present and not null
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upElementID != null)
									{
									if(objElementLayer1up.KeyClientBenefits != null)
										{
										//- insert the Heading if not inserted yet.
										if(!layerHeadingWritten)
											{
											objBody.Append(objParagraph);
											layerHeadingWritten = true;
											}

										//- Check if a hyperlink must be inserted
										if(documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												objElementLayer1up.IDsp;
											}
										else
											currentListURI = "";

										//- Set the Content Layer Colour Coding
										
										if(this.ColorCodingLayer1 && objElementLayer1up.ContentLayer == "Layer 1")
											currentContentLayer = "Layer1";
										else
											currentContentLayer = "None";

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
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											//-| A Table content error occurred, record it in the error log.
											this.LogError("Error: The Service Element ID: " + node.NodeID
												+ " contains an error in the Enhance Rich Text column Key Client Benefits. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}
									} //- if(layer1upElementID != null)

								//-| Insert Layer0up if not null
								if(objElement.KeyClientBenefits != null)
									{
									//- insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}

									//-| Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElement.IDsp;
										}
									else
										currentListURI = "";

									//- Set the Content Layer Colour Coding
									
									if (this.ColorCodingLayer2 && objElement.ContentLayer == "Layer 2")
										currentContentLayer = "Layer2";
									else
										currentContentLayer = "None";
									
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
											parHyperlinkID: ref hyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										//-| A Table content error occurred, record it in the error log.
										this.LogError("Error: The Service Element ID: " + node.NodeID
											+ " contains an error in the Enhance Rich Text column Key client Benefits. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
										objRun = oxmlDocument.Construct_RunText(
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
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}
								} //- if(this.Service_Element_KeyClientBenefits)

							//+ Insert the Key DD Benefits
							//- Check if the user specified to include the Service  Element Key DD Benefits
							if(this.Service_Element_Key_DD_Benefits)
								{
								//-| Insert the heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Element_KeyDDBenefits,
									parIsNewSection: false);
								objParagraph.Append(objRun);
								layerHeadingWritten = false;

								//-| Insert Layer1up if present and not null
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upElementID != null)
									{
									if(objElementLayer1up.KeyDDbenefits != null)
										{
										//- insert the Heading if not inserted yet.
										if(!layerHeadingWritten)
											{
											objBody.Append(objParagraph);
											layerHeadingWritten = true;
											}
										//- Check if a hyperlink must be inserted
										if(documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												objElementLayer1up.IDsp;
											}
										else
											currentListURI = "";

										//- Set the Content Layer Colour Coding
										
										if(this.ColorCodingLayer1 && objElementLayer1up.ContentLayer == "Layer 1")
											currentContentLayer = "Layer1";
										else
											currentContentLayer = "None";

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
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											//-| A Table content error occurred, record it in the error log.
											this.LogError("Error: The Service Element ID: " + node.NodeID
												+ " contains an error in the Enhance Rich Text column Key DD Benefits. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}
									} //- if(layer1upElementID != null)

								//-| Insert Layer0up if not null
								if(objElement.KeyDDbenefits != null)
									{
									//- insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}
									//- Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElement.IDsp;
										}
									else
										currentListURI = "";

									//- Set the Content Layer Colour Coding
									
									if (this.ColorCodingLayer2 && objElement.ContentLayer == "Layer 2")
										currentContentLayer = "Layer2";
									else
										currentContentLayer = "None";

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
											parHyperlinkID: ref hyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										//-| A Table content error occurred, record it in the error log.
										this.LogError("Error: The Service Element ID: " + node.NodeID
											+ " contains an error in the Enhance Rich Text column Key DD Benefits. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
										objRun = oxmlDocument.Construct_RunText(
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
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}
								} //- if(this.Service_Element_KeyDDbenefits)

							//+ Insert the Key Performance Indicators
							//- Check if the user specified to include the Service Element Key Performance Indicators
							if(this.Service_Element_Description)
								{
								//- Insert the heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Element_KPI,
									parIsNewSection: false);
								objParagraph.Append(objRun);
								layerHeadingWritten = false;

								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upElementID != null)
									{
									if(objElementLayer1up.KeyPerformanceIndicators != null)
										{
										//- Insert the Heading if not inserted yet.
										if(!layerHeadingWritten)
											{
											objBody.Append(objParagraph);
											layerHeadingWritten = true;
											}
										//- Check if a hyperlink must be inserted
										if(documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												objElementLayer1up.IDsp;
											}
										else
											currentListURI = "";

										//- Set the Content Layer Colour Coding
										
										if(this.ColorCodingLayer1 && objElementLayer1up.ContentLayer == "Layer 1")
											currentContentLayer = "Layer1";
										else
											currentContentLayer = "None";

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
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											//-| A Table content error occurred, record it in the error log.
											this.LogError("Error: The Service Element ID: " + node.NodeID
												+ " contains an error in the Enhance Rich Text column Key Performance Indicators. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}
									} //- if(layer1upElementID != null)

								//-| Insert Layer0up if not null
								if(objElement.KeyPerformanceIndicators != null)
									{
									//- insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}
									//- Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElement.IDsp;
										}
									else
										currentListURI = "";

									//- Set the Content Layer Colour Coding
									
									if (this.ColorCodingLayer2 && objElement.ContentLayer == "Layer 2")
										currentContentLayer = "Layer2";
									else
										currentContentLayer = "None";

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
											parHyperlinkID: ref hyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightDxa: this.PageHeight,
											parPageWidthDxa: this.PageWith, parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
										}
									catch(InvalidContentFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										//-| A Table content error occurred, record it in the error log.
										this.LogError("Error: The Service Element ID: " + node.NodeID
											+ " contains an error in the Enhance Rich Text column Key Performance Indicators. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
										objRun = oxmlDocument.Construct_RunText(
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
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}
								} //- if(this.Service_Element_KeyPerformanceIndicators)

							//+ Insert the High Level Process
							//- Check if the user specified to include the Service  Element High Level Process
							if(this.Service_Element_High_Level_Process)
								{
								//- Insert the heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Element_HighLevelProcess,
									parIsNewSection: false);
								objParagraph.Append(objRun);
								layerHeadingWritten = false;

								//-| Insert Layer1up if present and not null
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upElementID != null)
									{
									if(objElementLayer1up.ProcessLink != null)
										{
										//- insert the Heading if not inserted yet.
										if(!layerHeadingWritten)
											{
											objBody.Append(objParagraph);
											layerHeadingWritten = true;
											}
										//- Check if a hyperlink must be inserted
										if(documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												objElementLayer1up.IDsp;
											}
										else
											currentListURI = "";

										//- Set the Content Layer Colour Coding
										
										if(this.ColorCodingLayer1 && objElementLayer1up.ContentLayer == "Layer 1")
											currentContentLayer = "Layer1";
										else
											currentContentLayer = "None";

										try
											{
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 5,
												parHTML2Decode: HTMLdecoder.CleanHTML(objElementLayer1up.ProcessLink, parClientName),
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
											//-| A Table content error occurred, record it in the error log.
											this.LogError("Error: The Service Element ID: " + node.NodeID
												+ " contains an error in the Enhance Rich Text column Process Link. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}
									} //- if(layer1upElementID != null)

								//-| Insert Layer0up if not null
								if(objElement.ProcessLink != null)
									{
									//- insert the Heading if not inserted yet.
									if(!layerHeadingWritten)
										{
										objBody.Append(objParagraph);
										layerHeadingWritten = true;
										}
									//- Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
											Properties.AppResources.List_ServiceElementsURI +
											currentHyperlinkViewEditURI +
											objElement.IDsp;
										}
									else
										currentListURI = "";

									//- Set the Content Layer Colour Coding
									if (this.ColorCodingLayer2 && objElement.ContentLayer == "Layer 2")
										currentContentLayer = "Layer2";
									else
										currentContentLayer = "None";

									try
										{
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: HTMLdecoder.CleanHTML(objElement.ProcessLink, parClientName),
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
										//-| A Table content error occurred, record it in the error log.
										this.LogError("Error: The Service Element ID: " + node.NodeID
											+ " contains an error in the Enhance Rich Text column Process Link. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
										objRun = oxmlDocument.Construct_RunText(
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
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}
								} //-| if(this.Service_Element_HighLevelProcess)
							}
						else
							{
							//-| If the entry is not found - write an error in the document and record
							//-| an error in the error log.
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

						//+ Deliverable, Reports, Meetings
						case enumNodeTypes.ELD:  //-| Deliverable associated with Element
						case enumNodeTypes.ELR:  //-| Report deliverable associated with Element
						case enumNodeTypes.ELM:  //-| Meeting deliverable associated with Element
							{
							if(this.DRM_Heading)
								{
								if(drmHeading == false)
									{
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableReportsMeetings_Heading);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									drmHeading = true;
									}
								}
							//-| Get the entry from the Database
							objDeliverable = Deliverable.Read(parIDsp: node.NodeID);
							if (objDeliverable != null)
								{
								Console.Write("\t + {0} - {1}", objDeliverable.IDsp, objDeliverable.Title);

								//- Insert the Deliverable ISD Heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
								objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.ISDheading);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);

								//- Add the deliverable/report/meeting to the Dictionary for inclusion in the DRM section
								if(node.NodeType == enumNodeTypes.ELD) //-| Deliverable
									{
									if(dictDeliverables.ContainsKey(objDeliverable.IDsp) != true)
										dictDeliverables.Add(objDeliverable.IDsp, objDeliverable.ISDheading);
									}
								else if(node.NodeType == enumNodeTypes.ELR) //-| Report
									{
									if(dictReports.ContainsKey(objDeliverable.IDsp) != true)
										dictReports.Add(objDeliverable.IDsp, objDeliverable.ISDheading);
									}
								else if(node.NodeType == enumNodeTypes.ELM) //-| Meeting
									{
									if(dictMeetings.ContainsKey(objDeliverable.IDsp) != true)
										dictMeetings.Add(objDeliverable.IDsp, objDeliverable.ISDheading);
									}

								//- Check if the Deliverable Layer0up has a Content Predecessors
								if (objDeliverable.ContentLayer == "Layer 2")
									{
									if (objDeliverable.ContentPredecessorDeliverableIDsp == null)
										{
										layer1upDeliverableID = null;
										objDeliverableLayer1up = null;
										}
									else
										{
										//-| Get the entry from the Database
										objDeliverableLayer1up = Deliverable.Read(parIDsp: Convert.ToInt16(objDeliverable.ContentPredecessorDeliverableIDsp));
										if (objDeliverableLayer1up != null)
											layer1upDeliverableID = objDeliverableLayer1up.IDsp;
										else
											layer1upDeliverableID = null;
										}
									}
								else
									{
									objDeliverableLayer1up = null;
									layer1upDeliverableID = null;
									}

								//+ Include the Deliverable Summary
								if (this.DRM_Summary)
									{
									//- Insert Layer1up if present and not null
									if (this.PresentationMode == enumPresentationMode.Layered
									&& objDeliverableLayer1up != null
									&& objDeliverableLayer1up.ISDsummary != null)
										{
										//- Check for Colour coding Layers and add if necessary
										if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
											currentContentLayer = "Layer1";
										else
											currentContentLayer = "None";

										//- Insert the Summary Text
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: HTMLdecoder.CleanText(objDeliverableLayer1up.ISDsummary, parClientName),
											parContentLayer: currentContentLayer);

										//-| Check if a hyperlink must be inserted
										if (documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverableLayer1up.IDsp;

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
										} //- if(objDeliverableLayer1up.ISDsummary != null)

									//-| Insert Layer0up if present and not null
									if (objDeliverable.ISDsummary != null)
										{
										//- Check for Colour coding Layers and add if necessary
										if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
											currentContentLayer = "Layer2";
										else
											currentContentLayer = "None";

										if (documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverable.IDsp;
											}
										else
											currentListURI = "";

										//- Insert the Summary Text
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
										objRun = oxmlDocument.Construct_RunText(parText2Write: HTMLdecoder.CleanText(
											objDeliverable.ISDsummary, parClientName),
											parContentLayer: currentContentLayer);

										if (documentCollection_HyperlinkURL != "")
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
										} //- if(objDeliverable.ISDsummary != null)
									} //- if (this.DRM_Summary)

								//-| Insert the hyperlink to the **bookmark to the Deliverable's
								//-| relevant position** in the DRM Section.
								objParagraph = oxmlDocument.Construct_BookmarkHyperlink(
								parBodyTextLevel: 6,
								parBookmarkValue: "Deliverable_" + objDeliverable.IDsp);
								objBody.Append(objParagraph);
								}
							else
								{
								//-| If the entry is not found - write an error in the document and
								//-| record an error in the error log.
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

						//+ Activities
						case enumNodeTypes.EAC:  //-| Activity associated with Deliverable pertaining to Service Element
							{
							currentContentLayer = "None";
							if(this.Activities)
								{
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Activities_Heading);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);

								//-| Get the entry from the Database
								objActivity = Activity.Read(parIDsp: node.NodeID);
								if (objActivity != null)
									{
									Console.WriteLine("\t\t + {0} - {1}", objActivity.IDsp, objActivity.Title);

									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
									objRun = oxmlDocument.Construct_RunText(parText2Write: objActivity.ISDheading);
									//-| Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion +
												Properties.AppResources.List_ActvitiesURI +
												currentHyperlinkViewEditURI + objActivity.IDsp,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									//-| Check if the user specified to include the Deliverable Description
									if(this.Activity_Description_Table)
										{
										objActivityTable = CommonProcedures.BuildActivityTable(
											parWidthColumn1: Convert.ToInt16(this.PageWith * 0.25),
											parWidthColumn2: Convert.ToInt16(this.PageWith * 0.75),
											parActivityDesciption: objActivity.ISDdescription,
											parActivityInput: objActivity.Inputs,
											parActivityOutput: objActivity.Outputs,
											parActivityAssumptions: objActivity.Assumptions,
											parActivityOptionality: objActivity.Optionality);
										objBody.Append(objActivityTable);
										} //-| if (this.Activity_Description_Table)
									}
								else
									{
									//-| If the entry is not found - write an error in the document and
									//-| record an error in the error log.
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
								} //-| if (this.Activities)
							break;
							}

						//+ Service Levels
						case enumNodeTypes.ESL:  //-| Service Level associated with Deliverable pertaining to Service Element
							{
							currentContentLayer = "None";
							if(this.Service_Level_Heading)
								{
								//-| Populate the Service Level Heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_ServiceLevels_Heading_Text);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);

								//-| Check if the user specified to include the Deliverable Description
								if(this.Service_Level_Commitments_Table)
									{
									//-| Prepare the data which to insert into the Service Level Table
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

												//-| Add the Service Level entry to the Service Level
												//-| Dictionay (list)
												if(dictSLAs.ContainsKey(objServiceLevel.IDsp) != true)
													{
													//-| NOTE: the DeliverableServiceLevel ID is used NOT the ServiceLevel ID.
													dictSLAs.Add(objDeliverableServiceLevel.IDsp, objServiceLevel.ISDheading);
													}

												//-| Insert the Service Level ISD Description
												objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
												objRun = oxmlDocument.Construct_RunText(parText2Write: objServiceLevel.ISDheading);
												//-| Check if a hyperlink must be inserted
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
												//-| Populate the Service Level Table
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
												} //if(parDatabase.dsServiceLevels.TryGetValue(
											else
												{
												//-| If the entry is not found - write an error in the
												//-| document and record an error in the error log.
												this.LogError("Error: The Service Level ID " + node.NodeID
													+ " doesn't exist in SharePoint and it couldn't be retrieved.");
												objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "Error: Service Level: " + node.NodeID + " is missing.",
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												break;
												}
											} //if(objDeliverableServiceLevel.AssociatedServiceLevelID != null)
										} //-| if(parDatabase.dsDeliverableServiceLevels.TryGetValue(
									else
										{
										//-| If the entry is not found - write an error in the document
										//-| and record an error in the error log.
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
										} //-| else if(parDatabase.dsDeliverableServiceLevels.TryGetValue(
									} //-| if (this.Service Level_Description_Table)
								} //-| if (this.Service_Level_Heading)
							break;
							} //case enumNodeTypes.ESL:
						} //switch (node.NodeType)
					} //-| foreach(Hierarchy node in this.SelectedNodes)

				//++ Insert the Deliverable, Report, Meeting (DRM) Section
				Console.Write("\nGenerating Deliverable, Report, Meeting sections...\n");

				if(this.DRM_Section)
					{
					Console.Write("\nGenerating Deliverable, Report, Meeting sections...\n");

					//-| Insert the Deliverables, Reports and Meetings Section if there were any selected...
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

					//+ First the Deliverables
					if(dictDeliverables.Count == 0 || this.Deliverables == false)
						goto Process_Reports;

					Console.Write("\n\tDeliverables:\n");
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Deliverables_Heading_Text);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					string deliverableBookMark = "Deliverable_";
					//+ _Insert the individual Deliverables in the section_
					foreach(var deliverableItem in dictDeliverables.OrderBy(dD => dD.Value))
						{
						if(this.Deliverable_Heading)
							{
							Console.WriteLine("\n\t{0} - {1}", deliverableItem.Key, deliverableItem.Value);
							//-| Get the entry from the Database
							objDeliverable = Deliverable.Read(parIDsp: deliverableItem.Key);
							if (objDeliverable != null)
								{
								Console.Write("\n\t\t + {0} - {1}", objDeliverable.IDsp, objDeliverable.Title);

								//+ Insert the Deliverable's ISD Heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3, parBookMark: deliverableBookMark + objDeliverable.IDsp);
								objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.ISDheading);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);

								//-| Check if the Deliverable contain a Content Layer/Content Predecessors
								if (objDeliverable.ContentLayer == "Layer 2")
									{
									if (objDeliverable.ContentPredecessorDeliverableIDsp == null)
										{
										layer1upDeliverableID = null;
										objDeliverableLayer1up = null;
										}
									else
										{
										//- Get the layer1up entry from the Database
										objDeliverableLayer1up = Deliverable.Read(parIDsp: Convert.ToInt16(objDeliverable.ContentPredecessorDeliverableIDsp));
										if (objDeliverableLayer1up == null)
											{
											layer1upDeliverableID = null;
											objDeliverableLayer1up = null;
											}
										else
											{
											layer1upDeliverableID = objDeliverableLayer1up.IDsp;
											}											
										}
									}
								else
									{
									objDeliverableLayer1up = null;
									layer1upDeliverableID = null;
									}

								//+ Include the Deliverable ISD Description
								if(this.Deliverable_Description)
									{
									//- Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null
									&& objDeliverableLayer1up.ISDdescription != null)
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
										if(this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
											currentContentLayer = "Layer1";
										else
											currentContentLayer = "None";

										try
											{
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverableLayer1up.ISDdescription, parClientName),
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
											//-| A Table content error occurred, record it in the error log.
											this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
												+ " contains an error in the Enhance Rich Text column ISD Description "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //- if(this.PresentationMode == enumPresentationMode.Layered && layer1upDeliverableID != null)

									//-| Insert Layer0up if not null
									if(objDeliverable.ISDdescription != null)
										{
										//-| Check if a hyperlink must be inserted
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
										if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
											currentContentLayer = "Layer2";
										else
											currentContentLayer = "None";

										try
											{
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverable.ISDdescription, parClientName),
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
											//-| A Table content error occurred, record it in the error log.
											this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in the Enhance Rich Text column ISD Description. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}
									} //- if(this.Deliverable_Description)

								//+ Insert Deliverable Inputs
								//-| Check if the user specified to include the Deliverable Inputs
								if(this.Deliverable_Inputs)
									{
									if(objDeliverable.Inputs != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.Inputs != null))
										{
										//-| Insert the Inputs Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										//- Insert Layer1up if present and not null
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
											if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
												currentContentLayer = "Layer1";
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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
													+ " contains an error in the Enhance Rich Text column Inputs. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											}

										//- Insert Layer0up if not null
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
											if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in the Enhance Rich Text column Inputs. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											}
										} //- if(this.Deliverable_Inputs)
									} //- if(this.Deliverable_Inputs)

								//+ Insert the Deliverable Outputs
								if(this.Deliverable_Outputs)
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

										//- Insert Layer1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.Outputs != null)
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
												if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

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
													//-| A Table content error occurred, record it in
													//-| the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
														+ " contains an error in the Enhance Rich Text column Outputs. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
													objRun = oxmlDocument.Construct_RunText(
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
														objRun.Append(objDrawing);
														}
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);
													}
												} //- if(recDeliverable.Layer1up.Outputs != null)
											} //- if(layer1upDeliverableID != null)

										//-| Insert Layer0up if not null
										if(objDeliverable.Outputs != null)
											{
											//-| Check if a hyperlink must be inserted
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
											if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in the Enhance Rich Text column Outputs. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} //- if(recDeliverable.Outputs != null)
										} //- if(recDeliverables.Outputs !== null &&)
									} //- if(this.Deliverable_Outputs)

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

										//-| Insert Layer1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.DDobligations != null)
												{
												//-| Check if a hyperlink must be inserted
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
												if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

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
													//-| A Table content error occurred, record it in
													//-| the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
														+ " contains an error in the Enhance Rich Text column DD's Obligations. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
													objRun = oxmlDocument.Construct_RunText(
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
														objRun.Append(objDrawing);
														}
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);
													}
												} //- if(recDeliverable.Layer1up.DDobligations != null)
											} //- if(layer2upDeliverableID != null)

										//- Insert Layer0up if not null
										if(objDeliverable.DDobligations != null)
											{
											//-| Check if a hyperlink must be inserted
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
											if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in the Enhance Rich Text column DD's Obligations. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} //- if(recDeliverable.DDobligations != null)
										} //- if(recDeliverable.DDoblidations != null &&)
									} //- if(this.DDs_Deliverable_Obligations)

								//+ Insert the Client Responsibilities
								if(this.Clients_Deliverable_Responsibilities)
									{
									if(objDeliverable.ClientResponsibilities != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.ClientResponsibilities != null))
										{
										//- Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										//-| Insert Layer1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.ClientResponsibilities != null)
												{
												//-| Check if a hyperlink must be inserted
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
												if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

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
													//-| A Table content error occurred, record it in
													//-| the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
														+ " contains an error in the Enhance Rich Text column Client's Responsibilities. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
													objRun = oxmlDocument.Construct_RunText(
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
														objRun.Append(objDrawing);
														}
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);
													}
												} //-| if(recDeliverable.Layer1up.ClientResponsibilities != null)
											} //-| if(layer1upDeliverableID != null)

										//-| Insert Layer0up if not null
										if(objDeliverable.ClientResponsibilities != null)
											{
											//-| Check if a hyperlink must be inserted
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

											//- //- Check for Colour coding Layers and add if necessary
											if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in the Enhance Rich Text column Client's Responsibilities. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
										//+ Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										//-| Insert Layer1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.Exclusions != null)
												{
												//-| Check if a hyperlink must be inserted
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
												if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

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
													//-| A Table content error occurred, record it in
													//-| the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
														+ " contains an error in the Enhance Rich Text column Exclusions. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
													objRun = oxmlDocument.Construct_RunText(
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
														objRun.Append(objDrawing);
														}
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);
													}
												} //- if(recDeliverable.Layer1up.Exclusions != null)
											} //- if(layer2upDeliverableID != null)

										//-| Insert Layer0up if not null
										if(objDeliverable.ClientResponsibilities != null)
											{
											//-| Check if a hyperlink must be inserted
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
											if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in the Enhance Rich Text column Exclusions. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} //-| if(recDeliverable.Exclusions != null)
										} //-| if(recDeliverable.Exclusions != null &&)
									} //if(this.Deliverable_Exclusions)

								//+ Insert the Governance Controls
								if(this.Deliverable_Governance_Controls)
									{
									if(objDeliverable.GovernanceControls != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.GovernanceControls != null))
										{
										//- Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableGovernanceControls_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										//-| Insert Layer1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.GovernanceControls != null)
												{
												//-| Check if a hyperlink must be inserted
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
												if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
													currentContentLayer = "Layer1";
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
													//-| A Table content error occurred, record it in
													//-| the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
														+ " contains an error in the Enhance Rich Text column Governance Controls. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
													objRun = oxmlDocument.Construct_RunText(
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
														objRun.Append(objDrawing);
														}
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);
													}
												} //- if(recDeliverable.Layer1up.GovernanceControls != null)
											} //- if(layer2upDeliverableID != null)

										//-| Insert Layer0up if not null
										if(objDeliverable.GovernanceControls != null)
											{
											//-| Check if a hyperlink must be inserted
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
											if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in the Enhance Rich Text column Governance Controls. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} //-| if(recDeliverable.GovernanceControls != null)
										} //-| if(recDeliverable.GovernanceControls != null &&)
									} //if(this.Deliverable_GovernanceControls)

								//+ Check if there are any Glossary Terms or Acronyms associated with the Deliverable(s).
								if(this.Acronyms_Glossary_of_Terms_Section)
									{
									//-| if there are GlossaryAndAcronyms to add from layer0up
									if(objDeliverable.GlossaryAndAcronyms != null)
										{
										if(objDeliverable.GlossaryAndAcronyms != null)
											{
											foreach(var entry in objDeliverable.GlossaryAndAcronyms)
												{
												if(this.ListGlossaryAndAcronyms.Contains(entry) != true)
													ListGlossaryAndAcronyms.Add(entry);
												}
											}
										}
									//-| if there are GlossaryAndAcronyms to add from layer1up
									if(layer1upDeliverableID != null)
										{
										if(objDeliverableLayer1up.GlossaryAndAcronyms != null)
											{
											foreach(var entry in objDeliverableLayer1up.GlossaryAndAcronyms)
												{
												if(this.ListGlossaryAndAcronyms.Contains(entry) != true)
													ListGlossaryAndAcronyms.Add(entry);
												}
											}
										}
									}
								}
							else
								{
								//- If the entry is not found - write an error in the document and record an error in the error log.
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
							} //- if(this.DeliverableHeading
						} //- foreach (KeyValuePair<int, String>.....

					//+ Process Reports
Process_Reports:
					if(dictReports.Count == 0 || this.Reports == false)
						goto Process_Meetings;

					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Reports_Heading_Text);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);
					deliverableBookMark = "Report_";

					//+ Insert the individual Reports in the section
					foreach(KeyValuePair<int, string> reportItem in dictReports.OrderBy(key => key.Value))
						{
						if(this.Report_Heading)
							{
							//- Get the Deliverable(Report) entry from the Database
							objDeliverable = Deliverable.Read(parIDsp: reportItem.Key);
							if (objDeliverable != null)
								{
								Console.Write("\n\t + {0} - {1}", objDeliverable.IDsp, objDeliverable.Title);

								//+ Insert the Reports's ISD Heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3, parBookMark: deliverableBookMark + objDeliverable.IDsp);
								objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.ISDheading);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);

								//-| Check if the Deliverable contain a Content Layer/Content Predecessors
								if (objDeliverable.ContentLayer == "Layer 2")
									{
									if (objDeliverable.ContentPredecessorDeliverableIDsp == null)
										{
										layer1upDeliverableID = null;
										objDeliverableLayer1up = null;
										}
									else
										{
										//- Get the layer1up entry from the Database
										objDeliverableLayer1up = Deliverable.Read(parIDsp: Convert.ToInt16(objDeliverable.ContentPredecessorDeliverableIDsp));
										if (objDeliverableLayer1up == null)
											{
											layer1upDeliverableID = null;
											objDeliverableLayer1up = null;
											}
										else
											{
											layer1upDeliverableID = objDeliverableLayer1up.IDsp;
											}
										}
									}
								else
									{
									objDeliverableLayer1up = null;
									layer1upDeliverableID = null;
									}

								//+ Insert the Deliverable ISD Description
								if (this.Report_Description)
									{
									//- Insert Layer 1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null)
										{
										if(objDeliverableLayer1up.ISDdescription != null)
											{
											//-| Check if a hyperlink must be inserted
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
											if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(parClientName: parClientName,
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverableLayer1up.ISDdescription, parClientName),
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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
													+ " contains an error in the Enhance Rich Text column ISD Description. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											}
										} //- if(layer2upDeliverableID != null)

									//-| Insert Layer0up if not null
									if(objDeliverable.ISDdescription != null)
										{
										//-| Check if a hyperlink must be inserted
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
										if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
											currentContentLayer = "Layer2";
										else
											currentContentLayer = "None";

										try
											{
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverable.ISDdescription, parClientName),
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
											//-| A Table content error occurred, record it in the error log.
											this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in the Enhance Rich Text column ISD Description. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}
									} //- if(this.Report_Description)

								//+ Insert the Report Inputs
								if(this.Report_Inputs)
									{
									if(objDeliverable.Inputs != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.Inputs != null))
										{
										//-| Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										//-| Insert Layer1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.Inputs != null)
												{
												//-| Check if a hyperlink must be inserted
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
												if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
													currentContentLayer = "Layer1";
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
													//-| A Table content error occurred, record it in
													//-| the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
														+ " contains an error in the Enhance Rich Text column Inputs. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
													objRun = oxmlDocument.Construct_RunText(
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
														objRun.Append(objDrawing);
														}
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);
													}
												}
											} //-| if(layer2upDeliverableID != null)

										//-| Insert Layer0up if not null
										if(objDeliverable.Inputs != null)
											{
											//-| Check if a hyperlink must be inserted
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
											if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in the Enhance Rich Text column Inputs. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} //- if(recReport.Inputs != null)
										} //if- (recReports.Inputs != null &&)
									} //- if(this.Report_Inputs)

								//+ Insert the Deliverable Outputs
								if(this.Report_Outputs)
									{
									if(objDeliverable.Outputs != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.Outputs != null))
										{
										//-| Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableOutputs_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										//-| Insert Laye1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.Outputs != null)
												{
												//-| Check if a hyperlink must be inserted
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
												if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

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
													//-| A Table content error occurred, record it in
													//-| the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
														+ " contains an error in the Enhance Rich Text column Outputs. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
													objRun = oxmlDocument.Construct_RunText(
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
														objRun.Append(objDrawing);
														}
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);
													}
												} //- if(recReport.Layer1up.Outputs != null)
											} //- if(layer2upDeliverableID != null)

										//-| Insert Layer0up if not null
										if(objDeliverable.Outputs != null)
											{
											//-| Check if a hyperlink must be inserted
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
											if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in the Enhance Rich Text column Outputs. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} //- if(recReport.Outputs != null)
										} //- if(recReport.Outputs !== null &&)
									} //if- (this.Report_Outputs)

								//+ Insert the Report DD's Obligations
								if(this.DDs_Report_Obligations)
									{
									if(objDeliverable.DDobligations != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.DDobligations != null))
										{
										//-| Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										//-| Insert Layer 2up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.DDobligations != null)
												{
												//-| Check if a hyperlink must be inserted
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
												if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

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
													//-| A Table content error occurred, record it in
													//-| the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
														+ " contains an error in the Enhance Rich Text column DD's Obligations. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
													objRun = oxmlDocument.Construct_RunText(
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
														objRun.Append(objDrawing);
														}
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);
													}
												} //- if(recReport.Layer1up.DDobligations != null)
											} //- if(layer1upDeliverableID != null)

										//-| Insert Layer0up if not null
										if(objDeliverable.DDobligations != null)
											{
											//-| Check if a hyperlink must be inserted
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
											if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in the Enhance Rich Text column DD's Obligations. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} //- if(recReport.DDobligations != null)
										} //- if(recReport.DDoblidations != null &&)
									} //- if(this.DDs_Report_Obligations)

								//+ Insert the Client Responsibilities
								if(this.Clients_Report_Responsibilities)
									{
									if(objDeliverable.ClientResponsibilities != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.ClientResponsibilities != null))
										{
										//-| Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										//-| Insert Layer1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.ClientResponsibilities != null)
												{
												//-| Check if a hyperlink must be inserted
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
												if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

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
													//-| A Table content error occurred, record it in
													//-| the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
														+ " contains an error in the Enhance Rich Text column Client's Responsibilities. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
													objRun = oxmlDocument.Construct_RunText(
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
														objRun.Append(objDrawing);
														}
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);
													}
												} //- if(recReport.Layer1up.ClientResponsibilities != null)
											} //- if(layer1upDeliverableID != null)

										//-| Insert Layer0up if not null
										if(objDeliverable.ClientResponsibilities != null)
											{
											//-| Check if a hyperlink must be inserted
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
											if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in the Enhance Rich Text column Client's Responsibilities. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} //- if(recReport.ClientResponsibilities != null)
										} //- if(recReport.ClientResponsibilities != null &&)
									} //- if(this.Clients_Report_Responsibilities)

								//+ Insert the Deliverable Exclusions
								if(this.Report_Exclusions)
									{
									if(objDeliverable.Exclusions != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.Exclusions != null))
										{
										//-| Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										//-| Insert Layer1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.Exclusions != null)
												{
												//-| Check if a hyperlink must be inserted
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
												if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

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
													//-| A Table content error occurred, record it in
													//-| the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
														+ " contains an error in the Enhance Rich Text column Exclusions. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
													objRun = oxmlDocument.Construct_RunText(
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
														objRun.Append(objDrawing);
														}
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);
													}
												} //- if(recReport.Layer1up.Exclusions != null)
											} //- if(layer2upDeliverableID != null)

										//-| Insert Layer0up if not null
										if(objDeliverable.ClientResponsibilities != null)
											{
											//-| Check if a hyperlink must be inserted
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
											if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
													+ " contains an error in the Enhance Rich Text column Exclusions. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} //- if(recReport.Exclusions != null)
										} //- if(recReport.Exclusions != null &&)
									} //- if(this.Report_Exclusions)

								//+ Insert the Governance Controls
								if(this.Report_Governance_Controls)
									{
									if(objDeliverable.GovernanceControls != null
									|| (layer1upDeliverableID != null && objDeliverableLayer1up.GovernanceControls != null))
										{
										//-| Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableGovernanceControls_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										//-| Insert Layer 1up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered
										&& layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.GovernanceControls != null)
												{
												//-| Check if a hyperlink must be inserted
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
												if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
													currentContentLayer = "Layer1";
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
													//-| A Table content error occurred, record it in
													//-| the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
														+ " contains an error in the Enhance Rich Text column Governance Controls. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
													objRun = oxmlDocument.Construct_RunText(
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
														objRun.Append(objDrawing);
														}
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);
													}
												} //- if(recReport.Layer1up.GovernanceControls != null)
											} //- if(layer2upDeliverableID != null)

										//-| Insert Layer0up if not null
										if(objDeliverable.GovernanceControls != null)
											{
											//-| Check if a hyperlink must be inserted
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
											if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
													+ " contains an error in the Enhance Rich Text column Goverance Controls. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} //- if(recReport.GovernanceControls != null)
										} //- if(recReport.GovernanceControls != null &&)
									} //- if(this.Report_GovernanceControls)

								//+ Check if there are any Glossary Terms or Acronyms associated with the Deliverable(s).
								if(this.Acronyms_Glossary_of_Terms_Section)
									{
									//-| if there are GlossaryAndAcronyms to add from layer0up
									if(objDeliverable.GlossaryAndAcronyms != null)
										{
										if(objDeliverable.GlossaryAndAcronyms != null)
											{
											foreach(var entry in objDeliverable.GlossaryAndAcronyms)
												{
												if(this.ListGlossaryAndAcronyms.Contains(entry) != true)
													ListGlossaryAndAcronyms.Add(entry);
												}
											}
										}
									//-| if there are GlossaryAndAcronyms to add from layer1up
									if(layer1upDeliverableID != null)
										{
										if(objDeliverableLayer1up.GlossaryAndAcronyms != null)
											{
											foreach(var entry in objDeliverableLayer1up.GlossaryAndAcronyms)
												{
												if(this.ListGlossaryAndAcronyms.Contains(entry) != true)
													ListGlossaryAndAcronyms.Add(entry);
												}
											}
										}
									}
								}
							else
								{
								//-| If the entry is not found - write an error in the document and
								//-| record an error in the error log.
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
							} //- if(this.ReportHeading
						} //- foreach (KeyValuePair<int, String>.....

Process_Meetings:   //+ Meetings
					if(dictMeetings.Count == 0 || this.Meetings == false)
						goto Process_ServiceLevels;

					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Meetings_Heading_Text);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);
					deliverableBookMark = "Meeting_";
					//-| Insert the individual Meetings in the section
					foreach(KeyValuePair<int, string> meetingItem in dictMeetings.OrderBy(key => key.Value))
						{
						//-| Get the entry from the Database
						objDeliverable = Deliverable.Read(parIDsp: meetingItem.Key);
						if (objDeliverable != null)
							{
							Console.Write("\n\t + {0} - {1}", objDeliverable.IDsp, objDeliverable.Title);

							//-| Insert the Reports's ISD Heading
							objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3, parBookMark: deliverableBookMark + objDeliverable.IDsp);
							objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.ISDheading);
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);

							//-Check if the Meeting's Layer0up has Content Layers and Content Predecessors
							if (objDeliverable.ContentLayer == "Layer 2")
								{
								if (objDeliverable.ContentPredecessorDeliverableIDsp == null)
									{
									layer1upDeliverableID = null;
									objDeliverableLayer1up = null;
									}
								else
									{
									//- Get the layer1up entry from the Database
									objDeliverableLayer1up = Deliverable.Read(parIDsp: Convert.ToInt16(objDeliverable.ContentPredecessorDeliverableIDsp));
									if (objDeliverableLayer1up == null)
										{
										layer1upDeliverableID = null;
										objDeliverableLayer1up = null;
										}
									else
										{
										layer1upDeliverableID = objDeliverableLayer1up.IDsp;
										}
									}
								}
							else
								{
								objDeliverableLayer1up = null;
								layer1upDeliverableID = null;
								}

							//-| Check if the user specified to include the Deliverable ISD Description
							if (this.Meeting_Description)
								{
								//-| Insert Layer1up if present and not null
								if(this.PresentationMode == enumPresentationMode.Layered
								&& layer1upDeliverableID != null)
									{
									if(objDeliverableLayer1up.ISDdescription != null)
										{
										//-| Check if a hyperlink must be inserted
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
										if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
											currentContentLayer = "Layer1";
										else
											currentContentLayer = "None";

										try
											{
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverableLayer1up.ISDdescription, parClientName),
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
											//-| A Table content error occurred, record it in the error log.
											this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
												+ " contains an error in the Enhance Rich Text column ISD Description. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										}
									} //- if(layer2upDeliverableID != null)

								//-| Insert Layer0up if not null
								if(objDeliverable.ISDdescription != null)
									{
									//-| Check if a hyperlink must be inserted
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
									if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
										currentContentLayer = "Layer2";
									else
										currentContentLayer = "None";

									try
										{
										objHTMLdecoder.DecodeHTML(parClientName: parClientName,
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 4,
											parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverable.ISDdescription, parClientName),
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
										//-| A Table content error occurred, record it in the error log.
										this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
											+ " contains an error in the Enhance Rich Text column ISD Description. "
											+ exc.Message);
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
										objRun = oxmlDocument.Construct_RunText(
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
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									}
								} //- if(this.Meeting_Description)

							//+ Insert the Report Inputs
							if(this.Meeting_Inputs)
								{
								if(objDeliverable.Inputs != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.Inputs != null))
									{
									//-| Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									//-| Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null)
										{
										if(objDeliverableLayer1up.Inputs != null)
											{
											//-| Check if a hyperlink must be inserted
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
											if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
												currentContentLayer = "Layer1";
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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
													+ " contains an error in the Enhance Rich Text column Inputs. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											}
										} //- if(layer2upDeliverableID != null)

									//-| Insert Layer0up if not null
									if(objDeliverable.Inputs != null)
										{
										//-| Check if a hyperlink must be inserted
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
										if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
											currentContentLayer = "Layer2";
										else
											currentContentLayer = "None";

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
											//-| A Table content error occurred, record it in the error log.
											this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in the Enhance Rich Text column Inputs. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //- if(recMeeting.Inputs != null)
									} //- if(recMeeting.Inputs != null &&)
								} //- if(this.Meeting_Inputs)

							//+ Insert the Deliverable Outputs
							if(this.Meeting_Outputs)
								{
								if(objDeliverable.Outputs != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.Outputs != null))
									{
									//-| Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableOutputs_Heading_Text);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									//-| Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null)
										{
										if(objDeliverableLayer1up.Outputs != null)
											{
											//-| Check if a hyperlink must be inserted
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
											if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
													+ " contains an error in the Enhance Rich Text column Outputs. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} //- if(recMeeting.Layer1up.Outputs != null)
										} //- if(layer2upDeliverableID != null)

									//-| Insert Layer0up if not null
									if(objDeliverable.Outputs != null)
										{
										//-| Check if a hyperlink must be inserted
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
										if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
											currentContentLayer = "Layer2";
										else
											currentContentLayer = "None";

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
											//-| A Table content error occurred, record it in the error log.
											this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
												+ " contains an error in the Enhance Rich Text column Outputs. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //- if(recMeeting.Outputs != null)
									} //- if(recMeeting.Outputs !== null &&)
								} //- if(this.Meeting_Outputs)

							//+ Insert the Report DD's Obligations
							if(this.DDs_Report_Obligations)
								{
								if(objDeliverable.DDobligations != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.DDobligations != null))
									{
									//-| Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									//-| Insert Layer 1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null)
										{
										if(objDeliverableLayer1up.DDobligations != null)
											{
											//-| Check if a hyperlink must be inserted
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
											if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(parClientName: parClientName,
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverableLayer1up.DDobligations,parClientName),
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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
													+ " contains an error in the Enhance Rich Text column DD's Obligations. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} //- if(recMeeting.Layer1up.DDobligations != null)
										} //- if(layer2upDeliverableID != null)

									//-| Insert Layer0up if not null
									if(objDeliverable.DDobligations != null)
										{
										//-| Check if a hyperlink must be inserted
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
										if (this.ColorCodingLayer1 && objDeliverable.ContentLayer == "Layer 2")
											currentContentLayer = "Layer2";
										else
											currentContentLayer = "None";

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
											//-| A Table content error occurred, record it in the error log.
											this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in the Enhance Rich Text column DD's Obligations. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //-| if(recMeeting.DDobligations != null)
									} //if(recMeeting.DDoblidations != null &&)
								} //if(this.DDs_Report_Obligations)

							//+ Insert the Client Responsibilities
							if(this.Clients_Report_Responsibilities)
								{
								if(objDeliverable.ClientResponsibilities != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.ClientResponsibilities != null))
									{
									//-| Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									//-| Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null)
										{
										if(objDeliverableLayer1up.ClientResponsibilities != null)
											{
											//-| Check if a hyperlink must be inserted
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
											if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
													+ " contains an error in the Enhance Rich Text column Client's Responsibilities. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} //- if(recMeeting.Layer1up.ClientResponsibilities != null)
										} //- if(layer2upDeliverableID != null)

									//-| Insert Layer0up if not null
									if(objDeliverable.ClientResponsibilities != null)
										{
										//-| Check if a hyperlink must be inserted
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
										if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
											currentContentLayer = "Layer2";
										else
											currentContentLayer = "None";

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
											//-| A Table content error occurred, record it in the error log.
											this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in the Enhance Rich Text column Client's Responsibilities. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //- if(recMeeting.ClientResponsibilities != null)
									} //- if(recMeeting.ClientResponsibilities != null &&)
								} //- if(this.Clients_Report_Responsibilities)

							//+ Insert the Deliverable Exclusions
							if(this.Meeting_Exclusions)
								{
								if(objDeliverable.Exclusions != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.Exclusions != null))
									{
									//-| Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									//-| Insert Layer1up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered
									&& layer1upDeliverableID != null)
										{
										if(objDeliverableLayer1up.Exclusions != null)
											{
											//-| Check if a hyperlink must be inserted
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
											if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
													+ " contains an error in the Enhance Rich Text column Exclusions. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} //- if(recMeeting.Layer1up.Exclusions != null)
										} //- if(layer2upDeliverableID != null)

									//-| Insert Layer0up if not null
									if(objDeliverable.ClientResponsibilities != null)
										{
										//-| Check if a hyperlink must be inserted
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
										if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
											currentContentLayer = "Layer2";
										else
											currentContentLayer = "None";

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
											//-| A Table content error occurred, record it in the error log.
											this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in the Enhance Rich Text column Exclsuions. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //-| if(recMeeting.Exclusions != null)
									} //-| if(recMeeting.Exclusions != null &&)
								} //if(this.Deliverable_Exclusions)

							//+ Insert the Governance Controls
							if(this.Deliverable_Governance_Controls)
								{
								if(objDeliverable.GovernanceControls != null
								|| (layer1upDeliverableID != null && objDeliverableLayer1up.GovernanceControls != null))
									{
									//-| Insert the Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableGovernanceControls_Heading_Text);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									//-| Insert Layer1up if present and not null
									if(layer1upDeliverableID != null)
										{
										if(objDeliverableLayer1up.GovernanceControls != null)
											{
											//-| Check if a hyperlink must be inserted
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
											if (this.ColorCodingLayer1 && objDeliverableLayer1up.ContentLayer == "Layer 1")
												currentContentLayer = "Layer1";
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
												//-| A Table content error occurred, record it in the
												//-| error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.IDsp
													+ " contains an error in the Enhance Rich Text column Governance Controls. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
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
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} //- if(recMeeting.Layer1up.GovernanceControls != null)
										} //- if(layer2upDeliverableID != null)

									//-| Insert Layer0up if not null
									if(objDeliverable.GovernanceControls != null)
										{
										//-| Check if a hyperlink must be inserted
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
										if (this.ColorCodingLayer2 && objDeliverable.ContentLayer == "Layer 2")
											currentContentLayer = "Layer2";
										else
											currentContentLayer = "None";

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
											//-| A Table content error occurred, record it in the error log.
											this.LogError("Error: The Deliverable ID: " + objDeliverable.IDsp
												+ " contains an error in the Enhance Rich Text column Governance Controls. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
											objRun = oxmlDocument.Construct_RunText(
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
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //- if(recMeeting.GovernanceControls != null)
									} //- if(recMeeting.GovernanceControls != null &&)
								} //- if(this.Deliverable_GovernanceControls)

							//+ Check if there are any Glossary Terms or Acronyms associated with the Deliverable(s).
							if(this.Acronyms_Glossary_of_Terms_Section)
								{
								//-| if there are GlossaryAndAcronyms to add from layer0up
								if(objDeliverable.GlossaryAndAcronyms != null)
									{
									if(objDeliverable.GlossaryAndAcronyms != null)
										{
										foreach(var entry in objDeliverable.GlossaryAndAcronyms)
											{
											if(this.ListGlossaryAndAcronyms.Contains(entry) != true)
												ListGlossaryAndAcronyms.Add(entry);
											}
										}
									}
								//-| if there are GlossaryAndAcronyms to add from layer1up
								if(layer1upDeliverableID != null)
									{
									if(objDeliverableLayer1up.GlossaryAndAcronyms != null)
										{
										foreach(var entry in objDeliverableLayer1up.GlossaryAndAcronyms)
											{
											if(this.ListGlossaryAndAcronyms.Contains(entry) != true)
												ListGlossaryAndAcronyms.Add(entry);
											}
										}
									}
								} //-| if(this.Acronyms_Glossary_of_Terms_Section)
							}
						else
							{
							//-| If the entry is not found - write an error in the document and record
							//-| an error in the error log.
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
						} //-| foreach.....
					} //if(this.DRM_Section)

Process_ServiceLevels: //++ Insert the Service Levels Section
				if(this.Service_Level_Section)
					{
					//-| Insert the Service If any are relevant
					if(dictSLAs.Count > 0)
						{
						objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
						objRun = oxmlDocument.Construct_RunText(
							parText2Write: Properties.AppResources.Document_ServiceLevel_Section_Text,
							parIsNewSection: true);
						objParagraph.Append(objRun);
						objBody.Append(objParagraph);

						string servicelevelBookMark = "ServiceLevel_";
						//-| Insert the individual Service Levels in the section
						foreach(KeyValuePair<int, string> servicelevelItem in dictSLAs.OrderBy(sortkey => sortkey.Value))
							{
							//-| Prepare the data which to insert into the Service Level Table
							objDeliverableServiceLevel = DeliverableServiceLevel.Read(parIDsp: servicelevelItem.Key);
							if (objDeliverableLayer1up != null)
								{
								Console.WriteLine("\t\t + Deliverable ServiceLevel: {0} - {1}", objDeliverableServiceLevel.IDsp,
									objDeliverableServiceLevel.Title);

								//-| Get the Service Level entry from the Database
								if(objDeliverableServiceLevel.AssociatedServiceLevelIDsp != null)
									{
									objServiceLevel = ServiceLevel.Read(parIDsp: Convert.ToInt16(objDeliverableServiceLevel.AssociatedServiceLevelIDsp));
									if (objServiceLevel != null)
										{
										Console.WriteLine("\t\t\t + Service Level: {0} - {1}", objServiceLevel.IDsp,objServiceLevel.Title);
										Console.WriteLine("\t\t\t + Service Hour.: {0}", objServiceLevel.ServiceHours);

										if(this.Service_Level_Heading_in_Section)
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2,
												parBookMark: servicelevelBookMark + objServiceLevel.IDsp);
										objRun = oxmlDocument.Construct_RunText(parText2Write: objServiceLevel.ISDheading);
										//-| Check if a hyperlink must be inserted
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

										if(this.Service_Level_Table_in_Section)
											{
											if(objServiceLevel.ISDdescription != null)
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
														parHTML2Decode: HTMLdecoder.CleanHTML(objServiceLevel.ISDdescription, parClientName),
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
													//-| A Table content error occurred, record it in
													//-| the error log.
													this.LogError("Error: The Service Level ID: " + objServiceLevel.IDsp
														+ " contains an error in the Enhance Rich Text column ISD Description. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
													objRun = oxmlDocument.Construct_RunText(
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
														objRun.Append(objDrawing);
														}
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);
													}
												}

											List<string> listErrorMessagesParameter = this.ErrorMessages;
											//-| Populate the Service Level Table
											objServiceLevelTable = CommonProcedures.BuildSLAtable(
												parMainDocumentPart: ref objMainDocumentPart,
												parClientName:	parClientName,
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
												parNumberingCounter: ref  numberingCounter);

											if(listErrorMessagesParameter.Count != this.ErrorMessages.Count)
												this.ErrorMessages = listErrorMessagesParameter;

											objBody.Append(objServiceLevelTable);
											} //if(this.Service_Level_Commitments_Table)
										} //if(this.Service_Level_Heading_in_Section)
									}
								else
									{
									//-| If the entry is not found - write an error in the document and
									//-| record an error in the error log.
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
									}//(dictSLAs.Count >0)
								}
							} //foreach ...
						}
					} //- if(this.Service_Level_Section)


Process_Glossary_and_Acronyms: //++Glossary & Acronyms
	
				if(this.ListGlossaryAndAcronyms.Count == 0)
					goto Process_Document_Acceptance_Section;

				//-| Insert the Acronyms and Glossary of Terms scetion
				if(this.Acronyms_Glossary_of_Terms_Section)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_GlossaryAndAcronymSection_HeadingText,
						parIsNewSection: true);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					//-| Insert a blank paragrpah
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


Process_Document_Acceptance_Section: //++Document Acceptance

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
								parHyperlinkID: ref hyperlinkCounter,
								parSharePointSiteURL: Properties.Settings.Default.CurrentURLSharePoint);
							}
						catch(InvalidContentFormatException exc)
							{
							Console.WriteLine("\n\nException occurred: {0}", exc.Message);
							//-| A Table content error occurred, record it in the error log.
							this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
								+ " contains an error in the Enhance Rich Text column Document Acceptance. "
								+ exc.Message);
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "A content error occurred at this position and valid content could "
								+ "not be interpreted and inserted here. Please review the content in the SharePoint "
								+ "system and correct it. Error Detail: " + exc.Message,
								parIsNewSection: false,
								parIsError: true);
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);
							}
						}
					}

Close_Document:     //++Error Section

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
				//-| Save and close the Document
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
				} //-| end Try

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
			} //-| end of Generate method
		} //-| end of ISD_Document_DRM_Sections class
	}