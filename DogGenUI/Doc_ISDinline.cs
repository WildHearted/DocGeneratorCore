using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Services.Client;
using System.Linq;
using System.Net;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Validation;
using DocGeneratorCore.SDDPServiceReference;

namespace DocGeneratorCore
	{
	/// <summary>
	/// This class represent the Internal Service Definition (ISD) with inline DRM (Deliverable Report Meeting) 
	/// It inherits from the Internal_DRM_Inline Class.
	/// </summary>
	class ISD_Document_DRM_Inline:Internal_DRM_Inline
		{
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

		public bool Generate(
			ref CompleteDataSet  parDataSet,
			DesignAndDeliveryPortfolioDataContext parSDDPdatacontext)
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
			int? intLayer1upElementID = 0;
			int? intLayer2upElementID = 0;
			int? intLayer1upDeliverableID = 0;
			int? intLayer2upDeliverableID = 0;
			int intTableCaptionCounter = 0;
			int intImageCaptionCounter = 0;
			int iPictureNo = 49;
			int intHyperlinkCounter = 9;

			if(this.HyperlinkEdit)
				{
				documentCollection_HyperlinkURL = Properties.AppResources.SharePointSiteURL +
					Properties.AppResources.List_DocumentCollectionLibraryURI +
					Properties.AppResources.EditFormURI + this.DocumentCollectionID;
				currentHyperlinkViewEditURI = Properties.AppResources.EditFormURI;
				}
			if(this.HyperlinkView)
				{
				documentCollection_HyperlinkURL = Properties.AppResources.SharePointSiteURL +
					Properties.AppResources.List_DocumentCollectionLibraryURI +
					Properties.AppResources.DisplayFormURI + this.DocumentCollectionID;
				currentHyperlinkViewEditURI = Properties.AppResources.DisplayFormURI;
				}

			// define a new objOpenXMLdocument
			oxmlDocument objOXMLdocument = new oxmlDocument();
			// use CreateDocumentFromTemplate method to create a new MS Word Document based on the relevant template
			if(objOXMLdocument.CreateDocWbkFromTemplate(
				parDocumentOrWorkbook: enumDocumentOrWorkbook.Document,
				parTemplateURL: this.Template, 
				parDocumentType: this.DocumentType))
				{
				Console.WriteLine("\t\t objOXMLdocument:\n" +
				"\t\t\t+ LocalDocumentPath: {0}\n" +
				"\t\t\t+ DocumentFileName.: {1}\n" +
				"\t\t\t+ DocumentURI......: {2}", objOXMLdocument.LocalPath, objOXMLdocument.Filename, objOXMLdocument.LocalURI);
				}
			else
				{
				// if the creation failed.
				Console.WriteLine("An ERROR occurred and the new MS Word Document could not be created due to above stated ERROR conditions.");
				this.ErrorMessages.Add("Application was unable to create the document based on the template - Check the Output log.");
				this.DocumentStatus = enumDocumentStatusses.Failed;
				return false;
				}

			this.LocalDocumentURI = objOXMLdocument.LocalURI;
			this.FileName = objOXMLdocument.Filename;

			if(this.SelectedNodes == null || this.SelectedNodes.Count < 1)
				{
				Console.WriteLine("\t\t\t *** There are 0 selected nodes to generate");
				this.ErrorMessages.Add("There are no Selected Nodes to generate.");
				this.DocumentStatus = enumDocumentStatusses.Failed;
				return false;
				}

			try
				{
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
				// Declare the HTMLdecoder object and assign the document's WordProcessing Body to the WPbody property.
				HTMLdecoder objHTMLdecoder = new HTMLdecoder();
				objHTMLdecoder.WPbody = objBody;

				// Determine the Page Size for the current Body object.
				SectionProperties objSectionProperties = new SectionProperties();
				this.PageWith = Convert.ToUInt32(Properties.AppResources.DefaultPageWidth);
				this.PageHight = Convert.ToUInt32(Properties.AppResources.DefaultPageHeight);

				if(objBody.GetFirstChild<SectionProperties>() != null)
					{
					objSectionProperties = objBody.GetFirstChild<SectionProperties>();
					PageSize objPageSize = objSectionProperties.GetFirstChild<PageSize>();
					PageMargin objPageMargin = objSectionProperties.GetFirstChild<PageMargin>();
					if(objPageSize != null)
						{
						this.PageWith = objPageSize.Width;
						this.PageHight = objPageSize.Height;
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
							this.PageHight -= Convert.ToUInt32(tempTop);
							}
						if(objPageMargin.Bottom != null)
							{
							string tempBottom = objPageMargin.Bottom.ToString();
							//Console.WriteLine("\t\t\t - Bottom Margin: {0} twips", tempBottom);
							this.PageHight -= Convert.ToUInt32(tempBottom);
							}
						}
					}
				// Subtract the Table/Image Left indentation value from the Page width to ensure the table/image fits in the available space.
				this.PageWith -= Convert.ToUInt16(Properties.AppResources.Document_Table_Left_Indent);
				Console.WriteLine("\t\t Effective pageWidth x pageHeight.: {0} x {1} twips", this.PageWith, this.PageHight);

				// Check whether Hyperlinks need to be included and add the image to the Document Body
				if(this.HyperlinkEdit || this.HyperlinkView)
					{
					//Insert and embed the hyperlink image in the document and keep the Image's Relationship ID in a variable for repeated use
					hyperlinkImageRelationshipID = oxmlDocument.InsertHyperlinkImage(parMainDocumentPart: ref objMainDocumentPart);
					}

				//Check is Content Layering was requested and add a Ledgend for the colour coding of content
				if(this.ColorCodingLayer1 || this.ColorCodingLayer2 || this.ColorCodingLayer3)
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

					objParagraph = oxmlDocument.Construct_BulletNumberParagraph(parBulletLevel: 0, parIsBullet: true);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_ColourCodingLedgend_Layer1,
						parContentLayer: "Layer1");
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					objParagraph = oxmlDocument.Construct_BulletNumberParagraph(parBulletLevel: 0, parIsBullet: true);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_ColourCodingLedgend_Layer2,
						parContentLayer: "Layer2");
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					objParagraph = oxmlDocument.Construct_BulletNumberParagraph(parBulletLevel: 0, parIsBullet: true);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_ColourCodingLedgend_Layer3,
						parContentLayer: "Layer3");
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
				ServiceElement objElementLayer2up = new ServiceElement();
				Deliverable objDeliverable = new Deliverable();
				Deliverable objDeliverableLayer1up = new Deliverable();
				Deliverable objDeliverableLayer2up = new Deliverable();
				DeliverableActivity objDeliverableActivity = new DeliverableActivity();
				DeliverableServiceLevel objDeliverableServiceLevel = new DeliverableServiceLevel();
				Activity objActivity = new Activity();
				ServiceLevel objServiceLevel = new ServiceLevel();

				this.DocumentStatus = enumDocumentStatusses.Building;
				//--------------------------------------------------
				// Insert the Introductory Section
				if(this.Introductory_Section)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_IntruductorySection_HeadingText,
						parIsNewSection: true);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);
					}
				//--------------------------------------------------
				// Insert the Introduction
				if(this.Introduction)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Introduction_HeadingText);
					// Check if a hyperlink must be inserted
					if(documentCollection_HyperlinkURL != "")
						{
						intHyperlinkCounter += 1;
						Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
							objHTMLdecoder.DecodeHTML(
								parMainDocumentPart: ref objMainDocumentPart,
								parDocumentLevel: 2,
								parHTML2Decode: this.IntroductionRichText,
								parTableCaptionCounter: ref intTableCaptionCounter,
								parImageCaptionCounter: ref intImageCaptionCounter,
								parPictureNo: ref iPictureNo,
								parHyperlinkID: ref intHyperlinkCounter,
								parPageHeightTwips: this.PageHight,
								parPageWidthTwips: this.PageWith);
							}
						catch(InvalidTableFormatException exc)
							{
							Console.WriteLine("\n\nException occurred: {0}", exc.Message);
							// A Table content error occurred, record it in the error log.
							this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
								+ " contains an error in Introduction's Enhance Rich Text. " + exc.Message);
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "A content error occurred at this position and valid content could " +
								"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it.",
								parIsNewSection: false,
								parIsError: true);
							if(documentCollection_HyperlinkURL != "")
								{
								intHyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
				//--------------------------------------------------
				// Insert the Executive Summary
				if(this.Executive_Summary)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_ExecutiveSummary_HeadingText);
					// Check if a hyperlink must be inserted
					if(documentCollection_HyperlinkURL != "")
						{
						intHyperlinkCounter += 1;
						Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
							objHTMLdecoder.DecodeHTML(
							parMainDocumentPart: ref objMainDocumentPart,
							parDocumentLevel: 2,
							parHTML2Decode: this.ExecutiveSummaryRichText,
							parTableCaptionCounter: ref intTableCaptionCounter,
							parImageCaptionCounter: ref intImageCaptionCounter,
							parPictureNo: ref iPictureNo,
							parHyperlinkID: ref intHyperlinkCounter,
							parPageHeightTwips: this.PageHight,
							parPageWidthTwips: this.PageWith);
							}
						catch(InvalidTableFormatException exc)
							{
							Console.WriteLine("\n\nException occurred: {0}", exc.Message);
							// A Table content error occurred, record it in the error log.
							this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
								+ " contains an error in Executive Summary's Enhance Rich Text. " + exc.Message);
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "A content error occurred at this position and valid content could " +
								"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it.",
								parIsNewSection: false,
								parIsError: true);
							if(documentCollection_HyperlinkURL != "")
								{
								intHyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
				//-----------------------------------
				// Insert the user selected content
				//-----------------------------------
				if(this.SelectedNodes.Count <= 0)
					goto Process_Glossary_and_Acronyms;

				foreach(Hierarchy node in this.SelectedNodes)
					{
					Console.Write("\nNode: Seq:{0} LeveL:{1} Type:{2} ID:{3}", node.Sequence, node.Level, node.NodeType, node.NodeID);

					switch(node.NodeType)
						{
						//--------------------------------------------
						case enumNodeTypes.FRA:  // Service Framework
						case enumNodeTypes.POR:  //Service Portfolio
							{
							if(this.Service_Portfolio_Section)
								{
								if(parDataSet.dsPortfolios.TryGetValue(
									key: node.NodeID,
									value: out objPortfolio))
										{
										Console.Write("\t\t + {0} - {1}", objPortfolio.ID, objPortfolio.Title);
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: objPortfolio.ISDheading,
											parIsNewSection: true);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										intHyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServicePortfoliosURI +
												currentHyperlinkViewEditURI + objPortfolio.ID,
											parHyperlinkID: intHyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Porfolio Description
									try
										{
										objHTMLdecoder.DecodeHTML(
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 1,
											parHTML2Decode: objPortfolio.ISDdescription,
											parHyperlinkID: ref intHyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref intTableCaptionCounter,
											parImageCaptionCounter: ref intImageCaptionCounter,
											parPictureNo: ref iPictureNo,
											parPageHeightTwips: this.PageHight,
											parPageWidthTwips: this.PageWith);
										}
									catch(InvalidTableFormatException exc)
										{
										Console.WriteLine("\n\nException occurred: {0}", exc.Message);
										// A Table content error occurred, record it in the error log.
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
										if(documentCollection_HyperlinkURL != "")
											{
											intHyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parHyperlinkID: intHyperlinkCounter,
												parClickLinkURL: documentCollection_HyperlinkURL);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									} //Try
								else
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
									this.LogError("Error: The Service Portfolio ID " + node.NodeID +
										" doesn't exist in SharePoint and couldn't be retrieved.");
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Error: Service Portfolio " + node.NodeID + " is missing.",
										parIsNewSection: true,
										parIsError: true);
									objParagraph.Append(objRun);
									}
								} // //if(this.Service_Portfolio_Section)
							break;
							}
						//-----------------------------------------
						case enumNodeTypes.FAM:  // Service Family
							{
							if(this.Service_Family_Heading)
								{
								// Get the entry from the DataSet
								if(parDataSet.dsFamilies.TryGetValue(
									key: node.NodeID,
									value: out objFamily))
									{
									Console.Write("\t\t + {0} - {1}", objFamily.ID, objFamily.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: objFamily.ISDheading,
										parIsNewSection: false);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										intHyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
											Properties.AppResources.List_ServiceFamiliesURI +
											currentHyperlinkViewEditURI + objFamily.ID,
											parHyperlinkID: intHyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Family Description
									if(this.Service_Family_Description)
										{
										if(objFamily.ISDdescription != null)
											{
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 2,
													parHTML2Decode: objFamily.ISDdescription,
													parHyperlinkID: ref intHyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref intTableCaptionCounter,
													parImageCaptionCounter: ref intImageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											catch(InvalidTableFormatException exc)
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
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
									} 
								else
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
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
								} // //if(this.Service_Portfolio_Section)
							break;
							}
						//------------------------------------------
						case enumNodeTypes.PRO:  // Service Product
							{
							if(this.Service_Product_Heading)
								{
								// Get the entry from the DataSet
								if(parDataSet.dsProducts.TryGetValue(
									key: node.NodeID,
									value: out objProduct))
									{
									Console.Write("\t\t + {0} - {1}", objProduct.ID, objProduct.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: objProduct.ISDheading,
										parIsNewSection: false);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										intHyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
											Properties.AppResources.List_ServiceProductsURI +
											currentHyperlinkViewEditURI + objProduct.ID,
											parHyperlinkID: intHyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Product Description
									if(this.Service_Product_Description)
										{
										if(objProduct.ISDdescription != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceProductsURI +
												currentHyperlinkViewEditURI +
												objProduct.ID;

											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 3,
													parHTML2Decode: objProduct.ISDdescription,
													parHyperlinkID: ref intHyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref intTableCaptionCounter,
													parImageCaptionCounter: ref intImageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											catch(InvalidTableFormatException exc)
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
													+ "Please review the content in the SharePoint system and correct it.",
													parIsNewSection: false,
													parIsError: true);
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
										if(this.Service_Product_KeyDD_Benefits)
											{
											if(objProduct.KeyDDbenefits != null)
												{
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceProductsURI +
													currentHyperlinkViewEditURI +
													objProduct.ID;
												objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: Properties.AppResources.Document_Product_KeyDD_Benefits,
													parIsNewSection: false);
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: hyperlinkImageRelationshipID,
														parClickLinkURL: Properties.AppResources.SharePointURL +
														Properties.AppResources.List_ServiceProductsURI +
														currentHyperlinkViewEditURI + objProduct.ID,
														parHyperlinkID: intHyperlinkCounter);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);

												try
													{
													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 4,
														parHTML2Decode: objProduct.KeyDDbenefits,
														parHyperlinkID: ref intHyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref intTableCaptionCounter,
														parImageCaptionCounter: ref intImageCaptionCounter,
														parPictureNo: ref iPictureNo,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													}
												catch(InvalidTableFormatException exc)
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
														+ "Please review the content in the SharePoint system and correct it.",
														parIsNewSection: false,
														parIsError: true);
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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

										if(this.Service_Product_Key_Client_Benefits)
											{
											if(objProduct.KeyClientBenefits != null)
												{
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceProductsURI +
													currentHyperlinkViewEditURI +
													objProduct.ID;

												objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: Properties.AppResources.Document_Product_ClientKeyBenefits,
													parIsNewSection: false);
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: hyperlinkImageRelationshipID,
														parClickLinkURL: Properties.AppResources.SharePointURL +
														Properties.AppResources.List_ServiceProductsURI +
														currentHyperlinkViewEditURI + objProduct.ID,
														parHyperlinkID: intHyperlinkCounter);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);

												try
													{
													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 4,
														parHTML2Decode: objProduct.KeyClientBenefits,
														parHyperlinkID: ref intHyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref intTableCaptionCounter,
														parImageCaptionCounter: ref intImageCaptionCounter,
														parPictureNo: ref iPictureNo,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													}
												catch(InvalidTableFormatException exc)
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
														+ "Please review the content in the SharePoint system and correct it.",
														parIsNewSection: false,
														parIsError: true);
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
										}
									}
								else
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
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
						//------------------------------------------
						case enumNodeTypes.ELE:  // Service Element
							{
							if(this.Service_Element_Heading)
								{
								// Get the entry from the DataSet
								if(parDataSet.dsElements.TryGetValue(
									key: node.NodeID,
									value: out objElement))
									{
									Console.Write("\t\t + {0} - {1}", objElement.ID, objElement.Title);

									// Insert the Service Element ISD Heading...
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(parText2Write: objElement.ISDheading);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									//Check if the Element Layer0up has Content Layers and Content Predecessors
									if(objElement.ContentPredecessorElementID == null)
										{
										intLayer1upElementID = null;
										intLayer2upElementID = null;
										}
									else
										{
										intLayer1upElementID = objElement.ContentPredecessorElementID;
										// Get the entry from the DataSet
										if(parDataSet.dsElements.TryGetValue(
											key: Convert.ToInt16(intLayer1upElementID),
											value: out objElementLayer1up))
											{
											if(objElementLayer1up.ContentPredecessorElementID == null)
												{
												intLayer2upElementID = null;
												}
											else
												{
												intLayer2upElementID = objElementLayer1up.ContentPredecessorElementID;
												// Get the entry from the DataSet
												if(parDataSet.dsElements.TryGetValue(
													key: Convert.ToInt16(intLayer2upElementID),
													value: out objElementLayer2up))
													{
													intLayer2upElementID = objElementLayer2up.ContentPredecessorElementID;
													}
												else
													{
													intLayer2upDeliverableID = null;
													}
												}
											}
										else
											{
											intLayer2upElementID = null;
											}
										}

									// Check if the user specified to include the Service Element Description
									if(this.Service_Element_Description)
										{
										if(this.PresentationMode == enumPresentationMode.Layered)
											{
											// Insert Layer 2up if present and not null
											if(intLayer2upElementID != null)
												{
												if(objElementLayer1up.ISDdescription != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_ServiceElementsURI +
															currentHyperlinkViewEditURI +
															objElementLayer1up.ID;
														}
													else
														currentListURI = "";

													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer1";
													else
														currentContentLayer = "None";
													try
														{
														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objMainDocumentPart,
															parDocumentLevel: 4,
															parHTML2Decode: objElementLayer2up.ISDdescription,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref intTableCaptionCounter,
															parImageCaptionCounter: ref intImageCaptionCounter,
															parPictureNo: ref iPictureNo,
															parHyperlinkID: ref intHyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														}
													catch(InvalidTableFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: Service Element ID: " + objElementLayer2up.ID
															+ " contains an error the Enhance Rich Text column ISD Description. "
															+ exc.Message);
														objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
														objRun = oxmlDocument.Construct_RunText(
															parText2Write: "A content error occurred at this position and valid content could "
															+ "not be interpreted and inserted here. Please review the content "
															+ "in the SharePoint system and correct it.",
															parIsNewSection: false,
															parIsError: true);
														if(documentCollection_HyperlinkURL != "")
															{
															intHyperlinkCounter += 1;
															Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
																parMainDocumentPart: ref objMainDocumentPart,
																parImageRelationshipId: hyperlinkImageRelationshipID,
																parHyperlinkID: intHyperlinkCounter,
																parClickLinkURL: currentListURI);
															objRun.Append(objDrawing);
															}
														objParagraph.Append(objRun);
														objBody.Append(objParagraph);
														}
													} //if(objElementLayer2up.ISDdescription != null)
												} // if(layer2upElementID != null)
											} // if(this.PresentationMode == enumPresentationMode.Layered)	

										// Insert Layer 1up if present and not null
										if(intLayer1upElementID != null)
											{
											if(objElementLayer1up.ISDdescription != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_ServiceElementsURI +
														currentHyperlinkViewEditURI +
														objElementLayer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";
												try
													{
													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 4,
														parHTML2Decode: objElementLayer1up.ISDdescription,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref intTableCaptionCounter,
														parImageCaptionCounter: ref intImageCaptionCounter,
														parPictureNo: ref iPictureNo,
														parHyperlinkID: ref intHyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													}
												catch(InvalidTableFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: Service Element ID: " + objElementLayer1up.ID
														+ " contains an error in the Enhance Rich Text column ISD Description. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
													objRun = oxmlDocument.Construct_RunText(
														parText2Write: "A content error occurred at this position and valid content could "
														+ "not be interpreted and inserted here. Please review the content "
														+ "in the SharePoint system and correct it.",
														parIsNewSection: false,
														parIsError: true);
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
											} //// if(layer2upElementID != null)

										// Insert Layer 0up if not null
										if(objElement.ISDdescription != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceElementsURI +
													currentHyperlinkViewEditURI +
													objElement.ID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objElement.ISDdescription,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref intTableCaptionCounter,
													parImageCaptionCounter: ref intImageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref intHyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											catch(InvalidTableFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Element ID: " + objElement.ID
													+ " contains an error in the Enhance Rich Text column ISD Description. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content "
													+ "in the SharePoint system and correct it.",
													parIsNewSection: false,
													parIsError: true);
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
										} //if(this.Service_Element_Description)

									//--------------------------------------
									// Insert the Service Element Objectives
									// Check if the user specified to include the Service Service Element Objectives
									if(this.Service_Element_Objectives)
										{
										// Insert the heading
										layerHeadingWritten = false;
										//Prepeare the heading paragraph to be inserted, but only insert it if required...
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_Element_Objectives,
											parIsNewSection: false);
										objParagraph.Append(objRun);
										layerHeadingWritten = false;

										// Insert Layer 2up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered)
											{
											if(intLayer2upElementID != null)
												{
												if(objElementLayer2up.Objectives != null)
													{
													// insert the Heading if not inserted yet.
													if(!layerHeadingWritten)
														{
														objBody.Append(objParagraph);
														layerHeadingWritten = true;
														}
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_ServiceElementsURI +
															currentHyperlinkViewEditURI +
															objElementLayer2up.ID;
														}
													else
														currentListURI = "";

													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer1";
													else
														currentContentLayer = "None";

													try
														{
														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objMainDocumentPart,
															parDocumentLevel: 5,
															parHTML2Decode: objElementLayer2up.Objectives,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref intTableCaptionCounter,
															parImageCaptionCounter: ref intImageCaptionCounter,
															parPictureNo: ref iPictureNo,
															parHyperlinkID: ref intHyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														}
													catch(InvalidTableFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: Service Element ID: " + objElementLayer2up.ID
															+ " contains an error in the Enhance Rich Text column Objectives. "
															+ exc.Message);
														objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
														objRun = oxmlDocument.Construct_RunText(
															parText2Write: "A content error occurred at this position and valid content could "
															+ "not be interpreted and inserted here. Please review the content "
															+ "in the SharePoint system and correct it.",
															parIsNewSection: false,
															parIsError: true);
														if(documentCollection_HyperlinkURL != "")
															{
															intHyperlinkCounter += 1;
															Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
												} // if(layer2upElementID != null)
											} // if(this.PresentationMode == enumPresentationMode.Layered)

										// Insert Layer 1up if resent and not null
										if(intLayer1upElementID != null)
											{
											if(objElementLayer1up.Objectives != null)
												{
												// insert the Heading if not inserted yet.
												if(!layerHeadingWritten)
													{
													objBody.Append(objParagraph);
													layerHeadingWritten = true;
													}
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_ServiceElementsURI +
														currentHyperlinkViewEditURI +
														objElementLayer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";
												try
													{
													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 5,
														parHTML2Decode: objElementLayer1up.Objectives,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref intTableCaptionCounter,
														parImageCaptionCounter: ref intImageCaptionCounter,
														parPictureNo: ref iPictureNo,
														parHyperlinkID: ref intHyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													}
												catch(InvalidTableFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: Service Element ID: " + objElementLayer1up.ID
														+ " contains an error in the Enhance Rich Text column Objectives. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
													objRun = oxmlDocument.Construct_RunText(
														parText2Write: "A content error occurred at this position and valid content could "
														+ "not be interpreted and inserted here. Please review the content "
														+ "in the SharePoint system and correct it.",
														parIsNewSection: false,
														parIsError: true);
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
											} // if(layer2upElementID != null)

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
											if(documentCollection_HyperlinkURL != "")
												{
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceElementsURI +
													currentHyperlinkViewEditURI +
													objElement.ID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objElement.Objectives,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref intTableCaptionCounter,
													parImageCaptionCounter: ref intImageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref intHyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											catch(InvalidTableFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Element ID: " + objElement.ID
													+ " contains an error in the Enhance Rich Text column Objectives. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content "
													+ "in the SharePoint system and correct it.",
													parIsNewSection: false,
													parIsError: true);
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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

									//--------------------------------------
									// Insert the Critical Success Factors
									// Check if the user specified to include the Service Service Element Critical Success Factors
									if(this.Service_Element_Critical_Success_Factors)
										{
										// Prepare the heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_Element_CriticalSuccessFactors,
											parIsNewSection: false);
										objParagraph.Append(objRun);
										layerHeadingWritten = false;
										// Insert Layer 2up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered)
											{
											if(intLayer2upElementID != null)
												{
												if(objElementLayer2up.CriticalSuccessFactors != null)
													{
													// insert the Heading if not inserted yet.
													if(!layerHeadingWritten)
														{
														objBody.Append(objParagraph);
														layerHeadingWritten = true;
														}
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_ServiceElementsURI +
															currentHyperlinkViewEditURI +
															objElementLayer2up.ID;
														}
													else
														currentListURI = "";

													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer1";
													else
														currentContentLayer = "None";
													try
														{
														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objMainDocumentPart,
															parDocumentLevel: 5,
															parHTML2Decode: objElementLayer1up.CriticalSuccessFactors,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref intTableCaptionCounter,
															parImageCaptionCounter: ref intImageCaptionCounter,
															parPictureNo: ref iPictureNo,
															parHyperlinkID: ref intHyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														}
													catch(InvalidTableFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: Service Element ID: " + objElementLayer2up.ID
															+ " contains an error in the Enhance Rich Text column Critical Success Factors. "
															+ exc.Message);
														objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
														objRun = oxmlDocument.Construct_RunText(
															parText2Write: "A content error occurred at this position and valid content could "
															+ "not be interpreted and inserted here. Please review the content "
															+ "in the SharePoint system and correct it.",
															parIsNewSection: false,
															parIsError: true);
														if(documentCollection_HyperlinkURL != "")
															{
															intHyperlinkCounter += 1;
															Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
												} // if(layer2upElementID != null)
											} // if (this.PresentationMode == Layered)

										// Insert Layer 1up if resent and not null
										if(intLayer1upElementID != null)
											{
											if(objElementLayer1up.CriticalSuccessFactors != null)
												{
												// insert the Heading if not inserted yet.
												if(!layerHeadingWritten)
													{
													objBody.Append(objParagraph);
													layerHeadingWritten = true;
													}
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_ServiceElementsURI +
														currentHyperlinkViewEditURI +
														objElementLayer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";

												try
													{
													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 5,
														parHTML2Decode: objElementLayer1up.CriticalSuccessFactors,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref intTableCaptionCounter,
														parImageCaptionCounter: ref intImageCaptionCounter,
														parPictureNo: ref iPictureNo,
														parHyperlinkID: ref intHyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													}
												catch(InvalidTableFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: Service Element ID: " + objElementLayer1up.ID
														+ " contains an error in the Enhance Rich Text column Critical Success Factors. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
													objRun = oxmlDocument.Construct_RunText(
														parText2Write: "A content error occurred at this position and valid content could "
														+ "not be interpreted and inserted here. Please review the content "
														+ "in the SharePoint system and correct it.",
														parIsNewSection: false,
														parIsError: true);
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
											} //// if(layer2upElementID != null)

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
											if(documentCollection_HyperlinkURL != "")
												{
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceElementsURI +
													currentHyperlinkViewEditURI +
													objElement.ID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objElement.CriticalSuccessFactors,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref intTableCaptionCounter,
													parImageCaptionCounter: ref intImageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref intHyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											catch(InvalidTableFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Element ID: " + objElement.ID
													+ " contains an error in the Enhance Rich Text column Critical Success Factors. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content "
													+ "in the SharePoint system and correct it.",
													parIsNewSection: false,
													parIsError: true);
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
										} //if(this.Service_Element_CriticalSuccessFactors)

									// Insert the Key Client Advantages
									// Check if the user specified to include the Service Service Key Client Advantages
									if(this.Service_Element_Key_Client_Advantages)
										{
										// Insert the heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_Element_ClientKeyAdvantages,
											parIsNewSection: false);
										objParagraph.Append(objRun);
										layerHeadingWritten = false;

										// Insert Layer2up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered)
											{
											if(intLayer2upElementID != null)
												{
												if(objElementLayer2up.KeyClientAdvantages != null)
													{
													// insert the Heading if not inserted yet.
													if(!layerHeadingWritten)
														{
														objBody.Append(objParagraph);
														layerHeadingWritten = true;
														}
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_ServiceElementsURI +
															currentHyperlinkViewEditURI +
															objElementLayer2up.ID;
														}
													else
														currentListURI = "";

													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer1";
													else
														currentContentLayer = "None";

													try
														{
														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objMainDocumentPart,
															parDocumentLevel: 5,
															parHTML2Decode: objElementLayer2up.KeyClientAdvantages,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref intTableCaptionCounter,
															parImageCaptionCounter: ref intImageCaptionCounter,
															parPictureNo: ref iPictureNo,
															parHyperlinkID: ref intHyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														}
													catch(InvalidTableFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: Service Element ID: " + objElementLayer2up.ID
															+ " contains an error in the Enhance Rich Text column Key Client Advantages. "
															+ exc.Message);
														objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
														objRun = oxmlDocument.Construct_RunText(
															parText2Write: "A content error occurred at this position and valid content could "
															+ "not be interpreted and inserted here. Please review the content "
															+ "in the SharePoint system and correct it.",
															parIsNewSection: false,
															parIsError: true);
														if(documentCollection_HyperlinkURL != "")
															{
															intHyperlinkCounter += 1;
															Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
												} // if(layer2upElementID != null)
											} // if(this.PresentationMode == Layered)

										// Insert Layer 1up if resent and not null
										if(intLayer1upElementID != null)
											{
											if(objElementLayer1up.KeyClientAdvantages != null)
												{
												// insert the Heading if not inserted yet.
												if(!layerHeadingWritten)
													{
													objBody.Append(objParagraph);
													layerHeadingWritten = true;
													}
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_ServiceElementsURI +
														currentHyperlinkViewEditURI +
														objElementLayer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";
												try
													{
													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 5,
														parHTML2Decode: objElementLayer1up.KeyClientAdvantages,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref intTableCaptionCounter,
														parImageCaptionCounter: ref intImageCaptionCounter,
														parPictureNo: ref iPictureNo,
														parHyperlinkID: ref intHyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													}
												catch(InvalidTableFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: Service Element ID: " + objElementLayer1up.ID
														+ " contains an error in the Enhance Rich Text column Key Client Advantages. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
													objRun = oxmlDocument.Construct_RunText(
														parText2Write: "A content error occurred at this position and valid content could "
														+ "not be interpreted and inserted here. Please review the content "
														+ "in the SharePoint system and correct it.",
														parIsNewSection: false,
														parIsError: true);
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
											} //// if(layer2upElementID != null)

										// Insert Layer 0up if not null
										if(objElement.KeyClientAdvantages != null)
											{
											// insert the Heading if not inserted yet.
											if(!layerHeadingWritten)
												{
												objBody.Append(objParagraph);
												layerHeadingWritten = true;
												}
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceElementsURI +
													currentHyperlinkViewEditURI +
													objElement.ID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objElement.KeyClientAdvantages,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref intTableCaptionCounter,
													parImageCaptionCounter: ref intImageCaptionCounter, 
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref intHyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											catch(InvalidTableFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Element ID: " + objElement.ID
													+ " contains an error in the Enhance Rich Text column Key Client Advantages. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content "
													+ "in the SharePoint system and correct it.",
													parIsNewSection: false,
													parIsError: true);
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
										} //if(this.Service_Element_Key Client Advantages)

									// Insert Key Client Benefits
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

										// Insert Layer 2up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered)
											{
											if(intLayer2upElementID != null)
												{
												if(objElementLayer2up.KeyClientBenefits != null)
													{
													// insert the Heading if not inserted yet.
													if(!layerHeadingWritten)
														{
														objBody.Append(objParagraph);
														layerHeadingWritten = true;
														}
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_ServiceElementsURI +
															currentHyperlinkViewEditURI +
															objElementLayer2up.ID;
														}
													else
														currentListURI = "";

													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer1";
													else
														currentContentLayer = "None";

													try
														{
														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objMainDocumentPart,
															parDocumentLevel: 5,
															parHTML2Decode: objElementLayer2up.KeyClientBenefits,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref intTableCaptionCounter,
															parImageCaptionCounter: ref intImageCaptionCounter,
															parPictureNo: ref iPictureNo,
															parHyperlinkID: ref intHyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														}
													catch(InvalidTableFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: Service Element ID: " + objElementLayer2up.ID
															+ " contains an error in the Enhance Rich Text column Key Client Benefits. "
															+ exc.Message);
														objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
														objRun = oxmlDocument.Construct_RunText(
															parText2Write: "A content error occurred at this position and valid content could "
															+ "not be interpreted and inserted here. Please review the content "
															+ "in the SharePoint system and correct it.",
															parIsNewSection: false,
															parIsError: true);
														if(documentCollection_HyperlinkURL != "")
															{
															intHyperlinkCounter += 1;
															Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
												} //// if(layer2upElementID != null)
											}
										// Insert Layer 1up if resent and not null
										if(intLayer1upElementID != null)
											{
											if(objElementLayer1up.KeyClientBenefits != null)
												{
												// insert the Heading if not inserted yet.
												if(!layerHeadingWritten)
													{
													objBody.Append(objParagraph);
													layerHeadingWritten = true;
													}
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_ServiceElementsURI +
														currentHyperlinkViewEditURI +
														objElementLayer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";
												try
													{
													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 5,
														parHTML2Decode: objElementLayer1up.KeyClientBenefits,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref intTableCaptionCounter,
														parImageCaptionCounter: ref intImageCaptionCounter,
														parPictureNo: ref iPictureNo,
														parHyperlinkID: ref intHyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													}
												catch(InvalidTableFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: Service Element ID: " + objElementLayer1up.ID
														+ " contains an error in the Enhance Rich Text column Key Client Benefits. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
													objRun = oxmlDocument.Construct_RunText(
														parText2Write: "A content error occurred at this position and valid content could "
														+ "not be interpreted and inserted here. Please review the content "
														+ "in the SharePoint system and correct it.",
														parIsNewSection: false,
														parIsError: true);
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
											} //// if(layer2upElementID != null)

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
											if(documentCollection_HyperlinkURL != "")
												{
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceElementsURI +
													currentHyperlinkViewEditURI +
													objElement.ID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objElement.KeyClientBenefits,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref intTableCaptionCounter,
													parImageCaptionCounter: ref intImageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref intHyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith); }
											catch(InvalidTableFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Element ID: " + objElement.ID
													+ " contains an error in the Enhance Rich Text column Key Client Benefits. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content "
													+ "in the SharePoint system and correct it.",
													parIsNewSection: false,
													parIsError: true);
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
										} //if(this.Service_Element_KeyClientBenefits)

									//----------------------------------------------------------------------------
									// Insert the Key DD Benefits
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

										// Insert Layer 2up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered)
											{
											if(intLayer2upElementID != null)
												{
												if(objElementLayer2up.KeyDDbenefits != null)
													{
													// insert the Heading if not inserted yet.
													if(!layerHeadingWritten)
														{
														objBody.Append(objParagraph);
														layerHeadingWritten = true;
														}
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_ServiceElementsURI +
															currentHyperlinkViewEditURI +
															objElementLayer2up.ID;
														}
													else
														currentListURI = "";

													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer1";
													else
														currentContentLayer = "None";

													try
														{
														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objMainDocumentPart,
															parDocumentLevel: 5,
															parHTML2Decode: objElementLayer2up.KeyDDbenefits,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref intTableCaptionCounter,
															parImageCaptionCounter: ref intImageCaptionCounter,
															parPictureNo: ref iPictureNo,
															parHyperlinkID: ref intHyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														}
													catch(InvalidTableFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: Service Element ID: " + objElementLayer2up.ID
															+ " contains an error in the Enhance Rich Text column Key DD Benefits. "
															+ exc.Message);
														objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
														objRun = oxmlDocument.Construct_RunText(
															parText2Write: "A content error occurred at this position and valid content could "
															+ "not be interpreted and inserted here. Please review the content "
															+ "in the SharePoint system and correct it.",
															parIsNewSection: false,
															parIsError: true);
														if(documentCollection_HyperlinkURL != "")
															{
															intHyperlinkCounter += 1;
															Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
												} // if(layer2upElementID != null)
											}
										// Insert Layer 1up if resent and not null
										if(intLayer1upElementID != null)
											{
											if(objElementLayer1up.KeyDDbenefits != null)
												{
												// insert the Heading if not inserted yet.
												if(!layerHeadingWritten)
													{
													objBody.Append(objParagraph);
													layerHeadingWritten = true;
													}
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_ServiceElementsURI +
														currentHyperlinkViewEditURI +
														objElementLayer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";
												try
													{
													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 5,
														parHTML2Decode: objElementLayer1up.KeyDDbenefits,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref intTableCaptionCounter,
														parImageCaptionCounter: ref intImageCaptionCounter,
														parPictureNo: ref iPictureNo,
														parHyperlinkID: ref intHyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													}
												catch(InvalidTableFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: Service Element ID: " + objElementLayer1up.ID
														+ " contains an error in the Enhance Rich Text column Key DD Benefits. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
													objRun = oxmlDocument.Construct_RunText(
														parText2Write: "A content error occurred at this position and valid content could "
														+ "not be interpreted and inserted here. Please review the content "
														+ "in the SharePoint system and correct it.",
														parIsNewSection: false,
														parIsError: true);
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
											} //// if(layer2upElementID != null)

										// Insert Layer 0up if not null
										if(objElement.KeyDDbenefits != null)
											{
											// insert the Heading if not inserted yet.
											if(!layerHeadingWritten)
												{
												objBody.Append(objParagraph);
												layerHeadingWritten = true;
												}
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceElementsURI +
													currentHyperlinkViewEditURI +
													objElement.ID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objElement.KeyDDbenefits,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref intTableCaptionCounter,
													parImageCaptionCounter: ref intImageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref intHyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											catch(InvalidTableFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Element ID: " + objElement.ID
													+ " contains an error in the Enhance Rich Text column Key DD Benefits. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content "
													+ "in the SharePoint system and correct it.",
													parIsNewSection: false,
													parIsError: true);
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
										} //if(this.Service_Element_KeyDDbenefits)

									//--------------------------------------------------------------------------------------
									// Insert the Key Performance Indicators
									// Check if the user specified to include the Service Element Key Performance Indicators
									if(this.Service_Element_Key_Performance_Indicators)
										{
										// Insert the heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_Element_KPI,
											parIsNewSection: false);
										objParagraph.Append(objRun);
										layerHeadingWritten = false;

										if(this.PresentationMode == enumPresentationMode.Layered)
											{
											// Insert Layer 2up if present and not null
											if(intLayer2upElementID != null)
												{
												if(objElementLayer2up.KeyPerformanceIndicators != null)
													{
													// insert the Heading if not inserted yet.
													if(!layerHeadingWritten)
														{
														objBody.Append(objParagraph);
														layerHeadingWritten = true;
														}
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_ServiceElementsURI +
															currentHyperlinkViewEditURI +
															objElementLayer1up.ID;
														}
													else
														currentListURI = "";

													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer1";
													else
														currentContentLayer = "None";

													try
														{
														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objMainDocumentPart,
															parDocumentLevel: 5,
															parHTML2Decode: objElementLayer2up.KeyPerformanceIndicators,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref intTableCaptionCounter,
															parImageCaptionCounter: ref intImageCaptionCounter,
															parPictureNo: ref iPictureNo,
															parHyperlinkID: ref intHyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														}
													catch(InvalidTableFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: Service Element ID: " + objElementLayer2up.ID
															+ " contains an error in the Enhance Rich Text column Key Performance Indicators. "
															+ exc.Message);
														objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
														objRun = oxmlDocument.Construct_RunText(
															parText2Write: "A content error occurred at this position and valid content could "
															+ "not be interpreted and inserted here. Please review the content "
															+ "in the SharePoint system and correct it.",
															parIsNewSection: false,
															parIsError: true);
														if(documentCollection_HyperlinkURL != "")
															{
															intHyperlinkCounter += 1;
															Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
												} // if(layer2upElementID != null)
											}
										// Insert Layer 1up if resent and not null
										if(intLayer1upElementID != null)
											{
											if(objElementLayer1up.KeyPerformanceIndicators != null)
												{
												// insert the Heading if not inserted yet.
												if(!layerHeadingWritten)
													{
													objBody.Append(objParagraph);
													layerHeadingWritten = true;
													}
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_ServiceElementsURI +
														currentHyperlinkViewEditURI +
														objElementLayer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";

												try
													{
													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 5,
														parHTML2Decode: objElementLayer1up.KeyPerformanceIndicators,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref intTableCaptionCounter,
														parImageCaptionCounter: ref intImageCaptionCounter,
														parPictureNo: ref iPictureNo,
														parHyperlinkID: ref intHyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													}
												catch(InvalidTableFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: Service Element ID: " + objElementLayer1up.ID
														+ " contains an error in the Enhance Rich Text column Key Performance Indicators. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
													objRun = oxmlDocument.Construct_RunText(
														parText2Write: "A content error occurred at this position and valid content could "
														+ "not be interpreted and inserted here. Please review the content "
														+ "in the SharePoint system and correct it.",
														parIsNewSection: false,
														parIsError: true);
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
											} //// if(layer2upElementID != null)

										// Insert Layer 0up if not null
										if(objElement.KeyPerformanceIndicators != null)
											{
											// insert the Heading if not inserted yet.
											if(!layerHeadingWritten)
												{
												objBody.Append(objParagraph);
												layerHeadingWritten = true;
												}
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceElementsURI +
													currentHyperlinkViewEditURI +
													objElement.ID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objElement.KeyPerformanceIndicators,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref intTableCaptionCounter,
													parImageCaptionCounter: ref intImageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref intHyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											catch(InvalidTableFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Element ID: " + objElement.ID
													+ " contains an error in the Enhance Rich Text column Key Performance Indicators. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content "
													+ "in the SharePoint system and correct it.",
													parIsNewSection: false,
													parIsError: true);
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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

									//---------------------------------------------------
									// Insert the High Level Process
									// Check if the user specified to include the Service  Element High Level Process
									if(this.Service_Element_High_Level_Process)
										{
										// Insert the heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_Element_HighLevelProcess,
											parIsNewSection: false);
										objParagraph.Append(objRun);
										layerHeadingWritten = false;

										// Insert Layer 2up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered)
											{
											if(intLayer2upElementID != null)
												{
												if(objElementLayer2up.ProcessLink != null)
													{
													// insert the Heading if not inserted yet.
													if(!layerHeadingWritten)
														{
														objBody.Append(objParagraph);
														layerHeadingWritten = true;
														}
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_ServiceElementsURI +
															currentHyperlinkViewEditURI +
															objElementLayer2up.ID;
														}
													else
														currentListURI = "";

													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer1";
													else
														currentContentLayer = "None";

													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
													objRun = oxmlDocument.Construct_RunText(
														parText2Write: objElementLayer2up.ProcessLink);
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);

													}
												} //// if(layer2upElementID != null)
											}
										// Insert Layer 1up if resent and not null
										if(intLayer1upElementID != null)
											{
											if(objElementLayer1up.ProcessLink != null)
												{
												// insert the Heading if not inserted yet.
												if(!layerHeadingWritten)
													{
													objBody.Append(objParagraph);
													layerHeadingWritten = true;
													}
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_ServiceElementsURI +
														currentHyperlinkViewEditURI +
														objElementLayer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";

												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: objElementLayer1up.ProcessLink);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} //// if(layer2upElementID != null)

										// Insert Layer 0up if not null
										if(objElement.ProcessLink != null)
											{
											// insert the Heading if not inserted yet.
											if(!layerHeadingWritten)
												{
												objBody.Append(objParagraph);
												layerHeadingWritten = true;
												}
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceElementsURI +
													currentHyperlinkViewEditURI +
													objElement.ID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";

											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: objElement.ProcessLink);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} //if(this.Service_Element_HighLevelProcess)
									drmHeading = false;
									}
								else
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
									this.LogError("Error: The Service Element ID " + node.NodeID
										+ " doesn't exist in SharePoint and couldn't be retrieved.");
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Error: Service Element " + node.NodeID + " is missing.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									}
								} // if (this.Service_Element_Heading)
							break;
							}
						//---------------------------------------
						case enumNodeTypes.ELD:  // Deliverable associated with Element
						case enumNodeTypes.ELR:  // Report deliverable associated with Element
						case enumNodeTypes.ELM:  // Meeting deliverable associated with Element
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

							// Get the entry from the DataSet
							if(parDataSet.dsDeliverables.TryGetValue(
								key: node.NodeID,
								value: out objDeliverable))
								{
								Console.Write("\t\t + {0} - {1}", objDeliverable.ID, objDeliverable.Title);

								// Insert the Deliverable ISD Heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
								objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.ISDheading);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);

								//Check if the Deliverable Layer0up has Content Layers and Content Predecessors
								if(objDeliverable.ContentPredecessorDeliverableID == null)
									{
									intLayer1upDeliverableID = null;
									intLayer2upDeliverableID = null;
									}
								else
									{
									intLayer1upDeliverableID = objDeliverable.ContentPredecessorDeliverableID;
									// Get the entry from the DataSet
									if(parDataSet.dsDeliverables.TryGetValue(
										key: Convert.ToInt16(intLayer1upDeliverableID),
										value: out objDeliverableLayer1up))
										{
										if(objDeliverableLayer1up.ContentPredecessorDeliverableID == null)
											{
											intLayer2upDeliverableID = null;
											}
										else
											{
											intLayer2upDeliverableID = objDeliverableLayer1up.ContentPredecessorDeliverableID;
											// Get the entry from the DataSet
											if(parDataSet.dsDeliverables.TryGetValue(
												key: Convert.ToInt16(intLayer2upDeliverableID),
												value: out objDeliverableLayer2up))
												{
												intLayer2upDeliverableID = objDeliverableLayer1up.ContentPredecessorDeliverableID;
												}
											else
												{
												intLayer2upDeliverableID = null;
												}
											}
										}
									else
										{
										intLayer2upDeliverableID = null;
										}
									}
								//---------------------------------------------------------------
								// Check if the user specified to include the Deliverable Summary
								if(this.DRM_Description)
									{
									// Insert Layer 2up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered)
										{
										if(intLayer2upDeliverableID != null)
											{
											if(objDeliverableLayer2up.ISDdescription != null)
												{
												// Check for Colour coding Layers and add if necessary
												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverableLayer2up.ID;
													}
												else
													currentListURI = "";
												try
													{
													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 6,
														parHTML2Decode: objDeliverableLayer2up.ISDdescription,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref intTableCaptionCounter,
														parImageCaptionCounter: ref intImageCaptionCounter,
														parPictureNo: ref iPictureNo,
														parHyperlinkID: ref intHyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													}
												catch(InvalidTableFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer2up.ID
														+ " contains an error in the Enhance Rich Text column ISD Description. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
													objRun = oxmlDocument.Construct_RunText(
														parText2Write: "A content error occurred at this position and valid content could "
														+ "not be interpreted and inserted here. Please review the content in the "
														+ "SharePoint system and correct it.",
														parIsNewSection: false,
														parIsError: true);
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
															parMainDocumentPart: ref objMainDocumentPart,
															parImageRelationshipId: hyperlinkImageRelationshipID,
															parHyperlinkID: intHyperlinkCounter,
															parClickLinkURL: currentListURI);
														objRun.Append(objDrawing);
														}
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);
													}
												} // if(objDeliverableLayer2up.ISDdescription != null)
											} // if(layer2upDeliverableID != null)
										} //if(this.PresentationMode == enumPresentationMode.Layered)

									// Insert Layer 1up if present and not null
									if(intLayer1upDeliverableID != null)
										{
										if(objDeliverableLayer1up.ISDdescription != null)
											{
											// Check for Colour coding Layers and add if necessary
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

											if(documentCollection_HyperlinkURL != "")
												{
												intHyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI +
													objDeliverableLayer1up.ID;
												}
											else
												currentListURI = "";

											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 6,
													parHTML2Decode: objDeliverableLayer1up.ISDdescription,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref intTableCaptionCounter,
													parImageCaptionCounter: ref intImageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref intHyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											catch(InvalidTableFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
													+ " contains an error in the Enhance Rich Text column ISD Description. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could " +
													"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it.",
													parIsNewSection: false,
													parIsError: true);
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: hyperlinkImageRelationshipID,
														parHyperlinkID: intHyperlinkCounter,
														parClickLinkURL: currentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											}// if(objDeliverableLayer2up.ISDdescription != null)
										} // if(layer2upDeliverableID != null)

									// Insert Layer 0up if present and not null
									if(objDeliverable.ISDdescription != null)
										{
										// Check for Colour coding Layers and add if necessary
										if(this.ColorCodingLayer1)
											currentContentLayer = "Layer3";
										else
											currentContentLayer = "None";

										if(documentCollection_HyperlinkURL != "")
											{
											intHyperlinkCounter += 1;
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverable.ID;
											}
										else
											currentListURI = "";
										try
											{
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 6,
												parHTML2Decode: objDeliverable.ISDdescription,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref intTableCaptionCounter,
												parImageCaptionCounter: ref intImageCaptionCounter,
												parPictureNo: ref iPictureNo,
												parHyperlinkID: ref intHyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										catch(InvalidTableFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
												+ " contains an error in the Enhance Rich Text column ISD Description. "
												+ exc.Message);
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: "A content error occurred at this position and valid content could "
												+ "not be interpreted and inserted here. Please review the content in the SharePoint "
												+ "system and correct it.",
												parIsNewSection: false,
												parIsError: true);
											if(documentCollection_HyperlinkURL != "")
												{
												intHyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parHyperlinkID: intHyperlinkCounter,
													parClickLinkURL: currentListURI);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} // if(objDeliverable.ISDdescription != null)
									} // if (this.DRM_Description)

								//--------------------------------------------------------------
								// Check if the user specified to include the Deliverable Inputs
								if(this.DRM_Inputs)
									{
									if(objDeliverable.Inputs != null
									|| (intLayer1upDeliverableID != null && objDeliverable.Inputs != null)
									|| (intLayer2upDeliverableID != null && objDeliverableLayer2up.Inputs != null))
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer 2up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered)
											{
											if(intLayer2upDeliverableID != null)
												{
												if(objDeliverableLayer2up.Inputs != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer2up.ID;
														}
													else
														currentListURI = "";

													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer1";
													else
														currentContentLayer = "None";
													try
														{
														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objMainDocumentPart,
															parDocumentLevel: 7,
															parHTML2Decode: objDeliverableLayer2up.Inputs,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref intTableCaptionCounter,
															parImageCaptionCounter: ref intImageCaptionCounter,
															parPictureNo: ref iPictureNo,
															parHyperlinkID: ref intHyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														}
													catch(InvalidTableFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: The Deliverable ID: " + objDeliverableLayer2up.ID
															+ " contains an error in the Enhance Rich Text column Inputs. "
															+ exc.Message);
														objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
														objRun = oxmlDocument.Construct_RunText(
															parText2Write: "A content error occurred at this position and valid content could "
															+ "not be interpreted and inserted here. Please review the content "
															+ "in the SharePoint system and correct it.",
															parIsNewSection: false,
															parIsError: true);
														if(documentCollection_HyperlinkURL != "")
															{
															intHyperlinkCounter += 1;
															Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
																parMainDocumentPart: ref objMainDocumentPart,
																parImageRelationshipId: hyperlinkImageRelationshipID,
																parHyperlinkID: intHyperlinkCounter,
																parClickLinkURL: currentListURI);
															objRun.Append(objDrawing);
															}
														objParagraph.Append(objRun);
														objBody.Append(objParagraph);
														}
													} //if(recDeliverableLayer2up.Inputs != null)
												} // if(layer2upDeliverableID != null)
											} // if(this.PresentationMode == enumPresentationMode.Layered)

										// Insert Layer 1up if present and not null
										if(intLayer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.Inputs != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverableLayer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";
												try
													{
													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 7,
														parHTML2Decode: objDeliverableLayer1up.Inputs,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref intTableCaptionCounter,
														parImageCaptionCounter: ref intImageCaptionCounter,
														parPictureNo: ref iPictureNo,
														parHyperlinkID: ref intHyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													}
												catch(InvalidTableFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: Deliverable ID: " + objDeliverableLayer1up.ID
														+ " contains an error in the Enhance Rich Text column Inputs. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
													objRun = oxmlDocument.Construct_RunText(
														parText2Write: "A content error occurred at this position and valid content could "
														+ "not be interpreted and inserted here. Please review the content "
														+ "in the SharePoint system and correct it.",
														parIsNewSection: false,
														parIsError: true);
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
											} // if(layer2upDeliverableID != null)

										// Insert Layer0up if not null
										if(objDeliverable.Inputs != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												intHyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI +
													objDeliverable.ID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: objDeliverable.Inputs,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref intTableCaptionCounter,
													parImageCaptionCounter: ref intImageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref intHyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											catch(InvalidTableFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Deliverable ID: " + objDeliverable.ID
													+ " contains an error in the Enhance Rich Text column Inputs. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content "
													+ "in the SharePoint system and correct it.",
													parIsNewSection: false,
													parIsError: true);
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: hyperlinkImageRelationshipID,
														parHyperlinkID: intHyperlinkCounter,
														parClickLinkURL: currentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}

											} // if(recDeliverable.Inputs != null)
										} //if(objDeliverable.Inputs  &&...)
									} //if(this.DRM_Inputs)

								//----------------------------------------------------------------
								// Check if the user specified to include the Deliverable Outputs
								if(this.DRM_Outputs)
									{
									if(objDeliverable.Outputs != null
									|| (intLayer1upDeliverableID != null && objDeliverableLayer1up.Outputs != null)
									|| (intLayer2upDeliverableID != null && objDeliverableLayer2up.Outputs != null))
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableOutputs_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer 2up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered)
											{
											if(intLayer2upDeliverableID != null)
												{
												if(objDeliverableLayer2up.Outputs != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer2up.ID;
														}
													else
														currentListURI = "";

													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer1";
													else
														currentContentLayer = "None";
													try
														{
														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objMainDocumentPart,
															parDocumentLevel: 7,
															parHTML2Decode: objDeliverableLayer2up.Outputs,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref intTableCaptionCounter,
															parImageCaptionCounter: ref intImageCaptionCounter,
															parPictureNo: ref iPictureNo,
															parHyperlinkID: ref intHyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														}
													catch(InvalidTableFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: Deliverable ID: " + objDeliverableLayer2up.ID
															+ " contains an error in the Enhance Rich Text column Outputs. "
															+ exc.Message);
														objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
														objRun = oxmlDocument.Construct_RunText(
															parText2Write: "A content error occurred at this position and valid content could "
															+ "not be interpreted and inserted here. Please review the content "
															+ "in the SharePoint system and correct it.",
															parIsNewSection: false,
															parIsError: true);
														if(documentCollection_HyperlinkURL != "")
															{
															intHyperlinkCounter += 1;
															Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
																parMainDocumentPart: ref objMainDocumentPart,
																parImageRelationshipId: hyperlinkImageRelationshipID,
																parHyperlinkID: intHyperlinkCounter,
																parClickLinkURL: currentListURI);
															objRun.Append(objDrawing);
															}
														objParagraph.Append(objRun);
														objBody.Append(objParagraph);
														}
													} //if(recDeliverableLayer2up.Outputs != null)
												} // if(layer2upDeliverableID != null)
											} // if(this.Presentation.....

										// Insert Layer 1up if present and not null
										if(intLayer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.Outputs != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverableLayer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";
												try
													{
													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 7,
														parHTML2Decode: objDeliverableLayer1up.Outputs,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref intTableCaptionCounter,
														parImageCaptionCounter: ref intImageCaptionCounter,
														parPictureNo: ref iPictureNo,
														parHyperlinkID: ref intHyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													}
												catch(InvalidTableFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: Deliverable ID: " + objDeliverableLayer1up.ID
														+ " contains an error in the Enhance Rich Text column Outputs. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
													objRun = oxmlDocument.Construct_RunText(
														parText2Write: "A content error occurred at this position and valid content could "
														+ "not be interpreted and inserted here. Please review the content "
														+ "in the SharePoint system and correct it.",
														parIsNewSection: false,
														parIsError: true);
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
															parMainDocumentPart: ref objMainDocumentPart,
															parImageRelationshipId: hyperlinkImageRelationshipID,
															parHyperlinkID: intHyperlinkCounter,
															parClickLinkURL: currentListURI);
														objRun.Append(objDrawing);
														}
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);
													}
												} // if(objDeliverableLayerup.Outputs != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer0up if not null
										if(objDeliverable.Outputs != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												intHyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI +
													objDeliverable.ID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: objDeliverable.Outputs,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref intTableCaptionCounter,
													parImageCaptionCounter: ref intImageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref intHyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											catch(InvalidTableFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Deliverable ID: " + objDeliverable.ID
													+ " contains an error in the Enhance Rich Text column Outputs. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content "
													+ "in the SharePoint system and correct it.",
													parIsNewSection: false,
													parIsError: true);
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: hyperlinkImageRelationshipID,
														parHyperlinkID: intHyperlinkCounter,
														parClickLinkURL: currentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} // if(objDeliverable.Outputs != null)
										} //if(objDeliverables.Outputs !== null &&)
									} //if(this.DRM_Outputs)

								//-----------------------------------------------------------------------
								// Check if the user specified to include the Deliverable DD's Obligations
								if(this.DDS_DRM_Obligations)
									{
									if(objDeliverable.DDobligations != null
									|| (intLayer1upDeliverableID != null && objDeliverableLayer1up.DDobligations != null)
									|| (intLayer2upDeliverableID != null && objDeliverableLayer2up.DDobligations != null))
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer 2up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered)
											{
											if(intLayer2upDeliverableID != null)
												{
												if(objDeliverableLayer2up.DDobligations != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer2up.ID;
														}
													else
														currentListURI = "";

													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer1";
													else
														currentContentLayer = "None";
													try
														{
														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objMainDocumentPart,
															parDocumentLevel: 7,
															parHTML2Decode: objDeliverableLayer2up.DDobligations,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref intTableCaptionCounter,
															parImageCaptionCounter: ref intImageCaptionCounter,
															parPictureNo: ref iPictureNo,
															parHyperlinkID: ref intHyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														}
													catch(InvalidTableFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: Deliverable ID: " + objDeliverableLayer2up.ID
															+ " contains an error in the Enhance Rich Text column DD's Obligations. "
															+ exc.Message);
														objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
														objRun = oxmlDocument.Construct_RunText(
															parText2Write: "A content error occurred at this position and valid content could "
															+ "not be interpreted and inserted here. Please review the content "
															+ "in the SharePoint system and correct it.",
															parIsNewSection: false,
															parIsError: true);
														if(documentCollection_HyperlinkURL != "")
															{
															intHyperlinkCounter += 1;
															Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
																parMainDocumentPart: ref objMainDocumentPart,
																parImageRelationshipId: hyperlinkImageRelationshipID,
																parHyperlinkID: intHyperlinkCounter,
																parClickLinkURL: currentListURI);
															objRun.Append(objDrawing);
															}
														objParagraph.Append(objRun);
														objBody.Append(objParagraph);
														}
													} //if(objDeliverableLayer2up.DDobligations != null)
												} // if(layer2upDeliverableID != null)
											} // if(this.Perentation....
										// Insert Layer 1up if present and not null
										if(intLayer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.DDobligations != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverableLayer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";

												try
													{
													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 7,
														parHTML2Decode: objDeliverableLayer1up.DDobligations,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref intTableCaptionCounter,
														parImageCaptionCounter: ref intImageCaptionCounter,
														parPictureNo: ref iPictureNo,
														parHyperlinkID: ref intHyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													}
												catch(InvalidTableFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: Deliverable ID: " + objDeliverableLayer1up.ID
														+ " contains an error in the Enhance Rich Text column DD's Obligations. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
													objRun = oxmlDocument.Construct_RunText(
														parText2Write: "A content error occurred at this position and valid content could "
														+ "not be interpreted and inserted here. Please review the content "
														+ "in the SharePoint system and correct it.",
														parIsNewSection: false,
														parIsError: true);
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
															parMainDocumentPart: ref objMainDocumentPart,
															parImageRelationshipId: hyperlinkImageRelationshipID,
															parHyperlinkID: intHyperlinkCounter,
															parClickLinkURL: currentListURI);
														objRun.Append(objDrawing);
														}
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);
													}
												} // if(objDeliverableLayerup.DDobligations != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer0up if not null
										if(objDeliverable.DDobligations != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												intHyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI +
													objDeliverable.ID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: objDeliverable.DDobligations,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref intTableCaptionCounter,
													parImageCaptionCounter: ref intImageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref intHyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											catch(InvalidTableFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Deliverable ID: " + objDeliverable.ID
													+ " contains an error in the Enhance Rich Text column DD's Obligations. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content "
													+ "in the SharePoint system and correct it.",
													parIsNewSection: false,
													parIsError: true);
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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

								//-------------------------------------------------------------------
								// Check if the user specified to include the Client Responsibilities
								if(this.Clients_DRM_Responsibilities)
									{
									if(objDeliverable.ClientResponsibilities != null
									|| (intLayer1upDeliverableID != null && objDeliverableLayer1up.ClientResponsibilities != null)
									|| (intLayer2upDeliverableID != null && objDeliverableLayer2up.ClientResponsibilities != null))
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer 2up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered)
											{
											if(intLayer2upDeliverableID != null)
												{
												if(objDeliverableLayer2up.ClientResponsibilities != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer2up.ID;
														}
													else
														currentListURI = "";

													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer1";
													else
														currentContentLayer = "None";
													try
														{
														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objMainDocumentPart,
															parDocumentLevel: 7,
															parHTML2Decode: objDeliverableLayer2up.ClientResponsibilities,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref intTableCaptionCounter,
															parImageCaptionCounter: ref intImageCaptionCounter,
															parPictureNo: ref iPictureNo,
															parHyperlinkID: ref intHyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														}
													catch(InvalidTableFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: Deliverable ID: " + objDeliverableLayer2up.ID
															+ " contains an error in the Enhance Rich Text column Client's Responsibilities. "
															+ exc.Message);
														objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
														objRun = oxmlDocument.Construct_RunText(
															parText2Write: "A content error occurred at this position and valid content could "
															+ "not be interpreted and inserted here. Please review the content "
															+ "in the SharePoint system and correct it.",
															parIsNewSection: false,
															parIsError: true);
														if(documentCollection_HyperlinkURL != "")
															{
															intHyperlinkCounter += 1;
															Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
																parMainDocumentPart: ref objMainDocumentPart,
																parImageRelationshipId: hyperlinkImageRelationshipID,
																parHyperlinkID: intHyperlinkCounter,
																parClickLinkURL: currentListURI);
															objRun.Append(objDrawing);
															}
														objParagraph.Append(objRun);
														objBody.Append(objParagraph);
														}
													} //if(recDeliverableLayer2up.ClientResponsibilities != null)
												} // if(layer2upDeliverableID != null)
											} // if(this.Presentation...

										// Insert Layer 1up if present and not null
										if(intLayer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.ClientResponsibilities != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverableLayer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";
												try
													{
													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 7,
														parHTML2Decode: objDeliverableLayer1up.ClientResponsibilities,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref intTableCaptionCounter,
														parImageCaptionCounter: ref intImageCaptionCounter,
														parPictureNo: ref iPictureNo,
														parHyperlinkID: ref intHyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													}
												catch(InvalidTableFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: Deliverable ID: " + objDeliverableLayer1up.ID
														+ " contains an error in the Enhance Rich Text column Client's Responsibilities. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
													objRun = oxmlDocument.Construct_RunText(
														parText2Write: "A content error occurred at this position and valid content could "
														+ "not be interpreted and inserted here. Please review the content "
														+ "in the SharePoint system and correct it.",
														parIsNewSection: false,
														parIsError: true);
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
															parMainDocumentPart: ref objMainDocumentPart,
															parImageRelationshipId: hyperlinkImageRelationshipID,
															parHyperlinkID: intHyperlinkCounter,
															parClickLinkURL: currentListURI);
														objRun.Append(objDrawing);
														}
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);
													}
												} // if(objDeliverableLayerup.ClientResponsibilities != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer0up if not null
										if(objDeliverable.ClientResponsibilities != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												intHyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI +
													objDeliverable.ID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: objDeliverable.ClientResponsibilities,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref intTableCaptionCounter,
													parImageCaptionCounter: ref intImageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref intHyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											catch(InvalidTableFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Deliverable ID: " + objDeliverable.ID
													+ " contains an error in the Enhance Rich Text column Client's Responsibilities. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content "
													+ "in the SharePoint system and correct it.",
													parIsNewSection: false,
													parIsError: true);
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: hyperlinkImageRelationshipID,
														parHyperlinkID: intHyperlinkCounter,
														parClickLinkURL: currentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} // if(objDeliverable.ClientResponsibilities != null)
										} // if(objDeliverable.ClientResponsibilities != null &&)
									} //if(this.Clients_DRM_Responsibilities)

								//------------------------------------------------------------------
								// Check if the user specified to include the Deliverable Exclusions
								if(this.DRM_Exclusions)
									{
									if(objDeliverable.Exclusions != null
									|| (intLayer1upDeliverableID != null && objDeliverableLayer1up.Exclusions != null)
									|| (intLayer2upDeliverableID != null && objDeliverableLayer2up.Exclusions != null))
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer 2up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered)
											{
											if(intLayer2upDeliverableID != null)
												{
												if(objDeliverableLayer2up.Exclusions != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer1up.ID;
														}
													else
														currentListURI = "";

													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer1";
													else
														currentContentLayer = "None";
													try
														{
														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objMainDocumentPart,
															parDocumentLevel: 7,
															parHTML2Decode: objDeliverableLayer2up.Exclusions,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref intTableCaptionCounter,
															parImageCaptionCounter: ref intImageCaptionCounter,
															parPictureNo: ref iPictureNo,
															parHyperlinkID: ref intHyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														}
													catch(InvalidTableFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: Deliverable ID: " + objDeliverableLayer2up.ID
															+ " contains an error in the Enhance Rich Text column Exclusions. "
															+ exc.Message);
														objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
														objRun = oxmlDocument.Construct_RunText(
															parText2Write: "A content error occurred at this position and valid content could "
															+ "not be interpreted and inserted here. Please review the content "
															+ "in the SharePoint system and correct it.",
															parIsNewSection: false,
															parIsError: true);
														if(documentCollection_HyperlinkURL != "")
															{
															intHyperlinkCounter += 1;
															Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
																parMainDocumentPart: ref objMainDocumentPart,
																parImageRelationshipId: hyperlinkImageRelationshipID,
																parHyperlinkID: intHyperlinkCounter,
																parClickLinkURL: currentListURI);
															objRun.Append(objDrawing);
															}
														objParagraph.Append(objRun);
														objBody.Append(objParagraph);
														}
													} //if(recDeliverableLayer2up.Exclusions != null)
												} // if(layer2upDeliverableID != null)
											} // if(this.PresentationMode....

										// Insert Layer 1up if present and not null
										if(intLayer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.Exclusions != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverableLayer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";
												try
													{
													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 7,
														parHTML2Decode: objDeliverableLayer1up.Exclusions,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref intTableCaptionCounter,
														parImageCaptionCounter: ref intImageCaptionCounter,
														parPictureNo: ref iPictureNo,
														parHyperlinkID: ref intHyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													}
												catch(InvalidTableFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: Deliverable ID: " + objDeliverableLayer1up.ID
														+ " contains an error in the Enhance Rich Text column Exclusions. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
													objRun = oxmlDocument.Construct_RunText(
														parText2Write: "A content error occurred at this position and valid content could "
														+ "not be interpreted and inserted here. Please review the content "
														+ "in the SharePoint system and correct it.",
														parIsNewSection: false,
														parIsError: true);
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
															parMainDocumentPart: ref objMainDocumentPart,
															parImageRelationshipId: hyperlinkImageRelationshipID,
															parHyperlinkID: intHyperlinkCounter,
															parClickLinkURL: currentListURI);
														objRun.Append(objDrawing);
														}
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);
													}
												} // if(objDeliverableLayerup.Exclusions != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer0up if not null
										if(objDeliverable.ClientResponsibilities != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												intHyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI +
													objDeliverable.ID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: objDeliverable.Exclusions,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref intTableCaptionCounter,
													parImageCaptionCounter: ref intImageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref intHyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											catch(InvalidTableFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Deliverable ID: " + objDeliverable.ID
													+ " contains an error in the Enhance Rich Text column Exclusions. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content "
													+ "in the SharePoint system and correct it.",
													parIsNewSection: false,
													parIsError: true);
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: hyperlinkImageRelationshipID,
														parHyperlinkID: intHyperlinkCounter,
														parClickLinkURL: currentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											} // if(objDeliverable.Exclusions != null)
										} // if(objDeliverable.Exclusions != null &&)	
									} //if(this.DRMe_Exclusions)

								//---------------------------------------------------------------
								// Check if the user specified to include the Governance Controls
								if(this.DRM_Governance_Controls)
									{
									if(objDeliverable.GovernanceControls != null
									|| (intLayer1upDeliverableID != null && objDeliverableLayer1up.GovernanceControls != null)
									|| (intLayer2upDeliverableID != null && objDeliverableLayer2up.GovernanceControls != null))
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableGovernanceControls_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer 2up if present and not null
										if(this.PresentationMode == enumPresentationMode.Layered)
											{
											if(intLayer2upDeliverableID != null)
												{
												if(objDeliverableLayer2up.GovernanceControls != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														currentListURI = Properties.AppResources.SharePointURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer2up.ID;
														}
													else
														currentListURI = "";

													if(this.ColorCodingLayer1)
														currentContentLayer = "Layer1";
													else
														currentContentLayer = "None";

													try
														{
														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objMainDocumentPart,
															parDocumentLevel: 7,
															parHTML2Decode: objDeliverableLayer2up.GovernanceControls,
															parContentLayer: currentContentLayer,
															parTableCaptionCounter: ref intTableCaptionCounter,
															parImageCaptionCounter: ref intImageCaptionCounter,
															parPictureNo: ref iPictureNo,
															parHyperlinkID: ref intHyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith);
														}
													catch(InvalidTableFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: Deliverable ID: " + objDeliverableLayer2up.ID
															+ " contains an error in the Enhance Rich Text column Governance Controls. "
															+ exc.Message);
														objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
														objRun = oxmlDocument.Construct_RunText(
															parText2Write: "A content error occurred at this position and valid content could "
															+ "not be interpreted and inserted here. Please review the content "
															+ "in the SharePoint system and correct it.",
															parIsNewSection: false,
															parIsError: true);
														if(documentCollection_HyperlinkURL != "")
															{
															intHyperlinkCounter += 1;
															Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
																parMainDocumentPart: ref objMainDocumentPart,
																parImageRelationshipId: hyperlinkImageRelationshipID,
																parHyperlinkID: intHyperlinkCounter,
																parClickLinkURL: currentListURI);
															objRun.Append(objDrawing);
															}
														objParagraph.Append(objRun);
														objBody.Append(objParagraph);
														}
													} //if(objDeliverableLayer2up.GovernanceControls != null)
												} // if(layer2upDeliverableID != null)
											} // if(this.PresentationMode = 

										// Insert Layer 1up if present and not null
										if(intLayer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.GovernanceControls != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverableLayer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";

												try
													{
													objHTMLdecoder.DecodeHTML(
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 7,
														parHTML2Decode: objDeliverableLayer1up.GovernanceControls,
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref intTableCaptionCounter,
														parImageCaptionCounter: ref intImageCaptionCounter, 
														parPictureNo: ref iPictureNo,
														parHyperlinkID: ref intHyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightTwips: this.PageHight,
														parPageWidthTwips: this.PageWith);
													}
												catch(InvalidTableFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: Deliverable ID: " + objDeliverableLayer1up.ID
														+ " contains an error in the Enhance Rich Text column Governance Controls. "
														+ exc.Message);
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
													objRun = oxmlDocument.Construct_RunText(
														parText2Write: "A content error occurred at this position and valid content could "
														+ "not be interpreted and inserted here. Please review the content "
														+ "in the SharePoint system and correct it.",
														parIsNewSection: false,
														parIsError: true);
													if(documentCollection_HyperlinkURL != "")
														{
														intHyperlinkCounter += 1;
														Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
															parMainDocumentPart: ref objMainDocumentPart,
															parImageRelationshipId: hyperlinkImageRelationshipID,
															parHyperlinkID: intHyperlinkCounter,
															parClickLinkURL: currentListURI);
														objRun.Append(objDrawing);
														}
													objParagraph.Append(objRun);
													objBody.Append(objParagraph);
													}
												} // if(objDeliverableLayer1up.GovernanceControls != null)
											} // if(layer1upDeliverableID != null)

										// Insert Layer0up if not null
										if(objDeliverable.GovernanceControls != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												intHyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI +
													objDeliverable.ID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: objDeliverable.GovernanceControls,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref intTableCaptionCounter,
													parImageCaptionCounter: ref intImageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref intHyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											catch(InvalidTableFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Deliverable ID: " + objDeliverable.ID
													+ " contains an error in the Enhance Rich Text column Governance Controls. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content "
													+ "in the SharePoint system and correct it.",
													parIsNewSection: false,
													parIsError: true);
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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

								//---------------------------------------------------
								// Check if there are any Glossary Terms or Acronyms associated with the Deliverable(s).
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
									if(intLayer1upDeliverableID != null && objDeliverableLayer1up.GlossaryAndAcronyms != null)
										{
										foreach(var entry in objDeliverableLayer1up.GlossaryAndAcronyms)
											{
											if(this.DictionaryGlossaryAndAcronyms.ContainsKey(entry.Key) != true)
												DictionaryGlossaryAndAcronyms.Add(entry.Key, entry.Value);
											}
										}
									// if there are GlossaryAndAcronyms to add from layer2up
									if(intLayer2upDeliverableID != null && objDeliverableLayer2up.GlossaryAndAcronyms != null)
										{
										foreach(var entry in objDeliverableLayer2up.GlossaryAndAcronyms)
											{
											if(this.DictionaryGlossaryAndAcronyms.ContainsKey(entry.Key) != true)
												DictionaryGlossaryAndAcronyms.Add(entry.Key, entry.Value);
											}
										}
									} // if(this.Acronyms_Glossary_of_Terms_Section)	

								}
							else
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
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
							}
						//-----------------------------------------------------------
						case enumNodeTypes.EAC:  // Activity associated with Deliverable pertaining to Service Element
							{
							if(this.Activities)
								{
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Activities_Heading);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);
								
								// Get the entry from the DataSet
								if(parDataSet.dsActivities.TryGetValue(
									key: node.NodeID,
									value: out objActivity))
									{
									Console.WriteLine("\t\t + {0} - {1}", objActivity.ID, objActivity.Title);

									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
									objRun = oxmlDocument.Construct_RunText(parText2Write: objActivity.ISDheading);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										intHyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ActvitiesURI +
												currentHyperlinkViewEditURI + objActivity.ID,
											parHyperlinkID: intHyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									// Check if the user specified to include the Deliverable Description
									if(this.Activity_Description_Table)
										{
										objActivityTable = CommonProcedures.BuildActivityTable(
											parWidthColumn1: Convert.ToUInt32(this.PageWith * 0.25),
											parWidthColumn2: Convert.ToUInt32(this.PageWith * 0.75),
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
									// If the entry is not found - write an error in the document and record an error in the error log.
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
								} // if (this.Activities)
							break;
							}
						//------------------------------------------------------
						case enumNodeTypes.ESL:  // Service Level associated with Deliverable pertaining to Service Element
							{
							if(this.Service_Level_Heading)
								{
								// Populate the Service Level Heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_ServiceLevels_Heading_Text);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);

								// Check if the user specified to include the Service Level Commitments Table
								if(this.Service_Level_Commitments_Table)
									{
									// Prepare the data which to insert into the Service Level Table
									if(parDataSet.dsDeliverableServiceLevels.TryGetValue(
										key: node.NodeID,
										value: out objDeliverableServiceLevel))
										{
										Console.WriteLine("\t\t + Deliverable ServiceLevel: {0} - {1}", objDeliverableServiceLevel.ID,
											objDeliverableServiceLevel.Title);

										// Get the Service Level entry from the DataSet
										if(objDeliverableServiceLevel.AssociatedServiceLevelID != null)
											{
											if(parDataSet.dsServiceLevels.TryGetValue(
												key: Convert.ToInt16(objDeliverableServiceLevel.AssociatedServiceLevelID),
												value: out objServiceLevel))
												{
												Console.WriteLine("\t\t\t + Service Level: {0} - {1}", objServiceLevel.ID,
													objServiceLevel.Title);
												Console.WriteLine("\t\t\t + Service Hour.: {0}", objServiceLevel.ServiceHours);

												// Insert the Service Level ISD Description
												objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 8);
												objRun = oxmlDocument.Construct_RunText(parText2Write: objServiceLevel.ISDheading);
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													intHyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: hyperlinkImageRelationshipID,
														parClickLinkURL: Properties.AppResources.SharePointURL +
															Properties.AppResources.List_ServiceLevelsURI +
															currentHyperlinkViewEditURI + objServiceLevel.ID,
														parHyperlinkID: intHyperlinkCounter);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);

												List<string> listErrorMessagesParameter = this.ErrorMessages;
												// Populate the Service Level Table
												objServiceLevelTable = CommonProcedures.BuildSLAtable(
													parServiceLevelID: objServiceLevel.ID,
													parWidthColumn1: Convert.ToUInt32(this.PageWith * 0.30),
													parWidthColumn2: Convert.ToUInt32(this.PageWith * 0.70),
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
													parErrorMessages: ref listErrorMessagesParameter);

												if(listErrorMessagesParameter.Count != this.ErrorMessages.Count)
													this.ErrorMessages = listErrorMessagesParameter;

												objBody.Append(objServiceLevelTable);
												} //if(parDataSet.dsServiceLevels.TryGetValue(
											} // if(objDeliverableServiceLevel.AssociatedServiceLevelID != null)

										} // if(parDataSet.dsDeliverableServiceLevels.TryGetValue(
									else
										{
										// If the entry is not found - write an error in the document and record an error in the error log.
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
									} // if (this.Service Level_Commitments_Table)
								} // if (this.Service_Level_Heading)
							break;
							} //case enumNodeTypes.ESL:
						} //switch (node.NodeType)
					} // foreach(Hierarchy node in this.SelectedNodes)

Process_Glossary_and_Acronyms:
//--------------------------------------------------
// Insert the Glossary of Terms and Acronym Section
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
							parSDDPdatacontext: parSDDPdatacontext,
							parDictionaryGlossaryAcronym: this.DictionaryGlossaryAndAcronyms,
							parWidthColumn1: Convert.ToUInt32(this.PageWith * 0.3),
							parWidthColumn2: Convert.ToUInt32(this.PageWith * 0.2),
							parWidthColumn3: Convert.ToUInt32(this.PageWith * 0.5),
							parErrorMessages: ref listErrors);
						objBody.Append(tableGlossaryAcronym);
						}     //if(this.TermAndAcronymList.Count > 0)
					} // if (this.Acronyms)

Process_Document_Acceptance_Section:
				// Generate the Document Acceptance Section if it was selected

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
						objHTMLdecoder.DecodeHTML(
							parMainDocumentPart: ref objMainDocumentPart,
							parDocumentLevel: 1,
							parPageWidthTwips: this.PageWith,
							parPageHeightTwips: this.PageHight,
							parHTML2Decode: this.DocumentAcceptanceRichText,
							parTableCaptionCounter: ref intTableCaptionCounter,
							parImageCaptionCounter: ref intImageCaptionCounter,
							parPictureNo: ref iPictureNo,
							parHyperlinkID: ref intHyperlinkCounter);
						}
					}

				if(this.ErrorMessages.Count > 0)
					{
					//--------------------------------------------------
					// Insert the Document Generation Error Section

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
						objParagraph = oxmlDocument.Construct_BulletNumberParagraph(parBulletLevel: 1, parIsBullet: false);
						objRun = oxmlDocument.Construct_RunText(parText2Write: errorMessageEntry, parIsError: true);
						objParagraph.Append(objRun);
						objBody.Append(objParagraph);
						}
					}


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
					timeStarted,
					DateTime.Now,
					(DateTime.Now - timeStarted));
				} // end Try

			catch(OpenXmlPackageException exc)
				{
				Console.WriteLine("*** ERROR ***\nOpenXmlPackageException occurred."
					+ "\nHresult: {0}\nMessage: {1}\nInnerException: {2}\nStackTrace: {3} ",
					exc.HResult, exc.Message, exc.InnerException, exc.StackTrace);
				this.UnhandledError = true;
				this.DocumentStatus = enumDocumentStatusses.Failed;
				return false;
				}
			catch(ArgumentNullException exc)
				{
				Console.WriteLine("*** ERROR ***\nArgumentNullException occurred."
					+ "\nHresult: {0}\nMessage: {1}\nParameterName: {2}\nInnerException: {3}\nStackTrace: {4} ",
					exc.HResult, exc.Message, exc.ParamName, exc.InnerException, exc.StackTrace);
				this.UnhandledError = true;
				this.DocumentStatus = enumDocumentStatusses.Failed;
				return false;
				}
			catch(Exception exc)
				{
				Console.WriteLine("*** ERROR ***\nArgumentNullException occurred."
					+ "\nHresult: {0}\nMessage: {1}\nInnerException: {2}\nStackTrace: {3} ",
					exc.HResult, exc.Message, exc.InnerException, exc.StackTrace);
				this.UnhandledError = true;
				this.DocumentStatus = enumDocumentStatusses.Failed;
				return false;
				}

			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			
			return true;
			}
		} // end of ISD_Document_DRM_Inline class
	}
