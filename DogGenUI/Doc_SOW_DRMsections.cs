using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Validation;

namespace DocGeneratorCore
	{
	/// <summary>
	/// This class represent the Statement of Work (SOW) with sperate DRM (Deliverable Report Meeting) sections
	/// It inherits from the DRM Sections Class.
	/// </summary>
	class Contract_SOW_Service_Description : External_DRM_Sections
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
						case 195:
							this.Introductory_Section = true;
							break;
						case 196:
							this.Introduction = true;
							break;
						case 197:
							this.Service_Portfolio_Section = true;
							break;
						case 198:
							this.Service_Portfolio_Description = true;
							break;
						case 199:
							this.Service_Family_Heading = true;
							break;
						case 200:
							this.Service_Family_Description = true;
							break;
						case 201:
							this.Service_Product_Heading = true;
							break;
						case 202:
							this.Service_Product_Description = true;
							break;
						case 203:
							this.Service_Feature_Heading = true;
							break;
						case 204:
							this.Service_Feature_Description = true;
							break;
						case 205:
							this.Deliverables_Reports_Meetings = true;
							break;
						case 206:
							this.DRM_Heading = true;
							break;
						case 207:
							this.DRM_Summary = true;
							break;
						// --- Not applicablke to this document ---
						//case 208:
						//	this.Service_Levels = true;
						//	break;
						//case 209:
						//	this.Service_Level_Heading = true;
						//	break;
						//case 210:
						//	this.Service_Level_Commitments_Table = true;
						//	break;
						// --- not applicable to this document ---
						case 211:
							this.DRM_Section = true;
							break;
						case 212:
							this.Deliverables = true;
							break;
						case 213:
							this.Deliverable_Heading = true;
							break;
						case 214:
							this.Deliverable_Description = true;
							break;
						case 215:
							this.DDs_Deliverable_Obligations = true;
							break;
						case 216:
							this.Clients_Deliverable_Responsibilities = true;
							break;
						case 217:
							this.Deliverable_Exclusions = true;
							break;
						case 218:
							this.Deliverable_Governance_Controls = true;
							break;
						case 219:
							this.Reports = true;
							break;
						case 220:
							this.Report_Heading = true;
							break;
						case 221:
							this.Report_Description = true;
							break;
						case 222:
							this.DDs_Report_Obligations = true;
							break;
						case 223:
							this.Clients_Report_Responsibilities = true;
							break;
						case 224:
							this.Report_Exclusions = true;
							break;
						case 225:
							this.Report_Governance_Controls = true;
							break;
						case 226:
							this.Meetings = true;
							break;
						case 227:
							this.Meeting_Heading = true;
							break;
						case 228:
							this.Meeting_Description = true;
							break;
						case 229:
							this.DDs_Meeting_Obligations = true;
							break;
						case 230:
							this.Clients_Meeting_Responsibilities = true;
							break;
						case 231:
							this.Meeting_Exclusions = true;
							break;
						case 232:
							this.Meeting_Governance_Controls = true;
							break;
						case 233:
							this.Acronyms_Glossary_of_Terms_Section = true;
							break;
						case 234:
							this.Acronyms = true;
							break;
						case 235:
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
			int pictureNo = 49;
			int hyperlinkCounter = 9;
			string strErrorText = "";

			try
				{
				if(this.HyperlinkEdit)
					{
					documentCollection_HyperlinkURL = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
						Properties.AppResources.List_DocumentCollectionLibraryURI +
						Properties.AppResources.EditFormURI + this.DocumentCollectionID;
					currentHyperlinkViewEditURI = Properties.AppResources.EditFormURI;
					}
				if(this.HyperlinkView)
					{
					documentCollection_HyperlinkURL = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
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
				// use CreateDocumentFromTemplate method to create a new MS Word Document based on the relevant template
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
				// Declare the HTMLdecoder object and assign the document's WordProcessing Body to the WPbody property.
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
				// Subtract the Table/Image Left indentation value from the Page width to ensure the table/image fits in the available space.
				this.PageWith -= Convert.ToUInt16(Properties.AppResources.Document_Table_Left_Indent);
				Console.WriteLine("\t\t Effective pageWidth x pageHeight.: {0} x {1} twips", this.PageWith, this.PageHeight);

				// Check whether Hyperlinks need to be included and add the image to the Document Body
				if(this.HyperlinkEdit || this.HyperlinkView)
					{
					//Insert and embed the hyperlink image in the document and keep the Image's Relationship ID in a variable for repeated use
					hyperlinkImageRelationshipID = oxmlDocument.Insert_HyperlinkImage(parMainDocumentPart: ref objMainDocumentPart,
						parDataSet: ref parDataSet);
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
							parPictureNo: ref pictureNo,
							parHyperlinkID: ref hyperlinkCounter,
							parPageHeightDxa: this.PageHeight,
							parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
							}
						catch(Exception exc)
							{
							strErrorText = "Content Error in Document Collection: " + this.ID 
								+ "Introduction Content"
								+ " Please review all content and correct it.";
							this.LogError(strErrorText);
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: strErrorText,
								parIsNewSection: false,
								parIsError: true);
							objParagraph.Append(objRun);
							objBody.Append(objParagraph);
							Console.WriteLine("\n\nException occurred: {0} - {1}", exc.HResult, exc.Message);
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
					Console.WriteLine("\nNode: SEQ:{0} LeveL:{1} NodeType:{2} NodeID:{3}", node.Sequence, node.Level, node.NodeType, node.NodeID);

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
									Console.Write("\t + {0} - {1}", objPortfolio.ID, objPortfolio.Title);

									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: objPortfolio.SOWheading,
										parIsNewSection: true);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_ServicePortfoliosURI +
												currentHyperlinkViewEditURI + objPortfolio.ID,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Porfolio Description
									if(this.Service_Portfolio_Description)
										{
										if(objPortfolio.SOWdescription != null)
											{
											currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_ServicePortfoliosURI +
												currentHyperlinkViewEditURI + objPortfolio.ID;
											try
												{
												objHTMLdecoder.DecodeHTML(parClientName: parClientName,
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 1,
													parHTML2Decode: HTMLdecoder.CleanHTML(objPortfolio.SOWdescription, parClientName),
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
													parPictureNo: ref pictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightDxa: this.PageHeight,
													parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}\n", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: The Service Portfolio ID: " + node.NodeID
													+ " contains an error in one of its Enhance Rich Text columns. Please review "
													+ " the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ " system and correct it. Error Detail: " 
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
										}
									}
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
								Console.WriteLine("\t\t + {0} - {1}", objPortfolio.ID, objPortfolio.Title);
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
									Console.WriteLine("\t + {0} - {1}", objFamily.ID, objFamily.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: objFamily.SOWheading,
										parIsNewSection: false);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
											Properties.AppResources.List_ServiceFamiliesURI +
											currentHyperlinkViewEditURI + objFamily.ID,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Family Description
									if(this.Service_Family_Description)
										{
										if(objFamily.SOWdescription != null)
											{
											currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_ServicePortfoliosURI +
												currentHyperlinkViewEditURI +
												objFamily.ID;
											try
												{
												objHTMLdecoder.DecodeHTML(parClientName: parClientName,
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 2,
													parHTML2Decode: HTMLdecoder.CleanHTML(objFamily.SOWdescription, parClientName),
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
													parPictureNo: ref pictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightDxa: this.PageHeight,
													parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
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
													+ "system and correct it.  Error Detail: "
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
									Console.Write("\t + {0} - {1}", objProduct.ID, objProduct.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: objProduct.SOWheading,
										parIsNewSection: false);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
											Properties.AppResources.List_ServiceProductsURI +
											currentHyperlinkViewEditURI + objProduct.ID,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Product Description
									if(this.Service_Product_Description)
										{
										if(objProduct.SOWdescription != null)
											{
											currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_ServiceProductsURI +
												currentHyperlinkViewEditURI +
												objProduct.ID;
											try
												{
												objHTMLdecoder.DecodeHTML(parClientName: parClientName,
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 3,
													parHTML2Decode: HTMLdecoder.CleanHTML(objProduct.SOWdescription, parClientName),
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
													parPictureNo: ref pictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightDxa: this.PageHeight,
													parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
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
						case enumNodeTypes.FEA:  // Service Feature
							{
							if(this.Service_Feature_Heading)
								{
								// Get the entry from the DataSet
								if(parDataSet.dsFeatures.TryGetValue(
									key: node.NodeID,
									value: out objFeature))
									{
									Console.Write("\t + {0} - {1}", objFeature.ID, objFeature.Title);

									// Insert the Service Feature SOW Heading...
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(parText2Write: objFeature.SOWheading);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									//Check if the Feature Layer0up has Content Layers and Content Predecessors
									if(objFeature.ContentPredecessorFeatureID == null)
										{
										layer1upFeatureID = null;
										}
									else
										{
										layer1upFeatureID = objFeature.ContentPredecessorFeatureID;
										// Get the entry from the DataSet
										if(!parDataSet.dsFeatures.TryGetValue(
											key: Convert.ToInt16(layer1upFeatureID),
											value: out objFeatureLayer1up))
											{
											layer1upFeatureID = null;
											}
										}

									// Check if the user specified to include the Service Feature Description
									if(this.Service_Feature_Description)
										{
										//-|Insert Layer 1up if present and not null
										if(layer1upFeatureID != null)
											{
											if(objFeatureLayer1up.SOWdescription != null)
												{
												//-|Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_ServiceFeaturesURI +
														currentHyperlinkViewEditURI +
														objFeatureLayer1up.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
														currentContentLayer = "Layer2";
													}

												try
													{
													objHTMLdecoder.DecodeHTML(parClientName: parClientName,
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 4,
														parHTML2Decode: HTMLdecoder.CleanHTML(objFeatureLayer1up.SOWdescription, parClientName),
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref tableCaptionCounter,
														parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Service Feature ID: " + objFeatureLayer1up.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
													objRun = oxmlDocument.Construct_RunText(
														parText2Write: "A content error occurred at this position and valid "
														+ "content could not be interpreted and inserted here. "
														+ "Please review the content in the SharePoint system and correct it. Error Detail: "
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
											}

										// Insert Layer 0up if not null
										if(objFeature.SOWdescription != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
													Properties.AppResources.List_ServiceFeaturesURI +
													currentHyperlinkViewEditURI +
													objFeature.ID;
												}
											else
												currentListURI = "";

											//- Set the Content Layer Colour Coding
											currentContentLayer = "None";
											if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
												{
												if (objFeatureLayer1up.ContentLayer.Contains("1"))
													currentContentLayer = "Layer1";
												else if (objFeatureLayer1up.ContentLayer.Contains("2"))
													currentContentLayer = "Layer2";
												}

											try
												{
												objHTMLdecoder.DecodeHTML(parClientName: parClientName,
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: HTMLdecoder.CleanHTML(objFeature.SOWdescription, parClientName),
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
													parPictureNo: ref pictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightDxa: this.PageHeight,
													parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
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
													+ "system and correct it.  Error Detail: "
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
										} //if(this.Service_Feature_Description)
									drmHeading = false;
									}
								else
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
									this.LogError("Error: The Service Feature ID " + node.NodeID
										+ " doesn't exist in SharePoint and couldn't be retrieved.");
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Error: Service Feature " + node.NodeID + " is missing.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									}
								} // if (this.Service_Feature_Heading)
							break;
							}
					//---------------------------------------
					case enumNodeTypes.FED:  // Deliverable associated with Feature
					case enumNodeTypes.FER:  // Report deliverable associated with Feature
					case enumNodeTypes.FEM:  // Meeting deliverable associated with Feature
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
								Console.Write("\t + {0} - {1}", objDeliverable.ID, objDeliverable.Title);

								// Insert the Deliverable SOW Heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
								objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.SOWheading);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);

								// Add the deliverable/report/meeting to the relevant Dictionary for inclusion in the DRM section
								if(node.NodeType == enumNodeTypes.FED) // Deliverable
									{
									if(dictDeliverables.ContainsKey(objDeliverable.ID) != true)
										dictDeliverables.Add(objDeliverable.ID, objDeliverable.SOWheading);
									}
								else if(node.NodeType == enumNodeTypes.FER) // Report
									{
									if(dictReports.ContainsKey(objDeliverable.ID) != true)
										dictReports.Add(objDeliverable.ID, objDeliverable.SOWheading);
									}
								else if(node.NodeType == enumNodeTypes.FEM) // Meeting
									{
									if(dictMeetings.ContainsKey(objDeliverable.ID) != true)
										dictMeetings.Add(objDeliverable.ID, objDeliverable.SOWheading);
									}

								//Check if the Deliverable Layer0up has Content Layers and Content Predecessors
								Console.Write("\n\t\t + Deliverable Layer 0..: {0} - {1}", objDeliverable.ID, objDeliverable.Title);
								if(objDeliverable.ContentPredecessorDeliverableID == null)
									{
									layer1upDeliverableID = null;
									}
								else
									{
									layer1upDeliverableID = objDeliverable.ContentPredecessorDeliverableID;
									// Get the entry from the DataSet
									if(!parDataSet.dsDeliverables.TryGetValue(
										key: Convert.ToInt16(layer1upDeliverableID),
										value: out objDeliverableLayer1up))
										{
										layer1upDeliverableID = null;
										}
									}

								// Check if the user specified to include the Deliverable Summary
								if(this.DRM_Summary)
									{
									//-| Insert Layer 1up if present and not null
									if(layer1upDeliverableID != null)
										{
										if(objDeliverableLayer1up.SOWsummary != null)
											{
											//- Set the Content Layer Colour Coding
											currentContentLayer = "None";
											if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
												{
												if (objFeatureLayer1up.ContentLayer.Contains("1"))
													currentContentLayer = "Layer1";
												else if (objFeatureLayer1up.ContentLayer.Contains("2"))
													currentContentLayer = "Layer2";
												}
											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
											objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverableLayer1up.SOWsummary,
												parContentLayer: currentContentLayer);
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI + objDeliverableLayer1up.ID;

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
										}

									// Insert Layer 0up if present and not null
									if(objDeliverable.SOWsummary != null)
										{
										// Check for Colour coding Layers and add if necessary
										//- Set the Content Layer Colour Coding
										currentContentLayer = "None";
										if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
											{
											if (objFeatureLayer1up.ContentLayer.Contains("1"))
												currentContentLayer = "Layer1";
											else if (objFeatureLayer1up.ContentLayer.Contains("2"))
												currentContentLayer = "Layer2";
											}
										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
										objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.SOWsummary,
											parContentLayer: currentContentLayer);

										// Check if a hyperlink must be inserted
										if(documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI + objDeliverable.ID;

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
										} // if(objDeliverable.SOWsummary != null)

									// Insert the hyperlink to the bookmark of the Deliverable's rlevant position in the DRM Section.
									objParagraph = oxmlDocument.Construct_BookmarkHyperlink(
									parBodyTextLevel: 6,
									parBookmarkValue: "Deliverable_" + objDeliverable.ID);
									objBody.Append(objParagraph);
									} // if (this.DRM_Summary)
								} //try
							else
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
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
						} //switch (node.NodeType)
					} // foreach(Hierarchy node in this.SelectedNodes)

				//======================================================
				// Insert the Deliverable, Report, Meeting (DRM) Section
				Console.Write("\nGenerating Deliverable, Report, Meeting sections...");
				if(this.DRM_Section)
					{
					// Insert the Deliverables, Reports and Meetings Section
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
						goto Save_and_Close_Document;

					if(dictDeliverables.Count == 0)
						goto Process_Reports;

					if(this.Deliverables)
						{
						Console.Write("\n Deliverables:");
						objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
						objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Deliverables_Heading_Text);
						objParagraph.Append(objRun);
						objBody.Append(objParagraph);
						string deliverableBookMark = "Deliverable_";
						// Insert the individual Deliverables in the section
						foreach(KeyValuePair<int, string> deliverableItem in dictDeliverables.OrderBy(key => key.Value))
							{
							if(this.Deliverable_Heading)
								{
								if(parDataSet.dsDeliverables.TryGetValue(
									key: deliverableItem.Key,
									value: out objDeliverable))
									{
									Console.Write("\n\t + {0} - {1}", objDeliverable.ID, objDeliverable.SOWheading);

									// Insert the Deliverable's SOW Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3,
										parBookMark: deliverableBookMark + objDeliverable.ID);
									objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.SOWheading);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									//Check if the Deliverable's Layer0up has Content Layers and Content Predecessors
									if(objDeliverable.ContentPredecessorDeliverableID == null)
										{
										layer1upDeliverableID = null;
										}
									else
										{
										layer1upDeliverableID = objDeliverable.ContentPredecessorDeliverableID;
										// Get the entry from the DataSet
										if(!parDataSet.dsDeliverables.TryGetValue(
											key: Convert.ToInt16(layer1upDeliverableID),
											value: out objDeliverableLayer1up))
											{
											layer1upDeliverableID = null;
											}
										}

									// Check if the user specified to include the Deliverable SOW Description
									if(this.Deliverable_Description)
										{
										//-|Insert Layer 1up if present and not null
										if(layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.SOWdescription != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverableLayer1up.ID;
													}
												else
													currentListURI = "";
												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
														currentContentLayer = "Layer2";
													}

												try
													{
													Console.Write("\n\t\t + Layer1up {0} - {1}", objDeliverableLayer1up.ID, objDeliverableLayer1up.Title);
													objHTMLdecoder.DecodeHTML(parClientName: parClientName,
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 4,
														parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverableLayer1up.SOWdescription, parClientName),
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref tableCaptionCounter,
														parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
													objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
													objRun = oxmlDocument.Construct_RunText(
														parText2Write: "A content error occurred at this position and valid content could "
														+ "not be interpreted and inserted here. Please review the content in the SharePoint "
														+ "system and correct it.  Error Detail: " + exc.Message,
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

										// Insert Layer0up if not null
										if(objDeliverable.SOWdescription != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI +
													objDeliverable.ID;
												}
											else
												currentListURI = "";

											//- Set the Content Layer Colour Coding
											currentContentLayer = "None";
											if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
												{
												if (objFeatureLayer1up.ContentLayer.Contains("1"))
													currentContentLayer = "Layer1";
												else if (objFeatureLayer1up.ContentLayer.Contains("2"))
													currentContentLayer = "Layer2";
												}

											try
												{
												Console.Write("\n\t\t + Layer0up{0} - {1}", objDeliverable.ID, objDeliverable.Title);
												objHTMLdecoder.DecodeHTML(parClientName: parClientName,
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverable.SOWdescription, parClientName),
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
													parPictureNo: ref pictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightDxa: this.PageHeight,
													parPageWidthDxa: this.PageWith, 
													parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it.  Error Detail: " + exc.Message,
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
										} //if(this.Deliverable_Description)

									//--------------------------------------------------------------
									// Check if the user specified to include the Deliverable Inputs
									if(this.Deliverable_Inputs)
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
											//-|Insert Layer 1up if present and not null
											if(layer1upDeliverableID != null)
												{
												if(objDeliverableLayer1up.Inputs != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer1up.ID;
														}
													else
														currentListURI = "";

													//- Set the Content Layer Colour Coding
													currentContentLayer = "None";
													if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
														{
														if (objFeatureLayer1up.ContentLayer.Contains("1"))
															currentContentLayer = "Layer1";
														else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
															parPictureNo: ref pictureNo,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightDxa: this.PageHeight,
															parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
														}
													catch(InvalidContentFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
															+ " contains an error in one of its Enhance Rich Text columns. "
															+ "Please review the content (especially tables).");
														objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
														objRun = oxmlDocument.Construct_RunText(
															parText2Write: "A content error occurred at this position and valid content could "
															+ "not be interpreted and inserted here. Please review the content in the SharePoint "
															+ "system and correct it.  Error Detail: " + exc.Message,
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
												} // if(layer2upDeliverableID != null)

											// Insert Layer0up if not null
											if(objDeliverable.Inputs != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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
												} // if(recDeliverable.Inputs != null)
											} //if(this.Deliverable_Inputs)
										} //if(this.Deliverable_Inputs)
										  //----------------------------------------------------------------
										  // Check if the user specified to include the Deliverable Outputs
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
											
											//-| Insert Layer 1up if present and not null
											if(layer1upDeliverableID != null)
												{
												if(objDeliverableLayer1up.Outputs != null)
													{
													//-| Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer1up.ID;
														}
													else
														currentListURI = "";

													//- Set the Content Layer Colour Coding
													currentContentLayer = "None";
													if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
														{
														if (objFeatureLayer1up.ContentLayer.Contains("1"))
															currentContentLayer = "Layer1";
														else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
															parPictureNo: ref pictureNo,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightDxa: this.PageHeight,
															parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
														}
													catch(InvalidContentFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
															+ " contains an error in one of its Enhance Rich Text columns. "
															+ "Please review the content (especially tables).");
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

											// Insert Layer0up if not null
											if(objDeliverable.Outputs != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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
												} // if(recDeliverable.Outputs != null)
											} //if(recDeliverables.Outputs !== null &&)
										} //if(this.Deliverable_Outputs)

									//-----------------------------------------------------------------------
									// Check if the user specified to include the Deliverable DD's Obligations
									if(this.DDs_Deliverable_Obligations)
										{
										if(objDeliverable.DDobligations != null
										|| (layer1upDeliverableID != null && objDeliverableLayer1up.DDobligations != null))
											{
											//-|Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											
											//-|Insert Layer 1up if present and not null
											if(layer1upDeliverableID != null)
												{
												if(objDeliverableLayer1up.DDobligations != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer1up.ID;
														}
													else
														currentListURI = "";

													//- Set the Content Layer Colour Coding
													currentContentLayer = "None";
													if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
														{
														if (objFeatureLayer1up.ContentLayer.Contains("1"))
															currentContentLayer = "Layer1";
														else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
															parPictureNo: ref pictureNo,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightDxa: this.PageHeight,
															parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
														}
													catch(InvalidContentFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
															+ " contains an error in one of its Enhance Rich Text columns. "
															+ "Please review the content (especially tables).");
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
													} // if(recDeliverable.Layer1up.DDobligations != null)
												} // if(layer2upDeliverableID != null)

											// Insert Layer0up if not null
											if(objDeliverable.DDobligations != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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
												} // if(recDeliverable.DDobligations != null)
											} //if(recDeliverable.DDoblidations != null &&)
										} //if(this.DDs_Deliverable_Obligations)
										  //-------------------------------------------------------------------
									//-|Check if the user specified to include the Client Responsibilities
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
											
											//-|Insert Layer 1up if present and not null
											if(layer1upDeliverableID != null)
												{
												if(objDeliverableLayer1up.ClientResponsibilities != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer1up.ID;
														}
													else
														currentListURI = "";

													//- Set the Content Layer Colour Coding
													currentContentLayer = "None";
													if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
														{
														if (objFeatureLayer1up.ContentLayer.Contains("1"))
															currentContentLayer = "Layer1";
														else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
															parPictureNo: ref pictureNo,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightDxa: this.PageHeight,
															parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
														}
													catch(InvalidContentFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
															+ " contains an error in one of its Enhance Rich Text columns. "
															+ "Please review the content (especially tables).");
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
													} // if(recDeliverable.Layer1up.ClientResponsibilities != null)
												} // if(layer2upDeliverableID != null)

											// Insert Layer0up if not null
											if(objDeliverable.ClientResponsibilities != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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
												} // if(recDeliverable.ClientResponsibilities != null)
											} // if(recDeliverable.ClientResponsibilities != null &&)
										} //if(this.Clients_Deliverable_Responsibilities)

									//------------------------------------------------------------------
									// Check if the user specified to include the Deliverable Exclusions
									if(this.Deliverable_Exclusions)
										{
										if(objDeliverable.Exclusions != null
										|| (layer1upDeliverableID != null && objDeliverableLayer1up.Exclusions != null))
											{
											//-|Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											//-|Insert Layer 1up if present and not null
											if(layer1upDeliverableID != null)
												{
												if(objDeliverableLayer1up.Exclusions != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer1up.ID;
														}
													else
														currentListURI = "";

													//- Set the Content Layer Colour Coding
													currentContentLayer = "None";
													if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
														{
														if (objFeatureLayer1up.ContentLayer.Contains("1"))
															currentContentLayer = "Layer1";
														else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
															parPictureNo: ref pictureNo,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightDxa: this.PageHeight,
															parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
														}
													catch(InvalidContentFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
															+ " contains an error in one of its Enhance Rich Text columns. "
															+ "Please review the content (especially tables).");
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
													} // if(recDeliverable.Layer1up.Exclusions != null)
												} // if(layer2upDeliverableID != null)

											// Insert Layer0up if not null
											if(objDeliverable.ClientResponsibilities != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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
												} // if(recDeliverable.Exclusions != null)
											} // if(recDeliverable.Exclusions != null &&)	
										} //if(this.Deliverable_Exclusions)

									// Check if the user specified to include the Governance Controls
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
											//-|Insert Layer 1up if present and not null
											if(layer1upDeliverableID != null)
												{
												if(objDeliverableLayer1up.GovernanceControls != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer1up.ID;
														}
													else
														currentListURI = "";

													//- Set the Content Layer Colour Coding
													currentContentLayer = "None";
													if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
														{
														if (objFeatureLayer1up.ContentLayer.Contains("1"))
															currentContentLayer = "Layer1";
														else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
															parPictureNo: ref pictureNo,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightDxa: this.PageHeight,
															parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
														}
													catch(InvalidContentFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
															+ " contains an error in one of its Enhance Rich Text columns. "
															+ "Please review the content (especially tables).");
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

											// Insert Layer0up if not null
											if(objDeliverable.GovernanceControls != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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
												} // if(recDeliverable.GovernanceControls != null)
											} // if(recDeliverable.GovernanceControls != null &&)	
										} //if(this.Deliverable_GovernanceControls)

									//---------------------------------------------------
									// Check if there are any Glossary Terms or Acronyms associated with the Deliverable(s).
									if(this.Acronyms_Glossary_of_Terms_Section)
										{
										// if there are GlossaryAndAcronyms to add from layer0up
										if(objDeliverable.GlossaryAndAcronyms != null && objDeliverable.GlossaryAndAcronyms  != null)
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
									// If the entry is not found - write an error in the document and record an error in the error log.
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

								} // if(this.DeliverableHeading
							} // foreach (KeyValuePair<int, String>.....
						} //if(this.Deliverables)
Process_Reports:
					if(dictReports.Count == 0)
						goto Process_Meetings;

					if(this.Reports)
						{
						Console.Write("\n Reports:");
						objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
						objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Reports_Heading_Text);
						objParagraph.Append(objRun);
						objBody.Append(objParagraph);
						string deliverableBookMark = "Report_";
						// Insert the individual Report in the section
						foreach(KeyValuePair<int, string> reportItem in dictReports.OrderBy(key => key.Value))
							{
							//------------------------------
							if(this.Report_Heading)
								{
								// Get the entry from the DataSet
								if(parDataSet.dsDeliverables.TryGetValue(
									key: reportItem.Key,
									value: out objDeliverable))
									{
									Console.Write("\t + {0} - {1}", objDeliverable.ID, objDeliverable.SOWheading);

									// Insert the Reports's SOW Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3,
										parBookMark: deliverableBookMark + objDeliverable.ID);
									objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.SOWheading);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									//Check if the Report's Layer0up has Content Layers and Content Predecessors
									if(objDeliverable.ContentPredecessorDeliverableID == null)
										{
										layer1upDeliverableID = null;
										}
									else
										{
										layer1upDeliverableID = objDeliverable.ContentPredecessorDeliverableID;
										// Get the entry from the DataSet
										if(!parDataSet.dsDeliverables.TryGetValue(
											key: Convert.ToInt16(layer1upDeliverableID),
											value: out objDeliverableLayer1up))
											{
											layer1upDeliverableID = null;
											}
										}

									// Check if the user specified to include the Deliverable SOW Description
									if(this.Report_Description)
										{
										//-|Insert Layer 1up if present and not null
										if(layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.SOWdescription != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverableLayer1up.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
														currentContentLayer = "Layer2";
													}

												try
													{
													Console.Write("\n\t\t + Layer1up {0} - {1}", objDeliverableLayer1up.ID, objDeliverableLayer1up.Title);
													objHTMLdecoder.DecodeHTML(parClientName: parClientName,
														parMainDocumentPart: ref objMainDocumentPart,
														parDocumentLevel: 4,
														parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverableLayer1up.SOWdescription, parClientName),
														parContentLayer: currentContentLayer,
														parTableCaptionCounter: ref tableCaptionCounter,
														parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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

										// Insert Layer0up if not null
										if(objDeliverable.SOWdescription != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI +
													objDeliverable.ID;
												}
											else
												currentListURI = "";

											//- Set the Content Layer Colour Coding
											currentContentLayer = "None";
											if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
												{
												if (objFeatureLayer1up.ContentLayer.Contains("1"))
													currentContentLayer = "Layer1";
												else if (objFeatureLayer1up.ContentLayer.Contains("2"))
													currentContentLayer = "Layer2";
												}

											try
												{
												Console.Write("\n\t\t + Layer0up {0} - {1}", objDeliverable.ID, objDeliverable.Title);
												objHTMLdecoder.DecodeHTML(parClientName: parClientName,
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverable.SOWdescription, parClientName),
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
													parPictureNo: ref pictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightDxa: this.PageHeight,
													parPageWidthDxa: this.PageWith, 
													parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
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
										} //if(this.Report_Description)

									//--------------------------------------------------------------
									// Check if the user specified to include the Report Inputs
									if(this.Report_Inputs)
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
											//-|Insert Layer 1up if present and not null
											if(layer1upDeliverableID != null)
												{
												if(objDeliverableLayer1up.Inputs != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer1up.ID;
														}
													else
														currentListURI = "";

													//- Set the Content Layer Colour Coding
													currentContentLayer = "None";
													if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
														{
														if (objFeatureLayer1up.ContentLayer.Contains("1"))
															currentContentLayer = "Layer1";
														else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
															parPictureNo: ref pictureNo,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightDxa: this.PageHeight,
															parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
														}
													catch(InvalidContentFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
															+ " contains an error in one of its Enhance Rich Text columns. "
															+ "Please review the content (especially tables).");
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
												} // if(layer2upDeliverableID != null)

											// Insert Layer0up if not null
											if(objDeliverable.Inputs != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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
												} // if(recReport.Inputs != null)
											} //if(recReports.Inputs != null &&)
										} //if(this.Report_Inputs)
										  //----------------------------------------------------------------
										  // Check if the user specified to include the Deliverable Outputs
									if(this.Report_Outputs)
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
											//-|Insert Layer 1up if present and not null
											if(layer1upDeliverableID != null)
												{
												if(objDeliverableLayer1up.Outputs != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer1up.ID;
														}
													else
														currentListURI = "";

													//- Set the Content Layer Colour Coding
													currentContentLayer = "None";
													if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
														{
														if (objFeatureLayer1up.ContentLayer.Contains("1"))
															currentContentLayer = "Layer1";
														else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
															parPictureNo: ref pictureNo,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightDxa: this.PageHeight,
															parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
														}
													catch(InvalidContentFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
															+ " contains an error in one of its Enhance Rich Text columns. "
															+ "Please review the content (especially tables).");
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
													} // if(recReport.Layer1up.Outputs != null)
												} // if(layer2upDeliverableID != null)

											// Insert Layer0up if not null
											if(objDeliverable.Outputs != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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
												} // if(recReport.Outputs != null)
											} //if(recReport.Outputs !== null &&)
										} //if(this.Report_Outputs)

									//-----------------------------------------------------------------------
									// Check if the user specified to include the Report DD's Obligations
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
											//-|Insert Layer 1up if present and not null
											if(layer1upDeliverableID != null)
												{
												if(objDeliverableLayer1up.DDobligations != null)
													{
													// Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer1up.ID;
														}
													else
														currentListURI = "";

													//- Set the Content Layer Colour Coding
													currentContentLayer = "None";
													if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
														{
														if (objFeatureLayer1up.ContentLayer.Contains("1"))
															currentContentLayer = "Layer1";
														else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
															parPictureNo: ref pictureNo,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightDxa: this.PageHeight,
															parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
														}
													catch(InvalidContentFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
															+ " contains an error in one of its Enhance Rich Text columns. "
															+ "Please review the content (especially tables).");
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
													} // if(recReport.Layer1up.DDobligations != null)
												} // if(layer2upDeliverableID != null)

											// Insert Layer0up if not null
											if(objDeliverable.DDobligations != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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
												} // if(recReport.DDobligations != null)
											} //if(recReport.DDoblidations != null &&)
										} //if(this.DDs_Report_Obligations)
										  //-------------------------------------------------------------------
										  // Check if the user specified to include the Client Responsibilities
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

											//-|Insert Layer 1up if present and not null
											if(layer1upDeliverableID != null)
												{
												if(objDeliverableLayer1up.ClientResponsibilities != null)
													{
													//-|Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer1up.ID;
														}
													else
														currentListURI = "";

													//- Set the Content Layer Colour Coding
													currentContentLayer = "None";
													if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
														{
														if (objFeatureLayer1up.ContentLayer.Contains("1"))
															currentContentLayer = "Layer1";
														else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
															parPictureNo: ref pictureNo,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightDxa: this.PageHeight,
															parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
														}
													catch(InvalidContentFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
															+ " contains an error in one of its Enhance Rich Text columns. "
															+ "Please review the content (especially tables).");
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

											//-|Insert Layer0up if not null
											if(objDeliverable.ClientResponsibilities != null)
												{
												//-|Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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
												} // if(recReport.ClientResponsibilities != null)
											} // if(recReport.ClientResponsibilities != null &&)
										} //if(this.Clients_Report_Responsibilities)

									//------------------------------------------------------------------
									// Check if the user specified to include the Deliverable Exclusions
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
											//-|Insert Layer 1up if present and not null
											if(layer1upDeliverableID != null)
												{
												if(objDeliverableLayer1up.Exclusions != null)
													{
													//-|Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer1up.ID;
														}
													else
														currentListURI = "";

													//- Set the Content Layer Colour Coding
													currentContentLayer = "None";
													if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
														{
														if (objFeatureLayer1up.ContentLayer.Contains("1"))
															currentContentLayer = "Layer1";
														else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
															parPictureNo: ref pictureNo,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightDxa: this.PageHeight,
															parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
														}
													catch(InvalidContentFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
															+ " contains an error in one of its Enhance Rich Text columns. "
															+ "Please review the content (especially tables).");
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
													} //-|if(recReport.Layer1up.Exclusions != null)
												} //-|if(layer1upDeliverableID != null)

											// Insert Layer0up if not null
											if(objDeliverable.ClientResponsibilities != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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
												} // if(recReport.Exclusions != null)
											} // if(recReport.Exclusions != null &&)	
										} //if(this.Report_Exclusions)

									// Check if the user specified to include the Governance Controls
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
											//-|Insert Layer 1up if present and not null
											if(layer1upDeliverableID != null)
												{
												if(objDeliverableLayer1up.GovernanceControls != null)
													{
													//-|Check if a hyperlink must be inserted
													if(documentCollection_HyperlinkURL != "")
														{
														hyperlinkCounter += 1;
														currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
															Properties.AppResources.List_DeliverablesURI +
															currentHyperlinkViewEditURI +
															objDeliverableLayer1up.ID;
														}
													else
														currentListURI = "";

													//- Set the Content Layer Colour Coding
													currentContentLayer = "None";
													if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
														{
														if (objFeatureLayer1up.ContentLayer.Contains("1"))
															currentContentLayer = "Layer1";
														else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
															parPictureNo: ref pictureNo,
															parHyperlinkID: ref hyperlinkCounter,
															parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
															parHyperlinkURL: currentListURI,
															parPageHeightDxa: this.PageHeight,
															parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
														}
													catch(InvalidContentFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
															+ " contains an error in one of its Enhance Rich Text columns. "
															+ "Please review the content (especially tables).");
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
												} //- if(layer1upDeliverableID != null)

											// Insert Layer0up if not null
											if(objDeliverable.GovernanceControls != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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
										} //-if(this.Report_GovernanceControls)

									//---------------------------------------------------
									// Check if there are any Glossary Terms or Acronyms associated with the Deliverable(s).
									if(this.Acronyms_Glossary_of_Terms_Section)
										{
										// if there are GlossaryAndAcronyms to add from layer0up
										if(objDeliverable.GlossaryAndAcronyms != null && objDeliverable.GlossaryAndAcronyms  != null)
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
									// If the entry is not found - write an error in the document and record an error in the error log.
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
						} //if(this.Reports)

Process_Meetings:
					if(dictMeetings.Count == 0)
						goto Save_and_Close_Document;

					if(this.Meetings)
						{
						Console.Write("\n Meetings:");
						objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
						objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Meetings_Heading_Text);
						objParagraph.Append(objRun);
						objBody.Append(objParagraph);
						string deliverableBookMark = "Meeting_";
						// Insert the individual Meetings in the section
						foreach(KeyValuePair<int, string> meetingItem in dictMeetings.OrderBy(key => key.Value))
							{
							// Get the entry from the DataSet
							if(parDataSet.dsDeliverables.TryGetValue(
								key: meetingItem.Key,
								value: out objDeliverable))
								{
								Console.Write("\t + {0} - {1}", objDeliverable.ID, objDeliverable.Title);

								// Insert the Reports's SOW Heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3, parBookMark: deliverableBookMark + objDeliverable.ID);
								objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.SOWheading);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);

								//Check if the Report's Layer0up has Content Layers and Content Predecessors
								if(objDeliverable.ContentPredecessorDeliverableID == null)
									{
									layer1upDeliverableID = null;
									}
								else
									{
									layer1upDeliverableID = objDeliverable.ContentPredecessorDeliverableID;
									// Get the entry from the DataSet
									if(!parDataSet.dsDeliverables.TryGetValue(
										key: Convert.ToInt16(layer1upDeliverableID),
										value: out objDeliverableLayer1up))
										{
										layer1upDeliverableID = null;
										}
									}

								//-|Check if the user specified to include the Deliverable SOW Description
								if(this.Meeting_Description)
									{
									//-|Insert Layer 1up if present and not null
									if(layer1upDeliverableID != null)
										{
										if(objDeliverableLayer1up.SOWdescription != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI +
													objDeliverableLayer1up.ID;
												}
											else
												currentListURI = "";

											//- Set the Content Layer Colour Coding
											currentContentLayer = "None";
											if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
												{
												if (objFeatureLayer1up.ContentLayer.Contains("1"))
													currentContentLayer = "Layer1";
												else if (objFeatureLayer1up.ContentLayer.Contains("2"))
													currentContentLayer = "Layer2";
												}

											try
												{
												Console.Write("\n\t\t + Layer1up {0} - {1}", objDeliverableLayer1up.ID, objDeliverableLayer1up.Title);
												objHTMLdecoder.DecodeHTML(parClientName: parClientName,
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverableLayer1up.SOWdescription, parClientName),
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
													parPictureNo: ref pictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightDxa: this.PageHeight,
													parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
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
										} // if(layer1upDeliverableID != null)

									// Insert Layer0up if not null
									if(objDeliverable.SOWdescription != null)
										{
										// Check if a hyperlink must be inserted
										if(documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverable.ID;
											}
										else
											currentListURI = "";

										//- Set the Content Layer Colour Coding
										currentContentLayer = "None";
										if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
											{
											if (objFeatureLayer1up.ContentLayer.Contains("1"))
												currentContentLayer = "Layer1";
											else if (objFeatureLayer1up.ContentLayer.Contains("2"))
												currentContentLayer = "Layer2";
											}

										try
											{
											Console.Write("\n\t\t + Layer0up {0} - {1}", objDeliverable.ID, objDeliverable.Title);
											objHTMLdecoder.DecodeHTML(parClientName: parClientName,
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: HTMLdecoder.CleanHTML(objDeliverable.SOWdescription, parClientName),
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter, parNumberingCounter: ref numberingCounter, 
												parPictureNo: ref pictureNo,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightDxa: this.PageHeight,
												parPageWidthDxa: this.PageWith, 
												parSharePointSiteURL: parDataSet.SharePointSiteURL);
											}
										catch(InvalidContentFormatException exc)
											{
											Console.WriteLine("\n\nException occurred: {0}", exc.Message);
											// A Table content error occurred, record it in the error log.
											this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
												+ " contains an error in one of its Enhance Rich Text columns. "
												+ "Please review the content (especially tables).");
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
									} //if(this.Meeting_Description)

								//--------------------------------------------------------------
								// Check if the user specified to include the Report Inputs
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
										//-|Insert Layer 1up if present and not null
										if(layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.Inputs != null)
												{
												//-|Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverableLayer1up.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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
											} // if(layer2upDeliverableID != null)

										// Insert Layer0up if not null
										if(objDeliverable.Inputs != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI +
													objDeliverable.ID;
												}
											else
												currentListURI = "";

											//- Set the Content Layer Colour Coding
											currentContentLayer = "None";
											if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
												{
												if (objFeatureLayer1up.ContentLayer.Contains("1"))
													currentContentLayer = "Layer1";
												else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
													parPictureNo: ref pictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightDxa: this.PageHeight,
													parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
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
											} // if(recMeeting.Inputs != null)
										} //if(recMeeting.Inputs != null &&)
									} //if(this.Meeting_Inputs)
									  //----------------------------------------------------------------
									  // Check if the user specified to include the Deliverable Outputs
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
										//-|Insert Layer 1up if present and not null
										if(layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.Outputs != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverableLayer1up.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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
												} // if(recMeeting.Layer1up.Outputs != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer0up if not null
										if(objDeliverable.Outputs != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI +
													objDeliverable.ID;
												}
											else
												currentListURI = "";

											//- Set the Content Layer Colour Coding
											currentContentLayer = "None";
											if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
												{
												if (objFeatureLayer1up.ContentLayer.Contains("1"))
													currentContentLayer = "Layer1";
												else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
													parPictureNo: ref pictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightDxa: this.PageHeight,
													parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
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
											} // if(recMeeting.Outputs != null)
										} //if(recMeeting.Outputs !== null &&)
									} //if(this.Meeting_Outputs)

								//-----------------------------------------------------------------------
								// Check if the user specified to include the Report DD's Obligations
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
										// Insert Layer 1up if present and not null
										if(layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.DDobligations != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverableLayer1up.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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
												} // if(recMeeting.Layer1up.DDobligations != null)
											} // if(layer1upDeliverableID != null)

										// Insert Layer0up if not null
										if(objDeliverable.DDobligations != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI +
													objDeliverable.ID;
												}
											else
												currentListURI = "";

											//- Set the Content Layer Colour Coding
											currentContentLayer = "None";
											if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
												{
												if (objFeatureLayer1up.ContentLayer.Contains("1"))
													currentContentLayer = "Layer1";
												else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
													parPictureNo: ref pictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightDxa: this.PageHeight,
													parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
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
											} // if(recMeeting.DDobligations != null)
										} //if(recMeeting.DDoblidations != null &&)
									} //if(this.DDs_Report_Obligations)
									  //-------------------------------------------------------------------
									  // Check if the user specified to include the Client Responsibilities
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
										//-|Insert Layer 1up if present and not null
										if(layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.ClientResponsibilities != null)
												{
												//-|Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverableLayer1up.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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
												} // if(recMeeting.Layer1up.ClientResponsibilities != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer0up if not null
										if(objDeliverable.ClientResponsibilities != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI +
													objDeliverable.ID;
												}
											else
												currentListURI = "";

											//- Set the Content Layer Colour Coding
											currentContentLayer = "None";
											if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
												{
												if (objFeatureLayer1up.ContentLayer.Contains("1"))
													currentContentLayer = "Layer1";
												else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
													parPictureNo: ref pictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightDxa: this.PageHeight,
													parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
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
											} // if(recMeeting.ClientResponsibilities != null)
										} // if(recMeeting.ClientResponsibilities != null &&)
									} //if(this.Clients_Report_Responsibilities)

								// Check if the user specified to include the Deliverable Exclusions
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
										
										// Insert Layer 1up if present and not null
										if(layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.Exclusions != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverableLayer1up.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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
												} // if(recMeeting.Layer1up.Exclusions != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer0up if not null
										if(objDeliverable.Exclusions != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI +
													objDeliverable.ID;
												}
											else
												currentListURI = "";

											//- Set the Content Layer Colour Coding
											currentContentLayer = "None";
											if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
												{
												if (objFeatureLayer1up.ContentLayer.Contains("1"))
													currentContentLayer = "Layer1";
												else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
													parPictureNo: ref pictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightDxa: this.PageHeight,
													parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
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
											} // if(recMeeting.Exclusions != null)
										} // if(recMeeting.Exclusions != null &&)	
									} //if(this.Deliverable_Exclusions)

								// Check if the user specified to include the Governance Controls
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
										// Insert Layer 1up if present and not null
										if(layer1upDeliverableID != null)
											{
											if(objDeliverableLayer1up.GovernanceControls != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverableLayer1up.ID;
													}
												else
													currentListURI = "";

												//- Set the Content Layer Colour Coding
												currentContentLayer = "None";
												if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
													{
													if (objFeatureLayer1up.ContentLayer.Contains("1"))
														currentContentLayer = "Layer1";
													else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
														parPictureNo: ref pictureNo,
														parHyperlinkID: ref hyperlinkCounter,
														parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
														parHyperlinkURL: currentListURI,
														parPageHeightDxa: this.PageHeight,
														parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
													}
												catch(InvalidContentFormatException exc)
													{
													Console.WriteLine("\n\nException occurred: {0}", exc.Message);
													// A Table content error occurred, record it in the error log.
													this.LogError("Error: The Deliverable ID: " + objDeliverableLayer1up.ID
														+ " contains an error in one of its Enhance Rich Text columns. "
														+ "Please review the content (especially tables).");
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
												} // if(recMeeting.Layer1up.GovernanceControls != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer0up if not null
										if(objDeliverable.GovernanceControls != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI +
													objDeliverable.ID;
												}
											else
												currentListURI = "";

											//- Set the Content Layer Colour Coding
											currentContentLayer = "None";
											if (this.ColorCodingLayer1 || this.ColorCodingLayer2)
												{
												if (objFeatureLayer1up.ContentLayer.Contains("1"))
													currentContentLayer = "Layer1";
												else if (objFeatureLayer1up.ContentLayer.Contains("2"))
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
													parPictureNo: ref pictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightDxa: this.PageHeight,
													parPageWidthDxa: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: The Deliverable ID: " + objDeliverable.ID
													+ " contains an error in one of its Enhance Rich Text columns. "
													+ "Please review the content (especially tables).");
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
											} // if(recMeeting.GovernanceControls != null)
										} // if(recMeeting.GovernanceControls != null &&)	
									} //if(this.Deliverable_GovernanceControls)

								//---------------------------------------------------
								// Check if there are any Glossary Terms or Acronyms associated with the Deliverable(s).
								if(this.Acronyms_Glossary_of_Terms_Section)
									{
									// if there are GlossaryAndAcronyms to add from layer0up
									if(objDeliverable.GlossaryAndAcronyms != null && objDeliverable.GlossaryAndAcronyms  != null)
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
								// If the entry is not found - write an error in the document and record an error in the error log.
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
							} // foreach.....
						} //if(this.Meetings)
					} //if(this.DRM_Section)


Process_Glossary_and_Acronyms:
//--------------------------------------------------
// Insert the Glossary of Terms and Acronym Section
				if(this.DictionaryGlossaryAndAcronyms.Count == 0)
					goto Save_and_Close_Document;

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

Save_and_Close_Document:

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
						objParagraph = oxmlDocument.Construct_Error(errorMessageEntry);
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


			} // end of Generate method
		} // end of Contract_Sow_ServiceDescription class
	}
