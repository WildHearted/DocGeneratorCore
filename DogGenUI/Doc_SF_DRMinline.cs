using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Services.Client;
using System.Linq;
using System.Net;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Validation;
using DocGenerator.SDDPServiceReference;

namespace DocGenerator
	{
	/// <summary>
	/// This class represent the Framework with inline DRM (Deliverable Report Meeting) Document object
	/// It inherits from the Internal_DRM_Inline Class.
	/// </summary>
	class Services_Framework_Document_DRM_Inline:Internal_DRM_Inline
		{
		/// <summary>
		/// this option takes the values passed into the method as a list of integers
		/// which represents the options the user selected and transpose the values by
		/// setting the object's.
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
						case 293:
							this.Introductory_Section = true;
							break;
						case 294:
							this.Introduction = true;
							break;
						case 295:
							this.Executive_Summary = true;
							break;
						case 296:
							this.Service_Portfolio_Section = true;
							break;
						case 297:
							this.Service_Portfolio_Description = true;
							break;
						case 298:
							this.Service_Family_Heading = true;
							break;
						case 299:
							this.Service_Family_Description = true;
							break;
						case 300:
							this.Service_Product_Heading = true;
							break;
						case 301:
							this.Service_Product_Description = true;
							break;
						case 302:
							this.Service_Product_Key_Client_Benefits = true;
							break;
						case 303:
							this.Service_Product_KeyDD_Benefits = true;
							break;
						case 304:
							this.Service_Element_Heading = true;
							break;
						case 305:
							this.Service_Element_Description = true;
							break;
						case 306:
							this.Service_Element_Objectives = true;
							break;
						case 307:
							this.Service_Element_Key_Client_Benefits = true;
							break;
						case 308:
							this.Service_Element_Key_Client_Advantages = true;
							break;
						case 309:
							this.Service_Element_Key_DD_Benefits = true;
							break;
						case 311:
							this.Service_Element_Critical_Success_Factors = true;
							break;
						case 312:
							this.Service_Element_Key_Performance_Indicators = true;
							break;
						case 313:
							this.Service_Element_High_Level_Process = true;
							break;
						case 314:
							this.Deliverables_Reports_Meetings = true;
							break;
						case 315:
							this.DRM_Heading = true;
							break;
						case 316:
							this.DRM_Description = true;
							break;
						case 317:
							this.DRM_Inputs = true;
							break;
						case 318:
							this.DRM_Outputs = true;
							break;
						case 319:
							this.DDS_DRM_Obligations = true;
							break;
						case 320:
							this.Clients_DRM_Responsibilities = true;
							break;
						case 321:
							this.DRM_Exclusions = true;
							break;
						case 322:
							this.DRM_Governance_Controls = true;
							break;
						case 323:
							this.Service_Levels = true;
							break;
						case 324:
							this.Service_Level_Heading = true;
							break;
						case 325:
							this.Service_Level_Commitments_Table = true;
							break;
						case 326:
							this.Activities = true;
							break;
						case 327:
							this.Activity_Heading = true;
							break;
						case 328:
							this.Activity_Description_Table = true;
							break;
						case 329:
							this.Acronyms_Glossary_of_Terms_Section = true;
							break;
						case 330:
							this.Acronyms = true;
							break;
						case 331:
							this.Glossary_of_Terms = true;
							break;
						case 332:
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

		public bool Generate()
			{
			Console.WriteLine("\t Begin to generate {0}", this.DocumentType);
			DateTime timeStarted = DateTime.Now;
			string hyperlinkImageRelationshipID = "";
			string documentCollection_HyperlinkURL = "";
			string currentListURI = "";
			string currentHyperlinkViewEditURI = "";
			string currentContentLayer = "None";
			bool drmHeading = false;
			Table objActivityTable = new Table();
			Table objServiceLevelTable = new Table();

			if(this.HyperlinkEdit)
				documentCollection_HyperlinkURL = Properties.AppResources.SharePointSiteURL +
					Properties.AppResources.List_DocumentCollectionLibraryURI +
					Properties.AppResources.EditFormURI + this.DocumentCollectionID;
			currentHyperlinkViewEditURI = Properties.AppResources.EditFormURI;
			if(this.HyperlinkView)
				documentCollection_HyperlinkURL = Properties.AppResources.SharePointSiteURL +
					Properties.AppResources.List_DocumentCollectionLibraryURI +
					Properties.AppResources.DisplayFormURI + this.DocumentCollectionID;
			currentHyperlinkViewEditURI = Properties.AppResources.DisplayFormURI;
			int tableCaptionCounter = 1;
			int imageCaptionCounter = 1;
			int hyperlinkCounter = 4;

			//Initialize the Data access to SharePoint
			DesignAndDeliveryPortfolioDataContext datacontexSDDP = new DesignAndDeliveryPortfolioDataContext(new
				Uri(Properties.AppResources.SharePointSiteURL + Properties.AppResources.SharePointRESTuri));
			datacontexSDDP.Credentials = CredentialCache.DefaultCredentials;
			datacontexSDDP.MergeOption = MergeOption.NoTracking;

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
				return false;
				}

			if(this.SelectedNodes == null || this.SelectedNodes.Count < 1)
				{
				Console.WriteLine("\t\t\t *** There are 0 selected nodes to generate");
				this.ErrorMessages.Add("There are no Selected Nodes to generate.");
				return false;
				}
			// Create and open the new Document
			try
				{
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
						Console.WriteLine("\t\t Page width x height: {0} x {1} twips", this.PageWith, this.PageHight);
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
							this.PageHight -= Convert.ToUInt32(tempTop);
							}
						if(objPageMargin.Bottom != null)
							{
							string tempBottom = objPageMargin.Bottom.ToString();
							Console.WriteLine("\t\t\t - Bottom Margin: {0} twips", tempBottom);
							this.PageHight -= Convert.ToUInt32(tempBottom);
							}
						}
					}
				// Subtract the Table/Image Left indentation value from the Page width to ensure the table/image fits in the available space.
				this.PageWith -= Convert.ToUInt16(Properties.AppResources.Document_Table_Left_Indent);
				Console.WriteLine("\t\t Effective pageWidth x pageHeight.: {0} x {1} twips", this.PageWith, this.PageHight);

				// Check whether Hyperlinks need to be included
				if(this.HyperlinkEdit || this.HyperlinkView)
					{
					//Insert and embed the hyperlink image in the document and keep the Image's Relationship ID in a variable for repeated use
					hyperlinkImageRelationshipID = oxmlDocument.InsertHyperlinkImage(parMainDocumentPart: ref objMainDocumentPart);
					}
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
						hyperlinkCounter += 1;
						DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
						objHTMLdecoder.DecodeHTML(
							parMainDocumentPart: ref objMainDocumentPart,
							parDocumentLevel: 2,
							parHTML2Decode: this.IntroductionRichText,
							parTableCaptionCounter: ref tableCaptionCounter,
							parImageCaptionCounter: ref imageCaptionCounter,
							parHyperlinkID: ref hyperlinkCounter,
							parPageHeightTwips: this.PageHight,
							parPageWidthTwips: this.PageWith);
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
						hyperlinkCounter += 1;
						DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
						objHTMLdecoder.DecodeHTML(
							parMainDocumentPart: ref objMainDocumentPart,
							parDocumentLevel: 2,
							parHTML2Decode: this.ExecutiveSummaryRichText,
							parTableCaptionCounter: ref tableCaptionCounter,
							parImageCaptionCounter: ref imageCaptionCounter,
							parHyperlinkID: ref hyperlinkCounter,
							parPageHeightTwips: this.PageHight,
							parPageWidthTwips: this.PageWith);
						}
					}
				//-----------------------------------
				// Insert the user selected content
				//-----------------------------------
				if(this.SelectedNodes.Count <= 0)
					goto Process_Glossary_and_Acronyms;
				foreach(Hierarchy node in this.SelectedNodes)
					{
					Console.WriteLine("Node: {0} - {1} {2} {3}", node.Sequence, node.Level, node.NodeType, node.NodeID);

					switch(node.NodeType)
						{
					//--------------------------------------------
					case enumNodeTypes.FRA:  // Service Framework
					case enumNodeTypes.POR:  //Service Portfolio
							{
							if(this.Service_Portfolio_Section)
								{
								try
									{
									var rsPortfolios =
									from dsPortfolio in datacontexSDDP.ServicePortfolios
									where dsPortfolio.Id == node.NodeID
									select new
										{
										dsPortfolio.Id,
										dsPortfolio.Title,
										dsPortfolio.ISDHeading,
										dsPortfolio.ISDDescription
										};

									var recPortfolio = rsPortfolios.FirstOrDefault();

									Console.WriteLine("\t\t + {0} - {1}", recPortfolio.Id, recPortfolio.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: recPortfolio.ISDHeading,
										parIsNewSection: true);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServicePortfoliosURI +
												currentHyperlinkViewEditURI + recPortfolio.Id,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Porfolio Description
									if(this.Service_Portfolio_Description)
										{
										if(recPortfolio.ISDDescription != null)
											{
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServicePortfoliosURI +
													currentHyperlinkViewEditURI +
													recPortfolio.Id;
												}
											else
												currentListURI = "";


											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 1,
												parHTML2Decode: recPortfolio.ISDDescription,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									} //Try
								catch(DataServiceQueryException)
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
								catch(InvalidTableFormatException exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									// A Table content error occurred, record it in the error log.
									this.LogError("Error: The Deliverable ID: " + node.NodeID
										+ " contains an error in one of its Enahnce Rich Text columns. Please review the content (especially tables).");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "A content error occurred at this position and valid content could " +
										"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
								catch(Exception exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									}
								} // //if(this.Service_Portfolio_Section)
							break;
							}
					//-----------------------------------------
					case enumNodeTypes.FAM:  // Service Family
							{
							if(this.Service_Family_Heading)
								{
								try
									{
									var rsFamilies =
										from rsFamily in datacontexSDDP.ServiceFamilies
										where rsFamily.Id == node.NodeID
										select new
											{
											rsFamily.Id,
											rsFamily.Title,
											rsFamily.ISDHeading,
											rsFamily.ISDDescription
											};

									var recFamily = rsFamilies.FirstOrDefault();
									Console.WriteLine("\t\t + {0} - {1}", recFamily.Id, recFamily.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: recFamily.ISDHeading,
										parIsNewSection: false);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
											Properties.AppResources.List_ServiceFamiliesURI +
											currentHyperlinkViewEditURI + recFamily.Id,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Family Description
									if(this.Service_Family_Description)
										{
										if(recFamily.ISDDescription != null)
											{
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceFamiliesURI +
													currentHyperlinkViewEditURI +
													recFamily.Id;
												}
											else
												currentListURI = "";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 2,
												parHTML2Decode: recFamily.ISDDescription,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									} // Try
								catch(DataServiceClientException)
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
								catch(InvalidTableFormatException exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									// A Table content error occurred, record it in the error log.
									this.LogError("Error: The Deliverable ID: " + node.NodeID
										+ " contains an error in one of its Enahnce Rich Text columns. Please review the content (especially tables).");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "A content error occurred at this position and valid content could " +
										"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
								catch(Exception exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									}
								} // //if(this.Service_Portfolio_Section)
							break;
							}
					//------------------------------------------
					case enumNodeTypes.PRO:  // Service Product
							{
							if(this.Service_Product_Heading)
								{
								try
									{
									var rsProducts =
										from rsProduct in datacontexSDDP.ServiceProducts
										where rsProduct.Id == node.NodeID
										select new
											{
											rsProduct.Id,
											rsProduct.Title,
											rsProduct.ISDHeading,
											rsProduct.ISDDescription,
											rsProduct.KeyClientBenefits,
											rsProduct.KeyDDBenefits
											};

									var recProduct = rsProducts.FirstOrDefault();

									Console.WriteLine("\t\t + {0} - {1}", recProduct.Id, recProduct.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: recProduct.ISDHeading,
										parIsNewSection: false);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Product Description
									if(this.Service_Product_Description)
										{
										if(recProduct.ISDDescription != null)
											{
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceProductsURI +
													currentHyperlinkViewEditURI + recProduct.Id;
												}
											else
												currentListURI = "";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 3,
												parHTML2Decode: recProduct.ISDDescription,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									if(this.Service_Product_KeyDD_Benefits)
										{
										if(recProduct.KeyDDBenefits != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceProductsURI +
												currentHyperlinkViewEditURI +
												recProduct.Id;
											Console.WriteLine("\t\t + {0} - {1}", recProduct.Id, Properties.AppResources.Document_Product_KeyDD_Benefits);
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Product_KeyDD_Benefits,
												parIsNewSection: false);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceProductsURI +
													currentHyperlinkViewEditURI + recProduct.Id;
												}
											else
												currentListURI = "";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recProduct.KeyDDBenefits,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}

									if(this.Service_Product_Key_Client_Benefits)
										{
										if(recProduct.KeyClientBenefits != null)
											{
											Console.WriteLine("\t\t + {0} - {1}", recProduct.Id,
												Properties.AppResources.Document_Product_ClientKeyBenefits);
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Product_ClientKeyBenefits,
												parIsNewSection: false);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceProductsURI +
													currentHyperlinkViewEditURI + recProduct.Id;
												}
											else
												currentListURI = "";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recProduct.KeyClientBenefits,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									}
								catch(DataServiceClientException exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
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
								catch(InvalidTableFormatException exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									// A Table content error occurred, record it in the error log.
									this.LogError("Error: The Deliverable ID: " + node.NodeID
										+ " contains an error in one of its Enahnce Rich Text columns. Please review the content (especially tables).");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "A content error occurred at this position and valid content could " +
										"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
								catch(Exception exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									}
								} //if(this.Service_Product_Heading)
							break;
							}
					//------------------------------------------
					case enumNodeTypes.ELE:  // Service Element
							{
							if(this.Service_Element_Heading)
								{
								try
									{
									// Obtain the Element info from SharePoint
									var rsElements =
										from dsElement in datacontexSDDP.ServiceElements
										where dsElement.Id == node.NodeID
										select new
											{
											dsElement.Id,
											dsElement.Title,
											dsElement.ISDHeading,
											dsElement.ISDDescription,
											dsElement.Objective,
											dsElement.KeyClientAdvantages,
											dsElement.KeyClientBenefits,
											dsElement.KeyDDBenefits,
											dsElement.KeyPerformanceIndicators,
											dsElement.CriticalSuccessFactors,
											dsElement.ProcessLink
											};

									var recElement = rsElements.FirstOrDefault();

									Console.WriteLine("\t\t + {0} - {1}", recElement.Id, recElement.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: recElement.ISDHeading,
										parIsNewSection: false);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI + recElement.Id,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Service Element Description
									if(this.Service_Element_Description)
										{
										if(recElement.ISDDescription != null)
											{
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												recElement.Id;
												}
											else
												currentListURI = "";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recElement.ISDDescription,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									if(this.Service_Element_Objectives)
										{
										if(recElement.Objective != null)
											{
											Console.WriteLine("\t\t + {0} - {1}", recElement.Id,
												Properties.AppResources.Document_Element_Objectives);
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_Objectives,
												parIsNewSection: false);

											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												recElement.Id;
												}
											else
												currentListURI = "";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 5,
												parHTML2Decode: recElement.Objective,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}

									if(this.Service_Element_Critical_Success_Factors)
										{
										if(recElement.CriticalSuccessFactors != null)
											{
											// Insert the heading
											Console.WriteLine("\t\t + {0} - {1}", recElement.Id,
												Properties.AppResources.Document_Element_CriticalSuccessFactors);
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_CriticalSuccessFactors,
												parIsNewSection: false);

											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												recElement.Id;
												}
											else
												currentListURI = "";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 5,
												parHTML2Decode: recElement.CriticalSuccessFactors,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									if(this.Service_Element_Key_Client_Advantages)
										{
										if(recElement.KeyClientAdvantages != null)
											{
											// Insert the heading
											Console.WriteLine("\t\t + {0} - {1}", recElement.Id,
												Properties.AppResources.Document_Element_ClientKeyAdvantages);
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_ClientKeyAdvantages,
												parIsNewSection: false);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												recElement.Id;
												}
											else
												currentListURI = "";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 5,
												parHTML2Decode: recElement.KeyClientAdvantages,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									if(this.Service_Element_Key_Client_Benefits)
										{
										if(recElement.KeyClientBenefits != null)
											{
											// Insert the heading
											Console.WriteLine("\t\t + {0} - {1}", recElement.Id,
												Properties.AppResources.Document_Element_ClientKeyBenefits);
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_ClientKeyBenefits,
												parIsNewSection: false);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												recElement.Id;
												}
											else
												currentListURI = "";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 5,
												parHTML2Decode: recElement.KeyClientBenefits,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									if(this.Service_Element_Key_DD_Benefits)
										{
										if(recElement.KeyDDBenefits != null)
											{
											// Insert the heading
											Console.WriteLine("\t\t + {0} - {1}", recElement.Id,
												Properties.AppResources.Document_Element_KeyDDBenefits);
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_KeyDDBenefits,
												parIsNewSection: false);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												recElement.Id;
												}
											else
												currentListURI = "";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 5,
												parHTML2Decode: recElement.KeyDDBenefits,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									if(this.Service_Element_Key_Performance_Indicators)
										{
										if(recElement.KeyPerformanceIndicators != null)
											{
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												recElement.Id;
												}
											else
												currentListURI = "";

											// Insert the heading
											Console.WriteLine("\t\t + {0} - {1}", recElement.Id,
												Properties.AppResources.Document_Element_KPI);
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_KPI,
												parIsNewSection: false);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 5,
												parHTML2Decode: recElement.KeyPerformanceIndicators,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									if(this.Service_Element_High_Level_Process)
										{
										if(recElement.ProcessLink != null)
											{
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												recElement.Id;
												}
											else
												currentListURI = "";

											// Insert the heading
											Console.WriteLine("\t\t + {0} - {1}", recElement.Id,
												Properties.AppResources.Document_Element_KPI);
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_HighLevelProcess,
												parIsNewSection: false);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											//TODO: Insert generate hypelink in oxmlEncoder

											}
										}
									drmHeading = false;
									}
								catch(DataServiceClientException)
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
								catch(InvalidTableFormatException exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									// A Table content error occurred, record it in the error log.
									this.LogError("Error: The Deliverable ID: " + node.NodeID
										+ " contains an error in one of its Enahnce Rich Text columns. Please review the content (especially tables).");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "A content error occurred at this position and valid content could " +
										"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}

								catch(Exception exc)
									{
									Console.WriteLine("Exception occurred: {0} - {1}", exc.HResult, exc.Message);
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
							try
								{
								// Obtain the Deliverable info from SharePoint
								Deliverable objDeliverable = new Deliverable();
								objDeliverable.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: node.NodeID, parGetLayer1up: true);
								
								Console.WriteLine("\t\t + {0} - {1}", objDeliverable.ID, objDeliverable.Title);
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
								objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.ISDheading);
								
								// Check if a hyperlink must be inserted
								if(documentCollection_HyperlinkURL != "")
									{
									hyperlinkCounter += 1;
									Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
										parMainDocumentPart: ref objMainDocumentPart,
										parImageRelationshipId: hyperlinkImageRelationshipID,
										parClickLinkURL: Properties.AppResources.SharePointURL +
											Properties.AppResources.List_DeliverablesURI +
											currentHyperlinkViewEditURI + objDeliverable.ID,
										parHyperlinkID: hyperlinkCounter);
									objRun.Append(objDrawing);
									}
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);
								// Check if the user specified to include the Deliverable Description
								if(this.DRM_Description)
									{
									if(objDeliverable.ISDdescription != null)
										{
										if(documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverable.ID;
											}
										else
											currentListURI = "";

										if(this.ColorCodingLayer1)
											currentContentLayer = "Layer1";
										else
											currentContentLayer = "None";

										// Insert the contents
										objHTMLdecoder.DecodeHTML(
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 6,
											parHTML2Decode: objDeliverable.ISDdescription,
											parContentLayer: currentContentLayer,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter,
											parHyperlinkID: ref hyperlinkCounter,
											parPageHeightTwips: this.PageHight,
											parPageWidthTwips: this.PageWith);
										} //// if(recDeliverable.ISDDescription != null)
									} //if(this.Deliverable_Description)

								if(this.DRM_Inputs)
									{
									if(objDeliverable.Inputs != null)
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										if(documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverable.ID;
											}
										else
											currentListURI = "";

										if(this.ColorCodingLayer1)
											currentContentLayer = "Layer1";
										else
											currentContentLayer = "None";

										// Insert the contents
										objHTMLdecoder.DecodeHTML(
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 7,
											parHTML2Decode: objDeliverable.Inputs,
											parContentLayer: currentContentLayer,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter,
											parHyperlinkID: ref hyperlinkCounter,
											parPageHeightTwips: this.PageHight,
											parPageWidthTwips: this.PageWith);
										} // if(recDeliverable.Inputs != null)
									} //if(this.Deliverable_Inputs)

								// Check if the user specified to include the Deliverable Outputs
								if(this.DRM_Outputs)
									{
									if(objDeliverable.Outputs != null)
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableOutputs_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										if(documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverable.ID;
											}
										else
											currentListURI = "";

										if(this.ColorCodingLayer1)
											currentContentLayer = "Layer1";
										else
											currentContentLayer = "None";

										// Insert the contents
										objHTMLdecoder.DecodeHTML(
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 7,
											parHTML2Decode: objDeliverable.Outputs,
											parContentLayer: currentContentLayer,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter,
											parHyperlinkID: ref hyperlinkCounter,
											parPageHeightTwips: this.PageHight,
											parPageWidthTwips: this.PageWith);
										} // if(recDeliverable.Outputs != null)
									} //if(this.Deliverable_Outputs)

								// Check if the user specified to include the Deliverable DD's Obligations
								if(this.DDS_DRM_Obligations)
									{
									if(objDeliverable.DDobligations != null)
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										if(documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverable.ID;
											}
										else
											currentListURI = "";

										if(this.ColorCodingLayer1)
											currentContentLayer = "Layer1";
										else
											currentContentLayer = "None";

										// Insert the contents
										objHTMLdecoder.DecodeHTML(
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 7,
											parHTML2Decode: objDeliverable.DDobligations,
											parContentLayer: currentContentLayer,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter,
											parHyperlinkID: ref hyperlinkCounter,
											parPageHeightTwips: this.PageHight,
											parPageWidthTwips: this.PageWith);
										} // if(recDeliverable.SPObligations != null)
									} //if(this.DDS_Deliverable_Oblidations)

								// Check if the user specified to include the Client Responsibilities
								if(this.Clients_DRM_Responsibilities)
									{
									if(objDeliverable.ClientResponsibilities != null)
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										if(documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverable.ID;
											}
										else
											currentListURI = "";

										if(this.ColorCodingLayer1)
											currentContentLayer = "Layer1";
										else
											currentContentLayer = "None";

										// Insert the contents
										objHTMLdecoder.DecodeHTML(
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 7,
											parHTML2Decode: objDeliverable.ClientResponsibilities,
											parContentLayer: currentContentLayer,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter,
											parHyperlinkID: ref hyperlinkCounter,
											parPageHeightTwips: this.PageHight,
											parPageWidthTwips: this.PageWith);
										} // if(recDeliverable.Client_Responsibilities != null)
									} //if(this.Clients_Deliverable_Responsibilities)

								// Check if the user specified to include the Deliverable Exclusions
								if(this.DRM_Exclusions)
									{
									if(objDeliverable.Exclusions != null)
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										if(documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverable.ID;
											}
										else
											currentListURI = "";

										if(this.ColorCodingLayer1)
											currentContentLayer = "Layer1";
										else
											currentContentLayer = "None";

										// Insert the contents
										objHTMLdecoder.DecodeHTML(
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 7,
											parHTML2Decode: objDeliverable.Exclusions,
											parContentLayer: currentContentLayer,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter,
											parHyperlinkID: ref hyperlinkCounter,
											parPageHeightTwips: this.PageHight,
											parPageWidthTwips: this.PageWith);
										} // if(recDeliverable.Exclusions != null)
									} //if(this.Deliverable_Exclusions)

								// Check if the user specified to include the Governance Controls
								if(this.DRM_Governance_Controls)
									{
									if(objDeliverable.GovernanceControls != null)
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableGovernanceControls_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										if(documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												objDeliverable.ID;
											}
										else
											currentListURI = "";

										if(this.ColorCodingLayer1)
											currentContentLayer = "Layer1";
										else
											currentContentLayer = "None";

										// Insert the contents
										objHTMLdecoder.DecodeHTML(
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 7,
											parHTML2Decode: objDeliverable.GovernanceControls,
											parContentLayer: currentContentLayer,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter,
											parHyperlinkID: ref hyperlinkCounter,
											parPageHeightTwips: this.PageHight,
											parPageWidthTwips: this.PageWith);
										} // if(recDeliverable.GovernanceControls != null)
									} //if(this.Deliverable_GovernanceControls)

								// Check if there are any Glossary Terms or Acronyms associated with the Deliverable.
								if(objDeliverable.GlossaryAndAcronyms.Count > 0)
									{
									// Check if the user selected Acronyms and Glossy of Terms are requied
									if(this.Acronyms_Glossary_of_Terms_Section)
										{
										if(this.Acronyms || this.Glossary_of_Terms)
											{
											foreach(var entry in objDeliverable.GlossaryAndAcronyms)
												{
												if(this.DictionaryGlossaryAndAcronyms.ContainsKey(entry.Key) != true)
													DictionaryGlossaryAndAcronyms.Add(entry.Key, entry.Value);
												Console.WriteLine("\t\t\t + Term & Acronym added: {0} - {1}", entry.Key, entry.Value);
												}
											} // if(this.Acronyms || this.Glossary_of_Terms)
										} // if(this.Acronyms_Glossary_of_Terms_Section)
									} //if(recDeliverable.GlossaryAndAcronyms.Count > 0)
								} //try
							catch(DataServiceClientException)
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
							catch(InvalidTableFormatException exc)
								{
								Console.WriteLine("Exception occurred: {0}", exc.Message);
								// A Table content error occurred, record it in the error log.
								this.LogError("Error: The Deliverable ID: " + node.NodeID
									+ " contains an error in one of its Enahnce Rich Text columns. Please review the content (especially tables).");
								objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 7);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: "A content error occurred at this position and valid content could " +
									"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it.",
									parIsNewSection: false,
									parIsError: true);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);
								}
							catch(Exception exc)
								{
								Console.WriteLine("Exception occurred: {0} - {1}", exc.HResult, exc.Message);
								}
							break;
							}
					//--------------------------------
					case enumNodeTypes.EAC:  // Activity associated with Deliverable pertaining to Service Element
							{
							if(this.Activities)
								{
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Activities_Heading);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);
								try
									{
									// Obtain the Activity info from SharePoint
									var rsActivities =
										from rsActivity in datacontexSDDP.Activities
										where rsActivity.Id == node.NodeID
										select new
											{
											rsActivity.Id,
											rsActivity.Title,
											rsActivity.ISDHeading,
											rsActivity.ISDDescription,
											rsActivity.ActivityInput,
											rsActivity.ActivityOutput,
											rsActivity.ActivityOptionalityValue,
											rsActivity.ActivityAssumptions
											};

									var recActivity = rsActivities.FirstOrDefault();
									Console.WriteLine("\t\t + {0} - {1}", recActivity.Id, recActivity.Title);

									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 8);
									objRun = oxmlDocument.Construct_RunText(parText2Write: recActivity.ISDHeading);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ActvitiesURI +
												currentHyperlinkViewEditURI + recActivity.Id,
											parHyperlinkID: hyperlinkCounter);
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
											parActivityDesciption: recActivity.ISDDescription,
											parActivityInput: recActivity.ActivityInput,
											parActivityOutput: recActivity.ActivityOutput,
											parActivityAssumptions: recActivity.ActivityAssumptions,
											parActivityOptionality: recActivity.ActivityOptionalityValue);
										objBody.Append(objActivityTable);
										} // if (this.Activity_Description_Table)
									} // try
								catch(DataServiceClientException)
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
									this.LogError("Error: The Activity ID " + node.NodeID
										+ " doesn't exist in SharePoint and it couldn't be retrieved.");
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 8);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Error: Activity " + node.NodeID + " is missing.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									break;
									}

								catch(Exception exc)
									{
									Console.WriteLine("Exception occurred: {0} - {1}", exc.HResult, exc.Message);
									}
								} // if (this.Activities)
							break;
							}
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

								// Check if the user specified to include the Deliverable Description
								if(this.Service_Level_Commitments_Table)
									{
									// Prepare the data which to insert into the Service Level Table
									try
										{
										// Obtain the Deliverable Service Level from SharePoint
										var rsDeliverableServiceLevels =
											from rsDeliverableServiceLevel in datacontexSDDP.DeliverableServiceLevels
											where rsDeliverableServiceLevel.Id == node.NodeID
											select new
												{
												rsDeliverableServiceLevel.Id,
												rsDeliverableServiceLevel.Title,
												rsDeliverableServiceLevel.Service_LevelId,
												rsDeliverableServiceLevel.AdditionalConditions
												};

										var recDeliverableServiceLevel = rsDeliverableServiceLevels.FirstOrDefault();
										Console.WriteLine("\t\t + Deliverable ServiceLevel: {0} - {1}", recDeliverableServiceLevel.Id,
											recDeliverableServiceLevel.Title);

										// Obtain the Service Level info from SharePoint
										var dsServiceLevels = datacontexSDDP.ServiceLevels
											.Expand(sl => sl.Service_Hour);

										var rsServiceLevels =
											from rsServiceLevel in dsServiceLevels
											where rsServiceLevel.Id == recDeliverableServiceLevel.Service_LevelId
											select rsServiceLevel;

										var recServiceLevel = rsServiceLevels.FirstOrDefault();
										Console.WriteLine("\t\t + Service Level: {0} - {1}", recServiceLevel.Id, recServiceLevel.Title);
										Console.WriteLine("\t\t + Service Hour.: {0}", recServiceLevel.Service_Hour.Title);

										// Obtain the Service Level Thresholds from SharePoint
										var rsServiceLevelThresholds =
											from dsSLthresholds in datacontexSDDP.ServiceLevelTargets
											where dsSLthresholds.Service_LevelId == recServiceLevel.Id
												&& dsSLthresholds.ThresholdOrTargetValue == "Threshold"
											orderby dsSLthresholds.Title
											select new
												{
												dsSLthresholds.Id,
												dsSLthresholds.Title
												};
										// load the SL Thresholds into a list - apckaging it in order to send it as a parameter later on.
										List<string> listServiceLevelThresholds = new List<string>();
										foreach(var recSLthreshold in rsServiceLevelThresholds)
											{
											listServiceLevelThresholds.Add(recSLthreshold.Title);
											Console.WriteLine("\t\t\t + Threshold: {0} - {1}", recSLthreshold.Id, recSLthreshold.Title);
											}

										// Obtain the Service Level Targets from SharePoint
										var rsServiceLevelTargets =
											from dsSLTargets in datacontexSDDP.ServiceLevelTargets
											where dsSLTargets.Service_LevelId == recServiceLevel.Id && dsSLTargets.ThresholdOrTargetValue == "Target"
											orderby dsSLTargets.Title
											select new
												{
												dsSLTargets.Id,
												dsSLTargets.Title
												};
										// load the SL Targets into a list - apckaging it in order to send it as a parameter later on.
										List<string> listServiceLevelTargets = new List<string>();
										foreach(var recSLtarget in rsServiceLevelTargets)
											{
											listServiceLevelTargets.Add(recSLtarget.Title);
											Console.WriteLine("\t\t\t + Threshold: {0} - {1}", recSLtarget.Id, recSLtarget.Title);
											}

										// Insert the Service Level ISD Description
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 8);
										objRun = oxmlDocument.Construct_RunText(parText2Write: recServiceLevel.ISDHeading);
										// Check if a hyperlink must be inserted
										if(documentCollection_HyperlinkURL != "")
											{
											hyperlinkCounter += 1;
											Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
												parMainDocumentPart: ref objMainDocumentPart,
												parImageRelationshipId: hyperlinkImageRelationshipID,
												parClickLinkURL: Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceLevelsURI +
													currentHyperlinkViewEditURI + recServiceLevel.Id,
												parHyperlinkID: hyperlinkCounter);
											objRun.Append(objDrawing);
											}
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										List<string> listErrorMessagesParameter = this.ErrorMessages;
										// Populate the Service Level Table
										objServiceLevelTable = CommonProcedures.BuildSLAtable(
											parServiceLevelID: recServiceLevel.Id,
											parWidthColumn1: (this.PageWith * 30) / 100,
											parWidthColumn2: (this.PageWith * 70) / 100,
											parMeasurement: recServiceLevel.ServiceLevelMeasurement,
											parMeasureMentInterval: recServiceLevel.MeasurementIntervalValue,
											parReportingInterval: recServiceLevel.ReportingIntervalValue,
											parServiceHours: recServiceLevel.Service_Hour.Title,
											parCalculationMethod: recServiceLevel.CalculationMethod,
											parCalculationFormula: recServiceLevel.CalculationFormula,
											parThresholds: listServiceLevelThresholds,
											parTargets: listServiceLevelTargets,
											parBasicServiceLevelConditions: recServiceLevel.BasicServiceLevelConditions,
											parAdditionalServiceLevelConditions: recDeliverableServiceLevel.AdditionalConditions,
											parErrorMessages: ref listErrorMessagesParameter);

										if(listErrorMessagesParameter.Count != this.ErrorMessages.Count)
											this.ErrorMessages = listErrorMessagesParameter;

										objBody.Append(objServiceLevelTable);
										} // try
									catch(DataServiceClientException)
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
									catch(Exception exc)
										{
										Console.WriteLine("Exception occurred: {0} - {1}", exc.HResult, exc.Message);
										}
									} // if (this.Service Level_Description_Table)
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

					List<string> listErrors = this.ErrorMessages;
					if(this.DictionaryGlossaryAndAcronyms.Count > 0)
						{
						Table tableGlossaryAcronym = new Table();
						tableGlossaryAcronym = CommonProcedures.BuildGlossaryAcronymsTable(
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
							parTableCaptionCounter: ref tableCaptionCounter,
							parImageCaptionCounter: ref imageCaptionCounter,
							parHyperlinkID: ref hyperlinkCounter);
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

				Console.WriteLine(
					"Generation started...: {0} \nGeneration completed: {1} \n Durarion..........: {2}",
					timeStarted,
					DateTime.Now,
					(DateTime.Now - timeStarted));
				} // end Try

			catch(OpenXmlPackageException exc)
				{
				//TODO: add code to catch exception.
				}
			catch(ArgumentNullException exc)
				{
				//TODO: add code to catch exception.
				}

			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}
		
		} // end of Services_Framework_Document_DRM_Inline class
	}
