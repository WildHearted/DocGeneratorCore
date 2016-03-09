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
	/// This class represent the Internal Service Definition (ISD) with sperate DRM (Deliverable Report Meeting) sections
	/// It inherits from the Internal DRM Sections Class.
	/// </summary>
	class ISD_Document_DRM_Sections:Internal_DRM_Sections
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
			bool layerHeadingWritten = false;
			bool drmHeading = false;
			Table objActivityTable = new Table();
			Table objServiceLevelTable = new Table();
			Dictionary<int, string> dictDeliverables = new Dictionary<int, string>();
			Dictionary<int, string> dictReports = new Dictionary<int, string>();
			Dictionary<int, string> dictMeetings = new Dictionary<int, string>();
			Dictionary<int, string> dictSLAs = new Dictionary<int, string>();
			int? layer1upElementID = 0;
			int? layer2upElementID = 0;
			int tableCaptionCounter = 1;
			int imageCaptionCounter = 1;
			int hyperlinkCounter = 4;


			if(this.HyperlinkEdit)
				documentCollection_HyperlinkURL = Properties.AppResources.SharePointSiteURL +
					Properties.AppResources.List_DocumentCollectionLibraryURI +
					Properties.AppResources.EditFormURI + this.DocumentCollectionID;
			currentHyperlinkViewEditURI = Properties.AppResources.EditFormURI;
			if(this.Hyperlink_View)
				documentCollection_HyperlinkURL = Properties.AppResources.SharePointSiteURL +
					Properties.AppResources.List_DocumentCollectionLibraryURI +
					Properties.AppResources.DisplayFormURI + this.DocumentCollectionID;
			currentHyperlinkViewEditURI = Properties.AppResources.DisplayFormURI;
			

			//Initialize the Data access to SharePoint
			DesignAndDeliveryPortfolioDataContext datacontexSDDP = new DesignAndDeliveryPortfolioDataContext(new
				Uri(Properties.AppResources.SharePointSiteURL + Properties.AppResources.SharePointRESTuri));
			datacontexSDDP.Credentials = CredentialCache.DefaultCredentials;
			datacontexSDDP.MergeOption = System.Data.Services.Client.MergeOption.NoTracking;

			// define a new objOpenXMLdocument
			oxmlDocument objOXMLdocument = new oxmlDocument();
			// use CreateDocumentFromTemplate method to create a new MS Word Document based on the relevant template
			if(objOXMLdocument.CreateDocumentFromTemplate(parTemplateURL: this.Template, parDocumentType: this.DocumentType))
				{
				Console.WriteLine("\t\t objOXMLdocument:\n" +
				"\t\t\t+ LocalDocumentPath: {0}\n" +
				"\t\t\t+ DocumentFileName.: {1}\n" +
				"\t\t\t+ DocumentURI......: {2}", objOXMLdocument.LocalDocumentPath, objOXMLdocument.DocumentFilename, objOXMLdocument.LocalDocumentURI);
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
				WordprocessingDocument objWPdocument = WordprocessingDocument.Open(path: objOXMLdocument.LocalDocumentURI, isEditable: true);
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

				Console.WriteLine("\t\t Effective pageWidth x pageHeight.: {0} x {1} twips", this.PageWith, this.PageHight);

				// Check whether Hyperlinks need to be included
				if(this.HyperlinkEdit || this.Hyperlink_View)
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
						Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
						Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
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
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServicePortfoliosURI +
												currentHyperlinkViewEditURI + recPortfolio.Id;
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 1,
												parHTML2Decode: recPortfolio.ISDDescription,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
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
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServicePortfoliosURI +
												currentHyperlinkViewEditURI +
												recFamily.Id;
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 2,
												parHTML2Decode: recFamily.ISDDescription,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
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
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
											Properties.AppResources.List_ServiceProductsURI +
											currentHyperlinkViewEditURI + recProduct.Id,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Product Description
									if(this.Service_Product_Description)
										{
										if(recProduct.ISDDescription != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceProductsURI +
												currentHyperlinkViewEditURI +
												recProduct.Id;

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 3,
												parHTML2Decode: recProduct.ISDDescription,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
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
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parClickLinkURL: Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceProductsURI +
													currentHyperlinkViewEditURI + recProduct.Id,
													parHyperlinkID: hyperlinkCounter);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recProduct.KeyDDBenefits,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}

									if(this.Service_Product_Key_Client_Benefits)
										{
										if(recProduct.KeyClientBenefits != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceProductsURI +
												currentHyperlinkViewEditURI +
												recProduct.Id;

											Console.WriteLine("\t\t + {0} - {1}", recProduct.Id,
												Properties.AppResources.Document_Product_ClientKeyBenefits);
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Product_ClientKeyBenefits,
												parIsNewSection: false);
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parClickLinkURL: Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceProductsURI +
													currentHyperlinkViewEditURI + recProduct.Id,
													parHyperlinkID: hyperlinkCounter);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recProduct.KeyClientBenefits,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
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
									ServiceElement objServiceElement = new ServiceElement();
									objServiceElement.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: node.NodeID, parGetLayer1up: true);
									
									// Insert the Service Element ISD Heading...
									Console.WriteLine("\t\t + Service Element Layer 0..: {0} - {1}", objServiceElement.ID, objServiceElement.Title);
									if(objServiceElement.ContentPredecessorElementID != null)
										{
										Console.WriteLine("\t\t + Service Element Layer 1up: {0} - {1}", 
											objServiceElement.Layer1up.ID, objServiceElement.Layer1up.Title);
										if(objServiceElement.Layer1up.ContentPredecessorElementID != null)
											{
											Console.WriteLine("\t\t + Service Element Layer 2up: {0} - {1}",
												objServiceElement.Layer1up.Layer1up.ID, objServiceElement.Layer1up.Layer1up.Title);
											}
										}

									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: objServiceElement.ISDheading,
										parIsNewSection: false);
										objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									
									//Check if the Element Layer0up has Content Layers and Content Predecessors
									if(objServiceElement.ContentPredecessorElementID == null)
										{
										layer1upElementID = null;
										layer2upElementID = null;
										}
									else
										{
										layer1upElementID = objServiceElement.ContentPredecessorElementID;
										if(objServiceElement.Layer1up.ContentPredecessorElementID == null)
											{layer2upElementID = null;}
										else
											{layer2upElementID = objServiceElement.Layer1up.ContentPredecessorElementID;}
										}

									// Check if the user specified to include the Service Service Element Description
									if(this.Service_Element_Description)
										{
										// Insert Layer 2up if present and not null
										if(layer2upElementID != null)
											{
											if(objServiceElement.Layer1up.Layer1up.ISDdescription != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_ServiceElementsURI +
														currentHyperlinkViewEditURI +
														layer2upElementID;
													}
												else
													{currentListURI = "";}

												if(this.ColorCodingLayer1)
													{currentContentLayer = "Layer1";}
												else
													{currentContentLayer = "None";}

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objServiceElement.Layer1up.Layer1up.ISDdescription,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											} //// if(layer2upElementID != null)
											
										// Insert Layer 1up if present and not null
										if(layer1upElementID != null)
											{
											if(objServiceElement.Layer1up.ISDdescription != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_ServiceElementsURI +
														currentHyperlinkViewEditURI +
														layer1upElementID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													{currentContentLayer = "Layer2";}
												else
													{currentContentLayer = "None";}

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objServiceElement.Layer1up.ISDdescription,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											} //// if(layer2upElementID != null)
											
										// Insert Layer 0up if not null
										if(objServiceElement.ISDdescription != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceElementsURI +
													currentHyperlinkViewEditURI +
													objServiceElement.ID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: objServiceElement.ISDdescription,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
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
										if(layer2upElementID != null)
											{
											if(objServiceElement.Layer1up.Layer1up.Objectives != null)
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
														layer2upElementID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objServiceElement.Layer1up.Layer1up.Objectives,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											} // if(layer2upElementID != null)
										} // if(this.PresentationMode == enumPresentationMode.Layered)

									// Insert Layer 1up if resent and not null
									if(layer1upElementID != null)
										{
										if(objServiceElement.Layer1up.Objectives != null)
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
													layer1upElementID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 5,
												parHTML2Decode: objServiceElement.Layer1up.Objectives,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										} //// if(layer2upElementID != null)

									// Insert Layer 0up if not null
									if(objServiceElement.Objectives != null)
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
												objServiceElement.ID;
											}
										else
											currentListURI = "";

										if(this.ColorCodingLayer1)
											currentContentLayer = "Layer3";
										else
											currentContentLayer = "None";

										objHTMLdecoder.DecodeHTML(
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: objServiceElement.Objectives,
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter,
											parHyperlinkID: ref hyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightTwips: this.PageHight,
											parPageWidthTwips: this.PageWith);
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
										if(layer2upElementID != null)
											{
											if(objServiceElement.Layer1up.Layer1up.CriticalSuccessFactors != null)
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
														layer2upElementID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objServiceElement.Layer1up.Layer1up.CriticalSuccessFactors,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											} // if(layer2upElementID != null)
										} // if (this.PresentationMode == Layered)

									// Insert Layer 1up if resent and not null
									if(layer1upElementID != null)
										{
										if(objServiceElement.Layer1up.CriticalSuccessFactors != null)
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
													layer1upElementID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 5,
												parHTML2Decode: objServiceElement.Layer1up.CriticalSuccessFactors,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										} //// if(layer2upElementID != null)

									// Insert Layer 0up if not null
									if(objServiceElement.CriticalSuccessFactors != null)
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
												objServiceElement.ID;
											}
										else
											currentListURI = "";

										if(this.ColorCodingLayer1)
											currentContentLayer = "Layer3";
										else
											currentContentLayer = "None";

										objHTMLdecoder.DecodeHTML(
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: objServiceElement.CriticalSuccessFactors,
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter,
											parHyperlinkID: ref hyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightTwips: this.PageHight,
											parPageWidthTwips: this.PageWith);
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

									// Insert Layer 2up if present and not null
									if(this.PresentationMode == enumPresentationMode.Layered)
										{
										if(layer2upElementID != null)
											{
											if(objServiceElement.Layer1up.Layer1up.KeyClientAdvantages != null)
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
														layer2upElementID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objServiceElement.Layer1up.Layer1up.KeyClientAdvantages,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											} // if(layer2upElementID != null)
										} // if(this.PresentationMode == Layered)

									// Insert Layer 1up if resent and not null
									if(layer1upElementID != null)
										{
										if(objServiceElement.Layer1up.KeyClientAdvantages != null)
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
													layer1upElementID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 5,
												parHTML2Decode: objServiceElement.Layer1up.KeyClientAdvantages,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										} //// if(layer2upElementID != null)

									// Insert Layer 0up if not null
									if(objServiceElement.KeyClientAdvantages != null)
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
												objServiceElement.ID;
											}
										else
											currentListURI = "";

										if(this.ColorCodingLayer1)
											currentContentLayer = "Layer3";
										else
											currentContentLayer = "None";

										objHTMLdecoder.DecodeHTML(
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: objServiceElement.KeyClientAdvantages,
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter,
											parHyperlinkID: ref hyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightTwips: this.PageHight,
											parPageWidthTwips: this.PageWith);
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
									if (this.PresentationMode == enumPresentationMode.Layered)
										{
										if(layer2upElementID != null)
											{
											if(objServiceElement.Layer1up.Layer1up.KeyClientBenefits != null)
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
														layer2upElementID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objServiceElement.Layer1up.Layer1up.KeyClientBenefits,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											} //// if(layer2upElementID != null)
										}
									// Insert Layer 1up if resent and not null
									if(layer1upElementID != null)
										{
										if(objServiceElement.Layer1up.KeyClientBenefits != null)
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
													layer1upElementID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 5,
												parHTML2Decode: objServiceElement.Layer1up.KeyClientBenefits,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										} //// if(layer2upElementID != null)

									// Insert Layer 0up if not null
									if(objServiceElement.KeyClientBenefits != null)
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
												objServiceElement.ID;
											}
										else
											currentListURI = "";

										if(this.ColorCodingLayer1)
											currentContentLayer = "Layer3";
										else
											currentContentLayer = "None";

										objHTMLdecoder.DecodeHTML(
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: objServiceElement.KeyClientBenefits,
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter,
											parHyperlinkID: ref hyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightTwips: this.PageHight,
											parPageWidthTwips: this.PageWith);
										}
									} //if(this.Service_Element_KeyClientBenefits)

								//-----------------------------
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
										if(layer2upElementID != null)
											{
											if(objServiceElement.Layer1up.Layer1up.KeyDDbenefits != null)
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
														layer2upElementID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objServiceElement.Layer1up.Layer1up.KeyDDbenefits,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											} //// if(layer2upElementID != null)
										}
									// Insert Layer 1up if resent and not null
									if(layer1upElementID != null)
										{
										if(objServiceElement.Layer1up.KeyDDbenefits != null)
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
													layer1upElementID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 5,
												parHTML2Decode: objServiceElement.Layer1up.KeyDDbenefits,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										} //// if(layer2upElementID != null)

									// Insert Layer 0up if not null
									if(objServiceElement.KeyDDbenefits != null)
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
												objServiceElement.ID;
											}
										else
											currentListURI = "";

										if(this.ColorCodingLayer1)
											currentContentLayer = "Layer3";
										else
											currentContentLayer = "None";

										objHTMLdecoder.DecodeHTML(
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: objServiceElement.KeyDDbenefits,
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter,
											parHyperlinkID: ref hyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightTwips: this.PageHight,
											parPageWidthTwips: this.PageWith);
										}
									} //if(this.Service_Element_KeyDDbenefits)

								//--------------------------------------
								// Insert the Key Performance Indicators
								// Check if the user specified to include the Service Element Key Performance Indicators
								if(this.Service_Element_Description)
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
										if(layer2upElementID != null)
											{
											if(objServiceElement.Layer1up.Layer1up.KeyPerformanceIndicators != null)
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
														layer2upElementID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objServiceElement.Layer1up.Layer1up.KeyPerformanceIndicators,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											} //// if(layer2upElementID != null)
										}
									// Insert Layer 1up if resent and not null
									if(layer1upElementID != null)
										{
										if(objServiceElement.Layer1up.KeyPerformanceIndicators != null)
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
													layer1upElementID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 5,
												parHTML2Decode: objServiceElement.Layer1up.KeyPerformanceIndicators,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										} //// if(layer2upElementID != null)

									// Insert Layer 0up if not null
									if(objServiceElement.KeyPerformanceIndicators != null)
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
												objServiceElement.ID;
											}
										else
											currentListURI = "";

										if(this.ColorCodingLayer1)
											currentContentLayer = "Layer3";
										else
											currentContentLayer = "None";

										objHTMLdecoder.DecodeHTML(
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: objServiceElement.KeyPerformanceIndicators,
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter,
											parHyperlinkID: ref hyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightTwips: this.PageHight,
											parPageWidthTwips: this.PageWith);
										}
									} //if(this.Service_Element_KeyPerformanceIndicators)

								//--------------------------
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
										if(layer2upElementID != null)
											{
											if(objServiceElement.Layer1up.Layer1up.ProcessLink != null)
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
														layer2upElementID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objServiceElement.Layer1up.Layer1up.ProcessLink,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											} //// if(layer2upElementID != null)
										}
									// Insert Layer 1up if resent and not null
									if(layer1upElementID != null)
										{
										if(objServiceElement.Layer1up.ProcessLink != null)
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
													layer1upElementID;
												}
											else
												currentListURI = "";

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 5,
												parHTML2Decode: objServiceElement.Layer1up.ProcessLink,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										} //// if(layer2upElementID != null)

									// Insert Layer 0up if not null
									if(objServiceElement.ProcessLink != null)
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
												objServiceElement.ID;
											}
										else
											currentListURI = "";

										if(this.ColorCodingLayer1)
											currentContentLayer = "Layer3";
										else
											currentContentLayer = "None";

										objHTMLdecoder.DecodeHTML(
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 5,
											parHTML2Decode: objServiceElement.ProcessLink,
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter,
											parHyperlinkID: ref hyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightTwips: this.PageHight,
											parPageWidthTwips: this.PageWith);
										}
									} //if(this.Service_Element_HighLevelProcess)
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
								var rsDeliverables =
									from dsDeliverable in datacontexSDDP.Deliverables
									where dsDeliverable.Id == node.NodeID
									select new
										{
										dsDeliverable.Id,
										dsDeliverable.Title,
										dsDeliverable.ISDHeading,
										dsDeliverable.ISDSummary
										};

								var recDeliverable = rsDeliverables.FirstOrDefault();
								Console.WriteLine("\t\t + {0} - {1}", recDeliverable.Id, recDeliverable.Title);
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
								objRun = oxmlDocument.Construct_RunText(parText2Write: recDeliverable.ISDHeading);
								if(node.NodeType == enumNodeTypes.ELD)
									{
									if(dictDeliverables.ContainsKey(recDeliverable.Id) != true)
										dictDeliverables.Add(recDeliverable.Id, recDeliverable.ISDHeading);
									}
								else if(node.NodeType == enumNodeTypes.ELR)
									{
									if(dictReports.ContainsKey(recDeliverable.Id) != true)
										dictReports.Add(recDeliverable.Id, recDeliverable.ISDHeading);
									}
								else if(node.NodeType == enumNodeTypes.ELM)
									{
									if(dictMeetings.ContainsKey(recDeliverable.Id) != true)
										dictMeetings.Add(recDeliverable.Id, recDeliverable.ISDHeading);
									}
								// Check if a hyperlink must be inserted
								if(documentCollection_HyperlinkURL != "")
									{
									hyperlinkCounter += 1;
									Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
										parMainDocumentPart: ref objMainDocumentPart,
										parImageRelationshipId: hyperlinkImageRelationshipID,
										parClickLinkURL: Properties.AppResources.SharePointURL +
											Properties.AppResources.List_DeliverablesURI +
											currentHyperlinkViewEditURI + recDeliverable.Id,
										parHyperlinkID: hyperlinkCounter);
									objRun.Append(objDrawing);
									}
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);
								// Check if the user specified to include the Deliverable Description
								if(this.DRM_Summary)
									{
									if(recDeliverable.ISDSummary != null)
										{
										currentListURI = Properties.AppResources.SharePointURL +
											Properties.AppResources.List_DeliverablesURI +
											currentHyperlinkViewEditURI +
											recDeliverable.Id;
										if(this.ColorCodingLayer1)
											currentContentLayer = "Layer1";
										else
											currentContentLayer = "None";

										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
										objRun = oxmlDocument.Construct_RunText(parText2Write: recDeliverable.ISDSummary);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);
										}
									} // if(this.DeliverableSummary

								// Insert the Hyperlink to the relevant position in the DRM Section.
								objParagraph = oxmlDocument.Construct_BookmarkHyperlink(
									parBodyTextLevel: 6,
									parBookmarkValue: "Deliverable_" + recDeliverable.Id);
								objBody.Append(objParagraph);
								}
							catch(DataServiceClientException)
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
							catch(InvalidTableFormatException exc)
								{
								Console.WriteLine("Exception occurred: {0}", exc.Message);
								// A Table content error occurred, record it in the error log.
								this.LogError("Error: The Deliverable ID: " + node.NodeID
									+ " contains an error in one of its Enahnce Rich Text columns. Please review the content (especially tables).");
								objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
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
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
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

									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
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
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
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
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_ServiceLevels_Heading_Text);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);

								// Check if the user specified to include the Deliverable Description
								if(this.Activity_Description_Table)
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

										// Add the Service Level entry to the Service Level Dictionay (list)
										if(dictSLAs.ContainsKey(recServiceLevel.Id) != true)
											{
											// NOTE: the DeliverableServiceLevel ID is used NOT the ServiceLevel ID.
											dictSLAs.Add(recDeliverableServiceLevel.Id, recServiceLevel.ISDHeading);
											}

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
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
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
											parWidthColumn1: Convert.ToUInt32(this.PageWith * 0.30),
											parWidthColumn2: Convert.ToUInt32(this.PageWith * 0.70),
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
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
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

				//------------------------------------------------------
				// Insert the Deliverable, Report, Meeting (DRM) Section
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
						goto Process_ServiceLevels;

					if(dictDeliverables.Count == 0)
						goto Process_Reports;

					if(this.Deliverables)
						{
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
								try
									{
									// Obtain the Deliverable info from SharePoint
									var dsDeliverables = datacontexSDDP.Deliverables
										.Expand(p => p.GlossaryAndAcronyms);

									var rsDeliverables =
										from dsDeliverable in dsDeliverables
										where dsDeliverable.Id == deliverableItem.Key
										select dsDeliverable;

									var recDeliverable = rsDeliverables.FirstOrDefault();
									Console.WriteLine("\t\t + {0} - {1}", recDeliverable.Id, recDeliverable.Title);

									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3, parBookMark: deliverableBookMark + recDeliverable.Id);
									objRun = oxmlDocument.Construct_RunText(parText2Write: recDeliverable.ISDHeading);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI + recDeliverable.Id,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									// Check if the user specified to include the Deliverable Description
									if(this.Deliverable_Description)
										{
										if(recDeliverable.ISDDescription != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 3,
												parHTML2Decode: recDeliverable.ISDDescription,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.ISDDescription != null)
										} //if(this.Deliverable_Description)

									// Check if the user specified to include the Deliverable Inputs
									if(this.Deliverable_Inputs)
										{
										if(recDeliverable.Inputs != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.Inputs,
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
									if(this.Deliverable_Outputs)
										{
										if(recDeliverable.Outputs != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableOutputs_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.Outputs,
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
									if(this.DDs_Deliverable_Obligations)
										{
										if(recDeliverable.SPObligations != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.SPObligations,
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
									if(this.Clients_Deliverable_Responsibilities)
										{
										if(recDeliverable.ClientResponsibilities != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.ClientResponsibilities,
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
									if(this.Deliverable_Exclusions)
										{
										if(recDeliverable.Exclusions != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.Exclusions,
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
									if(this.Deliverable_Governance_Controls)
										{
										if(recDeliverable.GovernanceControls != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableGovernanceControls_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.GovernanceControls,
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
									if(recDeliverable.GlossaryAndAcronyms.Count > 0)
										{
										// Check if the user selected Acronyms and Glossy of Terms are requied
										if(this.Acronyms_Glossary_of_Terms_Section)
											{
											if(this.Acronyms || this.Glossary_of_Terms)
												{
												foreach(var entry in recDeliverable.GlossaryAndAcronyms)
													{
													if(this.DictionaryGlossaryAndAcronyms.ContainsKey(entry.Id) != true)
														DictionaryGlossaryAndAcronyms.Add(entry.Id, entry.Title);
													Console.WriteLine("\t\t\t + Term & Acronym added: {0} - {1}", entry.Id, entry.Title);
													}
												} // if(this.Acronyms || this.Glossary_of_Terms)
											} // if(this.Acronyms_Glossary_of_Terms_Section)
										} //if(recDeliverable.GlossaryAndAcronyms.Count > 0)
									} //Try
								catch(DataServiceClientException)
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
								catch(InvalidTableFormatException exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									// A Table content error occurred, record it in the error log.
									this.LogError("Error: The Deliverable ID: " + deliverableItem.Key
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
									this.LogError("Content Error in Deliverable " + deliverableItem.Key +
										" Please review all content for this deliverable and correct it.");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Content Error in Deliverable " + deliverableItem.Key +
										" Please review all content for this deliverable and correct it.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									Console.WriteLine("Exception occurred: {0} - {1}", exc.HResult, exc.Message);
									}

								} // if(this.DeliverableHeading
							} // foreach (KeyValuePair<int, String>.....
						} //if(this.Deliverables)
Process_Reports:
					if(dictReports.Count == 0)
						goto Process_Meetings;

					if(this.Reports)
						{
						objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
						objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Reports_Heading_Text);
						objParagraph.Append(objRun);
						objBody.Append(objParagraph);
						string deliverableBookMark = "Report_";
						// Insert the individual Reports in the section
						foreach(KeyValuePair<int, string> reportItem in dictReports.OrderBy(key => key.Value))
							{
							if(this.Deliverable_Heading)
								{
								try
									{
									// Obtain the Deliverable info from SharePoint
									var dsDeliverables = datacontexSDDP.Deliverables
										.Expand(p => p.GlossaryAndAcronyms);

									var rsDeliverables =
										from dsDeliverable in dsDeliverables
										where dsDeliverable.Id == reportItem.Key
										select dsDeliverable;

									var recDeliverable = rsDeliverables.FirstOrDefault();
									Console.WriteLine("\t\t + {0} - {1}", recDeliverable.Id, recDeliverable.Title);

									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3, parBookMark: deliverableBookMark + recDeliverable.Id);
									objRun = oxmlDocument.Construct_RunText(parText2Write: recDeliverable.ISDHeading);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI + recDeliverable.Id,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									// Check if the user specified to include the Report Description
									if(this.Report_Description)
										{
										if(recDeliverable.ISDDescription != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 3,
												parHTML2Decode: recDeliverable.ISDDescription,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.ISDDescription != null)
										} //if(this.Deliverable_Description)

									// Check if the user specified to include the Report Inputs
									if(this.Report_Inputs)
										{
										if(recDeliverable.Inputs != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.Inputs,
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

									// Check if the user specified to include the Report Outputs
									if(this.Report_Outputs)
										{
										if(recDeliverable.Outputs != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableOutputs_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.Outputs,
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

									// Check if the user specified to include the Report DD's Obligations
									if(this.DDs_Report_Obligations)
										{
										if(recDeliverable.SPObligations != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.SPObligations,
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
									if(this.Clients_Report_Responsibilities)
										{
										if(recDeliverable.ClientResponsibilities != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.ClientResponsibilities,
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

									// Check if the user specified to include the Report Exclusions
									if(this.Report_Exclusions)
										{
										if(recDeliverable.Exclusions != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.Exclusions,
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
									if(this.Deliverable_Governance_Controls)
										{
										if(recDeliverable.GovernanceControls != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableGovernanceControls_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.GovernanceControls,
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
									if(recDeliverable.GlossaryAndAcronyms.Count > 0)
										{
										// Check if the user selected Acronyms and Glossy of Terms are requied
										if(this.Acronyms_Glossary_of_Terms_Section)
											{
											if(this.Acronyms || this.Glossary_of_Terms)
												{
												foreach(var entry in recDeliverable.GlossaryAndAcronyms)
													{
													if(this.DictionaryGlossaryAndAcronyms.ContainsKey(entry.Id) != true)
														DictionaryGlossaryAndAcronyms.Add(entry.Id, entry.Title);
													Console.WriteLine("\t\t\t + Term & Acronym added: {0} - {1}", entry.Id, entry.Title);
													}
												} // if(this.Acronyms || this.Glossary_of_Terms)
											} // if(this.Acronyms_Glossary_of_Terms_Section)
										} //if(recDeliverable.GlossaryAndAcronyms.Count > 0)
									} //Try
								catch(DataServiceClientException)
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
								catch(InvalidTableFormatException exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									// A Table content error occurred, record it in the error log.
									this.LogError("Error: The Deliverable ID: " + reportItem.Key
										+ " contains an error in one of its Enhance Rich Text columns. Please review the content (especially tables).");
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
									this.LogError("Content Error in Deliverable " + reportItem.Key +
										" Please review all content for this deliverable and correct it.");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Content Error in Deliverable " + reportItem.Key +
										" Please review all content for this deliverable and correct it.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									Console.WriteLine("Exception occurred: {0} - {1}", exc.HResult, exc.Message);
									}
								} // if(this.DeliverableHeading
							}
						} //if(this.Reports)
Process_Meetings:
					if(dictMeetings.Count == 0)
						goto Process_ServiceLevels;

					if(this.Meetings)
						{
						objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
						objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Meetings_Heading_Text);
						objParagraph.Append(objRun);
						objBody.Append(objParagraph);
						string deliverableBookMark = "Meeting_";
						// Insert the individual Meetings in the section
						foreach(KeyValuePair<int, string> meetingItem in dictMeetings.OrderBy(key => key.Value))
							{
							if(this.Meeting_Heading)
								{
								try
									{
									// Obtain the Meeting info from SharePoint
									var dsDeliverables = datacontexSDDP.Deliverables
										.Expand(p => p.GlossaryAndAcronyms);

									var rsDeliverables =
										from dsDeliverable in dsDeliverables
										where dsDeliverable.Id == meetingItem.Key
										select dsDeliverable;

									var recDeliverable = rsDeliverables.FirstOrDefault();
									Console.WriteLine("\t\t + {0} - {1}", recDeliverable.Id, recDeliverable.Title);

									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3, parBookMark: deliverableBookMark + recDeliverable.Id);
									objRun = oxmlDocument.Construct_RunText(parText2Write: recDeliverable.ISDHeading);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI + recDeliverable.Id,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									// Check if the user specified to include the Meeting Description
									if(this.Meeting_Description)
										{
										if(recDeliverable.ISDDescription != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 3,
												parHTML2Decode: recDeliverable.ISDDescription,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.ISDDescription != null)
										} //if(this.Meeting_Description)

									// Check if the user specified to include the Meeting Inputs
									if(this.Meeting_Inputs)
										{
										if(recDeliverable.Inputs != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.Inputs,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.Inputs != null)
										} //if(this.Meeting_Inputs)

									// Check if the user specified to include the Meeting Outputs
									if(this.Meeting_Outputs)
										{
										if(recDeliverable.Outputs != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableOutputs_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.Outputs,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.Outputs != null)
										} //if(this.Meeting_Outputs)

									// Check if the user specified to include the Meeting DD's Obligations
									if(this.DDs_Meeting_Obligations)
										{
										if(recDeliverable.SPObligations != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.SPObligations,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.SPObligations != null)
										} //if(this.DDS_Report_Oblidations)

									// Check if the user specified to include the Client Responsibilities
									if(this.Clients_Meeting_Responsibilities)
										{
										if(recDeliverable.ClientResponsibilities != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";
											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.ClientResponsibilities,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.Client_Responsibilities != null)
										} //if(this.Clients_Report_Responsibilities)

									// Check if the user specified to include the Deliverable Exclusions
									if(this.Deliverable_Exclusions)
										{
										if(recDeliverable.Exclusions != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.Exclusions,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.Exclusions != null)
										} //if(this.Report_Exclusions)

									// Check if the user specified to include the Governance Controls
									if(this.Meeting_Governance_Controls)
										{
										if(recDeliverable.GovernanceControls != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableGovernanceControls_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.GovernanceControls,
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
									if(recDeliverable.GlossaryAndAcronyms.Count > 0)
										{
										// Check if the user selected Acronyms and Glossy of Terms are requied
										if(this.Acronyms_Glossary_of_Terms_Section)
											{
											if(this.Acronyms || this.Glossary_of_Terms)
												{
												foreach(var entry in recDeliverable.GlossaryAndAcronyms)
													{
													if(this.DictionaryGlossaryAndAcronyms.ContainsKey(entry.Id) != true)
														DictionaryGlossaryAndAcronyms.Add(entry.Id, entry.Title);
													Console.WriteLine("\t\t\t + Term & Acronym added: {0} - {1}", entry.Id, entry.Title);
													}
												} // if(this.Acronyms || this.Glossary_of_Terms)
											} // if(this.Acronyms_Glossary_of_Terms_Section)
										} //if(recDeliverable.GlossaryAndAcronyms.Count > 0)
									} //Try
								catch(DataServiceClientException)
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
								catch(InvalidTableFormatException exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									// A Table content error occurred, record it in the error log.
									this.LogError("Error: The Deliverable ID: " + meetingItem.Key
										+ " contains an error in one of its Enhance Rich Text columns. Please review the content (especially tables).");
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
									this.LogError("Content Error in Deliverable " + meetingItem.Key +
										" Please review all content for this deliverable and correct it.");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Content Error in Deliverable " + meetingItem.Key +
										" Please review all content for this deliverable and correct it.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									Console.WriteLine("Exception occurred: {0} - {1}", exc.HResult, exc.Message);
									}
								} // if(this.MeetingHeading
							} // foreach.....
						} //if(this.Meetings)
					} //if(this.DRM_Section)

//-------------------------------------------------------
// Insert the Service Levels Section
Process_ServiceLevels:
				if(this.Service_Level_Section)
					{
					// Insert the Service If any are relevant
					if(dictSLAs.Count > 0)
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
							try
								{
								// Obtain the Deliverable Service Level from SharePoint
								var rsDeliverableServiceLevels =
									from rsDeliverableServiceLevel in datacontexSDDP.DeliverableServiceLevels
									where rsDeliverableServiceLevel.Id == servicelevelItem.Key
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
								if(this.Service_Level_Heading_in_Section)
									{
									// Insert the Service Level ISD Heading
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2, parBookMark: servicelevelBookMark + recServiceLevel.Id);
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

									if(this.Service_Level_Table_in_Section)
										{
										if(recServiceLevel.ISDDescription != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceLevelsURI +
												currentHyperlinkViewEditURI +
												recServiceLevel.Id;
											currentContentLayer = "None";
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 2,
												parHTML2Decode: recServiceLevel.ISDDescription,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}

										List<string> listErrorMessagesParameter = this.ErrorMessages;
										// Populate the Service Level Table
										objServiceLevelTable = CommonProcedures.BuildSLAtable(
											parServiceLevelID: recServiceLevel.Id,
											parWidthColumn1: Convert.ToUInt32(this.PageWith * 0.30),
											parWidthColumn2: Convert.ToUInt32(this.PageWith * 0.70),
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
										} //if(this.Service_Level_Commitments_Table)
									} //if(this.Service_Level_Heading_in_Section)
								} // try
							catch(DataServiceClientException)
								{
								// If the entry is not found - write an error in the document and record an error in the error log.
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

							catch(Exception exc)
								{
								Console.WriteLine("Exception occurred: {0} - {1}", exc.HResult, exc.Message);
								}
							} //foreach 
						} //(dictSLAs.Count >0)
					else
						{
						goto Process_Glossary_and_Acronyms;
						}
					} //if(this.Service_Level_Section)

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
			} // end of Generate method
		} // end of ISD_Document_DRM_Sections class
	}
