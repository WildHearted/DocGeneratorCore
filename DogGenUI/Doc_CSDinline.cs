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
	/// This class contains all the Client Service Description (CSD) with inline DRM (Deliverable Report Meeting).
	/// </summary>
	class CSD_Document_DRM_Inline:External_Document
		{
		private bool _drm_Description = false;
		public bool DRM_Description
			{
			get
				{return this._drm_Description;}
			set
				{this._drm_Description = value;}
			}
		private bool _drm_Inputs = false;
		public bool DRM_Inputs
			{
			get{return this._drm_Inputs;}
			set{this._drm_Inputs = value;}
			}
		private bool _drm_Outputs = false;
		public bool DRM_Outputs
			{
			get{return this._drm_Outputs;}
			set{this._drm_Outputs = value;}
			}
		private bool _dds_DRM_Obligations = false;
		public bool DDS_DRM_Obligations
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
			{get{return this._drm_Exclusions;}
			set{this._drm_Exclusions = value;}
			}
		private bool _drm_Governance_Controls = false;
		public bool DRM_Governance_Controls
			{
			get{return this._drm_Governance_Controls;}
			set{this._drm_Governance_Controls = value;}
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
			int? layer1upFeatureID = 0;
			int? layer2upFeatureID = 0;
			int? layer1upDeliverableID = 0;
			int? layer2upDeliverableID = 0;
			int tableCaptionCounter = 0;
			int imageCaptionCounter = 0;
			int hyperlinkCounter = 9;

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
				parTemplateURL: this.Template, parDocumentType: this.DocumentType))
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
									// Obtain the Service Portfolio info from SharePoint
									ServicePortfolio objPortfolio = new ServicePortfolio();
									objPortfolio.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: node.NodeID);

									Console.WriteLine("\t\t + {0} - {1}", objPortfolio.ID, objPortfolio.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
									objRun = oxmlDocument.Construct_RunText(parText2Write: objPortfolio.CSDheading, parIsNewSection: true);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
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
										if(objPortfolio.CSDdescription != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServicePortfoliosURI +
												currentHyperlinkViewEditURI + objPortfolio.ID;
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 1,
												parHTML2Decode: objPortfolio.CSDdescription,
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
									Console.WriteLine("\n\nException occurred: {0}", exc.Message);
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
									Console.WriteLine("\n\nException occurred: {0}", exc.Message);
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
									// Obtain the Service Family info from SharePoint
									ServiceFamily objFamily = new ServiceFamily();
									objFamily.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: node.NodeID);

									Console.WriteLine("\t\t + {0} - {1}", objFamily.ID, objFamily.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: objFamily.CSDheading,
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
											currentHyperlinkViewEditURI + objFamily.ID,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Family Description
									if(this.Service_Family_Description)
										{
										if(objFamily.CSDdescription != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServicePortfoliosURI +
												currentHyperlinkViewEditURI +
												objFamily.ID;
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 2,
												parHTML2Decode: objFamily.CSDdescription,
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
									Console.WriteLine("\n\nException occurred: {0}", exc.Message);
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
									Console.WriteLine("\n\nException occurred: {0}", exc.Message);
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
									// Obtain the Service Product info from SharePoint
									ServiceProduct objProduct = new ServiceProduct();
									objProduct.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: node.NodeID);

									Console.WriteLine("\t\t + {0} - {1}", objProduct.ID, objProduct.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: objProduct.CSDheading,
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
											currentHyperlinkViewEditURI + objProduct.ID,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Product Description
									if(this.Service_Product_Description)
										{
										if(objProduct.CSDdescription != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceProductsURI +
												currentHyperlinkViewEditURI +
												objProduct.ID;

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 3,
												parHTML2Decode: objProduct.CSDdescription,
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
									Console.WriteLine("\n\nException occurred: {0}", exc.Message);
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
									Console.WriteLine("\n\nException occurred: {0}", exc.Message);
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
									Console.WriteLine("\n\nException occurred: {0}", exc.Message);
									}
								} //if(this.Service_Product_Heading)
							break;
							}
					//------------------------------------------
					case enumNodeTypes.FEA:  // Service Feature
							{
							if(this.Service_Feature_Heading)
								{
								try
									{
									// Obtain the Feature info from SharePoint
									ServiceFeature objServiceFeature = new ServiceFeature();
									objServiceFeature.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: node.NodeID, parGetLayer1up: true);

									// Insert the Service Feature CSD Heading...
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(parText2Write: objServiceFeature.CSDheading);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									//Check if the Feature Layer0up has Content Layers and Content Predecessors
									Console.WriteLine("\t\t + Service Feature Layer 0..: {0} - {1}", objServiceFeature.ID, objServiceFeature.Title);
									if(objServiceFeature.ContentPredecessorFeatureID == null)
										{
										layer1upFeatureID = null;
										layer2upFeatureID = null;
										}
									else
										{
										layer1upFeatureID = objServiceFeature.ContentPredecessorFeatureID;
										Console.WriteLine("\t\t + Service Feature Layer 1up: {0} - {1}",
											objServiceFeature.Layer1up.ID, objServiceFeature.Layer1up.Title);
										if(objServiceFeature.Layer1up.ContentPredecessorFeatureID == null)
											{
											layer2upFeatureID = null;
											}
										else
											{
											Console.WriteLine("\t\t + Service Feature Layer 2up: {0} - {1}",
												objServiceFeature.Layer1up.Layer1up.ID, objServiceFeature.Layer1up.Layer1up.Title);
											layer2upFeatureID = objServiceFeature.Layer1up.ContentPredecessorFeatureID;
											}
										}

									// Check if the user specified to include the Service Feature Description
									if(this.Service_Feature_Description)
										{
										// Insert Layer 2up if present and not null
										if(layer2upFeatureID != null)
											{
											if(objServiceFeature.Layer1up.Layer1up.CSDdescription != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_ServiceFeaturesURI +
														currentHyperlinkViewEditURI +
														objServiceFeature.Layer1up.Layer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objServiceFeature.Layer1up.Layer1up.CSDdescription,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												} //if(objServiceFeature.Layer1up.Layer1up.CSDdescription != null)
											} // if(layer2upFeatureID != null)

										// Insert Layer 1up if present and not null
										if(layer1upFeatureID != null)
											{
											if(objServiceFeature.Layer1up.CSDdescription != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_ServiceFeaturesURI +
														currentHyperlinkViewEditURI +
														objServiceFeature.Layer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objServiceFeature.Layer1up.CSDdescription,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											} //// if(layer2upFeatureID != null)

										// Insert Layer 0up if not null
										if(objServiceFeature.CSDdescription != null)
											{
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceFeaturesURI +
													currentHyperlinkViewEditURI +
													objServiceFeature.ID;
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
												parHTML2Decode: objServiceFeature.CSDdescription,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										} //if(this.Service_Feature_Description)
									drmHeading = false;
									}
								catch(DataServiceClientException)
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
									this.LogError("Error: The Service Feature ID " + node.NodeID
										+ " doesn't exist in SharePoint and couldn't be retrieved.");
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Error: Service Feature " + node.NodeID + " is missing.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									}
								catch(InvalidTableFormatException exc)
									{
									Console.WriteLine("\n\nException occurred: {0}", exc.Message);
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
									Console.WriteLine("\n\nException occurred: {0} - {1}", exc.HResult, exc.Message);
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
							try
								{
								// Obtain the Deliverable info from SharePoint
								Deliverable objDeliverable = new Deliverable();
								objDeliverable.PopulateObject(parDatacontexSDDP: datacontexSDDP, parID: node.NodeID, parGetLayer1up: true);

								// Insert the Deliverable CSD Heading
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
								objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.CSDheading);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);

								//Check if the Deliverable Layer0up has Content Layers and Content Predecessors
								Console.WriteLine("\t\t + Deliverable Layer 0..: {0} - {1}", objDeliverable.ID, objDeliverable.Title);
								if(objDeliverable.ContentPredecessorDeliverableID == null)
									{
									layer1upDeliverableID = null;
									layer2upDeliverableID = null;
									}
								else
									{
									Console.WriteLine("\t\t + Deliverable Layer 1up: {0} - {1}",
											objDeliverable.Layer1up.ID, objDeliverable.Layer1up.Title);
									layer1upDeliverableID = objDeliverable.ContentPredecessorDeliverableID;
									if(objDeliverable.Layer1up.ContentPredecessorDeliverableID == null)
										{
										layer2upFeatureID = null;
										}
									else
										{
										Console.WriteLine("\t\t + Deliverable Layer 2up: {0} - {1}",
											objDeliverable.Layer1up.Layer1up.ID, objDeliverable.Layer1up.Layer1up.Title);
										layer2upDeliverableID = objDeliverable.Layer1up.ContentPredecessorDeliverableID;
										}
									}
								//---------------------------------------------------------------
								// Check if the user specified to include the Deliverable Summary
								if(this.DRM_Description)
									{
									// Insert Layer 2up if present and not null
									if(layer2upDeliverableID != null)
										{
										if(objDeliverable.Layer1up.Layer1up.CSDdescription != null)
											{
											// Check for Colour coding Layers and add if necessary
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI +
													objDeliverable.Layer1up.Layer1up.ID;
												}
											else
												currentListURI = "";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 6,
												parHTML2Decode: objDeliverable.Layer1up.Layer1up.CSDdescription,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(objDeliverable.Layer1up.Layer1up.CSDdescription != null)
										} // if(layer2upDeliverableID != null)

									// Insert Layer 1up if present and not null
									if(layer1upDeliverableID != null)
										{
										if(objDeliverable.Layer1up.CSDdescription != null)
											{
											// Check for Colour coding Layers and add if necessary
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer2";
											else
												currentContentLayer = "None";

											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_DeliverablesURI +
													currentHyperlinkViewEditURI +
													objDeliverable.Layer1up.ID;
												}
											else
												currentListURI = "";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 6,
												parHTML2Decode: objDeliverable.Layer1up.CSDdescription,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}// if(objDeliverable.Layer1up.Layer1up.CSDdescription != null)
										} // if(layer2upDeliverableID != null)

									// Insert Layer 0up if present and not null
									if(objDeliverable.CSDdescription != null)
										{
										// Check for Colour coding Layers and add if necessary
										if(this.ColorCodingLayer1)
											currentContentLayer = "Layer3";
										else
											currentContentLayer = "None";

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

										objHTMLdecoder.DecodeHTML(
											parMainDocumentPart: ref objMainDocumentPart,
											parDocumentLevel: 6,
											parHTML2Decode: objDeliverable.CSDdescription,
											parContentLayer: currentContentLayer,
											parTableCaptionCounter: ref tableCaptionCounter,
											parImageCaptionCounter: ref imageCaptionCounter,
											parHyperlinkID: ref hyperlinkCounter,
											parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
											parHyperlinkURL: currentListURI,
											parPageHeightTwips: this.PageHight,
											parPageWidthTwips: this.PageWith);
										} // if(objDeliverable.CSDdescription != null)

									// Insert the hyperlink to the bookmark of the Deliverable's rlevant position in the DRM Section.
									objParagraph = oxmlDocument.Construct_BookmarkHyperlink(
									parBodyTextLevel: 6,
									parBookmarkValue: "Deliverable_" + objDeliverable.ID);
									objBody.Append(objParagraph);
									} // if (this.DRM_Description)

								//--------------------------------------------------------------
								// Check if the user specified to include the Deliverable Inputs
								if(this.DRM_Inputs)
									{
									if(objDeliverable.Inputs != null
									|| (layer1upDeliverableID != null && objDeliverable.Layer1up.Inputs != null)
									|| (layer2upDeliverableID != null && objDeliverable.Layer1up.Layer1up.Inputs != null))
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer 2up if present and not null
										if(layer2upDeliverableID != null)
											{
											if(objDeliverable.Layer1up.Layer1up.Inputs != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.Layer1up.Layer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: objDeliverable.Layer1up.Layer1up.Inputs,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												} //if(recDeliverable.Layer1up.Layer1up.Inputs != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer 1up if present and not null
										if(layer1upDeliverableID != null)
											{
											if(objDeliverable.Layer1up.Inputs != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.Layer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: objDeliverable.Layer1up.Inputs,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												}
											} // if(layer2upDeliverableID != null)

										// Insert Layer0up if not null
										if(objDeliverable.Inputs != null)
											{
											// Check if a hyperlink must be inserted
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
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 7,
												parHTML2Decode: objDeliverable.Inputs,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.Inputs != null)
										} //if(objDeliverable.Inputs  &&...)
									} //if(this.DRM_Inputs)

								//----------------------------------------------------------------
								// Check if the user specified to include the Deliverable Outputs
								if(this.DRM_Outputs)
									{
									if(objDeliverable.Outputs != null
									|| (layer1upDeliverableID != null && objDeliverable.Layer1up.Outputs != null)
									|| (layer2upDeliverableID != null && objDeliverable.Layer1up.Layer1up.Outputs != null))
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableOutputs_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer 2up if present and not null
										if(layer2upDeliverableID != null)
											{
											if(objDeliverable.Layer1up.Layer1up.Outputs != null)
												{
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.Layer1up.Layer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: objDeliverable.Layer1up.Layer1up.Outputs,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												} //if(recDeliverable.Layer1up.Layer1up.Outputs != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer 1up if present and not null
										if(layer1upDeliverableID != null)
											{
											if(objDeliverable.Layer1up.Outputs != null)
												{
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.Layer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: objDeliverable.Layer1up.Outputs,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												} // if(objDeliverable.Layer1up.Outputs != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer0up if not null
										if(objDeliverable.Outputs != null)
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
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 7,
												parHTML2Decode: objDeliverable.Outputs,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(objDeliverable.Outputs != null)
										} //if(objDeliverables.Outputs !== null &&)
									} //if(this.DRM_Outputs)

								//-----------------------------------------------------------------------
								// Check if the user specified to include the Deliverable DD's Obligations
								if(this.DDS_DRM_Obligations)
									{
									if(objDeliverable.DDobligations != null
									|| (layer1upDeliverableID != null && objDeliverable.Layer1up.DDobligations != null)
									|| (layer2upDeliverableID != null && objDeliverable.Layer1up.Layer1up.DDobligations != null))
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer 2up if present and not null
										if(layer2upDeliverableID != null)
											{
											if(objDeliverable.Layer1up.Layer1up.DDobligations != null)
												{
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.Layer1up.Layer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: objDeliverable.Layer1up.Layer1up.DDobligations,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												} //if(objDeliverable.Layer1up.Layer1up.DDobligations != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer 1up if present and not null
										if(layer1upDeliverableID != null)
											{
											if(objDeliverable.Layer1up.DDobligations != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.Layer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: objDeliverable.Layer1up.DDobligations,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												} // if(objDeliverable.Layer1up.DDobligations != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer0up if not null
										if(objDeliverable.DDobligations != null)
											{
											// Check if a hyperlink must be inserted
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
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 7,
												parHTML2Decode: objDeliverable.DDobligations,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(objDeliverable.DDobligations != null)
										} //if(objDeliverable.DDoblidations != null &&)
									} //if(this.DDs_DRM_Obligations)

								//-------------------------------------------------------------------
								// Check if the user specified to include the Client Responsibilities
								if(this.Clients_DRM_Responsibilities)
									{
									if(objDeliverable.ClientResponsibilities != null
									|| (layer1upDeliverableID != null && objDeliverable.Layer1up.ClientResponsibilities != null)
									|| (layer2upDeliverableID != null && objDeliverable.Layer1up.Layer1up.ClientResponsibilities != null))
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer 2up if present and not null
										if(layer2upDeliverableID != null)
											{
											if(objDeliverable.Layer1up.Layer1up.ClientResponsibilities != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.Layer1up.Layer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: objDeliverable.Layer1up.Layer1up.ClientResponsibilities,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												} //if(recDeliverable.Layer1up.Layer1up.ClientResponsibilities != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer 1up if present and not null
										if(layer1upDeliverableID != null)
											{
											if(objDeliverable.Layer1up.ClientResponsibilities != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.Layer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: objDeliverable.Layer1up.ClientResponsibilities,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												} // if(objDeliverable.Layer1up.ClientResponsibilities != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer0up if not null
										if(objDeliverable.ClientResponsibilities != null)
											{
											// Check if a hyperlink must be inserted
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
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 7,
												parHTML2Decode: objDeliverable.ClientResponsibilities,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(objDeliverable.ClientResponsibilities != null)
										} // if(objDeliverable.ClientResponsibilities != null &&)
									} //if(this.Clients_DRM_Responsibilities)

								//------------------------------------------------------------------
								// Check if the user specified to include the Deliverable Exclusions
								if(this.DRM_Exclusions)
									{
									if(objDeliverable.Exclusions != null
									|| (layer1upDeliverableID != null && objDeliverable.Layer1up.Exclusions != null)
									|| (layer2upDeliverableID != null && objDeliverable.Layer1up.Layer1up.Exclusions != null))
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer 2up if present and not null
										if(layer2upDeliverableID != null)
											{
											if(objDeliverable.Layer1up.Layer1up.Exclusions != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.Layer1up.Layer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: objDeliverable.Layer1up.Layer1up.Exclusions,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												} //if(recDeliverable.Layer1up.Layer1up.Exclusions != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer 1up if present and not null
										if(layer1upDeliverableID != null)
											{
											if(objDeliverable.Layer1up.Exclusions != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.Layer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: objDeliverable.Layer1up.Exclusions,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												} // if(objDeliverable.Layer1up.Exclusions != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer0up if not null
										if(objDeliverable.ClientResponsibilities != null)
											{
											// Check if a hyperlink must be inserted
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
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 7,
												parHTML2Decode: objDeliverable.Exclusions,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(objDeliverable.Exclusions != null)
										} // if(objDeliverable.Exclusions != null &&)	
									} //if(this.DRMe_Exclusions)

								//---------------------------------------------------------------
								// Check if the user specified to include the Governance Controls
								if(this.DRM_Governance_Controls)
									{
									if(objDeliverable.GovernanceControls != null
									|| (layer1upDeliverableID != null && objDeliverable.Layer1up.GovernanceControls != null)
									|| (layer2upDeliverableID != null && objDeliverable.Layer1up.Layer1up.GovernanceControls != null))
										{
										// Insert the Heading
										objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: Properties.AppResources.Document_DeliverableGovernanceControls_Heading_Text);
										objParagraph.Append(objRun);
										objBody.Append(objParagraph);

										// Insert Layer 2up if present and not null
										if(layer2upDeliverableID != null)
											{
											if(objDeliverable.Layer1up.Layer1up.GovernanceControls != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.Layer1up.Layer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer1";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: objDeliverable.Layer1up.Layer1up.GovernanceControls,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												} //if(objDeliverable.Layer1up.Layer1up.GovernanceControls != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer 1up if present and not null
										if(layer1upDeliverableID != null)
											{
											if(objDeliverable.Layer1up.GovernanceControls != null)
												{
												// Check if a hyperlink must be inserted
												if(documentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													currentListURI = Properties.AppResources.SharePointURL +
														Properties.AppResources.List_DeliverablesURI +
														currentHyperlinkViewEditURI +
														objDeliverable.Layer1up.ID;
													}
												else
													currentListURI = "";

												if(this.ColorCodingLayer1)
													currentContentLayer = "Layer2";
												else
													currentContentLayer = "None";

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 7,
													parHTML2Decode: objDeliverable.Layer1up.GovernanceControls,
													parContentLayer: currentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
													parHyperlinkURL: currentListURI,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
												} // if(objDeliverable.Layer1up.GovernanceControls != null)
											} // if(layer2upDeliverableID != null)

										// Insert Layer0up if not null
										if(objDeliverable.GovernanceControls != null)
											{
											// Check if a hyperlink must be inserted
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
												currentContentLayer = "Layer3";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 7,
												parHTML2Decode: objDeliverable.GovernanceControls,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(objDeliverable.GovernanceControls != null)
										} // if(objDeliverable.GovernanceControls != null &&)	
									} //if(this.DRM_GovernanceControls)

								//---------------------------------------------------
								// Check if there are any Glossary Terms or Acronyms associated with the Deliverable(s).
								if(this.Acronyms_Glossary_of_Terms_Section)
									{
									// if there are GlossaryAndAcronyms to add from layer0up
									if(objDeliverable.GlossaryAndAcronyms.Count > 0)
										{
										foreach(var entry in objDeliverable.GlossaryAndAcronyms)
											{
											if(this.DictionaryGlossaryAndAcronyms.ContainsKey(entry.Key) != true)
												DictionaryGlossaryAndAcronyms.Add(entry.Key, entry.Value);
											}
										}
									// if there are GlossaryAndAcronyms to add from layer1up
									if(layer1upDeliverableID != null && objDeliverable.Layer1up.GlossaryAndAcronyms.Count > 0)
										{
										foreach(var entry in objDeliverable.Layer1up.GlossaryAndAcronyms)
											{
											if(this.DictionaryGlossaryAndAcronyms.ContainsKey(entry.Key) != true)
												DictionaryGlossaryAndAcronyms.Add(entry.Key, entry.Value);
											}
										}
									// if there are GlossaryAndAcronyms to add from layer2up
									if(layer2upDeliverableID != null && objDeliverable.Layer1up.Layer1up.GlossaryAndAcronyms.Count > 0)
										{
										foreach(var entry in objDeliverable.Layer1up.Layer1up.GlossaryAndAcronyms)
											{
											if(this.DictionaryGlossaryAndAcronyms.ContainsKey(entry.Key) != true)
												DictionaryGlossaryAndAcronyms.Add(entry.Key, entry.Value);
											}
										}
									} // if(this.Acronyms_Glossary_of_Terms_Section)	

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
								Console.WriteLine("\n\nException occurred: {0}", exc.Message);
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
								Console.WriteLine("\n\nException occurred: {0} - {1}", exc.HResult, exc.Message);
								}
							break;
							}
					
					case enumNodeTypes.FSL:  // Service Level associated with Deliverable pertaining to Service Feature
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
										List<ServiceLevelTarget> listServiceLevelThresholds = new List<ServiceLevelTarget>();
										foreach(var recSLthreshold in rsServiceLevelThresholds)
											{
											ServiceLevelTarget objSLthreshold = new ServiceLevelTarget();
											objSLthreshold.ID = recSLthreshold.Id;
											objSLthreshold.Title = recSLthreshold.Title;
											listServiceLevelThresholds.Add(objSLthreshold);
											Console.WriteLine("\t\t\t + Threshold: {0} - {1}", recSLthreshold.Id, recSLthreshold.Title);
											}

										// Obtain the Service Level Targets from SharePoint
										var rsServiceLevelTargets =
											from dsSLTargets in datacontexSDDP.ServiceLevelTargets
											where dsSLTargets.Service_LevelId == recServiceLevel.Id 
												&& dsSLTargets.ThresholdOrTargetValue == "Target"
											orderby dsSLTargets.Title
											select new
												{
												dsSLTargets.Id,
												dsSLTargets.Title
												};
										// load the SL Targets into a list - apckaging it in order to send it as a parameter later on.
										List<ServiceLevelTarget> listServiceLevelTargets = new List<ServiceLevelTarget>();
										foreach(var recSLtarget in rsServiceLevelTargets)
											{
											ServiceLevelTarget objSLtarget = new ServiceLevelTarget();
											objSLtarget.ID = recSLtarget.Id;
											objSLtarget.Title = recSLtarget.Title;
											listServiceLevelTargets.Add(objSLtarget);
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
										Console.WriteLine("\n\nException occurred: {0} - {1}", exc.HResult, exc.Message);
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
				if(this.Acronyms_Glossary_of_Terms_Section && this.DictionaryGlossaryAndAcronyms.Count == 0)
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
							parDictionaryGlossaryAcronym: this.DictionaryGlossaryAndAcronyms,
							parWidthColumn1: Convert.ToUInt32(this.PageWith * 0.3),
							parWidthColumn2: Convert.ToUInt32(this.PageWith * 0.2),
							parWidthColumn3: Convert.ToUInt32(this.PageWith * 0.5),
							parErrorMessages: ref listErrors);
						objBody.Append(tableGlossaryAcronym);
						}     //if(this.TermAndAcronymList.Count > 0)
					} // if (this.Acronyms)



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

		} // end of CSD_inline DRM class
	}
