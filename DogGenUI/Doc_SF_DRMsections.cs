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
		/// This class represent the Services Framework Document with sperate DRM (Deliverable Report Meeting) sections
		/// It inherits from the Internal DRM Sections Class.
		/// </summary>
	class Services_Framework_Document_DRM_Sections:Internal_DRM_Sections
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
					parOptions.Sort();
					foreach(int option in parOptions)
						{
						//Console.WriteLine(option);
						switch(option)
							{
						case 236:
							this.Introductory_Section = true;
							break;
						case 237:
							this.Introduction = true;
							break;
						case 238:
							this.Executive_Summary = true;
							break;
						case 239:
							this.Service_Portfolio_Section = true;
							break;
						case 240:
							this.Service_Portfolio_Description = true;
							break;
						case 241:
							this.Service_Family_Heading = true;
							break;
						case 242:
							this.Service_Family_Description = true;
							break;
						case 243:
							this.Service_Product_Heading = true;
							break;
						case 244:
							this.Service_Product_Description = true;
							break;
						case 245:
							this.Service_Product_Key_Client_Benefits = true;
							break;
						case 246:
							this.Service_Product_KeyDD_Benefits = true;
							break;
						case 247:
							this.Service_Element_Heading = true;
							break;
						case 248:
							this.Service_Element_Description = true;
							break;
						case 249:
							this.Service_Element_Objectives = true;
							break;
						case 250:
							this.Service_Element_Key_Client_Benefits = true;
							break;
						case 251:
							this.Service_Element_Key_Client_Advantages = true;
							break;
						case 252:
							this.Service_Element_Key_DD_Benefits = true;
							break;
						case 253:
							this.Service_Element_Critical_Success_Factors = true;
							break;
						case 254:
							this.Service_Element_Key_Performance_Indicators = true;
							break;
						case 255:
							this.Service_Element_High_Level_Process = true;
							break;
						case 256:
							this.Deliverables_Reports_Meetings = true;
							break;
						case 257:
							this.DRM_Heading = true;
							break;
						case 258:
							this.DRM_Summary = true;
							break;
						case 259:
							this.Service_Levels = true;
							break;
						case 260:
							this.Service_Level_Heading = true;
							break;
						case 261:
							this.Service_Level_Commitments_Table = true;
							break;
						case 262:
							this.Activities = true;
							break;
						case 263:
							this.Activity_Heading = true;
							break;
						case 264:
							this.Activity_Description_Table = true;
							break;
						case 267:
							this.DRM_Section = true;
							break;
						case 268:
							this.Deliverables = true;
							break;
						case 269:
							this.Deliverable_Heading = true;
							break;
						case 333:
							this.Deliverable_Description = true;
							break;
						case 270:
							this.Deliverable_Inputs = true;
							break;
						case 271:
							this.Deliverable_Outputs = true;
							break;
						case 272:
							this.DDs_Deliverable_Obligations = true;
							break;
						case 273:
							this.Clients_Deliverable_Responsibilities = true;
							break;
						case 274:
							this.Deliverable_Exclusions = true;
							break;
						case 275:
							this.Deliverable_Governance_Controls = true;
							break;
						case 276:
							this.Reports = true;
							break;
						case 277:
							this.Report_Heading = true;
							break;
						case 278:
							this.Report_Description = true;
							break;
						case 279:
							this.DDs_Report_Obligations = true;
							break;
						case 280:
							this.Clients_Report_Responsibilities = true;
							break;
						case 281:
							this.Report_Exclusions = true;
							break;
						case 282:
							this.Report_Governance_Controls = true;
							break;
						case 283:
							this.Meetings = true;
							break;
						case 284:
							this.Meeting_Heading = true;
							break;
						case 285:
							this.Meeting_Description = true;
							break;
						case 286:
							this.DDs_Meeting_Obligations = true;
							break;
						case 287:
							this.Clients_Meeting_Responsibilities = true;
							break;
						case 288:
							this.Meeting_Exclusions = true;
							break;
						case 289:
							this.Meeting_Governance_Controls = true;
							break;
						case 335:
							this.Service_Level_Section = true;
							break;
						case 336:
							this.Service_Level_Heading_in_Section = true;
							break;
						case 337:
							this.Service_Level_Table_in_Section = true;
							break;
						case 290:
							this.Acronyms_Glossary_of_Terms_Section = true;
							break;
						case 291:
							this.Acronyms = true;
							break;
						case 292:
							this.Glossary_of_Terms = true;
							break;
						case 334:
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
			int? parRequestingUserID)
			{
			Console.WriteLine("\t Begin to generate {0}", this.DocumentType);
			this.UnhandledError = false;
			DateTime timeStarted = DateTime.Now;
			string strErrorText = "";
			string strHyperlinkImageRelationshipID = "";
			string strDocumentCollection_HyperlinkURL = "";
			string strCurrentListURI = "";
			string strCurrentHyperlinkViewEditURI = "";
			string strCurrentContentLayer = "None";
			bool bDRMHeadingInserted = false;
			Table objActivityTable = new Table();
			Table objServiceLevelTable = new Table();

			try
				{
				if(this.HyperlinkEdit)
					{
					strDocumentCollection_HyperlinkURL = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
						Properties.AppResources.List_DocumentCollectionLibraryURI +
						Properties.AppResources.EditFormURI + this.DocumentCollectionID;
					strCurrentHyperlinkViewEditURI = Properties.AppResources.EditFormURI;
					}
				if(this.HyperlinkView)
					{
					strDocumentCollection_HyperlinkURL = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
						Properties.AppResources.List_DocumentCollectionLibraryURI +
						Properties.AppResources.DisplayFormURI + this.DocumentCollectionID;
					strCurrentHyperlinkViewEditURI = Properties.AppResources.DisplayFormURI;
					}
				int tableCaptionCounter = 0;
				int imageCaptionCounter = 0;
				int iPictureNo = 49;
				int hyperlinkCounter = 9;

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

	
				// Open the MS Word document in Edit mode
				this.DocumentStatus = enumDocumentStatusses.Creating;
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

				// Check whether Hyperlinks need to be included
				if(this.HyperlinkEdit || this.HyperlinkView)
					{
					//Insert and embed the hyperlink image in the document and keep the Image's Relationship ID in a variable for repeated use
					strHyperlinkImageRelationshipID = oxmlDocument.InsertHyperlinkImage(parMainDocumentPart: ref objMainDocumentPart,
						parDataSet: ref parDataSet);
					}

				Dictionary<int, string> dictDeliverables = new Dictionary<int, string>();
				Dictionary<int, string> dictReports = new Dictionary<int, string>();
				Dictionary<int, string> dictMeetings = new Dictionary<int, string>();
				Dictionary<int, string> dictSLAs = new Dictionary<int, string>();
				ServicePortfolio objPortfolio = new ServicePortfolio();
				ServiceFamily objFamily = new ServiceFamily();
				ServiceProduct objProduct = new ServiceProduct();
				ServiceElement objElement = new ServiceElement();
				ElementDeliverable objElementDeliverable = new ElementDeliverable();
				Deliverable objDeliverable = new Deliverable();
				DeliverableActivity objDeliverableActivity = new DeliverableActivity();
				DeliverableServiceLevel objDeliverableServiceLevel = new DeliverableServiceLevel();
				ServiceLevel objServiceLevel = new ServiceLevel();
				Activity objActivity = new Activity();


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
					if(strDocumentCollection_HyperlinkURL != "")
						{
						hyperlinkCounter += 1;
						Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
							parMainDocumentPart: ref objMainDocumentPart,
							parImageRelationshipId: strHyperlinkImageRelationshipID,
							parClickLinkURL: strDocumentCollection_HyperlinkURL,
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
								parMainDocumentPart: ref objMainDocumentPart,
								parDocumentLevel: 2,
								parHTML2Decode: this.IntroductionRichText,
								parTableCaptionCounter: ref tableCaptionCounter,
								parImageCaptionCounter: ref imageCaptionCounter,
								parPictureNo: ref iPictureNo,
								parHyperlinkID: ref hyperlinkCounter,
								parPageHeightTwips: this.PageHight,
								parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
							}
						catch(InvalidContentFormatException exc)
							{
							Console.WriteLine("\n\nException occurred: {0}", exc.Message);
							// A Table content error occurred, record it in the error log.
							this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
								+ " contains an error in Introduction Enhance Rich Text. "
								+ exc.Message);
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "A content error occurred at this position and valid content could "
								+ "not be interpreted and inserted here. Please review the content in the SharePoint "
								+ "system and correct it." + exc.Message,
								parIsNewSection: false,
								parIsError: true);
							if(strDocumentCollection_HyperlinkURL != "")
								{
								hyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
									parMainDocumentPart: ref objMainDocumentPart,
									parImageRelationshipId: strHyperlinkImageRelationshipID,
									parHyperlinkID: hyperlinkCounter,
									parClickLinkURL: strDocumentCollection_HyperlinkURL);
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
					if(strDocumentCollection_HyperlinkURL != "")
						{
						hyperlinkCounter += 1;
						Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
							parMainDocumentPart: ref objMainDocumentPart,
							parImageRelationshipId: strHyperlinkImageRelationshipID,
							parClickLinkURL: strDocumentCollection_HyperlinkURL,
							parHyperlinkID: hyperlinkCounter);
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
								parTableCaptionCounter: ref tableCaptionCounter,
								parImageCaptionCounter: ref imageCaptionCounter,
								parPictureNo: ref iPictureNo,
								parHyperlinkID: ref hyperlinkCounter,
								parPageHeightTwips: this.PageHight,
								parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
							}
						catch(InvalidContentFormatException exc)
							{
							Console.WriteLine("\n\nException occurred: {0}", exc.Message);
							// A Table content error occurred, record it in the error log.
							this.LogError("Error: The Document Collection ID: " + this.DocumentCollectionID
								+ " contains an error in Excutive Summary Enhance Rich Text. "
								+ exc.Message);
							objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: "A content error occurred at this position and valid content could "
								+ "not be interpreted and inserted here. Please review the content in the SharePoint "
								+ "system and correct it." + exc.Message,
								parIsNewSection: false,
								parIsError: true);
							if(strDocumentCollection_HyperlinkURL != "")
								{
								hyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
									parMainDocumentPart: ref objMainDocumentPart,
									parImageRelationshipId: strHyperlinkImageRelationshipID,
									parHyperlinkID: hyperlinkCounter,
									parClickLinkURL: strDocumentCollection_HyperlinkURL);
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
					Console.Write("\nNode SEQ: {0} Level:{1} Type:{2} ID:{3}", node.Sequence, node.Level, node.NodeType, node.NodeID);

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
										parText2Write: objPortfolio.ISDheading,
										parIsNewSection: true);
									// Check if a hyperlink must be inserted
									if(strDocumentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
											Properties.AppResources.List_ServicePortfoliosURI +
											strCurrentHyperlinkViewEditURI + objPortfolio.ID;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: strHyperlinkImageRelationshipID,
											parClickLinkURL: strCurrentListURI,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Porfolio Description
									if(this.Service_Portfolio_Description)
										{
										if(objPortfolio.ISDdescription != null)
											{
											try
												{
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
													Properties.AppResources.List_ServicePortfoliosURI +
													strCurrentHyperlinkViewEditURI + objPortfolio.ID;

												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 1,
													parHTML2Decode: objPortfolio.ISDdescription,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Portfolio ID: " + objPortfolio.ID
													+ " contains an error in ISD Description Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Service Portfolio: " + node.NodeID + "Introduction Content"
												+ " Please review all content for this deliverable and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText +
													" Please review all content for this Service Portfolio and correct it.",
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\n\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												} 
											} 
										} 
									} //if(parDataSet.dsPortfolios.TryGetValue(
                                        else
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
									this.LogError("Error: The Service Portfolio ID " + node.NodeID +
										" doesn't exist in SharePoint and couldn't be retrieved.");
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Error: Service Portfolio " + node.NodeID + " does not exist in SharePoint.",
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
									Console.Write("\t + {0} - {1}", objFamily.ID, objFamily.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: objFamily.ISDheading,
										parIsNewSection: false);
									// Check if a hyperlink must be inserted
									if(strDocumentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
											Properties.AppResources.List_ServiceFamiliesURI +
											strCurrentHyperlinkViewEditURI + objFamily.ID;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: strHyperlinkImageRelationshipID,
											parClickLinkURL: strCurrentListURI,
											parHyperlinkID: hyperlinkCounter);
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
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Family ID: " + objFamily.ID
													+ " contains an error in ISD Description Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Service Family: " + node.NodeID + "ISD Description Content"
												+ " Please review all content for this deliverable and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText +
													" Please review all ISD content for this Service Family and correct it.",
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\n\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											}
										}
									} // if(parDataSet.dsFamilies.TryGetValue(...
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
								if(parDataSet.dsProducts.TryGetValue(
									key: node.NodeID,
									value: out objProduct))
									{
									Console.Write("\t + {0} - {1}", objProduct.ID, objProduct.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: objProduct.ISDheading,
										parIsNewSection: false);
									// Check if a hyperlink must be inserted
									if(strDocumentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: strHyperlinkImageRelationshipID,
											parClickLinkURL: parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
											Properties.AppResources.List_ServiceProductsURI +
											strCurrentHyperlinkViewEditURI + objProduct.ID,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Product Description
									if(this.Service_Product_Description)
										{
										if(objProduct.ISDdescription != null)
											{
											strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_ServiceProductsURI +
												strCurrentHyperlinkViewEditURI +
												objProduct.ID;
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 3,
													parHTML2Decode: objProduct.ISDdescription,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Product ID: " + objProduct.ID
													+ " contains an error in ISD Description Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 3);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Service Product: " + node.NodeID + " ISD Description Content"
												+ " Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 3);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText +
													" Please review the content and correct it.",
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											}
										}
									if(this.Service_Product_KeyDD_Benefits)
										{
										if(objProduct.KeyDDbenefits != null)
											{
											strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_ServiceProductsURI +
												strCurrentHyperlinkViewEditURI +
												objProduct.ID;
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Product_KeyDD_Benefits,
												parIsNewSection: false);
											// Check if a hyperlink must be inserted
											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: strHyperlinkImageRelationshipID,
													parClickLinkURL: parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
													Properties.AppResources.List_ServiceProductsURI +
													strCurrentHyperlinkViewEditURI + objProduct.ID,
													parHyperlinkID: hyperlinkCounter);
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
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Product ID: " + objProduct.ID
													+ " contains an error in Key DD Benefits Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Service Product: " + node.NodeID + " Key DD Benefits content."
												+ " Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText +
													" Please review content and correct it.",
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											}
										}

									if(this.Service_Product_Key_Client_Benefits)
										{
										if(objProduct.KeyClientBenefits != null)
											{
											strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_ServiceProductsURI +
												strCurrentHyperlinkViewEditURI +
												objProduct.ID;
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Product_ClientKeyBenefits,
												parIsNewSection: false);
											// Check if a hyperlink must be inserted
											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: strHyperlinkImageRelationshipID,
													parClickLinkURL: parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
													Properties.AppResources.List_ServiceProductsURI +
													strCurrentHyperlinkViewEditURI + objProduct.ID,
													parHyperlinkID: hyperlinkCounter);
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
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Product ID: " + objProduct.ID
													+ " contains an error in Key Client Benefits Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Service Product: " + node.NodeID + " Key Client Benefits content."
												+ " Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText +
													" Please review content and correct it.",
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
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
									Console.Write("\t + {0} - {1}", objElement.ID, objElement.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: objElement.ISDheading,
										parIsNewSection: false);
									// Check if a hyperlink must be inserted
									if(strDocumentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
											Properties.AppResources.List_ServiceElementsURI +
											strCurrentHyperlinkViewEditURI + objElement.ID;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: strHyperlinkImageRelationshipID,
											parClickLinkURL: strCurrentListURI,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Service Element Description
									if(this.Service_Element_Description)
										{
										if(objElement.ISDdescription != null)
											{
											// Check if a hyperlink must be inserted
											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_ServiceElementsURI +
												strCurrentHyperlinkViewEditURI +
												objElement.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objElement.ISDdescription,
													parContentLayer: strCurrentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Element ID: " + objElement.ID
													+ " contains an error in ISD Description Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Service Element: " + node.NodeID + " ISD Description content."
												+ " Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText + " Please review content and correct it.",
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											}
										}
									if(this.Service_Element_Objectives)
										{
										if(objElement.Objectives != null)
											{
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_Objectives,
												parIsNewSection: false);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											// Check if a hyperlink must be inserted
											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
													Properties.AppResources.List_ServiceElementsURI +
													strCurrentHyperlinkViewEditURI +
													objElement.ID;
												}
											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objElement.Objectives,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parContentLayer: strCurrentContentLayer,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Element ID: " + objElement.ID
													+ " contains an error in Objectives Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Service Element: " + node.NodeID + " Objectives content."
												+ " Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText + " Please review content and correct it.",
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											}
										}

									if(this.Service_Element_Critical_Success_Factors)
										{
										if(objElement.CriticalSuccessFactors != null)
											{
											// Insert the heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_CriticalSuccessFactors,
												parIsNewSection: false);

											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											// Check if a hyperlink must be inserted
											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_ServiceElementsURI +
												strCurrentHyperlinkViewEditURI +
												objElement.ID;
												}
											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objElement.CriticalSuccessFactors,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Element ID: " + objElement.ID
													+ " contains an error in Critical Success Factors Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Service Element: " + node.NodeID + " - Critical Success Factor. "
												+ "Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											}
										}
									if(this.Service_Element_Key_Client_Advantages)
										{
										if(objElement.KeyClientAdvantages != null)
											{
											// Insert the heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_ClientKeyAdvantages,
												parIsNewSection: false);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											// Check if a hyperlink must be inserted
											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_ServiceElementsURI +
												strCurrentHyperlinkViewEditURI +
												objElement.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objElement.KeyClientAdvantages,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Element ID: " + objElement.ID
													+ " contains an error in Key Client Advantages Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Service Element: " + node.NodeID + " Key Client advantages "
												+ " content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}\n{2}", exc.HResult, exc.Message, strErrorText);
												}
											}
										}
									if(this.Service_Element_Key_Client_Benefits)
										{
										if(objElement.KeyClientBenefits != null)
											{
											// Insert the heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_ClientKeyBenefits,
												parIsNewSection: false);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											// Check if a hyperlink must be inserted
											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_ServiceElementsURI +
												strCurrentHyperlinkViewEditURI +
												objElement.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objElement.KeyClientBenefits,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Element ID: " + objElement.ID
													+ " contains an error in Key Client Benefits Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Service Element: " + node.NodeID + " Key Client Benefits content."
												+ " Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											}
										}
									if(this.Service_Element_Key_DD_Benefits)
										{
										if(objElement.KeyDDbenefits != null)
											{
											// Insert the heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_KeyDDBenefits,
												parIsNewSection: false);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											// Check if a hyperlink must be inserted
											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_ServiceElementsURI +
												strCurrentHyperlinkViewEditURI +
												objElement.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objElement.KeyDDbenefits,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Element ID: " + objElement.ID
													+ " contains an error in Key DD Benefits Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Service Element: " + node.NodeID + " Key DD Benefit content."
												+ " Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											}
										}
									if(this.Service_Element_Key_Performance_Indicators)
										{
										if(objElement.KeyPerformanceIndicators != null)
											{
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_KPI,
												parIsNewSection: false);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											// Check if a hyperlink must be inserted
											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_ServiceElementsURI +
												strCurrentHyperlinkViewEditURI +
												objElement.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 5,
													parHTML2Decode: objElement.KeyPerformanceIndicators,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Element ID: " + objElement.ID
													+ " contains an error in Key Performance Indicators Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Service Element: " + node.NodeID + " Key Performance Error "
												+ " content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText + " Please review content and correct it.",
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											}
										}
									if(this.Service_Element_High_Level_Process)
										{
										if(objElement.ProcessLink != null)
											{
											strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_ServiceElementsURI +
												strCurrentHyperlinkViewEditURI +
												objElement.ID;
											// Insert the heading
											Console.WriteLine("\t\t + {0} - {1}", objElement.ID,
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
									bDRMHeadingInserted = false;
									}
								else
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
									strErrorText = "Error: The Service Element ID " + node.NodeID + " could not be retrived from SharePoint.";
									this.LogError(strErrorText);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: strErrorText,
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									Console.WriteLine("\t" + strErrorText);
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
								if(bDRMHeadingInserted == false)
									{
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 5);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableReportsMeetings_Heading);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									bDRMHeadingInserted = true;
									}
								}
							// Get the entry from the DataSet
							if(parDataSet.dsDeliverables.TryGetValue(
								key: node.NodeID,
								value: out objDeliverable))
								{
								Console.Write("\t + {0} - {1}", objDeliverable.ID, objDeliverable.Title);
								objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 6);
								objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.ISDheading);
								if(node.NodeType == enumNodeTypes.ELD)
									{
									if(dictDeliverables.ContainsKey(objDeliverable.ID) != true)
										dictDeliverables.Add(objDeliverable.ID, objDeliverable.ISDheading);
									}
								else if(node.NodeType == enumNodeTypes.ELR)
									{
									if(dictReports.ContainsKey(objDeliverable.ID) != true)
										dictReports.Add(objDeliverable.ID, objDeliverable.ISDheading);
									}
								else if(node.NodeType == enumNodeTypes.ELM)
									{
									if(dictMeetings.ContainsKey(objDeliverable.ID) != true)
										dictMeetings.Add(objDeliverable.ID, objDeliverable.ISDheading);
									}
								// Check if a hyperlink must be inserted
								if(strDocumentCollection_HyperlinkURL != "")
									{
									hyperlinkCounter += 1;
									Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
										parMainDocumentPart: ref objMainDocumentPart,
										parImageRelationshipId: strHyperlinkImageRelationshipID,
										parClickLinkURL: parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
											Properties.AppResources.List_DeliverablesURI +
											strCurrentHyperlinkViewEditURI + objDeliverable.ID,
										parHyperlinkID: hyperlinkCounter);
									objRun.Append(objDrawing);
									}
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);
								// Check if the user specified to include the Deliverable Description
								if(this.DRM_Summary)
									{
									if(objDeliverable.ISDsummary != null)
										{
										strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
											Properties.AppResources.List_DeliverablesURI +
											strCurrentHyperlinkViewEditURI +
											objDeliverable.ID;
										if(this.ColorCodingLayer1)
											strCurrentContentLayer = "Layer1";
										else
											strCurrentContentLayer = "None";

										objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 6);
											objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.ISDsummary);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
										}
									} // if(this.DeliverableSummary

								// Insert the Hyperlink to the relevant position in the DRM Section.
								objParagraph = oxmlDocument.Construct_BookmarkHyperlink(
									parBodyTextLevel: 6,
									parBookmarkValue: "Deliverable_" + objDeliverable.ID);
								objBody.Append(objParagraph);
								}
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
								// Get the entry from the DataSet
								if(parDataSet.dsActivities.TryGetValue(
									key: node.NodeID,
									value: out objActivity))
									{
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 7);
									objRun = oxmlDocument.Construct_RunText(parText2Write: objActivity.ISDheading);
									// Check if a hyperlink must be inserted
									if(strDocumentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: strHyperlinkImageRelationshipID,
											parClickLinkURL: parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_ActvitiesURI +
												strCurrentHyperlinkViewEditURI + objActivity.ID,
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
											parActivityDesciption: objActivity.ISDdescription,
											parActivityInput: objActivity.Input,
											parActivityOutput: objActivity.Output,
											parActivityAssumptions: objActivity.Assumptions,
											parActivityOptionality: objActivity.Optionality);
										objBody.Append(objActivityTable);
										} // if (this.Activity_Description_Table)
									} // try
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

								// Check if the user specified to include the Service Level Commitments Table
								if(this.Service_Level_Commitments_Table)
									{
									// Prepare the data which to insert into the Service Level Table
									// Get the Service Level entry from the DataSet
									if(objDeliverableServiceLevel.AssociatedServiceLevelID != null)
										{
										if(parDataSet.dsServiceLevels.TryGetValue(
											key: Convert.ToInt16(objDeliverableServiceLevel.AssociatedServiceLevelID),
											value: out objServiceLevel))
											{
											Console.WriteLine("\t\t\t + Service Level: {0} - {1}", objServiceLevel.ID,
												objServiceLevel.Title);

											// Insert the Service Level ISD Description
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 8);
											objRun = oxmlDocument.Construct_RunText(parText2Write: objServiceLevel.ISDheading);
											// Check if a hyperlink must be inserted
											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: strHyperlinkImageRelationshipID,
													parClickLinkURL: parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_ServiceLevelsURI +
														strCurrentHyperlinkViewEditURI + objServiceLevel.ID,
													parHyperlinkID: hyperlinkCounter);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											// Add the Service Level entry to the Service Level Dictionay (list)
											if(!dictSLAs.ContainsKey(objServiceLevel.ID))
												{
												// NOTE: the DeliverableServiceLevel ID is used NOT the ServiceLevel ID.
												dictSLAs.Add(objDeliverableServiceLevel.ID, objServiceLevel.ISDheading);
												}

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
											}
										else
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
										} //if(parDataSet.dsDeliverableServiceLevels.TryGetValue(
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
						foreach(KeyValuePair<int, string> deliverableEntry in dictDeliverables.OrderBy(key => key.Value))
							{
							if(this.Deliverable_Heading)
								{
								// Get the entry from the DataSet
								if(parDataSet.dsDeliverables.TryGetValue(
									key: deliverableEntry.Key,
									value: out objDeliverable))
									{
									Console.WriteLine("\t Deliverable: {0} - {1}", objDeliverable.ID, objDeliverable.Title);
									objParagraph = oxmlDocument.Construct_Heading(
										parHeadingLevel: 3, 
										parBookMark: deliverableBookMark + objDeliverable.ID);
									objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.ISDheading);
									// Check if a hyperlink must be inserted
									if(strDocumentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: strHyperlinkImageRelationshipID,
											parClickLinkURL: parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI + objDeliverable.ID,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									// Check if the user specified to include the Deliverable Description
									if(this.Deliverable_Description)
										{
										if(objDeliverable.ISDdescription != null)
											{
											strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;

											// Check if a hyperlink must be inserted
											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 3,
													parHTML2Decode: objDeliverable.ISDdescription,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in ISD Description Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 3);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + deliverableEntry.Key 
													+ " ISD Description content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText + " Please review content and correct it.",
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.ISDDescription != null)
										} //if(this.Deliverable_Description)

									// Check if the user specified to include the Deliverable Inputs
									if(this.Deliverable_Inputs)
										{
										if(objDeliverable.Inputs != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											// Check if a hyperlink must be inserted
											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objDeliverable.Inputs,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in Inputs Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + deliverableEntry.Key
													+ " Input content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.Inputs != null)
										} //if(this.Deliverable_Inputs)

									// Check if the user specified to include the Deliverable Outputs
									if(this.Deliverable_Outputs)
										{
										if(objDeliverable.Outputs != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableOutputs_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											// Check if a hyperlink must be inserted
											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objDeliverable.Outputs,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in Outputs Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + deliverableEntry.Key
													+ " Output content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.Outputs != null)
										} //if(this.Deliverable_Outputs)

									// Check if the user specified to include the Deliverable DD's Obligations
									if(this.DDs_Deliverable_Obligations)
										{
										if(objDeliverable.DDobligations != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											// Check if a hyperlink must be inserted
											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objDeliverable.DDobligations,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in DD Obligations Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + deliverableEntry.Key
													+ " DD Obligations content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.SPObligations != null)
										} //if(this.DDS_Deliverable_Oblidations)

									// Check if the user specified to include the Client Responsibilities
									if(this.Clients_Deliverable_Responsibilities)
										{
										if(objDeliverable.ClientResponsibilities != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											// Check if a hyperlink must be inserted
											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}
											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objDeliverable.ClientResponsibilities,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in Client Responsibilities Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + deliverableEntry.Key
													+ " Client Responsibilities content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.Client_Responsibilities != null)
										} //if(this.Clients_Deliverable_Responsibilities)

									// Check if the user specified to include the Deliverable Exclusions
									if(this.Deliverable_Exclusions)
										{
										if(objDeliverable.Exclusions != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objDeliverable.Exclusions,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in Exclusions Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + deliverableEntry.Key
													+ " Exclusions content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.Exclusions != null)
										} //if(this.Deliverable_Exclusions)

									// Check if the user specified to include the Governance Controls
									if(this.Deliverable_Governance_Controls)
										{
										if(objDeliverable.GovernanceControls != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableGovernanceControls_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objDeliverable.GovernanceControls,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in Governance Controls Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + deliverableEntry.Key
													+ " Governance Controls content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.GovernanceControls != null)
										} //if(this.Deliverable_GovernanceControls)

									// Check if there are any Glossary Terms or Acronyms associated with the Deliverable.
									if(objDeliverable.GlossaryAndAcronyms != null)
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
										} //if(objDeliverable.GlossaryAndAcronyms  != null)
									} 
								else
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
									this.LogError("Error: The Deliverable ID " + deliverableEntry.Key
										+ " doesn't exist in SharePoint and couldn't be retrieved.");
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Error: Deliverable " + deliverableEntry.Key + " is missing.",
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
						objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
						objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Reports_Heading_Text);
						objParagraph.Append(objRun);
						objBody.Append(objParagraph);
						string deliverableBookMark = "Report_";
						// Insert the individual Reports in the section
						foreach(KeyValuePair<int, string> reportEntry in dictReports.OrderBy(key => key.Value))
							{
							if(this.Deliverable_Heading)
								{
								// Get the entry from the DataSet
								if(parDataSet.dsDeliverables.TryGetValue(
									key: reportEntry.Key,
									value: out objDeliverable))
									{
									Console.WriteLine("\t\t + {0} - {1}", objDeliverable.ID, objDeliverable.Title);
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3, 
										parBookMark: deliverableBookMark + objDeliverable.ID);
									objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.ISDheading);
									// Check if a hyperlink must be inserted
									if(strDocumentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: strHyperlinkImageRelationshipID,
											parClickLinkURL: parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI + objDeliverable.ID,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									// Check if the user specified to include the Report Description
									if(this.Report_Description)
										{
										if(objDeliverable.ISDdescription != null)
											{

											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 3,
													parHTML2Decode: objDeliverable.ISDdescription,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in ISD Description Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + reportEntry.Key
													+ " ISD Description content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.ISDDescription != null)
										} //if(this.Deliverable_Description)

									// Check if the user specified to include the Report Inputs
									if(this.Report_Inputs)
										{
										if(objDeliverable.Inputs != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objDeliverable.Inputs,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in Inputs Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + reportEntry.Key
													+ " Input content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.Inputs != null)
										} //if(this.Deliverable_Inputs)

									// Check if the user specified to include the Report Outputs
									if(this.Report_Outputs)
										{
										if(objDeliverable.Outputs != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableOutputs_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objDeliverable.Outputs,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in Outputs Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + reportEntry.Key
													+ " Outputs content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.Outputs != null)
										} //if(this.Deliverable_Outputs)

									// Check if the user specified to include the Report DD's Obligations
									if(this.DDs_Report_Obligations)
										{
										if(objDeliverable.DDobligations != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objDeliverable.DDobligations,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in DD's Obligations Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + reportEntry.Key
													+ " DD obligations content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.SPObligations != null)
										} //if(this.DDS_Deliverable_Oblidations)

									// Check if the user specified to include the Client Responsibilities
									if(this.Clients_Report_Responsibilities)
										{
										if(objDeliverable.ClientResponsibilities != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objDeliverable.ClientResponsibilities,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in Client Responsibilities Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + reportEntry.Key
													+ " Client Responsibilities content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.Client_Responsibilities != null)
										} //if(this.Clients_Deliverable_Responsibilities)

									// Check if the user specified to include the Report Exclusions
									if(this.Report_Exclusions)
										{
										if(objDeliverable.Exclusions != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objDeliverable.Exclusions,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in Exclusions Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + reportEntry.Key
													+ " Exclusions content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.Exclusions != null)
										} //if(this.Deliverable_Exclusions)

									// Check if the user specified to include the Governance Controls
									if(this.Deliverable_Governance_Controls)
										{
										if(objDeliverable.GovernanceControls != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableGovernanceControls_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objDeliverable.GovernanceControls,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in Governance Controls Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + reportEntry.Key
													+ " Governance Controls content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.GovernanceControls != null)
										} //if(this.Deliverable_GovernanceControls)

									// Check if there are any Glossary Terms or Acronyms associated with the Deliverable.
									if(objDeliverable.GlossaryAndAcronyms  != null)
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
										} //if(objDeliverable.GlossaryAndAcronyms  != null)
									} 
								else
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
									this.LogError("Error: The Deliverable ID " + reportEntry.Key
										+ " doesn't exist in SharePoint and couldn't be retrieved.");
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Error: Deliverable " + reportEntry.Key + " is missing.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
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
						foreach(KeyValuePair<int, string> meetingEntry in dictMeetings.OrderBy(key => key.Value))
							{
							if(this.Meeting_Heading)
								{
								// Get the entry from the DataSet
								if(parDataSet.dsDeliverables.TryGetValue(
									key: meetingEntry.Key,
									value: out objDeliverable))
									{
									Console.WriteLine("\t\t + {0} - {1}", objDeliverable.ID, objDeliverable.Title);

									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3, parBookMark: deliverableBookMark + objDeliverable.ID);
									objRun = oxmlDocument.Construct_RunText(parText2Write: objDeliverable.ISDheading);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									// Check if the user specified to include the Meeting Description
									if(this.Meeting_Description)
										{
										if(objDeliverable.ISDdescription != null)
											{

											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 3,
													parHTML2Decode: objDeliverable.ISDdescription,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in ISD Description Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 3);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + meetingEntry.Key
													+ " ISD Description content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 3);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.ISDDescription != null)
										} //if(this.Meeting_Description)

									// Check if the user specified to include the Meeting Inputs
									if(this.Meeting_Inputs)
										{
										if(objDeliverable.Inputs != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{												
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objDeliverable.Inputs,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in Inputs Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + meetingEntry.Key
													+ " Inputs content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.Inputs != null)
										} //if(this.Meeting_Inputs)

									// Check if the user specified to include the Meeting Outputs
									if(this.Meeting_Outputs)
										{
										if(objDeliverable.Outputs != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableOutputs_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objDeliverable.Outputs,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in Ouputs Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + meetingEntry.Key
													+ " Outputs content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.Outputs != null)
										} //if(this.Meeting_Outputs)

									// Check if the user specified to include the Meeting DD's Obligations
									if(this.DDs_Meeting_Obligations)
										{
										if(objDeliverable.DDobligations != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objDeliverable.DDobligations,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in DD's Obligations Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + meetingEntry.Key
													+ " DD Obligations content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.SPObligations != null)
										} //if(this.DDS_Report_Oblidations)

									// Check if the user specified to include the Client Responsibilities
									if(this.Clients_Meeting_Responsibilities)
										{
										if(objDeliverable.ClientResponsibilities != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objDeliverable.ClientResponsibilities,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}

											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in Client Responsibilities Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + meetingEntry.Key
													+ " Content Responsibilities content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.Client_Responsibilities != null)
										} //if(this.Clients_Report_Responsibilities)

									// Check if the user specified to include the Deliverable Exclusions
									if(this.Deliverable_Exclusions)
										{
										if(objDeliverable.Exclusions != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objDeliverable.Exclusions,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in Exclusions Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + meetingEntry.Key
													+ " Exclusions content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.Exclusions != null)
										} //if(this.Report_Exclusions)

									// Check if the user specified to include the Governance Controls
									if(this.Meeting_Governance_Controls)
										{
										if(objDeliverable.GovernanceControls != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableGovernanceControls_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
												Properties.AppResources.List_DeliverablesURI +
												strCurrentHyperlinkViewEditURI +
												objDeliverable.ID;
												}

											if(this.ColorCodingLayer1)
												strCurrentContentLayer = "Layer1";
											else
												strCurrentContentLayer = "None";

											// Insert the contents
											try
												{
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 4,
													parHTML2Decode: objDeliverable.GovernanceControls,
													parContentLayer: strCurrentContentLayer,
													parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
													parHyperlinkURL: strCurrentListURI,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parPictureNo: ref iPictureNo,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
												}
											catch(InvalidContentFormatException exc)
												{
												Console.WriteLine("\n\nException occurred: {0}", exc.Message);
												// A Table content error occurred, record it in the error log.
												this.LogError("Error: Service Deliverable ID: " + objDeliverable.ID
													+ " contains an error in Governance Controls Enhance Rich Text. "
													+ exc.Message);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: "A content error occurred at this position and valid content could "
													+ "not be interpreted and inserted here. Please review the content in the SharePoint "
													+ "system and correct it." + exc.Message,
													parIsNewSection: false,
													parIsError: true);
												if(strDocumentCollection_HyperlinkURL != "")
													{
													hyperlinkCounter += 1;
													Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
														parMainDocumentPart: ref objMainDocumentPart,
														parImageRelationshipId: strHyperlinkImageRelationshipID,
														parHyperlinkID: hyperlinkCounter,
														parClickLinkURL: strCurrentListURI);
													objRun.Append(objDrawing);
													}
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												}
											catch(Exception exc)
												{
												strErrorText = "Content Error in Deliverable: " + meetingEntry.Key
													+ " Governance Controls content. Please review the content and correct it.";
												this.LogError(strErrorText);
												objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: strErrorText,
													parIsNewSection: false,
													parIsError: true);
												objParagraph.Append(objRun);
												objBody.Append(objParagraph);
												Console.WriteLine("\nException occurred: {0} - {1}", exc.HResult, exc.Message);
												}
											} // if(objDeliverable.GovernanceControls != null)
										} //if(this.Deliverable_GovernanceControls)

									// Check if there are any Glossary Terms or Acronyms associated with the Deliverable.
									if(objDeliverable.GlossaryAndAcronyms  != null)
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
										} //if(objDeliverable.GlossaryAndAcronyms  != null)
									} 
								else
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
									this.LogError("Error: The Deliverable ID " + meetingEntry.Key
										+ " doesn't exist in SharePoint and couldn't be retrieved.");
									objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Error: Deliverable " + meetingEntry.Key + " is missing.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
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
							// Obtain the Deliverable Service Level from SharePoint
							if(parDataSet.dsDeliverableServiceLevels.TryGetValue(
								key: servicelevelItem.Key,
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
										Console.WriteLine("\t\t + Deliverable ServiceLevel: {0} - {1}",
											objDeliverableServiceLevel.ID, objDeliverableServiceLevel.Title);
										Console.WriteLine("\t\t + Service Level: {0} - {1}", objServiceLevel.ID, objServiceLevel.Title);
										Console.WriteLine("\t\t + Service Hour.: {0}", objServiceLevel.ServiceHours);

										if(this.Service_Level_Heading_in_Section)
											{
											// Insert the Service Level ISD Heading
											objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2,
												parBookMark: servicelevelBookMark + objServiceLevel.ID);
											objRun = oxmlDocument.Construct_RunText(parText2Write: objServiceLevel.ISDheading);
											// Check if a hyperlink must be inserted
											if(strDocumentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: strHyperlinkImageRelationshipID,
													parClickLinkURL: parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_ServiceLevelsURI +
														strCurrentHyperlinkViewEditURI + objServiceLevel.ID,
													parHyperlinkID: hyperlinkCounter);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(this.Service_Level_Table_in_Section)
												{
												if(objServiceLevel.ISDdescription != null)
													{
													strCurrentListURI = parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL +
														Properties.AppResources.List_ServiceLevelsURI +
														strCurrentHyperlinkViewEditURI +
														objServiceLevel.ID;

													strCurrentContentLayer = "None";
													try
														{
														objHTMLdecoder.DecodeHTML(
															parMainDocumentPart: ref objMainDocumentPart,
															parDocumentLevel: 2,
															parHTML2Decode: objServiceLevel.ISDdescription,
															parContentLayer: strCurrentContentLayer,
															parHyperlinkImageRelationshipID: strHyperlinkImageRelationshipID,
															parHyperlinkURL: strCurrentListURI,
															parTableCaptionCounter: ref tableCaptionCounter,
															parImageCaptionCounter: ref imageCaptionCounter,
															parPictureNo: ref iPictureNo,
															parHyperlinkID: ref hyperlinkCounter,
															parPageHeightTwips: this.PageHight,
															parPageWidthTwips: this.PageWith, parSharePointSiteURL: parDataSet.SharePointSiteURL);
														}
													catch(InvalidContentFormatException exc)
														{
														Console.WriteLine("\n\nException occurred: {0}", exc.Message);
														// A Table content error occurred, record it in the error log.
														this.LogError("Error: Service Level ID: " + objServiceLevel.ID
															+ " contains an error in ISD Description Enhance Rich Text. "
															+ exc.Message);
														objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
														objRun = oxmlDocument.Construct_RunText(
															parText2Write: "A content error occurred at this position and valid content could "
															+ "not be interpreted and inserted here. Please review the content in the SharePoint "
															+ "system and correct it." + exc.Message,
															parIsNewSection: false,
															parIsError: true);
														if(strDocumentCollection_HyperlinkURL != "")
															{
															hyperlinkCounter += 1;
															Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
																parMainDocumentPart: ref objMainDocumentPart,
																parImageRelationshipId: strHyperlinkImageRelationshipID,
																parHyperlinkID: hyperlinkCounter,
																parClickLinkURL: strCurrentListURI);
															objRun.Append(objDrawing);
															}
														objParagraph.Append(objRun);
														objBody.Append(objParagraph);
														}
													catch(Exception exc)
														{
														this.LogError("Content Error in Deliverable " + servicelevelItem.Key +
															" Please review all content for this deliverable and correct it.");
														objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
														objRun = oxmlDocument.Construct_RunText(
															parText2Write: "Content Error in ServiceLevel " + servicelevelItem.Key +
															" Please review all content for this ServiceLevel and correct it.",
															parIsNewSection: false,
															parIsError: true);
														objParagraph.Append(objRun);
														objBody.Append(objParagraph);
														Console.WriteLine("\n\nException occurred: {0} - {1}", exc.HResult, exc.Message);
														}
													}

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
												} //if(this.Service_Level_Commitments_Table)
											} //if(this.Service_Level_Heading_in_Section)
										} //if(parDataSet.dsServiceLevels.TryGetValue(
									} // if(objDeliverableServiceLevel.AssociatedServiceLevelID != null)
								else
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
								} //if(parDataSet.dsDeliverableServiceLevels.TryGetValue(
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
							parSDDPdatacontext: parDataSet.SDDPdatacontext,
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
							parPictureNo: ref iPictureNo,
							parHyperlinkID: ref hyperlinkCounter,
							parSharePointSiteURL: parDataSet.SharePointSiteURL);
						}
					}
				//----------------------------------------------
				// Insert the Document Generation Error Section
				// ---------------------------------------------

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

			//+ -------------------
			//+ Handle Exceptions
			//+ -------------------
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

		} // end of Services_Framework_Document_DRM_Sections class
	}
