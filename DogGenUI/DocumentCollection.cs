﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Dynamic;
using System.Data.Services.Client;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint;
using DocGenerator.SDDPServiceReference;

namespace DocGenerator
	{
	/// <summary>
	///	Mapped to the [Content Layer Colour Coding Option] column in SharePoint List
	/// </summary>
	public enum enumContent_Layer_Colour_Coding_Options
		{
		Colour_Code_Layer_1=1,
		Colour_Code_Layer_2=2,
		Colour_Code_Layer_3=3
		}
	/// <summary>
	///	Mapped to the [Generate Action] column in SharePoint List
	/// </summary>
	public enum enumGenerate_Actions
		{
		Save_but_dont_generate_the_documents_yet=1,
		Submit_to_the_generate_queue=2,
		Schedule_for_a_specific_date_and_time=3
		}
	/// <summary>
	/// Mapped to the [Generate Schedule Option] column in SharePoint
	/// </summary>
	public enum enumGenerateScheduleOptions
		{
		Do_NOT_Repeat=0,
		Repeat_every=1
		}
	/// <summary>
	/// Mapped to the values of the [Generate Repeat Interval] column in SharePoint;
	/// </summary>
	public enum enumGenerateRepeatIntervals
		{
		Day=1,
		Week=2,
		Month=3
		}
	/// <summary>
	/// Mapped to the values of the [Hyperlink Options] column in SharePoint;
	/// </summary>
	public enum enumHyperlinkOptions
		{
		Do_NOT_Include_Hyperlinks=0,
		Include_EDIT_Hyperlinks=1,
		Include_VIEW_Hyperlinks=2
		}

	public enum enumPresentationMode
		{
		Layered=0,
		Expanded=1
		}

	/// <summary>
	/// This list contains the documents that the user selected which needs to be generated.
	/// </summary>
	public class DocumentCollection
		{
		// Object Properties

		public int ID{get; set;}
		public string ClientName{get; set;}
		public string Title{get; set;}
		public bool ColourCodingLayer1{get; set;}
		public bool ColourCodingLayer2{get; set;}
		public bool ColourCodingLayer3{get; set;}
		public enumHyperlinkOptions HyperLinkOption{get; set;}
		public int Mapping{get; set;}
		public enumPresentationMode PresentationMode{get; set;}
		public int PricingWorkbook{get; set;}
		public List<enumDocumentTypes> DocumentsToGenerate{get; set;}
		public bool NotifyMe{get; set;}
		public string NotificationEmail{get; set;}
		public int? RequestingUserID{get; set;}
		public enumGenerateScheduleOptions GenerateScheduleOption{get; set;}
		public DateTime GenerateOnDateTime{get; set;}
		public enumGenerateRepeatIntervals GenerateRepeatInterval{get; set;}
		public int GenerateRepeatIntervalValue{get; set;}
		public List<Hierarchy> SelectedNodes{get; set;}
		public List<dynamic> Document_and_Workbook_objects{get; set;}
		public bool UnexpectedErrors{get; set;}

		// Object Methods
		public void SetBasicProperties(int parID, string parTitle)
			{
			this.ID = parID;
			this.Title = parTitle;
			}

		public bool SetGenerateStatus(int parID, string parStatus)
			{
			// add the code to set the status of the Document Collection
			return false;
			}

		/// <summary>
		/// Method which obtains all the Document Collections from the [Document Collection Library] List
		/// that still need to be generated.
		/// The Method returns a List collection consisting of Document Collection objects that 
		/// must be generated.
		/// </summary>
		public static string GetCollectionsToGenerate(
			ref List<DocumentCollection> parCollectionsToGenerate,
			DesignAndDeliveryPortfolioDataContext parSDDPdatacontext)
			{
			List<int> optionsWorkList = new List<int>();
			string enumWorkString;
			//datacontexSDDP.MergeOption = MergeOption.AppendOnly;			//Use only if data is added
			//datacontexSDDP.MergeOption = MergeOption.OverwriteChanges;	//use when data is updated
			//datacontexSDDP.MergeOption = MergeOption.NoTracking;

			try
				{
				var dsDocCollectionLibrary = parSDDPdatacontext.DocumentCollectionLibrary
						.Expand(dc => dc.Client_)
						.Expand(dc => dc.ContentLayerColourCodingOption)
						.Expand(dc => dc.GenerateFrameworkDocuments)
						.Expand(dc => dc.GenerateInternalDocuments)
						.Expand(dc => dc.GenerateExternalDocuments)
						//.Expand(dc => dc.GenerateRepeatInterval)
						.Expand(dc => dc.HyperlinkOptions)
						.Expand(dc => dc.ModifiedBy);

				var dsDocumentCollections = 
					from docCollection in dsDocCollectionLibrary
					where docCollection.GenerateActionValue != null && docCollection.GenerateActionValue != "Save but don't generate the documents yet"
					orderby docCollection.Id select docCollection;	

				foreach(var recDocCollsToGen in dsDocumentCollections)
					{
					Console.WriteLine("\r\nDocumentCollection ID: {0}  Title: {1} Client Name: [{2}] - Client Title:[{3}] ", recDocCollsToGen.Id, recDocCollsToGen.Title, recDocCollsToGen.Client_.DocGenClientName, recDocCollsToGen.Client_.Title);

					// Create a new Instance for the DocumentCollection into which the object properties are loaded
					DocumentCollection objDocumentCollection = new DocumentCollection();
					//Set the basic object properties
					objDocumentCollection.ID = recDocCollsToGen.Id;
					Console.WriteLine("\t ID: {0} ", objDocumentCollection.ID);

					if(recDocCollsToGen.Client_.DocGenClientName == null)
						objDocumentCollection.ClientName = "the Client";
					else
						objDocumentCollection.ClientName = recDocCollsToGen.Client_.DocGenClientName;
					Console.WriteLine("\t ClientName: {0} ", objDocumentCollection.ClientName);

					if(recDocCollsToGen.Title == null)
						objDocumentCollection.Title = "Collection Title for entry " + recDocCollsToGen.Id;
					else
						objDocumentCollection.Title = recDocCollsToGen.Title;
					Console.WriteLine("\t Title: {0}", objDocumentCollection.Title);

					if(recDocCollsToGen.GenerateNotifyMe == null)
						objDocumentCollection.NotifyMe = false;
					else
						objDocumentCollection.NotifyMe = recDocCollsToGen.GenerateNotifyMe.Value;
					Console.WriteLine("\t NotifyMe: {0} ", objDocumentCollection.NotifyMe);

					objDocumentCollection.RequestingUserID = recDocCollsToGen.ModifiedById;
					Console.WriteLine("\t User who LAST requested the documents: {0} - {1}", recDocCollsToGen.ModifiedBy.Id, recDocCollsToGen.ModifiedBy.Name);

					if(recDocCollsToGen.GenerateNotificationEMail == null)
						objDocumentCollection.NotificationEmail = "None";
					else
						if(recDocCollsToGen.GenerateNotificationEMail == null)
							{
							objDocumentCollection.NotificationEmail = recDocCollsToGen.GenerateNotificationEMail;
							}
						else
							{
						objDocumentCollection.NotificationEmail = recDocCollsToGen.ModifiedBy.WorkEmail;
							}
						
					Console.WriteLine("\t NotificationEmail: {0} ", objDocumentCollection.NotificationEmail);
					// Set the GenerateOnDateTime value
					if(recDocCollsToGen.GenerateOnDateTime == null)
						objDocumentCollection.GenerateOnDateTime = DateTime.Now;
					else
						objDocumentCollection.GenerateOnDateTime = recDocCollsToGen.GenerateOnDateTime.Value;
					Console.WriteLine("\t GenerateOnDateTime: {0} ", objDocumentCollection.GenerateOnDateTime);
					// Set the Mapping value
					if(recDocCollsToGen.Mapping_Id != null)
						{
						try
							{
							objDocumentCollection.Mapping = Convert.ToInt32(recDocCollsToGen.Mapping_Id);
							}
						catch(OverflowException ex)
							{
							Console.WriteLine("Overflow Exception occurred when converting the Mappin value to a Integer.\n Error Description: {0}", ex.Message);
							objDocumentCollection.Mapping = 0;
							}
						}
					else
						{
						objDocumentCollection.Mapping = 0;
						}
					//Console.WriteLine("\t Mapping: {0} ", objDocumentCollection.Mapping);
					// Set the PricingWorkbook value
					if(recDocCollsToGen.PricingWorkbookId != null)
						try
							{
							objDocumentCollection.PricingWorkbook = Convert.ToInt32(recDocCollsToGen.PricingWorkbookId);
							}
						catch(OverflowException ex)
							{
							Console.WriteLine("Overflow Exception occurred when converting the Pricing Workbook value to a Integer."
								+ "\n Error Description: {0}", ex.Message);
							objDocumentCollection.Mapping = 0;
							}
					else
						objDocumentCollection.PricingWorkbook = 0;
					//Console.WriteLine("\t PricingWorkbook: {0} ", objDocumentCollection.PricingWorkbook);
					// Set the Generate Schedule Options
					enumGenerateScheduleOptions generateSchdlOption;
					if(recDocCollsToGen.GenerateScheduleOptionValue != null)
						{
						if(PrepareStringForEnum(recDocCollsToGen.GenerateScheduleOptionValue, out enumWorkString))
							{
							if(Enum.TryParse<enumGenerateScheduleOptions>(enumWorkString, out generateSchdlOption))
								objDocumentCollection.GenerateScheduleOption = generateSchdlOption;
							else
								objDocumentCollection.GenerateScheduleOption = enumGenerateScheduleOptions.Do_NOT_Repeat;
							}
						else
							objDocumentCollection.GenerateScheduleOption = enumGenerateScheduleOptions.Do_NOT_Repeat;
						}
					else
						{
						objDocumentCollection.GenerateScheduleOption = enumGenerateScheduleOptions.Do_NOT_Repeat;
						}
					Console.WriteLine("\t Generate ScheduleOption: {0} ", objDocumentCollection.GenerateScheduleOption);
					// Set the GenerateRepeatInterval
					enumGenerateRepeatIntervals generateRepeatIntrvl;
					if(recDocCollsToGen.GenerateRepeatIntervalValue0 != null)
						{
						if(PrepareStringForEnum(recDocCollsToGen.GenerateRepeatIntervalValue0, out enumWorkString))
							{
							if(Enum.TryParse<enumGenerateRepeatIntervals>(enumWorkString, out generateRepeatIntrvl))
								{
								objDocumentCollection.GenerateRepeatInterval = generateRepeatIntrvl;
								}
							else
								{
								objDocumentCollection.GenerateRepeatInterval = enumGenerateRepeatIntervals.Month;
								}
							}
						else
							{
							objDocumentCollection.GenerateRepeatInterval = enumGenerateRepeatIntervals.Month;
							}
						}
					else
						{
						objDocumentCollection.GenerateRepeatInterval = enumGenerateRepeatIntervals.Month;
						}
					Console.WriteLine("\t GenerateRepeatInterval: {0} ", objDocumentCollection.GenerateRepeatInterval);
					// Set the GenerateRepeatInterval Value
					if(recDocCollsToGen.GenerateRepeatIntervalValue != null)
						{
						try
							{
							objDocumentCollection.GenerateRepeatIntervalValue = Convert.ToInt32(recDocCollsToGen.GenerateRepeatIntervalValue.Value);
							}
						catch(OverflowException ex)
							{
							Console.WriteLine("Overflow Exception occurred when converting the Generate Repeat Interval to a Integer.\n Error Description: {0}", ex.Message);
							objDocumentCollection.GenerateRepeatIntervalValue = 0;
							}
						}
					else
						{
						objDocumentCollection.GenerateRepeatIntervalValue = 0;
						}
					Console.WriteLine("\t GenerateRepeatIntervalValue: {0} ", objDocumentCollection.GenerateRepeatIntervalValue);
					// Set the Hyperlink Options
					if(recDocCollsToGen.HyperlinkOptionsValue != null)
						{
						enumHyperlinkOptions hyperLnkOption;
						if(PrepareStringForEnum(recDocCollsToGen.HyperlinkOptionsValue, out enumWorkString))
							{
							if(Enum.TryParse<enumHyperlinkOptions>(enumWorkString, out hyperLnkOption))
								{
								objDocumentCollection.HyperLinkOption = hyperLnkOption;
								}
							else
								{
								objDocumentCollection.HyperLinkOption = enumHyperlinkOptions.Do_NOT_Include_Hyperlinks;
								}
							}
						else
							{
							objDocumentCollection.HyperLinkOption = enumHyperlinkOptions.Do_NOT_Include_Hyperlinks;
							}
						}
					else
						{
						objDocumentCollection.HyperLinkOption = enumHyperlinkOptions.Do_NOT_Include_Hyperlinks;
						}
					Console.WriteLine("\t HyperlinkOption: {0} ", objDocumentCollection.HyperLinkOption);
					// Get the Content Layer Colour Coding Option
					// Console.WriteLine("\t Content Layer Colour Coding has {0} entries.", DCsToGen.ContentLayerColourCodingOption.Count.ToString());
					objDocumentCollection.ColourCodingLayer1 = false;
					objDocumentCollection.ColourCodingLayer2 = false;
					objDocumentCollection.ColourCodingLayer3 = false;
					if(recDocCollsToGen.ContentLayerColourCodingOption.Count > 0)
						{
						foreach(var entry in recDocCollsToGen.ContentLayerColourCodingOption)
							{
							//Console.WriteLine("\t\t {0}", entry.Value);
							enumContent_Layer_Colour_Coding_Options CLCCOptions;
							if(PrepareStringForEnum(entry.Value, out enumWorkString))
								{
								if(Enum.TryParse<enumContent_Layer_Colour_Coding_Options>(enumWorkString, out CLCCOptions))
									{
									if(CLCCOptions.Equals(enumContent_Layer_Colour_Coding_Options.Colour_Code_Layer_1))
										{
										objDocumentCollection.ColourCodingLayer1 = true;
										}
									if(CLCCOptions.Equals(enumContent_Layer_Colour_Coding_Options.Colour_Code_Layer_2))
										{
										objDocumentCollection.ColourCodingLayer2 = true;
										}
									if(CLCCOptions.Equals(enumContent_Layer_Colour_Coding_Options.Colour_Code_Layer_3))
										{
										objDocumentCollection.ColourCodingLayer3 = true;
										}
									}
								}
							} //Foreach Loop
						}
					Console.WriteLine("\t ContentColourCodingLayer1: {0} ", objDocumentCollection.ColourCodingLayer1);
					Console.WriteLine("\t ContentColourCodingLayer2: {0} ", objDocumentCollection.ColourCodingLayer2);
					Console.WriteLine("\t ContentColourCodingLayer3: {0} ", objDocumentCollection.ColourCodingLayer3);

					//Set the PresentationMode
					if(recDocCollsToGen.PresentationModeValue == "Layered")
						objDocumentCollection.PresentationMode = enumPresentationMode.Layered;
					else
						objDocumentCollection.PresentationMode = enumPresentationMode.Expanded;
					
					int noOfDocsToGenerateInCollection = 0;
					List<enumDocumentTypes> listOfDocumentTypesToGenerate = new List<enumDocumentTypes>();
					enumDocumentTypes docType;
					// Set the FrameworkDocuments that must be generated
					Console.WriteLine("\t Generate Framework Documents: {0} entries.", recDocCollsToGen.GenerateFrameworkDocuments.Count.ToString());
					if(recDocCollsToGen.GenerateFrameworkDocuments.Count > 0)
						{
						foreach(var entry in recDocCollsToGen.GenerateFrameworkDocuments)
							{
							if(PrepareStringForEnum(entry.Value, out enumWorkString))
								{
								if(Enum.TryParse<enumDocumentTypes>(enumWorkString, out docType))
									{
									listOfDocumentTypesToGenerate.Add(docType);
									//Console.WriteLine("\t\t + [{0}]", docType);
									noOfDocsToGenerateInCollection += 1;
									}
								else
									if(Enum.TryParse<enumDocumentTypes>(enumWorkString, out docType))
									listOfDocumentTypesToGenerate.Add(docType);
								else
									Console.WriteLine("\t\t [{0}] Not found as enumeration [{1}]", enumWorkString, docType);
								}
							}
						}
					// Set the Internal Documents that must be generated
					Console.WriteLine("\t Generate Internal Documents: {0} entries.", recDocCollsToGen.GenerateInternalDocuments.Count.ToString());
					if(recDocCollsToGen.GenerateInternalDocuments.Count > 0)
						{
						foreach(var entry in recDocCollsToGen.GenerateInternalDocuments)
							{
							if(PrepareStringForEnum(entry.Value, out enumWorkString))
								{
								if(Enum.TryParse<enumDocumentTypes>(enumWorkString, out docType))
									{
									listOfDocumentTypesToGenerate.Add(docType);
									Console.WriteLine("\t\t + [{0}]", docType);
									noOfDocsToGenerateInCollection += 1;
									}
								}
							}
						}
					// Set the External Documents that must be generated
					Console.WriteLine("\t Generate External Documents: {0} entries.", recDocCollsToGen.GenerateExternalDocuments.Count.ToString());
					if(recDocCollsToGen.GenerateExternalDocuments.Count > 0)
						{
						foreach(var entry in recDocCollsToGen.GenerateExternalDocuments)
							{
							if(PrepareStringForEnum(entry.Value, out enumWorkString))
								{
								if(Enum.TryParse<enumDocumentTypes>(enumWorkString, out docType))
									{
									listOfDocumentTypesToGenerate.Add(docType);
									//Console.WriteLine("\t\t + [{0}]", docType);
									}
								}
							}
						}
					objDocumentCollection.DocumentsToGenerate = listOfDocumentTypesToGenerate;
					Console.WriteLine("\t {0} document to be generated for the Document Collection.", objDocumentCollection.DocumentsToGenerate.Count);

					//Load the nodes that need to be generated.
					//Set the Selected Nodes which must be generated by building a hierchical List with Hierarchy objects
					Console.WriteLine("\t Loading the Nodes that the user selected.");
					if(recDocCollsToGen.SelectedNodes != null)
						{
						List<Hierarchy> listOfNodesToGenerate = new List<Hierarchy>();
						if(Hierarchy.ConstructHierarchy(recDocCollsToGen.SelectedNodes, ref listOfNodesToGenerate))
							{
							objDocumentCollection.SelectedNodes = listOfNodesToGenerate;
							//Console.WriteLine("\t {0} nodes successfully loaded by ConstructHierarchy method.", listOfNodesToGenerate.Count);
							}
						else //there was an error during the Construct of the Hierarchy method
							{
							Console.WriteLine("An error occurred when the Hierarchy was constructed.");
							}
						}
					else
						{
						Console.WriteLine("\t There are no selected content to generate for Document Collection {0} - {1}", recDocCollsToGen.Id, recDocCollsToGen.Title);
						}
					//-----------------------------------------------------------------
					// Load options for each of the documents that need to be generated
					//-----------------------------------------------------------------
					Console.WriteLine("\t Creating the Document object(s) for {0} document.", objDocumentCollection.DocumentsToGenerate.Count);
					
					if(objDocumentCollection.DocumentsToGenerate.Count > 0)
						{
						string strTemplateURL = ""; // variable used to store the individual Template URLs
						// Declare a new List of Document_and_Workbook objects that can hold all the object entries
						List<dynamic> listDocumentWorkbookObjects = new List<dynamic>();
						
						foreach(enumDocumentTypes objDocsToGenerate in objDocumentCollection.DocumentsToGenerate)
							{
							Console.WriteLine("\n\t Busy constructing Document object for {0}...", objDocsToGenerate.ToString());
							switch(objDocsToGenerate)
								{

								//====================================================
								case enumDocumentTypes.Activity_Effort_Workbook:
									{
									//NOT_AVAILABLE: not currently implemented - Activities and Effort Drivers removed from SharePoint
									break;
									}
								//====================================================
								// Client Requirement Mapping workbook
								case enumDocumentTypes.Client_Requirement_Mapping_Workbook:
									{
									Client_Requirements_Mapping_Workbook objClientRequirementsMappingWorkbook = new Client_Requirements_Mapping_Workbook();
									objClientRequirementsMappingWorkbook.DocumentCollectionID = objDocumentCollection.ID;
									objClientRequirementsMappingWorkbook.DocumentCollectionTitle = objDocumentCollection.Title;
									objClientRequirementsMappingWorkbook.DocumentStatus = enumDocumentStatusses.New;
									objClientRequirementsMappingWorkbook.DocumentType = enumDocumentTypes.Client_Requirement_Mapping_Workbook;
									strTemplateURL = GetTheDocumentTemplate(parSDDPdatacontext, "Client Requirements Mapping Workbook");
                                             switch (strTemplateURL)
										{
										case "None":
											objClientRequirementsMappingWorkbook.Template = "";
											objClientRequirementsMappingWorkbook.LogError("The template could not be found.");
                                                       break;
										case "Error":
											objClientRequirementsMappingWorkbook.Template = "";
											objClientRequirementsMappingWorkbook.LogError("The template could not be accessed.");
                                                       break;
										default:
											objClientRequirementsMappingWorkbook.Template = strTemplateURL;
											break;
										}
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objClientRequirementsMappingWorkbook.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objClientRequirementsMappingWorkbook.HyperlinkView = true;

									// The Hierarchical nodes from the Document Collection is not applicable on this Document object.
									objClientRequirementsMappingWorkbook.SelectedNodes = null;
									// Instead, set the Client Requirements Mapping value
									objClientRequirementsMappingWorkbook.CRM_Mapping = recDocCollsToGen.Mapping_Id;

									// add the object to the Document Collection's DocumentsWorkbooks to be generated.
									listDocumentWorkbookObjects.Add(objClientRequirementsMappingWorkbook);
									break;
									}
								//================================================
								// Content Status Workbook
								case enumDocumentTypes.Content_Status_Workbook:
									{
									Content_Status_Workbook objContentStatus_Workbook = new Content_Status_Workbook();
									objContentStatus_Workbook.DocumentCollectionID = objDocumentCollection.ID;
									objContentStatus_Workbook.DocumentCollectionTitle = objDocumentCollection.Title;
									objContentStatus_Workbook.DocumentStatus = enumDocumentStatusses.New;
									objContentStatus_Workbook.DocumentType = enumDocumentTypes.Content_Status_Workbook;
									strTemplateURL = GetTheDocumentTemplate(parSDDPdatacontext, "Content Status Workbook");
									switch(strTemplateURL)
										{
										case "None":
											objContentStatus_Workbook.Template = "";
											objContentStatus_Workbook.LogError("The template could not be found.");
                                                       break;
										case "Error":
											objContentStatus_Workbook.Template = "";
											objContentStatus_Workbook.LogError("The template could not be accessed.");
                                                       break;
										default:
											objContentStatus_Workbook.Template = Properties.AppResources.SharePointSiteURL.Substring(0, 
												Properties.AppResources.SharePointSiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									//Console.WriteLine("\t Template: {0}", objContentStatus_Workbook.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objContentStatus_Workbook.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objContentStatus_Workbook.HyperlinkView = true;
									
									// Add the Hierarchical nodes from the Document Collection obect to the Document object.
									objContentStatus_Workbook.SelectedNodes = objDocumentCollection.SelectedNodes;
									// add the object to the Document Collection's DocumentsWorkbooks to be generated.
									listDocumentWorkbookObjects.Add(objContentStatus_Workbook);
									break;
									}
								//================================================
								// Contract SoW Service Description
								case enumDocumentTypes.Contract_SoW_Service_Description:
									{
									Contract_SoW_Service_Description objContractSoWServiceDescription = new Contract_SoW_Service_Description();
									objContractSoWServiceDescription.DocumentCollectionID = objDocumentCollection.ID;
									objContractSoWServiceDescription.DocumentCollectionTitle = objDocumentCollection.Title;
									objContractSoWServiceDescription.DocumentStatus = enumDocumentStatusses.New;
									objContractSoWServiceDescription.DocumentType = enumDocumentTypes.Contract_SoW_Service_Description;
									objContractSoWServiceDescription.IntroductionRichText = recDocCollsToGen.ContractSDIntroduction;
									strTemplateURL = GetTheDocumentTemplate(parSDDPdatacontext, "Contract: Service Description (Appendix F)");
									switch(strTemplateURL)
										{
										case "None":
											objContractSoWServiceDescription.Template = "";
											objContractSoWServiceDescription.LogError("The template could not be found.");
											break;
										case "Error":
											objContractSoWServiceDescription.Template = "";
											objContractSoWServiceDescription.LogError("Unable to access the template.");
											break;
										default:
											objContractSoWServiceDescription.Template = Properties.AppResources.SharePointSiteURL.Substring(0, 
												Properties.AppResources.SharePointSiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									//Console.WriteLine("\t Template: {0}", objContractSoWServiceDescription.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objContractSoWServiceDescription.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objContractSoWServiceDescription.HyperlinkView = true;

									objContractSoWServiceDescription.ColorCodingLayer1 = objDocumentCollection.ColourCodingLayer1;
									objContractSoWServiceDescription.ColorCodingLayer2 = objDocumentCollection.ColourCodingLayer2;
									objContractSoWServiceDescription.ColorCodingLayer3 = objDocumentCollection.ColourCodingLayer3;

									// Load the Presentation Layer
									objContractSoWServiceDescription.PresentationMode = objDocumentCollection.PresentationMode;

									// Load the Document Options
									if(recDocCollsToGen.SoWSDOptions != null)
										{
										if(ConvertOptionsToList(recDocCollsToGen.SoWSDOptions, ref optionsWorkList)) // conversion is successful
											{
											objContractSoWServiceDescription.TransposeDocumentOptions(ref optionsWorkList);
											}
										else // the conversion failed
											{
											objContractSoWServiceDescription.LogError("Invalid format in the Document Options :. unable to generate the document.");
											//Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
											}
										}
									else  // == Null
										{
										objContractSoWServiceDescription.LogError("No document options were specified - cannot generate blank documents.");
										//Console.WriteLine("No document options were selected - cannot generate blank documents.");
										}

									// Add the Hierarchical nodes from the Document Collection obect to the Document object.
									objContractSoWServiceDescription.SelectedNodes = objDocumentCollection.SelectedNodes;
									// add the object to the Document Collection's DocumentsWorkbooks to be generated.
									listDocumentWorkbookObjects.Add(objContractSoWServiceDescription);
									break;
									}
								//================================================
								// CSD based on Client Requirements Mapping
								case enumDocumentTypes.CSD_based_on_Client_Requirements_Mapping:
									{
									CSD_based_on_ClientRequirementsMapping objCSDbasedonCRM = new CSD_based_on_ClientRequirementsMapping();
									objCSDbasedonCRM.DocumentCollectionID = objDocumentCollection.ID;
									objCSDbasedonCRM.DocumentCollectionTitle = objDocumentCollection.Title;
									objCSDbasedonCRM.DocumentStatus = enumDocumentStatusses.New;
									objCSDbasedonCRM.DocumentType = enumDocumentTypes.CSD_based_on_Client_Requirements_Mapping;
									objCSDbasedonCRM.IntroductionRichText = recDocCollsToGen.CSDDocumentIntroduction;
									objCSDbasedonCRM.ExecutiveSummaryRichText = recDocCollsToGen.CSDDocumentExecSummary;
									strTemplateURL = GetTheDocumentTemplate(parSDDPdatacontext, "Client Service Description");
									switch(strTemplateURL)
										{
										case "None":
											objCSDbasedonCRM.Template = "";
											objCSDbasedonCRM.LogError("The template could not be found.");
											break;
										case "Error":
											objCSDbasedonCRM.Template = "";
											objCSDbasedonCRM.LogError("Unable to access the template.");
											break;
										default:
											objCSDbasedonCRM.Template = Properties.AppResources.SharePointSiteURL.Substring(0, 
												Properties.AppResources.SharePointSiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									//Console.WriteLine("\t Template: {0}", objCSDbasedonCRM.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objCSDbasedonCRM.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objCSDbasedonCRM.HyperlinkView = true;

									objCSDbasedonCRM.ColorCodingLayer1 = objDocumentCollection.ColourCodingLayer1;
									objCSDbasedonCRM.ColorCodingLayer2 = objDocumentCollection.ColourCodingLayer2;
									objCSDbasedonCRM.ColorCodingLayer3 = objDocumentCollection.ColourCodingLayer3;

									// Load the Presentation Layer
									objCSDbasedonCRM.PresentationMode = objDocumentCollection.PresentationMode;

									// Load the Document Options
									if(recDocCollsToGen.CSDDocumentBasedOnCRMOptions != null)
										{
										if(ConvertOptionsToList(recDocCollsToGen.CSDDocumentBasedOnCRMOptions, ref optionsWorkList)) // conversion is successful
											{
											objCSDbasedonCRM.TransposeDocumentOptions(ref optionsWorkList);
											}
										else // the conversion failed
											{
											objCSDbasedonCRM.LogError("Invalid format in the Document Options :. unable to generate the document.");
											//Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
											}
										}
									else  // == Null
										{
										objCSDbasedonCRM.LogError("No document options were specified - cannot generate blank documents.");
										//Console.WriteLine("No document options were selected - cannot generate blank documents.");
										}

									// The Hierarchical nodes from the Document Collection is not applicable on this Document object.
									objCSDbasedonCRM.SelectedNodes = null;

									objCSDbasedonCRM.CRM_Mapping = recDocCollsToGen.Mapping_Id;
									// add the object to the Document Collection's DocumentsWorkbooks to be generated.
									listDocumentWorkbookObjects.Add(objCSDbasedonCRM);
									break;
									}
								//=====================================================
								// CSD Document DRM Inline
								case enumDocumentTypes.CSD_Document_DRM_Inline:
									{
									CSD_Document_DRM_Inline objCSDdrmInline = new CSD_Document_DRM_Inline();
									objCSDdrmInline.DocumentCollectionID = objDocumentCollection.ID;
									objCSDdrmInline.DocumentCollectionTitle = objDocumentCollection.Title;
									objCSDdrmInline.DocumentStatus = enumDocumentStatusses.New;
									objCSDdrmInline.DocumentType = enumDocumentTypes.CSD_Document_DRM_Inline;
									objCSDdrmInline.IntroductionRichText = recDocCollsToGen.CSDDocumentIntroduction;
									objCSDdrmInline.ExecutiveSummaryRichText = recDocCollsToGen.CSDDocumentExecSummary;
									strTemplateURL = GetTheDocumentTemplate(parSDDPdatacontext, "Client Service Description");
									switch(strTemplateURL)
										{
										case "None":
											objCSDdrmInline.Template = "";
											objCSDdrmInline.LogError("The template could not be found.");
											break;
										case "Error":
											objCSDdrmInline.Template = "";
											objCSDdrmInline.LogError("Unable to access the template.");
											break;
										default:
											objCSDdrmInline.Template = Properties.AppResources.SharePointSiteURL.Substring(0, 
												Properties.AppResources.SharePointSiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									//Console.WriteLine("\t Template: {0}", objCSDdrmInline.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objCSDdrmInline.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objCSDdrmInline.HyperlinkView = true;
																		
									objCSDdrmInline.ColorCodingLayer1 = objDocumentCollection.ColourCodingLayer1;
									objCSDdrmInline.ColorCodingLayer2 = objDocumentCollection.ColourCodingLayer2;
									objCSDdrmInline.ColorCodingLayer3 = objDocumentCollection.ColourCodingLayer3;

									// Load the Presentation Layer
									objCSDdrmInline.PresentationMode = objDocumentCollection.PresentationMode;

									// Load the Document Options
									if(recDocCollsToGen.CSDDocumentDRMInlineOptions != null)
										{
										if(ConvertOptionsToList(recDocCollsToGen.CSDDocumentDRMInlineOptions, ref optionsWorkList)) // conversion is successful
											{
											objCSDdrmInline.TransposeDocumentOptions(ref optionsWorkList);
											}
										else // the conversion failed
											{
											objCSDdrmInline.LogError("Invalid format in the Document Options :. unable to generate the document.");
											//Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
											}
										}
									else  // == Null
										{
										objCSDdrmInline.LogError("No document options were specified - cannot generate blank documents.");
										//Console.WriteLine("No document options were selected - cannot generate blank documents.");
										}

									// Add the Hierarchical nodes from the Document Collection obect to the Document object.
									objCSDdrmInline.SelectedNodes = objDocumentCollection.SelectedNodes;
									// add the object to the Document Collection's DocumentsWorkbooks to be generated.
									listDocumentWorkbookObjects.Add(objCSDdrmInline);
									break;
									}
								//================================================
								// CSD Document DRM Sections
								case enumDocumentTypes.CSD_Document_DRM_Sections:
									{
									CSD_Document_DRM_Sections objCSDdrmSections = new CSD_Document_DRM_Sections();
									objCSDdrmSections.DocumentCollectionID = objDocumentCollection.ID;
									objCSDdrmSections.DocumentCollectionTitle = objDocumentCollection.Title;
									objCSDdrmSections.DocumentStatus = enumDocumentStatusses.New;
									objCSDdrmSections.DocumentType = enumDocumentTypes.CSD_Document_DRM_Sections;
									objCSDdrmSections.IntroductionRichText = recDocCollsToGen.CSDDocumentIntroduction;
									objCSDdrmSections.ExecutiveSummaryRichText = recDocCollsToGen.CSDDocumentExecSummary;
									strTemplateURL = GetTheDocumentTemplate(parSDDPdatacontext, "Client Service Description");
									switch(strTemplateURL)
										{
										case "None":
											objCSDdrmSections.Template = "";
											objCSDdrmSections.LogError("The template could not be found.");
											break;
										case "Error":
											objCSDdrmSections.Template = "";
											objCSDdrmSections.LogError("Unable to access the template.");
											break;
										default:
											objCSDdrmSections.Template = Properties.AppResources.SharePointSiteURL.Substring(0, 
												Properties.AppResources.SharePointSiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									//Console.WriteLine("\t Template: {0}", objCSDdrmSections.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objCSDdrmSections.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objCSDdrmSections.HyperlinkView = true;

									objCSDdrmSections.ColorCodingLayer1 = objDocumentCollection.ColourCodingLayer1;
									objCSDdrmSections.ColorCodingLayer2 = objDocumentCollection.ColourCodingLayer2;
									objCSDdrmSections.ColorCodingLayer3 = objDocumentCollection.ColourCodingLayer3;

									// Load the Presentation Layer
									objCSDdrmSections.PresentationMode = objDocumentCollection.PresentationMode;

									// Load the Document Options
									if(recDocCollsToGen.CSDDocumentDRMSectionsOptions != null)
										{
										if(ConvertOptionsToList(recDocCollsToGen.CSDDocumentDRMSectionsOptions, ref optionsWorkList)) // conversion is successful
											{
											objCSDdrmSections.TransposeDocumentOptions(ref optionsWorkList);
											}
										else // the conversion failed
											{
											objCSDdrmSections.LogError("Invalid format in the Document Options :. unable to generate the document.");
											//Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
											}
										}
									else  // == Null
										{
										objCSDdrmSections.LogError("No document options were specified - cannot generate blank documents.");
										//Console.WriteLine("No document options were selected - cannot generate blank documents.");
										}

									// Add the Hierarchical nodes from the Document Collection obect to the Document object.
									objCSDdrmSections.SelectedNodes = objDocumentCollection.SelectedNodes;
									// add the object to the Document Collection's DocumentsWorkbooks to be generated.
									listDocumentWorkbookObjects.Add(objCSDdrmSections);
									break;
									}
								//==============================================================
								// External Technology Coverage Dashboard.
								case enumDocumentTypes.External_Technology_Coverage_Dashboard:
									{
									External_Technology_Coverage_Dashboard_Workbook objExtTechCoverDasboard = new External_Technology_Coverage_Dashboard_Workbook();
									objExtTechCoverDasboard.DocumentCollectionID = objDocumentCollection.ID;
									objExtTechCoverDasboard.DocumentCollectionTitle = objDocumentCollection.Title;
									objExtTechCoverDasboard.DocumentStatus = enumDocumentStatusses.New;
									objExtTechCoverDasboard.DocumentType = enumDocumentTypes.External_Technology_Coverage_Dashboard;
									strTemplateURL = GetTheDocumentTemplate(parSDDPdatacontext, "Technology Roadmap Workbook");
									switch(strTemplateURL)
										{
										case "None":
											objExtTechCoverDasboard.Template = "";
											objExtTechCoverDasboard.LogError("The template could not be found.");
                                                       break;
										case "Error":
											objExtTechCoverDasboard.Template = "";
											objExtTechCoverDasboard.LogError("The template could not be accessed.");
                                                       break;
										default:
											objExtTechCoverDasboard.Template = Properties.AppResources.SharePointSiteURL.Substring(0, 
												Properties.AppResources.SharePointSiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									//Console.WriteLine("\t Template: {0}", objExtTechCoverDasboard.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objExtTechCoverDasboard.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objExtTechCoverDasboard.HyperlinkView = true;

									// Add the Hierarchical nodes from the Document Collection obect to the Document object.
									objExtTechCoverDasboard.SelectedNodes = objDocumentCollection.SelectedNodes;
									// add the object to the Document Collection's DocumentsWorkbooks to be generated.
									listDocumentWorkbookObjects.Add(objExtTechCoverDasboard);
									break;
									}
								//=================================================
								// Internal Technology Coverage Dashboard
								case enumDocumentTypes.Internal_Technology_Coverage_Dashboard:
									{
									Internal_Technology_Coverage_Dashboard_Workbook objIntTechCoverDashboard = new Internal_Technology_Coverage_Dashboard_Workbook();
									objIntTechCoverDashboard.DocumentCollectionID = objDocumentCollection.ID;
									objIntTechCoverDashboard.DocumentCollectionTitle = objDocumentCollection.Title;
									objIntTechCoverDashboard.DocumentStatus = enumDocumentStatusses.New;
									objIntTechCoverDashboard.DocumentType = enumDocumentTypes.Internal_Technology_Coverage_Dashboard;
									strTemplateURL = GetTheDocumentTemplate(parSDDPdatacontext, "Technology Roadmap Workbook");
									switch(strTemplateURL)
										{
										case "None":
											objIntTechCoverDashboard.Template = "";
											objIntTechCoverDashboard.LogError("The template could not be found.");
                                                       break;
										case "Error":
											objIntTechCoverDashboard.Template = "";
											objIntTechCoverDashboard.LogError("The template could not be accessed.");
                                                       break;
										default:
											objIntTechCoverDashboard.Template = Properties.AppResources.SharePointSiteURL.Substring(0, 
												Properties.AppResources.SharePointSiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}

									//Console.WriteLine("\t Template: {0}", objIntTechCoverDashboard.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objIntTechCoverDashboard.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objIntTechCoverDashboard.HyperlinkView = true;

									// Add the Hierarchical nodes from the Document Collection obect to the Document object.
									objIntTechCoverDashboard.SelectedNodes = objDocumentCollection.SelectedNodes;
									// add the object to the Document Collection's DocumentsWorkbooks to be generated.
									listDocumentWorkbookObjects.Add(objIntTechCoverDashboard);
									break;
									}
								//========================================================
								// ISD Document DRM Inline
								case enumDocumentTypes.ISD_Document_DRM_Inline:
									{
									ISD_Document_DRM_Inline objISDdrmInline = new ISD_Document_DRM_Inline();
									objISDdrmInline.DocumentCollectionID = objDocumentCollection.ID;
									objISDdrmInline.DocumentCollectionTitle = objDocumentCollection.Title;
									objISDdrmInline.DocumentStatus = enumDocumentStatusses.New;
									objISDdrmInline.DocumentType = enumDocumentTypes.ISD_Document_DRM_Inline;
									objISDdrmInline.IntroductionRichText = recDocCollsToGen.ISDDocumentIntroduction;
									objISDdrmInline.ExecutiveSummaryRichText = recDocCollsToGen.ISDDocumentExecSummary;
									objISDdrmInline.DocumentAcceptanceRichText = recDocCollsToGen.ISDDocumentAcceptance;
									strTemplateURL = GetTheDocumentTemplate(parSDDPdatacontext, "Internal Service Description");
									switch(strTemplateURL)
										{
										case "None":
											objISDdrmInline.Template = "";
											objISDdrmInline.LogError("The template could not be found.");
											break;
										case "Error":
											objISDdrmInline.Template = "";
											objISDdrmInline.LogError("Unable to access the template.");
											break;
										default:
											objISDdrmInline.Template = Properties.AppResources.SharePointSiteURL.Substring(0, 
												Properties.AppResources.SharePointSiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									//Console.WriteLine("\t Template: {0}", objISDdrmInline.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objISDdrmInline.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objISDdrmInline.HyperlinkView = true;

									objISDdrmInline.ColorCodingLayer1 = objDocumentCollection.ColourCodingLayer1;
									objISDdrmInline.ColorCodingLayer2 = objDocumentCollection.ColourCodingLayer2;
									objISDdrmInline.ColorCodingLayer3 = objDocumentCollection.ColourCodingLayer3;

									// Load the Presentation Layer
									objISDdrmInline.PresentationMode = objDocumentCollection.PresentationMode;

									// Load the Document Options
									if(recDocCollsToGen.ISDDocumentDRMInlineOptions != null)
										{
										if(ConvertOptionsToList(recDocCollsToGen.ISDDocumentDRMInlineOptions, ref optionsWorkList))
											{
											objISDdrmInline.TransposeDocumentOptions(ref optionsWorkList);
											}
										else // the conversion failed
											{
											objISDdrmInline.LogError("Invalid format in the Document Options :. unable to generate the document.");
											//Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
											}
										}
									else  // == Null
										{
										objISDdrmInline.LogError("No document options were specified - cannot generate blank documents.");
										//Console.WriteLine("No document options were selected - cannot generate blank documents.");
										}

									// Add the Hierarchical nodes from the Document Collection obect to the Document object.
									objISDdrmInline.SelectedNodes = objDocumentCollection.SelectedNodes;
									// add the object to the Document Collection's DocumentsWorkbooks to be generated.
									listDocumentWorkbookObjects.Add(objISDdrmInline);
									break;
									}
								//============================
								// ISD Document DRM Sections
								case enumDocumentTypes.ISD_Document_DRM_Sections:
									{
									ISD_Document_DRM_Sections objISDdrmSections = new ISD_Document_DRM_Sections();
									objISDdrmSections.DocumentCollectionID = objDocumentCollection.ID;
									objISDdrmSections.DocumentCollectionTitle = objDocumentCollection.Title;
									objISDdrmSections.DocumentStatus = enumDocumentStatusses.New;
									objISDdrmSections.DocumentType = enumDocumentTypes.ISD_Document_DRM_Sections;
									objISDdrmSections.IntroductionRichText = recDocCollsToGen.ISDDocumentIntroduction;
									objISDdrmSections.ExecutiveSummaryRichText = recDocCollsToGen.ISDDocumentExecSummary;
									objISDdrmSections.DocumentAcceptanceRichText = recDocCollsToGen.ISDDocumentAcceptance;
									strTemplateURL = GetTheDocumentTemplate(parSDDPdatacontext, "Internal Service Description");
									switch(strTemplateURL)
										{
										case "None":
											objISDdrmSections.Template = "";
											objISDdrmSections.LogError("The template could not be found.");
											break;
										case "Error":
											objISDdrmSections.Template = "";
											objISDdrmSections.LogError("Unable to access the template.");
											break;
										default:
											objISDdrmSections.Template = Properties.AppResources.SharePointSiteURL.Substring(0, 
												Properties.AppResources.SharePointSiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									//Console.WriteLine("\t Template: {0}", objISDdrmSections.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objISDdrmSections.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objISDdrmSections.HyperlinkView = true;
									
									
									objISDdrmSections.ColorCodingLayer1 = objDocumentCollection.ColourCodingLayer1;
									objISDdrmSections.ColorCodingLayer2 = objDocumentCollection.ColourCodingLayer2;
									objISDdrmSections.ColorCodingLayer3 = objDocumentCollection.ColourCodingLayer3;

									// Load the Presentation Layer
									objISDdrmSections.PresentationMode = objDocumentCollection.PresentationMode;

									// Load the Document Options
									if(recDocCollsToGen.ISDDocumentDRMSectionsOptions != null)
										{
										if(ConvertOptionsToList(recDocCollsToGen.ISDDocumentDRMSectionsOptions, ref optionsWorkList))
											{
											objISDdrmSections.TransposeDocumentOptions(ref optionsWorkList);
											}
										else // the conversion failed
											{
											objISDdrmSections.LogError("Invalid format in the Document Options :. unable to generate the document.");
											//Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
											}
										}
									else  // == Null
										{
										objISDdrmSections.LogError("No document options were specified - cannot generate blank documents.");
										//Console.WriteLine("No document options were selected - cannot generate blank documents.");
										}

									// Add the Hierarchical nodes from the Document Collection obect to the Document object.
									objISDdrmSections.SelectedNodes = objDocumentCollection.SelectedNodes;
									// add the object to the Document Collection's DocumentsWorkbooks to be generated.
									listDocumentWorkbookObjects.Add(objISDdrmSections);
									break;
									}
								//===========================
								// Pricing Addendum Document
								case enumDocumentTypes.Pricing_Addendum_Document:
									{
									//NOT_AVAILABLE: not currently implemented - Activities and Effort Drivers removed from SharePoint.
									break;
									}
								//====================================
								// RACI Matrix Workbook per Deliverable
								case enumDocumentTypes.RACI_Matrix_Workbook_per_Deliverable:
									{
									RACI_Matrix_Workbook_per_Deliverable objRACIperDeliverable = new RACI_Matrix_Workbook_per_Deliverable();
									objRACIperDeliverable.DocumentCollectionID = objDocumentCollection.ID;
									objRACIperDeliverable.DocumentCollectionTitle = objDocumentCollection.Title;
									objRACIperDeliverable.DocumentStatus = enumDocumentStatusses.New;
									objRACIperDeliverable.DocumentType = enumDocumentTypes.RACI_Matrix_Workbook_per_Deliverable;
									strTemplateURL = GetTheDocumentTemplate(parSDDPdatacontext, "RACI Matrix Workbook");
									switch(strTemplateURL)
										{
										case "None":
											objRACIperDeliverable.Template = "";
											objRACIperDeliverable.LogError("The template could not be found.");
											break;
										case "Error":
											objRACIperDeliverable.Template = "";
											objRACIperDeliverable.LogError("The template could not be accessed.");
											break;
										default:
											objRACIperDeliverable.Template = Properties.AppResources.SharePointSiteURL.Substring(0, 
												Properties.AppResources.SharePointSiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									//Console.WriteLine("\t Template: {0}", objRACIperDeliverable.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objRACIperDeliverable.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objRACIperDeliverable.HyperlinkView = true;

									// Add the Hierarchical nodes from the Document Collection obect to the Document object.
									objRACIperDeliverable.SelectedNodes = objDocumentCollection.SelectedNodes;
									// add the object to the Document Collection's DocumentsWorkbooks to be generated.
									listDocumentWorkbookObjects.Add(objRACIperDeliverable);
									//Console.WriteLine("\t {0} object added to listDocumentWorkbookObjects", objRACIperDeliverable.GetType());
									break;
									}
								//==================================================
								// RACI Workbook per Role
								case enumDocumentTypes.RACI_Workbook_per_Role:
									{
									RACI_Workbook_per_Role objRACIperRole = new RACI_Workbook_per_Role();
									objRACIperRole.DocumentCollectionID = objDocumentCollection.ID;
									objRACIperRole.DocumentCollectionTitle = objDocumentCollection.Title;
									objRACIperRole.DocumentStatus = enumDocumentStatusses.New;
									objRACIperRole.DocumentType = enumDocumentTypes.RACI_Workbook_per_Role;
									strTemplateURL = GetTheDocumentTemplate(parSDDPdatacontext, "RACI Workbook");
									switch(strTemplateURL)
										{
										case "None":
											objRACIperRole.Template = "";
											objRACIperRole.LogError(("The template could not be found."));
											break;
										case "Error":
											objRACIperRole.Template = "";
											objRACIperRole.LogError(("The template could not be accessed."));
											break;
										default:
											objRACIperRole.Template = Properties.AppResources.SharePointSiteURL.Substring(0, 
												Properties.AppResources.SharePointSiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}

									//Console.WriteLine("\t Template: {0}", objRACIperRole.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										{
										objRACIperRole.HyperlinkEdit = true;
										}
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										{
										objRACIperRole.HyperlinkView = true;
										}

									// Add the Hierarchical nodes from the Document Collection object to the Document object.
									objRACIperRole.SelectedNodes = objDocumentCollection.SelectedNodes;
									// add the object to the Document Collection's DocumentsWorkbooks to be generated.
									listDocumentWorkbookObjects.Add(objRACIperRole);
									//Console.WriteLine("\t {0} object added to listDocumentWorkbookObjects", objRACIperRole.GetType());
									break;
									}
								//=============================================================
								// Service Framework Document DRM inline
								case enumDocumentTypes.Service_Framework_Document_DRM_inline:
									{
									Services_Framework_Document_DRM_Inline objSFdrmInline = new Services_Framework_Document_DRM_Inline();
									objSFdrmInline.DocumentCollectionID = objDocumentCollection.ID;
									objSFdrmInline.DocumentCollectionTitle = objDocumentCollection.Title;
									objSFdrmInline.DocumentStatus = enumDocumentStatusses.New;
									objSFdrmInline.DocumentType = enumDocumentTypes.Service_Framework_Document_DRM_inline;
									objSFdrmInline.IntroductionRichText = recDocCollsToGen.ISDDocumentIntroduction;
									objSFdrmInline.ExecutiveSummaryRichText = recDocCollsToGen.ISDDocumentExecSummary;
									objSFdrmInline.DocumentAcceptanceRichText = recDocCollsToGen.ISDDocumentAcceptance;
									strTemplateURL = GetTheDocumentTemplate(parSDDPdatacontext, "Services Framework Description");
									switch(strTemplateURL)
										{
										case "None":
											objSFdrmInline.Template = "";
											objSFdrmInline.LogError("The template could not be found.");
											break;
										case "Error":
											objSFdrmInline.Template = "";
											objSFdrmInline.LogError("Unable to access the template.");
											break;
										default:
											objSFdrmInline.Template = Properties.AppResources.SharePointSiteURL.Substring(0, 
												Properties.AppResources.SharePointSiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									//Console.WriteLine("\t Template: {0}", objSFdrmInline.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objSFdrmInline.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objSFdrmInline.HyperlinkView = true;

									objSFdrmInline.ColorCodingLayer1 = objDocumentCollection.ColourCodingLayer1;
									objSFdrmInline.ColorCodingLayer2 = objDocumentCollection.ColourCodingLayer2;
									objSFdrmInline.ColorCodingLayer3 = objDocumentCollection.ColourCodingLayer3;

									// Load the Presentation Layer
									objSFdrmInline.PresentationMode = objDocumentCollection.PresentationMode;

									// Load the Document Options
									if(recDocCollsToGen.ISDDocumentDRMInlineOptions != null)
										{
										if(ConvertOptionsToList(recDocCollsToGen.ISDDocumentDRMInlineOptions, ref optionsWorkList)) // conversion is successful
											{
											objSFdrmInline.TransposeDocumentOptions(ref optionsWorkList);
											}
										else // the conversion failed
											{
											objSFdrmInline.LogError("Invalid format in the Document Options :. unable to generate the document.");
											//Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
											}
										}
									else  // == Null
										{
										objSFdrmInline.LogError("No document options were specified - cannot generate blank documents.");
										//Console.WriteLine("No document options were selected - cannot generate blank documents.");
										}
									// Add the Hierarchical nodes from the Document Collection obect to the Document object.
									objSFdrmInline.SelectedNodes = objDocumentCollection.SelectedNodes;
									//Console.WriteLine("\t {0} object added to listDocumentWorkbookObjects", objSFdrmInline.ToString());
									listDocumentWorkbookObjects.Add(objSFdrmInline);
									break;
									}
								//=====================================================
								// Service Framework Document DRM sections
								case enumDocumentTypes.Service_Framework_Document_DRM_sections:
									{
									Services_Framework_Document_DRM_Sections objSFdrmSections = new Services_Framework_Document_DRM_Sections();
									objSFdrmSections.DocumentCollectionID = objDocumentCollection.ID;
									objSFdrmSections.DocumentCollectionTitle = objDocumentCollection.Title;
									objSFdrmSections.DocumentStatus = enumDocumentStatusses.New;
									objSFdrmSections.DocumentType = enumDocumentTypes.Service_Framework_Document_DRM_sections;
									objSFdrmSections.IntroductionRichText = recDocCollsToGen.ISDDocumentIntroduction;
									objSFdrmSections.ExecutiveSummaryRichText = recDocCollsToGen.ISDDocumentExecSummary;
									objSFdrmSections.DocumentAcceptanceRichText = recDocCollsToGen.ISDDocumentAcceptance;
									strTemplateURL = GetTheDocumentTemplate(parSDDPdatacontext, "Services Framework Description");
									switch(strTemplateURL)
										{
										case "None":
											objSFdrmSections.Template = "";
											objSFdrmSections.LogError("The template could not be found.");
											break;
										case "Error":
											objSFdrmSections.Template = "";
											objSFdrmSections.LogError("Unable to access the template.");
											break;
										default:
											objSFdrmSections.Template = Properties.AppResources.SharePointURL + strTemplateURL;
											break;
										}
									//Console.WriteLine("\t Template: {0}", objSFdrmSections.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objSFdrmSections.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objSFdrmSections.HyperlinkView = true;

									objSFdrmSections.ColorCodingLayer1 = objDocumentCollection.ColourCodingLayer1;
									objSFdrmSections.ColorCodingLayer2 = objDocumentCollection.ColourCodingLayer2;
									objSFdrmSections.ColorCodingLayer3 = objDocumentCollection.ColourCodingLayer3;

									// Load the Presentation Layer
									objSFdrmSections.PresentationMode = objDocumentCollection.PresentationMode;

									// Load the Document Options
									if(recDocCollsToGen.ISDDocumentDRMSectionsOptions != null)
										{
										if(ConvertOptionsToList(recDocCollsToGen.ISDDocumentDRMSectionsOptions, ref optionsWorkList))
											{
											objSFdrmSections.TransposeDocumentOptions(ref optionsWorkList);
											}
										else
											{
											objSFdrmSections.LogError("Invalid format in the Document Options :. unable to generate the document.");
											//Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
											}
										} // !=Null
									else
										{
										objSFdrmSections.LogError("No document options were specified - cannot generate a blank document.");
										//Console.WriteLine("No document options were selected - cannot generate blank documents.");
										}

									// Add the Hierarchical nodes from the Document Collection obect to the Document object.
									objSFdrmSections.SelectedNodes = objDocumentCollection.SelectedNodes;
									// add the object to the Document Collection's DocumentsWorkbooks to be generated.
									listDocumentWorkbookObjects.Add(objSFdrmSections);
									//Console.WriteLine("\t {0} object added to listDocumentWorkbookObjects", objSFdrmSections.GetType());
									break;
									}
								default:
									{
									break;
									}
								} // End Switch
							} // end ForEach loop
						
						// assign the list of DocumentWorkbooks to the collection of Documents_and_Workbooks of the DocumentCollection
						objDocumentCollection.Document_and_Workbook_objects = listDocumentWorkbookObjects;
                              }
					// Add the instance of the Document Collection Object to the List of Document Collection that must be generated
					parCollectionsToGenerate.Add(objDocumentCollection);
					Console.WriteLine(" Document Collection: {0} successfully loaded..\n Now there are {1} collections to generate.\n", recDocCollsToGen.Id, parCollectionsToGenerate.Count);
					} // Loop of the For Each DocColsToGenerate
				Console.WriteLine("All entries processed and added to List parCollectionsToGenerate) - {0} collections to generate...", parCollectionsToGenerate.Count);
				return "Good";
				} // end of Try
			catch(DataServiceClientException exc)
				{
				Console.Beep(2500, 750);
				Console.WriteLine("\n*** Exception ERROR ***\nHResult: {0}\nMessage: {1}\nStatusCode:{2}\nInnerException: {3}\nTargetSite: {4}\nStackTrace: {5}.", 
					exc.HResult, exc.Message, exc.StatusCode, exc.InnerException, exc.TargetSite, exc.StackTrace);
				return "Error: Cannot access site: " + Properties.AppResources.SharePointSiteURL + " Ensure the computer/server is connected to the Dimension Data Domain network";
				}
			catch(DataServiceQueryException exc)
				{
				Console.Beep(2500,750);
				Console.WriteLine("\n*** Exception ERROR ***\nHResult: {0}\nMessage: {1}\nResponse:{2}\nInnerException: {3}\nTargetSite: {4}\nStackTrace: {5}.",
					exc.HResult, exc.Message, exc.Response, exc.InnerException, exc.TargetSite, exc.StackTrace);
				return "Error: Cannot access SharePoint site: " + Properties.AppResources.SharePointSiteURL + " Ensure the computer/server is connected to the Dimension Data Domain network";
				}
			catch(DataServiceRequestException exc)
				{
				Console.WriteLine("*** Exception ERROR **** Accessing SharePoint Document Collection Library List\n"
					+ "DataServiceRequestException - HResult: {0}\nMessage: {1}\nTargetSite: {2}\nStackTrace: {3}",
					exc.HResult, exc.Message, exc.TargetSite, exc.StackTrace);
				return "Error: DataServiceRequestException: " + exc.Message;
				}
			catch(DataServiceTransportException exc)
				{
				Console.WriteLine("*** Exception ERROR **** Accessing SharePoint Document Collection Library List\n"
					+ "DataServiceTransportException - HResult: {0}\nMessage: {1}\nTargetSite: {2}\nStackTrace: {3}",
					exc.HResult, exc.Message, exc.TargetSite, exc.StackTrace);
				return "Error: DataServiceRequestException: " + exc.Message;
				}
			catch(Exception ex)
				{
				Console.Beep(2500, 750);
				Console.WriteLine("\n\nException: [{0}] occurred and was caught. \n{1}", ex.HResult.ToString(), ex.Message);

				if(ex.HResult == -2146330330)
					return "Error: Cannot access site: " + Properties.AppResources.SharePointSiteURL + " Ensure the computer is connected to the Dimension Data Domain network";
				else if(ex.HResult == -2146233033)
					return "Error: Input string missing to connect to " + Properties.AppResources.SharePointSiteURL + " Ensure the computer is connected to the Dimension Data Domain network";
				else
					return "Error: Unexpected error occurred. " + ex.HResult + " - " + ex.Message;
				}
			} // end of Method
				
		/// <summary>
		/// This method finds the relevant Document Template and if found returns the path to the template URL in a string 
		/// </summary>
		/// <param name="parDataContext">Pass the DataContext for the Template SharePoint Site.</param>
		/// <param name="parTemplateType">Pass the Template Type that need to be found as a string.</param>
		/// <returns></returns>
		public static string GetTheDocumentTemplate(DesignAndDeliveryPortfolioDataContext parDataContext, string parTemplateType)
			{
			string returnPath = "";
			try
				{
				//var DocumentTemplates = parDataContext.DocumentTemplates;
				var rsTemplate = from docTemplate in parDataContext.DocumentTemplates
							where docTemplate.TemplateTypeValue == parTemplateType
							select docTemplate;
				Console.WriteLine("\t\t\t + {0} templates found.", rsTemplate.Count());
				if(rsTemplate != null)
					{
					foreach(var tpl in rsTemplate)
						{
						//Console.WriteLine("\t\t\t - {0} - {1} [{2}]", tpl.Id, tpl.Title, tpl.TemplateTypeValue);
						if(tpl.TemplateTypeValue == parTemplateType)
							{
							returnPath =  tpl.Path + "/" + tpl.Name;
							break;
							}
						}
					}
				else // No template was found
					{
					returnPath = "None";
					}
				}
			catch(Exception ex)
				{
				Console.WriteLine("Error occurred: {0} \n {1}", ex.Message, ex.Data);
				returnPath = "Error";
				}
			
			return returnPath;

			}
		/// <summary>
		/// This method convers a comma delimited sting of numbers into a List of intergers for further processing later in the generation process.
		/// </summary>
		/// <param name="parStringOptions">The Document Options in string format.</param>
		/// <param name="parListOfOptions">a Reference to List<int> list which is returned</param>
		/// <returns></returns>
		public static bool ConvertOptionsToList(string parStringOptions, ref List<int> parListOfOptions)
			{
			int position = 0;
			int value = 0;
			int errors = 0;
			// Clear the parListOfOptions if it is not empty
			if(parListOfOptions.Count > 0)
				{
				parListOfOptions.RemoveRange(0, parListOfOptions.Count);
				}
			// read through the parStringOptions and load each of the values into parListOptions
			do
				{
				try
					{
					if(parStringOptions.IndexOf(",", position) < 1)  // there are no entries or the last entry was reached
						{
						if(position > 0) // entries are alreay processed, therefore it is probably the last entry...
							{
							if(int.TryParse(parStringOptions.Substring(position, (parStringOptions.Length - position)), out value))
								{
								//Console.WriteLine("\t\t + OptionID: {0}", parStringOptions.Substring(position, (parStringOptions.Length - position)));
								parListOfOptions.Add(value);
								position = parStringOptions.Length;
								}
                                   }
						}
					else // there are entries in the list :. process the next one...
						{
						if(int.TryParse(parStringOptions.Substring(position, (parStringOptions.IndexOf(",", position) - position)), out value))
							{
							//Console.WriteLine("\t\t + OptionID: {0}", parStringOptions.Substring(position, (parStringOptions.IndexOf(",", position) - position)));
							parListOfOptions.Add(value);
							position = parStringOptions.IndexOf(",", position) + 1;
							}
						else // unable to parse the string value to int32
							{
							Console.WriteLine("Option value is not numeric at position {0} in {1}.", position, parStringOptions);
							errors += 1;
							}
						}
					}
				catch (Exception exc)
					{
					if(!int.TryParse(parStringOptions.Substring(position, (parStringOptions.IndexOf(",", position) - position)), out value))
						{
						//Console.WriteLine("Option value is not numeric at position {0} in {1}.", position, parStringOptions);
						errors += 1;
						}
					Console.WriteLine("\n\nException Error: {0} - {1}", exc.HResult, exc.Message);
					}
				}
			while(position < parStringOptions.Length);
			if(errors > 0)
				return false;
			else
				return true;
			}

		/// <summary>
		/// This method converts a string value to the actual enumerator value.
		/// </summary>
		/// <param name="parStringValue">A string value that must be convered to an Enumerator, Must NOT be Null</param>
		/// <param name="parOutputEnumValue">returns (Output) an type object which is the actual converted enumerator value.</param>
		/// <returns>bool is returned as True if the conversion is successfull, else it returns False</returns>
		public static bool PrepareStringForEnum (string parStringValue, out String parOutputEnumValue)
			{
			if(parStringValue == null)
				{
				parOutputEnumValue = null;
				return false;
				}
			if(parStringValue.Length == 0)
				{
				parOutputEnumValue = null;
				return false;
				}
			// only pass this point if the first 2 parameters are not null
			// remove spaces and ( ), [ ], { },from the parStringValue
			string strValue = parStringValue.Replace(" ", "_");
			strValue = strValue.Replace("(", "");
			strValue = strValue.Replace(")", "");
			strValue = strValue.Replace("[", "");
			strValue = strValue.Replace("]", "");
			strValue = strValue.Replace("{", "");
			strValue = strValue.Replace("}", "");
			strValue = strValue.Replace(",", "");
			strValue = strValue.Replace(".", "");
			strValue = strValue.Replace("/", "");
			strValue = strValue.Replace("'", "");
               strValue = strValue.Replace("|", "");
			strValue = strValue.Replace("\\", "");
			strValue = strValue.Trim();

			if(parStringValue.Length == 0)
				{
				parOutputEnumValue = null;
				return false;
				}
			else
				{
				parOutputEnumValue = strValue;
				return true;
				}
			}
		}
	}