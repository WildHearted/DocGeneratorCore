using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Services.Client;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint;
using DogGenUI.SDDPServiceReference;

namespace DogGenUI
	{
	/// <summary>
	///	Mapped to the [Content Layer Colour Coding Option] column in SharePoint List
	/// </summary>
	enum enumContent_Layer_Colour_Coding_Options
		{
		Colour_Code_Layer_1=1,
		Colour_Code_Layer_2=2,
		Colour_Code_Layer_3=3
		}
	/// <summary>
	///	Mapped to the [Generate Action] column in SharePoint List
	/// </summary>
	enum enumGenerate_Actions
		{
		Save_but_dont_generate_the_documents_yet=1,
		Submit_to_the_generate_queue=2,
		Schedule_for_a_specific_date_and_time=3
		}
	/// <summary>
	/// Mapped to the [Generate Schedule Option] column in SharePoint
	/// </summary>
	enum enumGenerateScheduleOptions
		{
		Do_NOT_Repeat=0,
		Repeat_every=1
		}
	/// <summary>
	/// Mapped to the values of the [Generate Repeat Interval] column in SharePoint;
	/// </summary>
	enum enumGenerateRepeatIntervals
		{
		Day=1,
		Week=2,
		Month=3
		}
	/// <summary>
	/// Mapped to the values of the [Hyperlink Options] column in SharePoint;
	/// </summary>
	enum enumHyperlinkOptions
		{
		Do_NOT_Include_Hyperlinks=0,
		Include_EDIT_Hyperlinks=1,
		Include_VIEW_Hyperlinks=2
		}

	/// <summary>
	/// This list contains the documents that the user selected which needs to be generated.
	/// </summary>
	class DocumentCollection
		{
		
		// Object Properties
		private int _id = 0;
		public int ID
			{
			get
				{
				return _id;
				}
			private set
				{
				_id = value;
				}
			}
		private string _clientName;
		public string ClientName
			{
			get
				{
				return _clientName;
				}
			private set
				{
				_clientName = value;
				}
			}
		private string _title;
		public string Title
			{
			get
				{
				return _title;
				}
			private set
				{
				_title = value;
				}
			}
		private bool _colourCodingLayer1 = false;
		public bool ColourCodingLayer1
			{
			get
				{
				return this._colourCodingLayer1;
				}
			private set
				{
				this._colourCodingLayer1 = value;
				}
			}
		private bool _colourCodingLayer2 = false;
		public bool ColourCodingLayer2
			{
			get
				{
				return this._colourCodingLayer2;
				}
			private set
				{
				this._colourCodingLayer2 = value;
				}
			}
		private bool _colourCodingLayer3 = false;
		public bool ColourCodingLayer3
			{
			get
				{
				return this._colourCodingLayer3;
				}
			private set
				{
				this._colourCodingLayer3 = value;
				}
			}
		private List<enumDocumentTypes> _documentsToGenerate;
		public List<enumDocumentTypes> DocumentsToGenerate
			{
			get
				{
				return this._documentsToGenerate;
				}
			private set
				{
				this._documentsToGenerate = value;
				}
			}
		private bool _notifyMe;
		public bool NotifyMe
			{
			get
				{
				return this._notifyMe;
				}
			private set
				{
				this._notifyMe = value;
				}
			}
		private string _notificationEmail;
		public string NotificationEmail
			{
			get
				{
				return this._notificationEmail;
				}
			private set
				{
				this._notificationEmail = value;
				}
			}
		private enumGenerateScheduleOptions _generateScheduleOption;
		public enumGenerateScheduleOptions GenerateScheduleOption
			{
			get
				{
				return this._generateScheduleOption;
				}
			private set
				{
				this._generateScheduleOption = value;
				}
			}
		private DateTime _generateOnDateTime;
		public DateTime GenerateOnDateTime
			{
			get
				{
				return this._generateOnDateTime;
				}
			private set
				{
				this._generateOnDateTime = value;
				}
			}
		private enumGenerateRepeatIntervals _generateRepeatInterval;
		public enumGenerateRepeatIntervals GenerateRepeatInterval
			{
			get
				{
				return this._generateRepeatInterval;
				}
			private set
				{
				this._generateRepeatInterval = value;
				}
			}
		private int _GenerateRepeatIntervalValue;
		public int GenerateRepeatIntervalValue
			{
			get
				{
				return this._GenerateRepeatIntervalValue;
				}
			private set
				{
				this._GenerateRepeatIntervalValue = value;
				}
			}
		private enumHyperlinkOptions _hyperlinkOption;
		public enumHyperlinkOptions HyperLinkOption
			{
			get
				{
				return this._hyperlinkOption;
				}
			private set
				{
				this._hyperlinkOption = value;
				}
			}
		private int _mapping;
		public int Mapping
			{
			get
				{
				return this._mapping;
				}
			private set
				{
				this._mapping = value;
				}
			}
		private int _pricingWorkbook;
		public int PricingWorkbook
			{
			get
				{
				return this._pricingWorkbook;
				}
			private set
				{
				this._pricingWorkbook = value;
				}
			}
		private List<Hierarchy> _selectedNodes;
		public List<Hierarchy> SelectedNodes
			{
			get
				{
				return this._selectedNodes;
				}
			private set
				{
				this._selectedNodes = value;
				}
			}
		private List<Document_Workbook> _documents_and_Workbooks;
		public List<Document_Workbook> Documents_and_Workbooks
			{
			get
				{
				return _documents_and_Workbooks;
				}
			set
				{
				_documents_and_Workbooks = value;
				}
			}

		// Other Fields

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
		public static string GetCollectionsToGenerate(ref List<DocumentCollection> parCollectionsToGenerate)
			{
			List<int> optionsWorkList = new List<int>();
			string enumWorkString;
			string websiteURL = "https://teams.dimensiondata.com/sites/ServiceCatalogue";
			DesignAndDeliveryPortfolioDataContext datacontexSDDP = new DesignAndDeliveryPortfolioDataContext(new Uri(websiteURL + "/_vti_bin/listdata.svc"));
			datacontexSDDP.Credentials = CredentialCache.DefaultCredentials;
			//datacontexSDDP.MergeOption = MergeOption.AppendOnly;			//Use only if data is added
			//datacontexSDDP.MergeOption = MergeOption.OverwriteChanges;	//use when data is updated
			datacontexSDDP.MergeOption = MergeOption.NoTracking;
			try
				{
				var DocCollectionLib = datacontexSDDP.DocumentCollectionLibrary
						.Expand(p => p.Client_)
						.Expand(p => p.ContentLayerColourCodingOption)
						.Expand(p => p.GenerateFrameworkDocuments)
						.Expand(p => p.GenerateInternalDocuments)
						.Expand(p => p.GenerateExternalDocuments)
						.Expand(p => p.GenerateRepeatInterval)
						.Expand(p => p.HyperlinkOptions);

				var DocColsToGenerate = from dc in DocCollectionLib where dc.GenerateActionValue != null orderby dc.Id select dc;	
				// var DocColsToGenerate = from dc in DocCollectionLib orderby dc.Id select dc;

				Console.WriteLine("There are {0} Document Collections to create...", DocColsToGenerate.Count());

				foreach(var DCsToGen in DocColsToGenerate)
					{
					if(DCsToGen.GenerateActionValue.Substring(0, 4) == "Save")
						{
						Console.WriteLine("{0} Generate  Action value is {1}, therefore it will not be generated.", DCsToGen.Id, DCsToGen.GenerateActionValue);
						continue;
						}
					Console.WriteLine("\nID: {0}  Title: {1}\n\t DocGen Client Name: [{2}] - Client Title:[{3}] ", DCsToGen.Id, DCsToGen.Title, DCsToGen.Client_.DocGenClientName, DCsToGen.Client_.Title);

					// Create a new Instance for the Document Collection into which the object properties are loaded
					DocumentCollection objDocumentCollection = new DocumentCollection();
					//Set the basic object properties
					objDocumentCollection.ID = DCsToGen.Id;
					Console.WriteLine("\t ID: {0} ", objDocumentCollection.ID);

					if(DCsToGen.Client_.DocGenClientName == null)
						objDocumentCollection.ClientName = "the Client";
					else
						objDocumentCollection.ClientName = DCsToGen.Client_.DocGenClientName;
					Console.WriteLine("\t ClientName: {0} ", objDocumentCollection.ClientName);

					if(DCsToGen.Title == null)
						objDocumentCollection.Title = "Collection Title for entry " + DCsToGen.Id;
					else
						objDocumentCollection.Title = DCsToGen.Title;

					Console.WriteLine("\t Title: {0}", objDocumentCollection.Title);
					if(DCsToGen.GenerateNotifyMe == null)
						objDocumentCollection.NotifyMe = false;
					else
						objDocumentCollection.NotifyMe = DCsToGen.GenerateNotifyMe.Value;
					Console.WriteLine("\t NotifyMe: {0} ", objDocumentCollection.NotifyMe);

					if(DCsToGen.GenerateNotificationEMail == null)
						objDocumentCollection.NotificationEmail = "None";
					else
						objDocumentCollection.NotificationEmail = DCsToGen.GenerateNotificationEMail;
					Console.WriteLine("\t NotificationEmail: {0} ", objDocumentCollection.NotificationEmail);

					if(DCsToGen.GenerateOnDateTime == null)
						objDocumentCollection.GenerateOnDateTime = DateTime.Now;
					else
						objDocumentCollection.GenerateOnDateTime = DCsToGen.GenerateOnDateTime.Value;
					Console.WriteLine("\t GenerateOnDateTime: {0} ", objDocumentCollection.GenerateOnDateTime);

					// Set the Mapping value
					if(DCsToGen.Mapping_Id != null)
						{
						try
							{
							objDocumentCollection.Mapping = Convert.ToInt32(DCsToGen.Mapping_Id);
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
					Console.WriteLine("\t Mapping: {0} ", objDocumentCollection.Mapping);

					// Set the Pricing Workbook
					if(DCsToGen.PricingWorkbookId != null)
						try
							{
							objDocumentCollection.PricingWorkbook = Convert.ToInt32(DCsToGen.PricingWorkbookId);
							}
						catch(OverflowException ex)
							{
							Console.WriteLine("Overflow Exception occurred when converting the Pricing Workbook value to a Integer.\n Error Description: {0}", ex.Message);
							objDocumentCollection.Mapping = 0;
							}
					else
						objDocumentCollection.PricingWorkbook = 0;
					Console.WriteLine("\t PricingWorkbook: {0} ", objDocumentCollection.PricingWorkbook);

					// Set the Generate Schedule Options
					enumGenerateScheduleOptions generateSchdlOption;
					if(DCsToGen.GenerateScheduleOptionValue != null)
						{
						if(PrepareStringForEnum(DCsToGen.GenerateScheduleOptionValue, out enumWorkString))
							{
							if(Enum.TryParse<enumGenerateScheduleOptions>(enumWorkString, out generateSchdlOption))
								{
								objDocumentCollection.GenerateScheduleOption = generateSchdlOption;
								}
							else
								{
								objDocumentCollection.GenerateScheduleOption = enumGenerateScheduleOptions.Do_NOT_Repeat;
								}
							}
						else
							{
							objDocumentCollection.GenerateScheduleOption = enumGenerateScheduleOptions.Do_NOT_Repeat;
							}
						}
					else
						{
						objDocumentCollection.GenerateScheduleOption = enumGenerateScheduleOptions.Do_NOT_Repeat;
						}
					Console.WriteLine("\t Generate ScheduleOption: {0} ", objDocumentCollection.GenerateScheduleOption);

					// Set the Generate Repeat Intervals
					enumGenerateRepeatIntervals generateRepeatIntrvl;
					if(DCsToGen.GenerateRepeatIntervalValue0 != null)
						{
						if(PrepareStringForEnum(DCsToGen.GenerateRepeatIntervalValue0, out enumWorkString))
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

					// Set the Generate Repeat Interval Value
					if(DCsToGen.GenerateRepeatIntervalValue != null)
						{
						try
							{
							objDocumentCollection.GenerateRepeatIntervalValue = Convert.ToInt32(DCsToGen.GenerateRepeatIntervalValue.Value);
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
					if(DCsToGen.HyperlinkOptionsValue != null)
						{
						enumHyperlinkOptions hyperLnkOption;
						if(PrepareStringForEnum(DCsToGen.HyperlinkOptionsValue, out enumWorkString))
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
					if(DCsToGen.ContentLayerColourCodingOption.Count > 0)
						{
						foreach(var entry in DCsToGen.ContentLayerColourCodingOption)
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

					// Set the Framework Documents that must be generated
					int noOfDocsToGenerateInCollection = 0;
					List<enumDocumentTypes> listOfDocumentsToGenerate = new List<enumDocumentTypes>();
					enumDocumentTypes docType;
					Console.WriteLine("\t Generate Framework Documents: {0} entries.", DCsToGen.GenerateFrameworkDocuments.Count.ToString());
					if(DCsToGen.GenerateFrameworkDocuments.Count > 0)
						{
						foreach(var entry in DCsToGen.GenerateFrameworkDocuments)
							{
							if(PrepareStringForEnum(entry.Value, out enumWorkString))
								{
								if(Enum.TryParse<enumDocumentTypes>(enumWorkString, out docType))
									{
									listOfDocumentsToGenerate.Add(docType);
									Console.WriteLine("\t\t + [{0}]", docType);
									noOfDocsToGenerateInCollection += 1;
									}
								else
									if(Enum.TryParse<enumDocumentTypes>(enumWorkString, out docType))
									listOfDocumentsToGenerate.Add(docType);
								else
									Console.WriteLine("\t\t [{0}] Not found as enumeration [{1}]", enumWorkString, docType);
								}
							}
						}
					// Set the Internal Documents that must be generated
					Console.WriteLine("\t Generate Internal Documents: {0} entries.", DCsToGen.GenerateInternalDocuments.Count.ToString());
					if(DCsToGen.GenerateInternalDocuments.Count > 0)
						{
						foreach(var entry in DCsToGen.GenerateInternalDocuments)
							{
							if(PrepareStringForEnum(entry.Value, out enumWorkString))
								{
								if(Enum.TryParse<enumDocumentTypes>(enumWorkString, out docType))
									{
									listOfDocumentsToGenerate.Add(docType);
									Console.WriteLine("\t\t + [{0}]", docType);
									noOfDocsToGenerateInCollection += 1;
									}
								}
							}
						}
					// Set the External Documents that must be generated
					Console.WriteLine("\t Generate External Documents: {0} entries.", DCsToGen.GenerateExternalDocuments.Count.ToString());
					if(DCsToGen.GenerateExternalDocuments.Count > 0)
						{
						foreach(var entry in DCsToGen.GenerateExternalDocuments)
							{
							if(PrepareStringForEnum(entry.Value, out enumWorkString))
								{
								if(Enum.TryParse<enumDocumentTypes>(enumWorkString, out docType))
									{
									listOfDocumentsToGenerate.Add(docType);
									Console.WriteLine("\t\t + [{0}]", docType);
									}
								}
							}
						}
					objDocumentCollection.DocumentsToGenerate = listOfDocumentsToGenerate;
					Console.WriteLine("\t {0} document to be generated for the Document Collection.", objDocumentCollection.DocumentsToGenerate.Count);

					//Set the Selected Nodes that need to be generated by building a hierchical List with Hierarchy objects
					Console.WriteLine("\t Loading the Nodes that the user selected.");
					if(DCsToGen.SelectedNodes != null)
						{
						List<Hierarchy> listOfNodesToGenerate = new List<Hierarchy>();
						if(Hierarchy.ConstructHierarchy(DCsToGen.SelectedNodes, ref listOfNodesToGenerate))
							{
							objDocumentCollection.SelectedNodes = listOfNodesToGenerate;
							Console.WriteLine("\t {0} nodes successfully loaded by ConstructHierarchy method.", listOfNodesToGenerate.Count);
							}
						else //there was an error during the Construct of the Hierarchy method
							{
							Console.WriteLine("An error occurred when the Hierarchy was constructed.");
							}
						}
					else
						{
						Console.WriteLine("There are no selected content to generate for Document Collection {0} - {1}", DCsToGen.Id, DCsToGen.Title);
						}

					// Load options for each of the documents that need to be generated
					Console.WriteLine("\t Creating the Document object(s) for {0} document.", objDocumentCollection.DocumentsToGenerate.Count);
					
					if(objDocumentCollection.DocumentsToGenerate.Count > 0)
						{
						string strTemplateURL = ""; // variable used to store the individual Template URLs
						// Declare a new List of Document_and_Workbook objects that can hold all the object entries
						List<Document_Workbook> listDocumentsWorkbooks = new List<Document_Workbook>();
						
						foreach(enumDocumentTypes objDocsToGenerate in objDocumentCollection.DocumentsToGenerate)
							{
							Console.WriteLine("\t\t Busy constructing Document object for {0}...", objDocsToGenerate.ToString());
							switch(objDocsToGenerate)
								{
								case enumDocumentTypes.Activity_Effort_Workbook:
									{
									// ignore, not implemented now
									break;
									}
								case enumDocumentTypes.Client_Requirement_Mapping_Workbook:
									{
									Client_Requirements_Mapping_Workbook objClinetRequirementsMappingWorkbook = new Client_Requirements_Mapping_Workbook();
									objClinetRequirementsMappingWorkbook.DocumentCollectionID = objDocumentCollection.ID;
									objClinetRequirementsMappingWorkbook.DocumentStatus = enumDocumentStatusses.New;
									objClinetRequirementsMappingWorkbook.DocumentType = enumDocumentTypes.Client_Requirement_Mapping_Workbook;
									strTemplateURL = GetTheDocumentTemplate(datacontexSDDP, "Activity Effort Workbook");
                                             switch (strTemplateURL)
										{
										case "None":
											objClinetRequirementsMappingWorkbook.Template = "";
											objClinetRequirementsMappingWorkbook.LogError("The template could not be found.");
                                                       break;
										case "Error":
											objClinetRequirementsMappingWorkbook.Template = "";
											objClinetRequirementsMappingWorkbook.LogError("The template could not be accessed.");
                                                       break;
										default:
											objClinetRequirementsMappingWorkbook.Template = strTemplateURL;
											break;
										}
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objClinetRequirementsMappingWorkbook.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objClinetRequirementsMappingWorkbook.Hyperlink_View = true;
									// add to the list of DocumentOptions
									listDocumentsWorkbooks.Add(objClinetRequirementsMappingWorkbook);
									break;
									}
								case enumDocumentTypes.Content_Status_Workbook:
									{
									Content_Status_Workbook objContentStatus_Workbook = new Content_Status_Workbook();
									objContentStatus_Workbook.DocumentCollectionID = objDocumentCollection.ID;
									objContentStatus_Workbook.DocumentStatus = enumDocumentStatusses.New;
									objContentStatus_Workbook.DocumentType = enumDocumentTypes.Content_Status_Workbook;
									strTemplateURL = GetTheDocumentTemplate(datacontexSDDP, "Content Status Workbook");
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
											objContentStatus_Workbook.Template = websiteURL.Substring(0, websiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									Console.WriteLine("\t Template: {0}", objContentStatus_Workbook.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objContentStatus_Workbook.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objContentStatus_Workbook.Hyperlink_View = true;
									// add to the list of DocumentOptions
									listDocumentsWorkbooks.Add(objContentStatus_Workbook);
									break;
									}
								case enumDocumentTypes.Contract_SoW_Service_Description:
									{
									Contract_SoW_Service_Description objContractSoWServiceDescriptionDoc = new Contract_SoW_Service_Description();
									objContractSoWServiceDescriptionDoc.DocumentCollectionID = objDocumentCollection.ID;
									objContractSoWServiceDescriptionDoc.DocumentStatus = enumDocumentStatusses.New;
									objContractSoWServiceDescriptionDoc.DocumentType = enumDocumentTypes.ISD_Document_DRM_Sections;
									strTemplateURL = GetTheDocumentTemplate(datacontexSDDP, "Contract: Service Description (Appendix F)");
									switch(strTemplateURL)
										{
										case "None":
											objContractSoWServiceDescriptionDoc.Template = "";
											objContractSoWServiceDescriptionDoc.LogError("The template could not be found.");
											break;
										case "Error":
											objContractSoWServiceDescriptionDoc.Template = "";
											objContractSoWServiceDescriptionDoc.LogError("Unable to access the template.");
											break;
										default:
											objContractSoWServiceDescriptionDoc.Template = websiteURL.Substring(0, websiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									Console.WriteLine("\t Template: {0}", objContractSoWServiceDescriptionDoc.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objContractSoWServiceDescriptionDoc.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objContractSoWServiceDescriptionDoc.Hyperlink_View = true;
									// Load the Document Options
									if(DCsToGen.SoWSDOptions != null)
										{
										if(ConvertOptionsToList(DCsToGen.SoWSDOptions, ref optionsWorkList)) // conversion is successful
											{
											objContractSoWServiceDescriptionDoc.TransposeDocumentOptions(ref optionsWorkList);
											}
										else // the conversion failed
											{
											objContractSoWServiceDescriptionDoc.LogError("Invalid format in the Document Options :. unable to generate the document.");
											Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
											}
										}
									else  // == Null
										{
										objContractSoWServiceDescriptionDoc.LogError("No document options were specified - cannot generate blank documents.");
										Console.WriteLine("No document options were selected - cannot generate blank documents.");
										}
									// add to the list of DocumentToBeGenerated
									listDocumentsWorkbooks.Add(objContractSoWServiceDescriptionDoc);
									break;
									}
								case enumDocumentTypes.CSD_based_on_Client_Requirements_Mapping:
									{
									CSD_based_on_ClientRequirementsMapping objCSDbasedonCRMDoc = new CSD_based_on_ClientRequirementsMapping();
									objCSDbasedonCRMDoc.DocumentCollectionID = objDocumentCollection.ID;
									objCSDbasedonCRMDoc.DocumentStatus = enumDocumentStatusses.New;
									objCSDbasedonCRMDoc.DocumentType = enumDocumentTypes.ISD_Document_DRM_Sections;
									strTemplateURL = GetTheDocumentTemplate(datacontexSDDP, "Client Service Description");
									switch(strTemplateURL)
										{
										case "None":
											objCSDbasedonCRMDoc.Template = "";
											objCSDbasedonCRMDoc.LogError("The template could not be found.");
											break;
										case "Error":
											objCSDbasedonCRMDoc.Template = "";
											objCSDbasedonCRMDoc.LogError("Unable to access the template.");
											break;
										default:
											objCSDbasedonCRMDoc.Template = websiteURL.Substring(0, websiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									Console.WriteLine("\t Template: {0}", objCSDbasedonCRMDoc.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objCSDbasedonCRMDoc.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objCSDbasedonCRMDoc.Hyperlink_View = true;
									// Load the Document Options
									if(DCsToGen.CSDDocumentBasedOnCRMOptions != null)
										{
										if(ConvertOptionsToList(DCsToGen.CSDDocumentBasedOnCRMOptions, ref optionsWorkList)) // conversion is successful
											{
											objCSDbasedonCRMDoc.TransposeDocumentOptions(ref optionsWorkList);
											}
										else // the conversion failed
											{
											objCSDbasedonCRMDoc.LogError("Invalid format in the Document Options :. unable to generate the document.");
											Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
											}
										}
									else  // == Null
										{
										objCSDbasedonCRMDoc.LogError("No document options were specified - cannot generate blank documents.");
										Console.WriteLine("No document options were selected - cannot generate blank documents.");
										}
									// add to the list of DocumentToBeGenerated
									listDocumentsWorkbooks.Add(objCSDbasedonCRMDoc);
									break;
									}
								case enumDocumentTypes.CSD_Document_DRM_Inline:
									{
									CSD_Document_DRM_Inline objCSDdrmInlineDoc = new CSD_Document_DRM_Inline();
									objCSDdrmInlineDoc.DocumentCollectionID = objDocumentCollection.ID;
									objCSDdrmInlineDoc.DocumentStatus = enumDocumentStatusses.New;
									objCSDdrmInlineDoc.DocumentType = enumDocumentTypes.ISD_Document_DRM_Sections;
									strTemplateURL = GetTheDocumentTemplate(datacontexSDDP, "Client Service Description");
									switch(strTemplateURL)
										{
										case "None":
											objCSDdrmInlineDoc.Template = "";
											objCSDdrmInlineDoc.LogError("The template could not be found.");
											break;
										case "Error":
											objCSDdrmInlineDoc.Template = "";
											objCSDdrmInlineDoc.LogError("Unable to access the template.");
											break;
										default:
											objCSDdrmInlineDoc.Template = websiteURL.Substring(0, websiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									Console.WriteLine("\t Template: {0}", objCSDdrmInlineDoc.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objCSDdrmInlineDoc.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objCSDdrmInlineDoc.Hyperlink_View = true;
									// Load the Document Options
									if(DCsToGen.CSDDocumentDRMInlineOptions != null)
										{
										if(ConvertOptionsToList(DCsToGen.CSDDocumentDRMInlineOptions, ref optionsWorkList)) // conversion is successful
											{
											objCSDdrmInlineDoc.TransposeDocumentOptions(ref optionsWorkList);
											}
										else // the conversion failed
											{
											objCSDdrmInlineDoc.LogError("Invalid format in the Document Options :. unable to generate the document.");
											Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
											}
										}
									else  // == Null
										{
										objCSDdrmInlineDoc.LogError("No document options were specified - cannot generate blank documents.");
										Console.WriteLine("No document options were selected - cannot generate blank documents.");
										}
									// add to the list of DocumentToBeGenerated
									listDocumentsWorkbooks.Add(objCSDdrmInlineDoc);
									break;
									}
								case enumDocumentTypes.CSD_Document_DRM_Sections:
									{
									CSD_Document_DRM_Sections objCSDdrmSectionsDoc = new CSD_Document_DRM_Sections();
									objCSDdrmSectionsDoc.DocumentCollectionID = objDocumentCollection.ID;
									objCSDdrmSectionsDoc.DocumentStatus = enumDocumentStatusses.New;
									objCSDdrmSectionsDoc.DocumentType = enumDocumentTypes.ISD_Document_DRM_Sections;
									strTemplateURL = GetTheDocumentTemplate(datacontexSDDP, "Client Service Description");
									switch(strTemplateURL)
										{
										case "None":
											objCSDdrmSectionsDoc.Template = "";
											objCSDdrmSectionsDoc.LogError("The template could not be found.");
											break;
										case "Error":
											objCSDdrmSectionsDoc.Template = "";
											objCSDdrmSectionsDoc.LogError("Unable to access the template.");
											break;
										default:
											objCSDdrmSectionsDoc.Template = websiteURL.Substring(0, websiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									Console.WriteLine("\t Template: {0}", objCSDdrmSectionsDoc.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objCSDdrmSectionsDoc.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objCSDdrmSectionsDoc.Hyperlink_View = true;
									// Load the Document Options
									if(DCsToGen.CSDDocumentDRMSectionsOptions != null)
										{
										if(ConvertOptionsToList(DCsToGen.CSDDocumentDRMSectionsOptions, ref optionsWorkList)) // conversion is successful
											{
											objCSDdrmSectionsDoc.TransposeDocumentOptions(ref optionsWorkList);
											}
										else // the conversion failed
											{
											objCSDdrmSectionsDoc.LogError("Invalid format in the Document Options :. unable to generate the document.");
											Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
											}
										}
									else  // == Null
										{
										objCSDdrmSectionsDoc.LogError("No document options were specified - cannot generate blank documents.");
										Console.WriteLine("No document options were selected - cannot generate blank documents.");
										}
									// add to the list of DocumentToBeGenerated
									listDocumentsWorkbooks.Add(objCSDdrmSectionsDoc);
									break;
									}
								case enumDocumentTypes.External_Technology_Coverage_Dashboard:
									{
									External_Technology_Coverage_Dashboard_Workbook objExternalTechnologyCoverageDasboardWB = new External_Technology_Coverage_Dashboard_Workbook();
									objExternalTechnologyCoverageDasboardWB.DocumentCollectionID = objDocumentCollection.ID;
									objExternalTechnologyCoverageDasboardWB.DocumentStatus = enumDocumentStatusses.New;
									objExternalTechnologyCoverageDasboardWB.DocumentType = enumDocumentTypes.Client_Requirement_Mapping_Workbook;
									strTemplateURL = GetTheDocumentTemplate(datacontexSDDP, "Technology Roadmap Workbook");
									switch(strTemplateURL)
										{
										case "None":
											objExternalTechnologyCoverageDasboardWB.Template = "";
											objExternalTechnologyCoverageDasboardWB.LogError("The template could not be found.");
                                                       break;
										case "Error":
											objExternalTechnologyCoverageDasboardWB.Template = "";
											objExternalTechnologyCoverageDasboardWB.LogError("The template could not be accessed.");
                                                       break;
										default:
											objExternalTechnologyCoverageDasboardWB.Template = websiteURL.Substring(0, websiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									Console.WriteLine("\t Template: {0}", objExternalTechnologyCoverageDasboardWB.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objExternalTechnologyCoverageDasboardWB.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objExternalTechnologyCoverageDasboardWB.Hyperlink_View = true;
									// add to the list of DocumentOptions
									listDocumentsWorkbooks.Add(objExternalTechnologyCoverageDasboardWB);
									break;
									}
								case enumDocumentTypes.Internal_Technology_Coverage_Dashboard:
									{
									Internal_Technology_Coverage_Dashboard_Workbook objInternalTechnologyCoverageDashboardWB = new Internal_Technology_Coverage_Dashboard_Workbook();
									objInternalTechnologyCoverageDashboardWB.DocumentCollectionID = objDocumentCollection.ID;
									objInternalTechnologyCoverageDashboardWB.DocumentStatus = enumDocumentStatusses.New;
									objInternalTechnologyCoverageDashboardWB.DocumentType = enumDocumentTypes.Client_Requirement_Mapping_Workbook;
									strTemplateURL = GetTheDocumentTemplate(datacontexSDDP, "Technology Roadmap Workbook");
									switch(strTemplateURL)
										{
										case "None":
											objInternalTechnologyCoverageDashboardWB.Template = "";
											objInternalTechnologyCoverageDashboardWB.LogError("The template could not be found.");
                                                       break;
										case "Error":
											objInternalTechnologyCoverageDashboardWB.Template = "";
											objInternalTechnologyCoverageDashboardWB.LogError("The template could not be accessed.");
                                                       break;
										default:
											objInternalTechnologyCoverageDashboardWB.Template = websiteURL.Substring(0, websiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}

									Console.WriteLine("\t Template: {0}", objInternalTechnologyCoverageDashboardWB.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objInternalTechnologyCoverageDashboardWB.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objInternalTechnologyCoverageDashboardWB.Hyperlink_View = true;
									// add to the list of DocumentOptions
									listDocumentsWorkbooks.Add(objInternalTechnologyCoverageDashboardWB);
									break;
									}
								case enumDocumentTypes.ISD_Document_DRM_Inline:
									{
									ISD_Document_DRM_Inline objISDdrmInlineDoc = new ISD_Document_DRM_Inline();
									objISDdrmInlineDoc.DocumentCollectionID = objDocumentCollection.ID;
									objISDdrmInlineDoc.DocumentStatus = enumDocumentStatusses.New;
									objISDdrmInlineDoc.DocumentType = enumDocumentTypes.ISD_Document_DRM_Inline;
									strTemplateURL = GetTheDocumentTemplate(datacontexSDDP, "Internal Service Description");
									switch(strTemplateURL)
										{
										case "None":
											objISDdrmInlineDoc.Template = "";
											objISDdrmInlineDoc.LogError("The template could not be found.");
											break;
										case "Error":
											objISDdrmInlineDoc.Template = "";
											objISDdrmInlineDoc.LogError("Unable to access the template.");
											break;
										default:
											objISDdrmInlineDoc.Template = websiteURL.Substring(0, websiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									Console.WriteLine("\t Template: {0}", objISDdrmInlineDoc.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objISDdrmInlineDoc.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objISDdrmInlineDoc.Hyperlink_View = true;
									// Load the Document Options
									if(DCsToGen.ISDDocumentDRMInlineOptions != null)
										{
										if(ConvertOptionsToList(DCsToGen.ISDDocumentDRMInlineOptions, ref optionsWorkList)) // conversion is successful
											{
											objISDdrmInlineDoc.TransposeDocumentOptions(ref optionsWorkList);
											}
										else // the conversion failed
											{
											objISDdrmInlineDoc.LogError("Invalid format in the Document Options :. unable to generate the document.");
											Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
											}
										}
									else  // == Null
										{
										objISDdrmInlineDoc.LogError("No document options were specified - cannot generate blank documents.");
										Console.WriteLine("No document options were selected - cannot generate blank documents.");
										}
									// add to the list of DocumentToBeGenerated
									listDocumentsWorkbooks.Add(objISDdrmInlineDoc);
									break;
									}
								case enumDocumentTypes.ISD_Document_DRM_Sections:
									{
									ISD_Document_DRM_Sections objISDdrmSectionsDoc = new ISD_Document_DRM_Sections();
									objISDdrmSectionsDoc.DocumentCollectionID = objDocumentCollection.ID;
									objISDdrmSectionsDoc.DocumentStatus = enumDocumentStatusses.New;
									objISDdrmSectionsDoc.DocumentType = enumDocumentTypes.ISD_Document_DRM_Sections;
									strTemplateURL = GetTheDocumentTemplate(datacontexSDDP, "Internal Service Description");
									switch(strTemplateURL)
										{
										case "None":
											objISDdrmSectionsDoc.Template = "";
											objISDdrmSectionsDoc.LogError("The template could not be found.");
											break;
										case "Error":
											objISDdrmSectionsDoc.Template = "";
											objISDdrmSectionsDoc.LogError("Unable to access the template.");
											break;
										default:
											objISDdrmSectionsDoc.Template = websiteURL.Substring(0, websiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									Console.WriteLine("\t Template: {0}", objISDdrmSectionsDoc.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objISDdrmSectionsDoc.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objISDdrmSectionsDoc.Hyperlink_View = true;
									// Load the Document Options
									if(DCsToGen.ISDDocumentDRMSectionsOptions != null)
										{
										if(ConvertOptionsToList(DCsToGen.ISDDocumentDRMSectionsOptions, ref optionsWorkList)) // conversion is successful
											{
											objISDdrmSectionsDoc.TransposeDocumentOptions(ref optionsWorkList);
											}
										else // the conversion failed
											{
											objISDdrmSectionsDoc.LogError("Invalid format in the Document Options :. unable to generate the document.");
											Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
											}
										}
									else  // == Null
										{
										objISDdrmSectionsDoc.LogError("No document options were specified - cannot generate blank documents.");
										Console.WriteLine("No document options were selected - cannot generate blank documents.");
										}
									// add to the list of DocumentToBeGenerated
									listDocumentsWorkbooks.Add(objISDdrmSectionsDoc);
									break;
									}
								case enumDocumentTypes.Pricing_Addendum_Document:
									{
									// not currently implemented
									break;
									}
								case enumDocumentTypes.RACI_Matrix_Workbook_per_Deliverable:
									{
									RACI_Matrix_Workbook_per_Deliverable objRACIperDeliverableWB = new RACI_Matrix_Workbook_per_Deliverable();
									objRACIperDeliverableWB.DocumentCollectionID = objDocumentCollection.ID;
									objRACIperDeliverableWB.DocumentStatus = enumDocumentStatusses.New;
									objRACIperDeliverableWB.DocumentType = enumDocumentTypes.Client_Requirement_Mapping_Workbook;
									strTemplateURL = GetTheDocumentTemplate(datacontexSDDP, "RACI Matrix Workbook");
									switch(strTemplateURL)
										{
										case "None":
											objRACIperDeliverableWB.Template = "";
											objRACIperDeliverableWB.LogError("The template could not be found.");
											break;
										case "Error":
											objRACIperDeliverableWB.Template = "";
											objRACIperDeliverableWB.LogError("The template could not be accessed.");																break;
										default:
											objRACIperDeliverableWB.Template = websiteURL.Substring(0, websiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									Console.WriteLine("\t Template: {0}", objRACIperDeliverableWB.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objRACIperDeliverableWB.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objRACIperDeliverableWB.Hyperlink_View = true;
									// add to the list of DocumentOptions
									listDocumentsWorkbooks.Add(objRACIperDeliverableWB);
									break;
									}
								case enumDocumentTypes.RACI_Workbook_per_Role:
									{
									RACI_Workbook_per_Role objRACIperRoleWB = new RACI_Workbook_per_Role();
									objRACIperRoleWB.DocumentCollectionID = objDocumentCollection.ID;
									objRACIperRoleWB.DocumentStatus = enumDocumentStatusses.New;
									objRACIperRoleWB.DocumentType = enumDocumentTypes.Client_Requirement_Mapping_Workbook;
									strTemplateURL = GetTheDocumentTemplate(datacontexSDDP, "RACI Workbook");
									switch(strTemplateURL)
										{
										case "None":
											objRACIperRoleWB.Template = "";
											objRACIperRoleWB.LogError(("The template could not be found."));
											break;
										case "Error":
											objRACIperRoleWB.Template = "";
											objRACIperRoleWB.LogError(("The template could not be accessed."));
											break;
										default:
											objRACIperRoleWB.Template = websiteURL.Substring(0, websiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
							
									Console.WriteLine("\t Template: {0}", objRACIperRoleWB.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objRACIperRoleWB.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objRACIperRoleWB.Hyperlink_View = true;
									// add to the list of DocumentOptions
									listDocumentsWorkbooks.Add(objRACIperRoleWB);
									break;
									}
								case enumDocumentTypes.Service_Framework_Document_DRM_inline:
									{
									Services_Framework_Document_DRM_Inline objServicesFrameworkDRMinlineDoc = new Services_Framework_Document_DRM_Inline();
									objServicesFrameworkDRMinlineDoc.DocumentCollectionID = objDocumentCollection.ID;
									objServicesFrameworkDRMinlineDoc.DocumentStatus = enumDocumentStatusses.New;
									objServicesFrameworkDRMinlineDoc.DocumentType = enumDocumentTypes.ISD_Document_DRM_Inline;
									strTemplateURL = GetTheDocumentTemplate(datacontexSDDP, "Services Framework Description");
									switch(strTemplateURL)
										{
										case "None":
											objServicesFrameworkDRMinlineDoc.Template = "";
											objServicesFrameworkDRMinlineDoc.LogError("The template could not be found.");
											break;
										case "Error":
											objServicesFrameworkDRMinlineDoc.Template = "";
											objServicesFrameworkDRMinlineDoc.LogError("Unable to access the template.");
											break;
										default:
											objServicesFrameworkDRMinlineDoc.Template = websiteURL.Substring(0, websiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									Console.WriteLine("\t Template: {0}", objServicesFrameworkDRMinlineDoc.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objServicesFrameworkDRMinlineDoc.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objServicesFrameworkDRMinlineDoc.Hyperlink_View = true;
									// Load the Document Options
									if(DCsToGen.ISDDocumentDRMInlineOptions != null)
										{
										if(ConvertOptionsToList(DCsToGen.ISDDocumentDRMInlineOptions, ref optionsWorkList)) // conversion is successful
											{
											objServicesFrameworkDRMinlineDoc.TransposeDocumentOptions(ref optionsWorkList);
											}
										else // the conversion failed
											{
											objServicesFrameworkDRMinlineDoc.LogError("Invalid format in the Document Options :. unable to generate the document.");
											Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
											}
										}
									else  // == Null
										{
										objServicesFrameworkDRMinlineDoc.LogError("No document options were specified - cannot generate blank documents.");
										Console.WriteLine("No document options were selected - cannot generate blank documents.");
										}
									listDocumentsWorkbooks.Add(objServicesFrameworkDRMinlineDoc);
									break;
									}
								case enumDocumentTypes.Service_Framework_Document_DRM_sections:
									{
									Services_Framework_Document_DRM_Sections objServicesFrameworkDRMsectionsDoc = new Services_Framework_Document_DRM_Sections();
									objServicesFrameworkDRMsectionsDoc.DocumentCollectionID = objDocumentCollection.ID;
									objServicesFrameworkDRMsectionsDoc.DocumentStatus = enumDocumentStatusses.New;
									objServicesFrameworkDRMsectionsDoc.DocumentType = enumDocumentTypes.Service_Framework_Document_DRM_sections;
									strTemplateURL = GetTheDocumentTemplate(datacontexSDDP, "Services Framework Description");
									switch(strTemplateURL)
										{
										case "None":
											objServicesFrameworkDRMsectionsDoc.Template = "";
											objServicesFrameworkDRMsectionsDoc.LogError("The template could not be found.");
											break;
										case "Error":
											objServicesFrameworkDRMsectionsDoc.Template = "";
											objServicesFrameworkDRMsectionsDoc.LogError("Unable to access the template.");
											break;
										default:
											objServicesFrameworkDRMsectionsDoc.Template = websiteURL.Substring(0, websiteURL.IndexOf("/", 11)) + strTemplateURL;
											break;
										}
									Console.WriteLine("\t Template: {0}", objServicesFrameworkDRMsectionsDoc.Template);
									if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
										objServicesFrameworkDRMsectionsDoc.HyperlinkEdit = true;
									else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
										objServicesFrameworkDRMsectionsDoc.Hyperlink_View = true;
									// Load the Document Options
									if(DCsToGen.ISDDocumentDRMSectionsOptions != null)
										{
										if(ConvertOptionsToList(DCsToGen.ISDDocumentDRMSectionsOptions, ref optionsWorkList))
											{
											objServicesFrameworkDRMsectionsDoc.TransposeDocumentOptions(ref optionsWorkList);
											}
										else
											{
											objServicesFrameworkDRMsectionsDoc.LogError("Invalid format in the Document Options :. unable to generate the document.");
											Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
											}
										} // !=Null
									else
										{
										objServicesFrameworkDRMsectionsDoc.LogError("No document options were specified - cannot generate a blank document.");
										Console.WriteLine("No document options were selected - cannot generate blank documents.");
										}
									// add to the list of DocumentToBeGenerated
									listDocumentsWorkbooks.Add(objServicesFrameworkDRMsectionsDoc);
									break;
									}
								default:
									{
									break;
									}
								}
							}
						// assign the list of DocumentWorkbooks to the collection of Documents_and_Workbooks of the DocumentCollection
						objDocumentCollection.Documents_and_Workbooks = listDocumentsWorkbooks;
                              }
					// Add the instance of the Document Collection Object to the List of Document Collection that must be generated
					parCollectionsToGenerate.Add(objDocumentCollection);
					Console.WriteLine(" Document Collection: {0} successfully loaded..\n Now there are {1} collections to generate.\n", DCsToGen.Id, parCollectionsToGenerate.Count);
					} // Loop of the For Each DocColsToGenerate
				Console.WriteLine("All entries processed and added to List parCollectionsToGenerate) - {0} collections to generate...", parCollectionsToGenerate.Count);
				return "Good";
				} // end of Try
			catch(Exception ex)
				{
				Console.WriteLine("Exception: [{0}] occurred and was caught. \n{1}", ex.HResult.ToString(), ex.Message);
				if(ex.HResult == -2146330330)
					return "Error: Cannot access site: " + websiteURL + " Ensure the computer is connected to the Dimension Data Domain network";
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
				var DocumentTemplates = parDataContext.DocumentTemplates;
				var Template = from docTemplate in DocumentTemplates where docTemplate.TemplateTypeValue == parTemplateType
							select docTemplate;
				// Console.WriteLine("\t\t\t + {0} templates found.", Template.Count());
				if(Template.Count() > 0)
					{
					foreach(var tpl in Template)
						{
						Console.WriteLine("\t\t\t - {0} - {1} [{2}]", tpl.Id, tpl.Title, tpl.TemplateTypeValue);
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
				Console.WriteLine("\t\t + OptionID: {0}", parStringOptions.Substring(position, (parStringOptions.IndexOf(",", position) - position)));

				if(!int.TryParse(parStringOptions.Substring(position, (parStringOptions.IndexOf(",", position) - position)), out value))
					{
					Console.WriteLine("Option value is not numeric at position {0} in {1}.", position, parStringOptions);
					errors += 1;
					}
				else
					{
					parListOfOptions.Add(value);
					}

				if(parStringOptions.IndexOf(",", position) > 0)
					{
					position = parStringOptions.IndexOf(",", position) + 1;
					//Console.WriteLine("\t\t\t\t {0} of {1}", position, parStringOptions.Length);
					}
				else
					{
					Console.WriteLine("\t\t\t\t {0} of {1}", position, parStringOptions.Length);
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