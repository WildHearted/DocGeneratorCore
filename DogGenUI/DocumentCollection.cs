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

		// Other Fields
		
		// Object Methods
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
		public static bool GetCollectionsToGenerate(ref List<DocumentCollection> parCollectionsToGenerate)
			{
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
						// Only 7 expands are allowed with 1 data access.
						//.Expand(p => p.StandardPricingProduct)
						//.Expand(p => p.ISDDocumentDRMInlineOptions)
						//.Expand(p => p.ISDDocumentDRMSectionsOptions)
						//.Expand(p => p.CSDDocumentDRMInlineOptions);
						//.Expand(p => p.CSDDocumentDRMSectionsOptions);
						//.Expand(p => p.CSDDocumentBasedOnCRMOptions)
						//.Expand(p => p.SoWSDOptions);

				var DocColsToGenerate = from dc in DocCollectionLib where dc.GenerateActionValue != null & dc.GenerateActionValue.Substring(0,4) != "Save" orderby dc.Id select dc;

				Console.WriteLine("{0} Document Collections to create...", DocColsToGenerate.Count());

				foreach(var DCsToGen in DocColsToGenerate)
					{

					if(DCsToGen.GenerateActionValue.Substring(0,4) == "Save")
						{
						Console.WriteLine("{0} Generate  Action value is {1}, therefore it will not be generated.",DCsToGen.Id, DCsToGen.GenerateActionValue);
						continue;
                              }
					Console.WriteLine("\nID: {0}  Title: {1}\n\t DocGen Client Name: [{2}] - Client Title:[{3}] ", DCsToGen.Id, DCsToGen.Title, DCsToGen.Client_.DocGenClientName, DCsToGen.Client_.Title);

					// Create a new Instance for the Document Collection into which the object properties are loaded
					DocumentCollection iDocumentCollection = new DocumentCollection();
					//Set the basic object properties
					iDocumentCollection.ID = DCsToGen.Id;
					Console.WriteLine("\t ID: {0} ", iDocumentCollection.ID);

					if(DCsToGen.Client_.DocGenClientName == null)
						iDocumentCollection.ClientName = "the Client";
					else
						iDocumentCollection.ClientName = DCsToGen.Client_.DocGenClientName;
					Console.WriteLine("\t ClientName: {0} ", iDocumentCollection.ClientName);

					if(DCsToGen.Title == null)
						iDocumentCollection.Title = "Collection Title for entry " + DCsToGen.Id;
					else
						iDocumentCollection.Title = DCsToGen.Title;

					Console.WriteLine("\t Title: {0}", iDocumentCollection.Title);
					if(DCsToGen.GenerateNotifyMe == null)
						iDocumentCollection.NotifyMe = false;
					else
						iDocumentCollection.NotifyMe = DCsToGen.GenerateNotifyMe.Value;
					Console.WriteLine("\t NotifyMe: {0} ", iDocumentCollection.NotifyMe);

					if (DCsToGen.GenerateNotificationEMail == null)
						iDocumentCollection.NotificationEmail = "None";
					else
	                         iDocumentCollection.NotificationEmail = DCsToGen.GenerateNotificationEMail;
					Console.WriteLine("\t NotificationEmail: {0} ", iDocumentCollection.NotificationEmail);

					if(DCsToGen.GenerateOnDateTime == null)
						iDocumentCollection.GenerateOnDateTime = DateTime.Now;
					else
						iDocumentCollection.GenerateOnDateTime = DCsToGen.GenerateOnDateTime.Value;
					Console.WriteLine("\t GenerateOnDateTime: {0} ", iDocumentCollection.GenerateOnDateTime);

					// Set the Mapping value
					if(DCsToGen.Mapping_Id != null)
						{
						try
							{
							iDocumentCollection.Mapping = Convert.ToInt32(DCsToGen.Mapping_Id);
							}
						catch(OverflowException ex)
							{
							Console.WriteLine("Overflow Exception occurred when converting the Mappin value to a Integer.\n Error Description: {0}", ex.Message);
							iDocumentCollection.Mapping = 0;
							}
						}
					else
						{
						iDocumentCollection.Mapping = 0;
						}
					Console.WriteLine("\t Mapping: {0} ", iDocumentCollection.Mapping);

					// Set the Pricing Workbook
					if(DCsToGen.PricingWorkbookId != null)
						try
							{
							iDocumentCollection.PricingWorkbook = Convert.ToInt32(DCsToGen.PricingWorkbookId);
							}
						catch(OverflowException ex)
							{
							Console.WriteLine("Overflow Exception occurred when converting the Pricing Workbook value to a Integer.\n Error Description: {0}", ex.Message);
							iDocumentCollection.Mapping = 0;
							}
					else
						iDocumentCollection.PricingWorkbook = 0;
					Console.WriteLine("\t PricingWorkbook: {0} ", iDocumentCollection.PricingWorkbook);

					// Set the Generate Schedule Options
					enumGenerateScheduleOptions generateSchdlOption;
					if(DCsToGen.GenerateScheduleOptionValue != null)
						{
						if(PrepareStringForEnum(DCsToGen.GenerateScheduleOptionValue, out enumWorkString))
							{
							if(Enum.TryParse<enumGenerateScheduleOptions>(enumWorkString, out generateSchdlOption))
								{
								iDocumentCollection.GenerateScheduleOption = generateSchdlOption;
								}
							else
								{
								iDocumentCollection.GenerateScheduleOption = enumGenerateScheduleOptions.Do_NOT_Repeat;
								}
							}
						else
							{
							iDocumentCollection.GenerateScheduleOption = enumGenerateScheduleOptions.Do_NOT_Repeat;
							}
						}
					else
						{
						iDocumentCollection.GenerateScheduleOption = enumGenerateScheduleOptions.Do_NOT_Repeat;
						}
					Console.WriteLine("\t Generate ScheduleOption: {0} ", iDocumentCollection.GenerateScheduleOption);

					// Set the Generate Repeat Intervals
					enumGenerateRepeatIntervals generateRepeatIntrvl;
					if(DCsToGen.GenerateRepeatIntervalValue0 != null)
						{
						if(PrepareStringForEnum(DCsToGen.GenerateRepeatIntervalValue0, out enumWorkString))
							{
							if(Enum.TryParse<enumGenerateRepeatIntervals>(enumWorkString, out generateRepeatIntrvl))
								{
								iDocumentCollection.GenerateRepeatInterval = generateRepeatIntrvl;
								}
							else
								{
								iDocumentCollection.GenerateRepeatInterval = enumGenerateRepeatIntervals.Month;
								}
							}
						else
							{
							iDocumentCollection.GenerateRepeatInterval = enumGenerateRepeatIntervals.Month;
							}
						}
					else
						{
						iDocumentCollection.GenerateRepeatInterval = enumGenerateRepeatIntervals.Month;
						}
					Console.WriteLine("\t GenerateRepeatInterval: {0} ", iDocumentCollection.GenerateRepeatInterval);

					// Set the Generate Repeat Interval Value
					if(DCsToGen.GenerateRepeatIntervalValue != null)
						{
						try
							{
							iDocumentCollection.GenerateRepeatIntervalValue = Convert.ToInt32(DCsToGen.GenerateRepeatIntervalValue.Value);
							}
						catch(OverflowException ex)
							{
							Console.WriteLine("Overflow Exception occurred when converting the Generate Repeat Interval to a Integer.\n Error Description: {0}", ex.Message);
							iDocumentCollection.GenerateRepeatIntervalValue = 0;
							}
						}
					else
						{
						iDocumentCollection.GenerateRepeatIntervalValue = 0;
						}
					Console.WriteLine("\t GenerateRepeatIntervalValue: {0} ", iDocumentCollection.GenerateRepeatIntervalValue);
					// Set the Hyperlink Options
					if(DCsToGen.HyperlinkOptionsValue != null)
						{
						enumHyperlinkOptions hyperLnkOption;
						if(PrepareStringForEnum(DCsToGen.HyperlinkOptionsValue, out enumWorkString))
							{
							if(Enum.TryParse<enumHyperlinkOptions>(enumWorkString, out hyperLnkOption))
								{
								iDocumentCollection.HyperLinkOption = hyperLnkOption;
								}
							else
								{
								iDocumentCollection.HyperLinkOption = enumHyperlinkOptions.Do_NOT_Include_Hyperlinks;
								}
							}
						else
							{
							iDocumentCollection.HyperLinkOption = enumHyperlinkOptions.Do_NOT_Include_Hyperlinks;
							}
						}
					else
						{
						iDocumentCollection.HyperLinkOption = enumHyperlinkOptions.Do_NOT_Include_Hyperlinks;
						}
					Console.WriteLine("\t HyperlinkOption: {0} ", iDocumentCollection.HyperLinkOption);

					// Get the Content Layer Colour Coding Option
					Console.WriteLine("\t Content Layer Colour Coding has {0} entries.", DCsToGen.ContentLayerColourCodingOption.Count.ToString());
					iDocumentCollection.ColourCodingLayer1 = false;
					iDocumentCollection.ColourCodingLayer2 = false;
					iDocumentCollection.ColourCodingLayer3 = false;
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
										iDocumentCollection.ColourCodingLayer1 = true;
										}
									if(CLCCOptions.Equals(enumContent_Layer_Colour_Coding_Options.Colour_Code_Layer_2))
										{
										iDocumentCollection.ColourCodingLayer2 = true;
										}
									if(CLCCOptions.Equals(enumContent_Layer_Colour_Coding_Options.Colour_Code_Layer_3))
										{
										iDocumentCollection.ColourCodingLayer3 = true;
										}
									}
								}
							} //Foreach Loop
						}
					Console.WriteLine("\t ContentColourCodingLayer1: {0} ", iDocumentCollection.ColourCodingLayer1);
					Console.WriteLine("\t ContentColourCodingLayer2: {0} ", iDocumentCollection.ColourCodingLayer2);
					Console.WriteLine("\t ContentColourCodingLayer3: {0} ", iDocumentCollection.ColourCodingLayer3);

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
									if (Enum.TryParse<enumDocumentTypes>(enumWorkString,out docType))
										listOfDocumentsToGenerate.Add(docType);
									else
									Console.WriteLine("\t\t [{0}] Not found as enumeration [{1}]",enumWorkString, docType);
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
					iDocumentCollection.DocumentsToGenerate = listOfDocumentsToGenerate;
					Console.WriteLine("\t {0} document to be generated for the Document Collection.", iDocumentCollection.DocumentsToGenerate.Count);

					//Set the Selected Nodes that need to be generated by building a hierchical List with Hierarchy objects
					Console.WriteLine("\t Loading the Nodes that the user selected.");
					if(DCsToGen.SelectedNodes != null)
						{
                              List<Hierarchy> listOfNodesToGenerate = new List<Hierarchy>();
						if(Hierarchy.ConstructHierarchy(DCsToGen.SelectedNodes, ref listOfNodesToGenerate))
							{
							iDocumentCollection.SelectedNodes = listOfNodesToGenerate;
							Console.WriteLine("\t {0} nodes successfully loaded by ConstructHierarchy method.",listOfNodesToGenerate.Count);
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
					// Add the instance of the Document Collection Object to the List of Document Collection that must be generated
					parCollectionsToGenerate.Add(iDocumentCollection);
					Console.WriteLine("Document Collection: {0} successfully loaded..\n Now there are {1} collections to generate.\n", DCsToGen.Id, parCollectionsToGenerate.Count);
					} // Loop of the For Each DocColsToGenerate

				Console.WriteLine("All entries processed and added to List parCollectionsToGenerate) - {0} collections to generate...", parCollectionsToGenerate.Count);
				return true;
				}
			catch(SystemException ex)
				{
				Console.WriteLine("Exception [{0}] occurred and was handled.", ex.Message);
				return false;
				}
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