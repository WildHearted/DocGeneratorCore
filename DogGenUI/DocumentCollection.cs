using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Services.Client;
using Microsoft.SharePoint.Client;
using System.Net;
using System.Linq;
using DocGeneratorCore.SDDPServiceReference;

namespace DocGeneratorCore
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
	/// Mapped to the [Generate Schedule Option] column in SharePoint. It indicates whether the document generation must be repeated or Not repeated.
	/// </summary>
	public enum enumGenerateScheduleOptions
		{
		Do_NOT_Repeat=0,
		Repeat_every=1
		}
	/// <summary>
	/// Mapped to the values of the [Generate Repeat Interval] column in SharePoint. It indicates how often the document generation must repeat.
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

	public enum enumGenerationStatus
		{
		Pending = 0,
		Generating = 1,
		Failed = 8,
		Completed = 9
		}

	/// <summary>
	/// This list contains the documents that the user selected which needs to be generated.
	/// </summary>
	public class DocumentCollection
		{
		// Object Properties

		public int ID{get; set;}
		public bool DetailComplete{get; set;}
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
		
		//++ UpdateGenerateStatus

		public bool UpdateGenerateStatus(ref CompleteDataSet  parDataSet, enumGenerationStatus parGenerationStatus)
			{
			Console.WriteLine("Updating Generation Status of Document Collection: {0}", this.ID);
			string strExceptionMessage = string.Empty;
			try
				{
				Console.WriteLine("Updating status of the entry in Document Collection Libray");
				// Construct the SharePoint Client context and authentication...
				ClientContext objSPcontext = new ClientContext(webFullUrl: parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL + "/");
				objSPcontext.Credentials = parDataSet.SDDPdatacontext.Credentials;
				//objSPcontext.Credentials = new NetworkCredential(
				//	userName: Properties.AppResources.DocGenerator_AccountName,
				//	password: Properties.AppResources.DocGenerator_Account_Password,
				//	domain: Properties.AppResources.DocGenerator_AccountDomain);
				Web objWeb = objSPcontext.Web;

				// Obtain the Document Collection Library entry and its relevant fields/columns.
				List objDocumentCollectionList = objWeb.Lists.GetByTitle("Document Collection Library");
				FieldCollection objGeneratedDocumentsFields = objDocumentCollectionList.Fields;
				CamlQuery objCAMLquery = new CamlQuery();
				objCAMLquery.ViewXml = 
					@"<View>"
						+ "<Query>" 
							+ "<Where>" 
								+ "<Eq><FieldRef Name='ID'/>"
									+ "<Value Type='Counter'>"
									+ this.ID.ToString()
									+ "</Value>"
								+ "</Eq>"
							+ "</Where>" 
						+ "</Query>"
					+ "</View>";

				ListItemCollection objListEntries = objDocumentCollectionList.GetItems(objCAMLquery);
				objSPcontext.Load(objListEntries);

				objSPcontext.ExecuteQuery();

				ListItem objListItem = objListEntries[0];

				Console.WriteLine("{0} - {1}", objListItem["ID"], objListItem["Title"]);
				// update the Generation Status 
				//- Check if the Document Collection Entry must be generate again in future
				if(this.GenerateScheduleOption == enumGenerateScheduleOptions.Repeat_every)
					{//- Yes, the document collection generation must be repeated
					 //- Determine for WHEN it must be rescheduled, and set the next date and time when it must be generated.
					DateTime dtNextScheduleToGenerate;
					if(this.GenerateOnDateTime == null || this.GenerateOnDateTime.Equals(DateTime.MinValue))
						{
						dtNextScheduleToGenerate = DateTime.UtcNow;
						}
					else
						{
						dtNextScheduleToGenerate = this.GenerateOnDateTime;
						}
					switch(this.GenerateRepeatInterval)
						{
						case (enumGenerateRepeatIntervals.Day):
							{
							dtNextScheduleToGenerate.AddDays(Convert.ToDouble(this.GenerateRepeatIntervalValue));
							break;
							}
						case (enumGenerateRepeatIntervals.Week):
							{
							dtNextScheduleToGenerate.AddDays(Convert.ToDouble(this.GenerateRepeatIntervalValue) * 7);
							break;
							}
						case (enumGenerateRepeatIntervals.Month):
							{
							dtNextScheduleToGenerate.AddMonths(this.GenerateRepeatIntervalValue);
							break;
							}
						}
					//-  Clear the **[Generation Status]** in order for DocGenerator to process it again in future
					//- set the next date and time when the entry must be generated - **[Generate on Date Time]** 
					//- and set the **[Generate Action]** to "Schedule...
					objListItem["Generate_x0020_Action"] = "Schedule for a specific date and time";
					objListItem.Update();
					objListItem["Generate_x0020_on_x0020_Date_x00"] = dtNextScheduleToGenerate;
					objListItem.Update();
					objListItem["Generation_x0020_Status"] = null;
					objListItem.Update();
					}
				else
					{//- If (not a repeating Scheduled entry **No**, 
					 //- the Document Collection Entry must only be generated once, therefore set:
					 //- the **[Generation Status]** 
					 //- clear the ** [Generate on Date Time] **
					 //- and clear the **[Generate Action]** that the DocGenerator doesn't pick it up and generate it agian.
					objListItem["Generate_x0020_Action"] = null;
					objListItem.Update();
					objListItem["Generation_x0020_Status"] = parGenerationStatus.ToString();
					objListItem.Update();
					objListItem["Generate_x0020_on_x0020_Date_x00"] = null;
					objListItem.Update();
					}
				
				objSPcontext.ExecuteQuery();
				Console.WriteLine("\t + Successfully Updated ID: {0}", this.ID);
				objSPcontext.Dispose();
				}
			catch(InvalidQueryExpressionException exc)
				{
				Console.WriteLine("\n*** ERROR: Invalid Query Expression Exception ***\n{0} - {1}\nInnerException: {2}\nStackTrace: {3}.",
					exc.HResult, exc.Message, exc.InnerException, exc.StackTrace);
				return false;
				}

			catch(Exception exc)
				{
				Console.WriteLine("\n*** Exception ERROR ***\n{0} - {1}\nInnerException: {2}\nStackTrace: {3}.", exc.HResult, exc.Message, exc.InnerException, exc.StackTrace);
				return false;
				}

			Console.WriteLine("Update Successful...");
			return true;

			}

		//++ PopulateCollections
		/// <summary>
		/// Method which obtains all the Document Collections from the [Document Collection Library] List that still need to be generated.
		/// The Method returns a List collection consisting of Document Collection objects that must be generated.
		/// </summary>
		public static void PopulateCollections(
			ref CompleteDataSet parDataSet,
			ref List<DocumentCollection> parDocumentCollectionList)
			{
			List<int> optionsWorkList = new List<int>();
			string enumWorkString;
			string strExceptionMessage = string.Empty;
			//List<DocumentCollection> listDocumentCollection = new List<DocumentCollection>();

			try
				{
				var dsDocCollectionLibrary = parDataSet.SDDPdatacontext.DocumentCollectionLibrary
					.Expand(dc => dc.Client_)
					.Expand(dc => dc.ContentLayerColourCodingOption)
					.Expand(dc => dc.GenerateFrameworkDocuments)
					.Expand(dc => dc.GenerateInternalDocuments)
					.Expand(dc => dc.GenerateExternalDocuments)
					.Expand(dc => dc.ModifiedBy)
					.Expand(dc => dc.CreatedBy);

				//foreach(var recDocCollsToGen in dsDocumentCollections)
				foreach(DocumentCollection objDocumentCollection in parDocumentCollectionList)
					{
					Console.WriteLine("\r\nDocumentCollection ID: {0}  Title: {1}", 
						objDocumentCollection.ID, objDocumentCollection.Title);

					// Create a new Instance for the DocumentCollection into which the object properties are loaded
					//DocumentCollection objDocumentCollection = new DocumentCollection();
					//Set the basic object properties
					//objDocumentCollection.ID = recDocCollsToGen.Id;

					var dsDocumentCollections =
						from dsCollection in dsDocCollectionLibrary
						where dsCollection.Id == objDocumentCollection.ID
						select dsCollection;

					var objDocCollection = dsDocumentCollections.FirstOrDefault();

					//Chack if the DocumentCollection entry was retreived.
					if(objDocCollection == null)
						{
						objDocumentCollection.DetailComplete = false;
						}
					else
						{
						Console.WriteLine("\t ID: {0} ", objDocumentCollection.ID);

						if(objDocumentCollection.ClientName == null)
							objDocumentCollection.ClientName = "the Client";
						else
							objDocumentCollection.ClientName = objDocCollection.Client_.DocGenClientName;
						Console.WriteLine("\t ClientName: {0} ", objDocumentCollection.ClientName);

						if(objDocCollection.Title == null)
							objDocumentCollection.Title = "Collection Title for entry " + objDocCollection.Id;
						else
							objDocumentCollection.Title = objDocCollection.Title;
						Console.WriteLine("\t Title: {0}", objDocumentCollection.Title);

						if(objDocCollection.GenerateNotifyMe == null)
							objDocumentCollection.NotifyMe = false;
						else
							objDocumentCollection.NotifyMe = objDocCollection.GenerateNotifyMe.Value;
						Console.WriteLine("\t NotifyMe: {0} ", objDocumentCollection.NotifyMe);

						if(objDocCollection.ModifiedBy.Name.Contains("SDDP") == false)
							objDocumentCollection.RequestingUserID = objDocCollection.ModifiedById;
						else if(objDocCollection.CreatedBy.Name.Contains("SDDP") == false)
							objDocumentCollection.RequestingUserID = objDocCollection.CreatedById;
						else
							objDocumentCollection.RequestingUserID = objDocCollection.ModifiedById;

						
						if(objDocCollection.GenerateNotificationEMail == null)
							objDocumentCollection.NotificationEmail = null;
						else
							if(objDocCollection.GenerateNotificationEMail == null)
							{
							objDocumentCollection.NotificationEmail = objDocCollection.ModifiedBy.WorkEmail;
							}
						else
							{
							objDocumentCollection.NotificationEmail = objDocCollection.GenerateNotificationEMail;
							}

						Console.WriteLine("\t NotificationEmail: {0} ", objDocumentCollection.NotificationEmail);
						// Set the GenerateOnDateTime value
						if(objDocCollection.GenerateOnDateTime == null)
							objDocumentCollection.GenerateOnDateTime = DateTime.Now;
						else
							objDocumentCollection.GenerateOnDateTime = objDocCollection.GenerateOnDateTime.Value;
						Console.WriteLine("\t GenerateOnDateTime: {0} ", objDocumentCollection.GenerateOnDateTime);
						// Set the Mapping value
						if(objDocCollection.Mapping_Id != null)
							{
							try
								{
								objDocumentCollection.Mapping = Convert.ToInt32(objDocCollection.Mapping_Id);
								}
							catch(OverflowException ex)
								{
								Console.WriteLine("Overflow Exception occurred when converting the Mapping value to a Integer.\n Error Description: {0}", ex.Message);
								objDocumentCollection.Mapping = 0;
								}
							}
						else
							{
							objDocumentCollection.Mapping = 0;
							}
						//Console.WriteLine("\t Mapping: {0} ", objDocumentCollection.Mapping);
						// Set the PricingWorkbook value
						if(objDocCollection.PricingWorkbookId != null)
							try
								{
								objDocumentCollection.PricingWorkbook = Convert.ToInt32(objDocCollection.PricingWorkbookId);
								}
							catch(OverflowException ex)
								{
								Console.WriteLine("Overflow Exception occurred when converting the Pricing Workbook value to a Integer."
									+ "\n Error Description: {0}", ex.Message);
								objDocumentCollection.Mapping = 0;
								}
						else
							objDocumentCollection.PricingWorkbook = 0;

						// Set the [Generate Schedule Option]
						enumGenerateScheduleOptions generateSchdlOption;
						if(objDocCollection.GenerateScheduleOptionValue != null)
							{
							if(PrepareStringForEnum(objDocCollection.GenerateScheduleOptionValue, out enumWorkString))
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
						if(objDocCollection.GenerateRepeatIntervalValue0 != null)
							{
							if(PrepareStringForEnum(objDocCollection.GenerateRepeatIntervalValue0, out enumWorkString))
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
						if(objDocCollection.GenerateRepeatIntervalValue != null)
							{
							try
								{
								objDocumentCollection.GenerateRepeatIntervalValue = Convert.ToInt32(objDocCollection.GenerateRepeatIntervalValue.Value);
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
						if(objDocCollection.HyperlinkOptionsValue != null)
							{
							enumHyperlinkOptions hyperLnkOption;
							if(PrepareStringForEnum(objDocCollection.HyperlinkOptionsValue, out enumWorkString))
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
						if(objDocCollection.ContentLayerColourCodingOption.Count > 0)
							{
							foreach(var entry in objDocCollection.ContentLayerColourCodingOption)
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
						if(objDocCollection.PresentationModeValue == null
						|| objDocCollection.PresentationModeValue == "Layered")
							objDocumentCollection.PresentationMode = enumPresentationMode.Layered;
						else
							objDocumentCollection.PresentationMode = enumPresentationMode.Expanded;

						int noOfDocsToGenerateInCollection = 0;
						List<enumDocumentTypes> listOfDocumentTypesToGenerate = new List<enumDocumentTypes>();
						enumDocumentTypes docType;
						// Set the FrameworkDocuments that must be generated
						Console.WriteLine("\t Generate Framework Documents: {0} entries.", objDocCollection.GenerateFrameworkDocuments.Count.ToString());
						if(objDocCollection.GenerateFrameworkDocuments.Count > 0)
							{
							foreach(var entry in objDocCollection.GenerateFrameworkDocuments)
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
						Console.WriteLine("\t Generate Internal Documents: {0} entries.", objDocCollection.GenerateInternalDocuments.Count.ToString());
						if(objDocCollection.GenerateInternalDocuments.Count > 0)
							{
							foreach(var entry in objDocCollection.GenerateInternalDocuments)
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
						Console.WriteLine("\t Generate External Documents: {0} entries.", objDocCollection.GenerateExternalDocuments.Count.ToString());
						if(objDocCollection.GenerateExternalDocuments.Count > 0)
							{
							foreach(var entry in objDocCollection.GenerateExternalDocuments)
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
						if(objDocCollection.SelectedNodes != null)
							{
							List<Hierarchy> listOfNodesToGenerate = new List<Hierarchy>();
							if(Hierarchy.ConstructHierarchy(objDocCollection.SelectedNodes, ref listOfNodesToGenerate))
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
							Console.WriteLine("\t There are no selected content to generate for Document Collection {0} - {1}", objDocCollection.Id, objDocCollection.Title);
							}
						//---------------------------------------------------------------------------------
						//- Load options for each of the documents that need to be generated
						//---------------------------------------------------------------------------------
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
									//====================================
									//+ Activity_Effort_Workbook
									case enumDocumentTypes.Activity_Effort_Workbook:
										{
										//NOT_AVAILABLE: not currently implemented - Activities and Effort Drivers removed from SharePoint
										break;
										}

									//==================================================
									//+ Client_Requirement_Mapping_Workbook
									case enumDocumentTypes.Client_Requirement_Mapping_Workbook:
										{
										Client_Requirements_Mapping_Workbook objCRMworkbook = new Client_Requirements_Mapping_Workbook();
										objCRMworkbook.DocumentCollectionID = objDocumentCollection.ID;
										objCRMworkbook.DocumentCollectionTitle = objDocumentCollection.Title;
										objCRMworkbook.DocumentStatus = enumDocumentStatusses.New;
										objCRMworkbook.DocumentType = enumDocumentTypes.Client_Requirement_Mapping_Workbook;
										strTemplateURL = GetDocumentTemplate(parDataSet.SDDPdatacontext, "Client Requirements Mapping Workbook");
										switch(strTemplateURL)
											{
										case "None":
											objCRMworkbook.Template = "";
											objCRMworkbook.LogError("The template could not be found.");
											break;
										case "Error":
											objCRMworkbook.Template = "";
											objCRMworkbook.LogError("The template could not be accessed.");
											break;
										default:
											objCRMworkbook.Template = parDataSet.SharePointSiteURL + strTemplateURL;
											break;
											}
										if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
											objCRMworkbook.HyperlinkEdit = true;
										else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
											objCRMworkbook.HyperlinkView = true;

										// The Hierarchical nodes from the Document Collection is not applicable on this Document object.
										objCRMworkbook.SelectedNodes = null;
										// Instead, set the Client Requirements Mapping value
										objCRMworkbook.CRM_Mapping = objDocCollection.Mapping_Id;

										// add the object to the Document Collection's DocumentsWorkbooks to be generated.
										listDocumentWorkbookObjects.Add(objCRMworkbook);
										break;
										}
									//==================================
									//+ Content_Status_Workbook
									case enumDocumentTypes.Content_Status_Workbook:
										{
										Content_Status_Workbook objContentStatus_Workbook = new Content_Status_Workbook();
										objContentStatus_Workbook.DocumentCollectionID = objDocumentCollection.ID;
										objContentStatus_Workbook.DocumentCollectionTitle = objDocumentCollection.Title;
										objContentStatus_Workbook.DocumentStatus = enumDocumentStatusses.New;
										objContentStatus_Workbook.DocumentType = enumDocumentTypes.Content_Status_Workbook;
										strTemplateURL = GetDocumentTemplate(parDataSet.SDDPdatacontext, "Content Status Workbook");
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
											objContentStatus_Workbook.Template = parDataSet.SharePointSiteURL + strTemplateURL;
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
									//===============================================
									//+ Contract_SoW_Service_Description
									case enumDocumentTypes.Contract_SoW_Service_Description:
										{
										Contract_SoW_Service_Description objContractSoWServiceDescription = new Contract_SoW_Service_Description();
										objContractSoWServiceDescription.DocumentCollectionID = objDocumentCollection.ID;
										objContractSoWServiceDescription.DocumentCollectionTitle = objDocumentCollection.Title;
										objContractSoWServiceDescription.DocumentStatus = enumDocumentStatusses.New;
										objContractSoWServiceDescription.DocumentType = enumDocumentTypes.Contract_SoW_Service_Description;
										objContractSoWServiceDescription.IntroductionRichText = objDocCollection.ContractSDIntroduction;
										strTemplateURL = GetDocumentTemplate(parDataSet.SDDPdatacontext, "Contract: Service Description (Appendix F)");
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
											objContractSoWServiceDescription.Template = parDataSet.SharePointSiteURL + strTemplateURL;
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
										if(objDocCollection.SoWSDOptions != null)
											{
											if(ConvertOptionsToList(objDocCollection.SoWSDOptions, ref optionsWorkList)) // conversion is successful
												{
												objContractSoWServiceDescription.TransposeDocumentOptions(ref optionsWorkList);
												}
											else // the conversion failed
												{
												objContractSoWServiceDescription.LogError("Invalid format in the Document Options :. "
													+ "unable to generate the document.");
												//Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
												}
											}
										else  // == Null
											{
											objContractSoWServiceDescription.LogError("No document options were specified - "
												+ "cannot generate blank documents.");
											Console.WriteLine("No document options were selected - cannot generate blank documents.");
											}

										// Add the Hierarchical nodes from the Document Collection obect to the Document object.
										objContractSoWServiceDescription.SelectedNodes = objDocumentCollection.SelectedNodes;
										// add the object to the Document Collection's DocumentsWorkbooks to be generated.
										listDocumentWorkbookObjects.Add(objContractSoWServiceDescription);
										break;
										}
									//==========================================================
									//+ CSD_based_on_Client_Requirements_Mapping
									case enumDocumentTypes.CSD_based_on_Client_Requirements_Mapping:
										{
										CSD_based_on_ClientRequirementsMapping objCSDbasedonCRM = new CSD_based_on_ClientRequirementsMapping();
										objCSDbasedonCRM.DocumentCollectionID = objDocumentCollection.ID;
										objCSDbasedonCRM.DocumentCollectionTitle = objDocumentCollection.Title;
										objCSDbasedonCRM.DocumentStatus = enumDocumentStatusses.New;
										objCSDbasedonCRM.DocumentType = enumDocumentTypes.CSD_based_on_Client_Requirements_Mapping;
										objCSDbasedonCRM.IntroductionRichText = objDocCollection.CSDDocumentIntroduction;
										objCSDbasedonCRM.ExecutiveSummaryRichText = objDocCollection.CSDDocumentExecSummary;
										strTemplateURL = GetDocumentTemplate(parDataSet.SDDPdatacontext, "Client Service Description");
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
											objCSDbasedonCRM.Template = parDataSet.SharePointSiteURL + strTemplateURL;
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
										if(objDocCollection.CSDDocumentBasedOnCRMOptions != null)
											{
											if(ConvertOptionsToList(objDocCollection.CSDDocumentBasedOnCRMOptions, ref optionsWorkList))
												{
												objCSDbasedonCRM.TransposeDocumentOptions(ref optionsWorkList);
												}
											else // the conversion failed
												{
												objCSDbasedonCRM.LogError("Invalid format in the Document Options :. "
													+ "unable to generate the document.");
												//Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
												}
											}
										else  // == Null
											{
											objCSDbasedonCRM.LogError("No document options were specified - cannot generate blank documents.");
											Console.WriteLine("No document options were selected - cannot generate blank documents.");
											}

										// The Hierarchical nodes from the Document Collection is not applicable on this Document object.
										objCSDbasedonCRM.SelectedNodes = null;

										objCSDbasedonCRM.CRM_Mapping = objDocCollection.Mapping_Id;
										// add the object to the Document Collection's DocumentsWorkbooks to be generated.
										listDocumentWorkbookObjects.Add(objCSDbasedonCRM);
										break;
										}
									//=================================
									//+ CSD_Document_DRM_Inline
									case enumDocumentTypes.CSD_Document_DRM_Inline:
										{
										CSD_Document_DRM_Inline objCSDdrmInline = new CSD_Document_DRM_Inline();
										objCSDdrmInline.DocumentCollectionID = objDocumentCollection.ID;
										objCSDdrmInline.DocumentCollectionTitle = objDocumentCollection.Title;
										objCSDdrmInline.DocumentStatus = enumDocumentStatusses.New;
										objCSDdrmInline.DocumentType = enumDocumentTypes.CSD_Document_DRM_Inline;
										objCSDdrmInline.IntroductionRichText = objDocCollection.CSDDocumentIntroduction;
										objCSDdrmInline.ExecutiveSummaryRichText = objDocCollection.CSDDocumentExecSummary;
										strTemplateURL = GetDocumentTemplate(parDataSet.SDDPdatacontext, "Client Service Description");
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
											objCSDdrmInline.Template = parDataSet.SharePointSiteURL + strTemplateURL;
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
										if(objDocCollection.CSDDocumentDRMInlineOptions != null)
											{
											if(ConvertOptionsToList(objDocCollection.CSDDocumentDRMInlineOptions, ref optionsWorkList))
												{
												objCSDdrmInline.TransposeDocumentOptions(ref optionsWorkList);
												}
											else // the conversion failed
												{
												objCSDdrmInline.LogError("Invalid format in the Document Options :. "
													+ "unable to generate the document.");
												//Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
												}
											}
										else  // == Null
											{
											objCSDdrmInline.LogError("No document options were specified - cannot generate blank documents.");
											Console.WriteLine("No document options were selected - cannot generate blank documents.");
											}

										// Add the Hierarchical nodes from the Document Collection obect to the Document object.
										objCSDdrmInline.SelectedNodes = objDocumentCollection.SelectedNodes;
										// add the object to the Document Collection's DocumentsWorkbooks to be generated.
										listDocumentWorkbookObjects.Add(objCSDdrmInline);
										break;
										}
									//====================================
									//+ CSD_Document_DRM_Sections
									case enumDocumentTypes.CSD_Document_DRM_Sections:
										{
										CSD_Document_DRM_Sections objCSDdrmSections = new CSD_Document_DRM_Sections();
										objCSDdrmSections.DocumentCollectionID = objDocumentCollection.ID;
										objCSDdrmSections.DocumentCollectionTitle = objDocumentCollection.Title;
										objCSDdrmSections.DocumentStatus = enumDocumentStatusses.New;
										objCSDdrmSections.DocumentType = enumDocumentTypes.CSD_Document_DRM_Sections;
										objCSDdrmSections.IntroductionRichText = objDocCollection.CSDDocumentIntroduction;
										objCSDdrmSections.ExecutiveSummaryRichText = objDocCollection.CSDDocumentExecSummary;
										strTemplateURL = GetDocumentTemplate(parDataSet.SDDPdatacontext, "Client Service Description");
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
											objCSDdrmSections.Template = parDataSet.SharePointSiteURL + strTemplateURL;
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
										if(objDocCollection.CSDDocumentDRMSectionsOptions != null)
											{
											if(ConvertOptionsToList(objDocCollection.CSDDocumentDRMSectionsOptions, ref optionsWorkList))
												{
												objCSDdrmSections.TransposeDocumentOptions(ref optionsWorkList);
												}
											else // the conversion failed
												{
												objCSDdrmSections.LogError("Invalid format in the Document Options :. "
													+ "unable to generate the document.");
												//Console.WriteLine("Invalid format in the Document Options :. unable to generate the document.");
												}
											}
										else  // == Null
											{
											objCSDdrmSections.LogError("No document options were specified - cannot generate blank documents.");
											Console.WriteLine("No document options were selected - cannot generate blank documents.");
											}

										// Add the Hierarchical nodes from the Document Collection obect to the Document object.
										objCSDdrmSections.SelectedNodes = objDocumentCollection.SelectedNodes;
										// add the object to the Document Collection's DocumentsWorkbooks to be generated.
										listDocumentWorkbookObjects.Add(objCSDdrmSections);
										break;
										}
									//=======================================================
									//+ External_Technology_Coverage_Dashboard
									case enumDocumentTypes.External_Technology_Coverage_Dashboard:
										{
										External_Technology_Coverage_Dashboard_Workbook objExtTechCoverDasboard = new External_Technology_Coverage_Dashboard_Workbook();
										objExtTechCoverDasboard.DocumentCollectionID = objDocumentCollection.ID;
										objExtTechCoverDasboard.DocumentCollectionTitle = objDocumentCollection.Title;
										objExtTechCoverDasboard.DocumentStatus = enumDocumentStatusses.New;
										objExtTechCoverDasboard.DocumentType = enumDocumentTypes.External_Technology_Coverage_Dashboard;
										strTemplateURL = GetDocumentTemplate(parDataSet.SDDPdatacontext, "Technology Roadmap Workbook");
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
											objExtTechCoverDasboard.Template = parDataSet.SharePointSiteURL + strTemplateURL;
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
									//======================================================
									//+ Internal_Technology_Coverage_Dashboard
									case enumDocumentTypes.Internal_Technology_Coverage_Dashboard:
										{
										Internal_Technology_Coverage_Dashboard_Workbook objIntTechCoverDashboard = new Internal_Technology_Coverage_Dashboard_Workbook();
										objIntTechCoverDashboard.DocumentCollectionID = objDocumentCollection.ID;
										objIntTechCoverDashboard.DocumentCollectionTitle = objDocumentCollection.Title;
										objIntTechCoverDashboard.DocumentStatus = enumDocumentStatusses.New;
										objIntTechCoverDashboard.DocumentType = enumDocumentTypes.Internal_Technology_Coverage_Dashboard;
										strTemplateURL = GetDocumentTemplate(parDataSet.SDDPdatacontext, "Technology Roadmap Workbook");
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
											objIntTechCoverDashboard.Template = parDataSet.SharePointSiteURL + strTemplateURL;
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

								//=================================
								//+ Services_Model_Workbook
								case enumDocumentTypes.Services_Model_Workbook:
										{
										Services_Model_Workbook objInternalServicesModelWB = new Services_Model_Workbook();
										objInternalServicesModelWB.DocumentCollectionID = objDocumentCollection.ID;
										objInternalServicesModelWB.DocumentCollectionTitle = objDocumentCollection.Title;
										objInternalServicesModelWB.DocumentStatus = enumDocumentStatusses.New;
										objInternalServicesModelWB.DocumentType = enumDocumentTypes.Services_Model_Workbook;
										strTemplateURL = GetDocumentTemplate(parDataSet.SDDPdatacontext, "Services Model Workbook");
										switch(strTemplateURL)
											{
										case "None":
											objInternalServicesModelWB.Template = "";
											objInternalServicesModelWB.LogError("The workbook template could not be found.");
											break;
										case "Error":
											objInternalServicesModelWB.Template = "";
											objInternalServicesModelWB.LogError("The workbook template could not be accessed.");
											break;
										default:
											objInternalServicesModelWB.Template = parDataSet.SharePointSiteURL + strTemplateURL;
											break;
											}

										//Console.WriteLine("\t Template: {0}", objIntTechCoverDashboard.Template);
										if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_EDIT_Hyperlinks)
											objInternalServicesModelWB.HyperlinkEdit = true;
										else if(objDocumentCollection.HyperLinkOption == enumHyperlinkOptions.Include_VIEW_Hyperlinks)
											objInternalServicesModelWB.HyperlinkView = true;

										// Add the Hierarchical nodes from the Document Collection obect to the Document object.
										objInternalServicesModelWB.SelectedNodes = objDocumentCollection.SelectedNodes;
										// add the object to the Document Collection's DocumentsWorkbooks to be generated.
										listDocumentWorkbookObjects.Add(objInternalServicesModelWB);
										break;
										}

								//===================================
								//+ ISD_Document_DRM_Inline
								case enumDocumentTypes.ISD_Document_DRM_Inline:
										{
										ISD_Document_DRM_Inline objISDdrmInline = new ISD_Document_DRM_Inline();
										objISDdrmInline.DocumentCollectionID = objDocumentCollection.ID;
										objISDdrmInline.DocumentCollectionTitle = objDocumentCollection.Title;
										objISDdrmInline.DocumentStatus = enumDocumentStatusses.New;
										objISDdrmInline.DocumentType = enumDocumentTypes.ISD_Document_DRM_Inline;
										objISDdrmInline.IntroductionRichText = objDocCollection.ISDDocumentIntroduction;
										objISDdrmInline.ExecutiveSummaryRichText = objDocCollection.ISDDocumentExecSummary;
										objISDdrmInline.DocumentAcceptanceRichText = objDocCollection.ISDDocumentAcceptance;
										strTemplateURL = GetDocumentTemplate(parDataSet.SDDPdatacontext, "Internal Service Description");
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
											objISDdrmInline.Template = parDataSet.SharePointSiteURL + strTemplateURL;
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
										if(objDocCollection.ISDDocumentDRMInlineOptions != null)
											{
											if(ConvertOptionsToList(objDocCollection.ISDDocumentDRMInlineOptions, ref optionsWorkList))
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
											Console.WriteLine("No document options were selected - cannot generate blank documents.");
											}

										// Add the Hierarchical nodes from the Document Collection obect to the Document object.
										objISDdrmInline.SelectedNodes = objDocumentCollection.SelectedNodes;
										// add the object to the Document Collection's DocumentsWorkbooks to be generated.
										listDocumentWorkbookObjects.Add(objISDdrmInline);
										break;
										}
									//====================================
									//+ ISD_Document_DRM_Sections
									case enumDocumentTypes.ISD_Document_DRM_Sections:
										{
										ISD_Document_DRM_Sections objISDdrmSections = new ISD_Document_DRM_Sections();
										objISDdrmSections.DocumentCollectionID = objDocumentCollection.ID;
										objISDdrmSections.DocumentCollectionTitle = objDocumentCollection.Title;
										objISDdrmSections.DocumentStatus = enumDocumentStatusses.New;
										objISDdrmSections.DocumentType = enumDocumentTypes.ISD_Document_DRM_Sections;
										objISDdrmSections.IntroductionRichText = objDocCollection.ISDDocumentIntroduction;
										objISDdrmSections.ExecutiveSummaryRichText = objDocCollection.ISDDocumentExecSummary;
										objISDdrmSections.DocumentAcceptanceRichText = objDocCollection.ISDDocumentAcceptance;
										strTemplateURL = GetDocumentTemplate(parDataSet.SDDPdatacontext, "Internal Service Description");
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
											objISDdrmSections.Template = parDataSet.SharePointSiteURL + strTemplateURL;
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
										if(objDocCollection.ISDDocumentDRMSectionsOptions != null)
											{
											if(ConvertOptionsToList(objDocCollection.ISDDocumentDRMSectionsOptions, ref optionsWorkList))
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
											Console.WriteLine("No document options were selected - cannot generate blank documents.");
											}

										// Add the Hierarchical nodes from the Document Collection obect to the Document object.
										objISDdrmSections.SelectedNodes = objDocumentCollection.SelectedNodes;
										// add the object to the Document Collection's DocumentsWorkbooks to be generated.
										listDocumentWorkbookObjects.Add(objISDdrmSections);
										break;
										}
									//======================================
									//+ Pricing_Addendum_Document
									case enumDocumentTypes.Pricing_Addendum_Document:
										{
										//NOT_AVAILABLE: not currently implemented - Activities and Effort Drivers removed from SharePoint.
										break;
										}
									//====================================
									// RACI_Matrix_Workbook_per_Deliverable
									case enumDocumentTypes.RACI_Matrix_Workbook_per_Deliverable:
										{
										RACI_Matrix_Workbook_per_Deliverable objRACIperDeliverable = new RACI_Matrix_Workbook_per_Deliverable();
										objRACIperDeliverable.DocumentCollectionID = objDocumentCollection.ID;
										objRACIperDeliverable.DocumentCollectionTitle = objDocumentCollection.Title;
										objRACIperDeliverable.DocumentStatus = enumDocumentStatusses.New;
										objRACIperDeliverable.DocumentType = enumDocumentTypes.RACI_Matrix_Workbook_per_Deliverable;
										strTemplateURL = GetDocumentTemplate(parDataSet.SDDPdatacontext, "RACI Matrix Workbook");
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
											objRACIperDeliverable.Template = parDataSet.SharePointSiteURL + strTemplateURL;
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
									//=================================
									//+ RACI_Workbook_per_Role
									case enumDocumentTypes.RACI_Workbook_per_Role:
										{
										RACI_Workbook_per_Role objRACIperRole = new RACI_Workbook_per_Role();
										objRACIperRole.DocumentCollectionID = objDocumentCollection.ID;
										objRACIperRole.DocumentCollectionTitle = objDocumentCollection.Title;
										objRACIperRole.DocumentStatus = enumDocumentStatusses.New;
										objRACIperRole.DocumentType = enumDocumentTypes.RACI_Workbook_per_Role;
										strTemplateURL = GetDocumentTemplate(parDataSet.SDDPdatacontext, "RACI Workbook");
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
											objRACIperRole.Template = parDataSet.SharePointSiteURL + strTemplateURL;
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
									//+ Service_Framework_Document_DRM_inline
									case enumDocumentTypes.Service_Framework_Document_DRM_inline:
										{
										Services_Framework_Document_DRM_Inline objSFdrmInline = new Services_Framework_Document_DRM_Inline();
										objSFdrmInline.DocumentCollectionID = objDocumentCollection.ID;
										objSFdrmInline.DocumentCollectionTitle = objDocumentCollection.Title;
										objSFdrmInline.DocumentStatus = enumDocumentStatusses.New;
										objSFdrmInline.DocumentType = enumDocumentTypes.Service_Framework_Document_DRM_inline;
										objSFdrmInline.IntroductionRichText = objDocCollection.ISDDocumentIntroduction;
										objSFdrmInline.ExecutiveSummaryRichText = objDocCollection.ISDDocumentExecSummary;
										objSFdrmInline.DocumentAcceptanceRichText = objDocCollection.ISDDocumentAcceptance;
										strTemplateURL = GetDocumentTemplate(parDataSet.SDDPdatacontext, "Services Framework Description");
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
											objSFdrmInline.Template = parDataSet.SharePointSiteURL + strTemplateURL;
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
										if(objDocCollection.ISDDocumentDRMInlineOptions != null)
											{
											if(ConvertOptionsToList(objDocCollection.ISDDocumentDRMInlineOptions, ref optionsWorkList))
												{
												objSFdrmInline.TransposeDocumentOptions(ref optionsWorkList);
												}
											else // the conversion failed
												{
												objSFdrmInline.LogError("Invalid format in the Document Options :. "
													+ "unable to generate the document.");
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
									//========================================================
									//+ Service_Framework_Document_DRM_sections
									case enumDocumentTypes.Service_Framework_Document_DRM_sections:
										{
										Services_Framework_Document_DRM_Sections objSFdrmSections = new Services_Framework_Document_DRM_Sections();
										objSFdrmSections.DocumentCollectionID = objDocumentCollection.ID;
										objSFdrmSections.DocumentCollectionTitle = objDocumentCollection.Title;
										objSFdrmSections.DocumentStatus = enumDocumentStatusses.New;
										objSFdrmSections.DocumentType = enumDocumentTypes.Service_Framework_Document_DRM_sections;
										objSFdrmSections.IntroductionRichText = objDocCollection.ISDDocumentIntroduction;
										objSFdrmSections.ExecutiveSummaryRichText = objDocCollection.ISDDocumentExecSummary;
										objSFdrmSections.DocumentAcceptanceRichText = objDocCollection.ISDDocumentAcceptance;
										strTemplateURL = GetDocumentTemplate(parDataSet.SDDPdatacontext, "Services Framework Description");
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
											objSFdrmSections.Template = parDataSet.SharePointSiteURL 
												+ parDataSet.SharePointSiteSubURL + strTemplateURL;
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
										if(objDocCollection.ISDDocumentDRMSectionsOptions != null)
											{
											if(ConvertOptionsToList(objDocCollection.ISDDocumentDRMSectionsOptions, ref optionsWorkList))
												{
												objSFdrmSections.TransposeDocumentOptions(ref optionsWorkList);
												}
											else
												{
												objSFdrmSections.LogError("Invalid format in the Document Options :. "
													+ "unable to generate the document.");
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
							Console.WriteLine(" Document Collection: {0} successfully loaded...", objDocumentCollection.ID);
							}
						objDocumentCollection.DetailComplete = true;
						} //else // populate the DocumentCollectionObject
					} // Loop of the For Each DocColsToGenerate
				} // end of Try
			catch(DataServiceClientException exc)
				{
				Console.Beep(2500, 750);
				strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " 
					+ parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL  
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult + "\nStatusCode: " + exc.StatusCode
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
				Console.WriteLine(strExceptionMessage);
				throw new GeneralException(strExceptionMessage);
				}
			catch(DataServiceQueryException exc)
				{
				Console.Beep(2500, 750);
				strExceptionMessage = "*** Exception ERROR ***: Cannot access site: "
					+ parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult 
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
				Console.WriteLine(strExceptionMessage);
				throw new GeneralException(strExceptionMessage);
				}
			catch(DataServiceRequestException exc)
				{
				Console.Beep(2500, 750);
				strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " 
					+ parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult 
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
				Console.WriteLine(strExceptionMessage);
				throw new GeneralException(strExceptionMessage);
				}
			catch(DataServiceTransportException exc)
				{
				Console.Beep(2500, 750);
				strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " 
					+ parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
				Console.WriteLine(strExceptionMessage);
				throw new GeneralException(strExceptionMessage);
				}
			catch(Exception exc)
				{
				Console.Beep(2500, 750);

				if(exc.HResult == -2146330330)
					{
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: "
						+ parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					}
				else if(exc.HResult == -2146233033)
					{
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: "
						+ parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					}
				else
					{
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: "
						+ parDataSet.SharePointSiteURL + parDataSet.SharePointSiteSubURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					}

				Console.WriteLine(strExceptionMessage);
				throw new GeneralException(strExceptionMessage);
				}

			} // end of Method
		

		//++ GetDocumentTemplate
		/// <summary>
		/// This method finds the relevant Document Template and if found returns the path to the template URL in a string 
		/// </summary>
		/// <param name="parDataContext">Pass the DataContext for the Template SharePoint Site.</param>
		/// <param name="parTemplateType">Pass the Template Type that need to be found as a string.</param>
		/// <returns></returns>
		public static string GetDocumentTemplate(
			DesignAndDeliveryPortfolioDataContext parDataContext, 
			string parTemplateType)
			{
			string returnPath = "";
			try
				{
				//var DocumentTemplates = parDataContext.DocumentTemplates;
				var rsTemplate = from docTemplate in parDataContext.DocGeneratorTemplates
							where docTemplate.TemplateTypeValue == parTemplateType
							select docTemplate;
				Console.WriteLine("\t\t\t + {0} templates found.", rsTemplate.Count());
				if(rsTemplate != null)
					{
					foreach(var templateEntry in rsTemplate)
						{
						//Console.WriteLine("\t\t\t - {0} - {1} [{2}]", tpl.Id, tpl.Title, tpl.TemplateTypeValue);
						if(templateEntry.TemplateTypeValue == parTemplateType)
							{
							returnPath =  templateEntry.Path + "/" + templateEntry.Name;
							break;
							}
						}
					}
				else // No template was found
					{
					Console.WriteLine("Error occurred: Template could not be located.");
					returnPath = "None";
					}
				}
			catch(DataServiceQueryException exc)
				{
				Console.WriteLine("Error occurred: Template could not be located\n {0} \n {1}", exc.Message, exc.Data);
				returnPath = "None";
				}

			catch(Exception ex)
				{
				Console.WriteLine("Error occurred: {0} \n {1}", ex.Message, ex.Data);
				returnPath = "Error";
				}
			
			return returnPath;

			}

		//++ConvertOptionsToList
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


		//++ PrepareStringForEnum
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