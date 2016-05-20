using System;
using System.Collections.Generic;
using System.Data.Services.Client;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using DocGeneratorCore.SDDPServiceReference;

namespace DocGeneratorCore
	{
	//++ MainController Class
	public class MainController
		{

		public bool SuccessfulSentEmail{get; set;}

		public bool SuccessfullUpdatedDocCollection{get; set;}

		public string EmailBodyText{get; set;}

		public List<DocumentCollection> DocumentCollectionsToGenerate{get; set;}

		//-- -----------------------------------------------------------------------------------------------------------
		//- Object Variables
		//- CountdownEvent controller that is used for the Main Thread to wait the DataSet is build
		public static CountdownEvent mainThreadCountDown = new CountdownEvent(1);

		private string strErrorMessage = String.Empty;  //- a string that is used to construct eroror message that are recorded and displayed

		//++MainProcess method
		public void MainProcess(ref CompleteDataSet parDataSet)
			{
			Console.WriteLine("Begin to execute the MainProcess in the DocGeneratorCore module");
			//- Check if a dataset was passed into the app with **parDataset** parameter.
			//- If it was not passed, setup the DataContext with which to obtain data from SharePoint...
			if(parDataSet == null)
				{
				parDataSet = new CompleteDataSet();
				parDataSet.SDDPdatacontext = new DesignAndDeliveryPortfolioDataContext(new
					Uri(Properties.AppResources.SharePointSiteURL + Properties.AppResources.SharePointRESTuri));

				parDataSet.SDDPdatacontext.Credentials = new NetworkCredential(
					userName: Properties.AppResources.DocGenerator_AccountName,
					password: Properties.AppResources.DocGenerator_Account_Password,
					domain: Properties.AppResources.DocGenerator_AccountDomain);
				parDataSet.SDDPdatacontext.MergeOption = MergeOption.NoTracking;

				parDataSet.LastRefreshedOn = new DateTime(2000, 1, 1, 0, 0, 0);
				parDataSet.RefreshingDateTimeStamp = DateTime.UtcNow;
				parDataSet.IsDataSetComplete = false;
				}

			//+Check and Populate the Dataset
			//- To ensure optimal Document Generation performance, the complete dataset is loaded into memory.
			//- Check if the **complete DataSet** is ready *_AND_* not older than **60 seconds** before beginning to generate the documents.
			try
				{
				//- If the DataSet is **incomplete**, rebuild the dataset from scratch...
				if(parDataSet.IsDataSetComplete == false)
					{
					//- backdate the **LastRefreshedOn** property to a point in the past to ensure the complete dataset is loaded
					parDataSet.LastRefreshedOn = new DateTime(2000, 1, 1, 0, 0, 0);
					parDataSet.RefreshingDateTimeStamp = DateTime.UtcNow;
					parDataSet.IsDataSetComplete = false;
					}
				else
					{
					//- Check if the current **Complete Dataset** is older than 2 minutes, if it is, refesh any changes in the dataset
					TimeSpan timeDifference = DateTime.UtcNow.Subtract(parDataSet.LastRefreshedOn);
					if(timeDifference.TotalSeconds > 60 * 2)
						{
						parDataSet.RefreshingDateTimeStamp = DateTime.UtcNow;
						parDataSet.IsDataSetComplete = false;
						}
					}
				//- if the dataset is incomplete or outdated, build/and or refresh it
				if(parDataSet.IsDataSetComplete == false)
					{
					//- ---------------------------------------------------------------------------------------------------------------------
					//- Because the parDataSet was passed into the app by reference, it cannot be *passed* in threading instructions
					//- Therefore create a temporary DataSet and build it with multi-threads and then assign the new set to the parDataSet...
					//- ---------------------------------------------------------------------------------------------------------------------
					CompleteDataSet objDataSet = parDataSet;
					CompleteDataSet.threadCountDown.Reset(7);
					//- --------------------------------------------------------------------------------------------------
					//- Launch the **6 Threads** to build the Complete DataSet - concurrency means improved performance
					//- --------------------------------------------------------------------------------------------------
					Thread tThread1 = new Thread(() => objDataSet.PopulateBaseDataObjects());
					Thread tThread2 = new Thread(() => objDataSet.PopulateBaseDataObjects());
					Thread tThread3 = new Thread(() => objDataSet.PopulateBaseDataObjects());
					Thread tThread4 = new Thread(() => objDataSet.PopulateBaseDataObjects());
					Thread tThread5 = new Thread(() => objDataSet.PopulateBaseDataObjects());
					Thread tThread6 = new Thread(() => objDataSet.PopulateBaseDataObjects());
					Thread tThread7 = new Thread(() => objDataSet.PopulateBaseDataObjects());
					//- Set the **Name** for each **Thread** because the *PopulateDataSet* method use the names to direct the threads.
					tThread1.Name = "Data1";
					tThread1.Start();
					tThread2.Name = "Data2";
					tThread2.Start();
					tThread3.Name = "Data3";
					tThread3.Start();
					tThread4.Name = "Data4";
					tThread4.Start();
					tThread5.Name = "Data5";
					tThread5.Start();
					tThread6.Name = "Data6";
					tThread6.Start();
					tThread7.Name = "Data7";
					tThread7.Start();

					//- Pass the CurrentThread as the **Synchronisation Thread** which has to wait until all the DataSet Population threads completed,
					//- before it declare the DataSet to be "**Complete**" by setting the **IsDataSetComplete** property.
					objDataSet.PopulateBaseDataObjects();

					//- After populating the **objDataset**, chek if is complete.
					if(objDataSet.IsDataSetComplete == false)
						{//- Send an e-mail to Technical Support if the DataSet is not complete...
						this.EmailBodyText = "Please investigate, the DocGenerator was unable to successfully load the Complete DataSet from SharePoint.";
						Console.WriteLine("Error: ***" + this.EmailBodyText + "***");
						SuccessfulSentEmail = eMail.SendEmail(
							parRecipient: Properties.AppResources.Email_Technical_Support,
							parSubject: "SDDP: Unexpected DocGenerator Error occurred.)",
							parBody: EmailBodyText,
							parSendBcc: false);
						goto Procedure_Ends;
						}
					else
						{
						//- The **objDataset** is complete, therefore assign it to the **parDataset**
						parDataSet = objDataSet;
						}
					}
				}
			catch(GeneralException exc)
				{
				parDataSet.IsDataSetComplete = false;
				this.EmailBodyText = "Exception Error occurred during the loading of the complete DataSet: " 
					+ exc.Message + "\n HResult: " + exc.HResult + "\nInnerexception : " + exc.InnerException;
				Console.WriteLine(this.EmailBodyText);
				// Send the e-mail Technical Support
				SuccessfulSentEmail = eMail.SendEmail(
					parRecipient: Properties.AppResources.Email_Technical_Support,
					parSubject: "SDDP: Unexpected DocGenerator Error occurred.)",
					parBody: EmailBodyText,
					parSendBcc: false);
				goto Procedure_Ends;
				}

//---g
			//+ Obtain the details of the document Collections to be generated..
			string strDocWkbType = string.Empty;
			Console.WriteLine("{0} Document Collections to generate...", this.DocumentCollectionsToGenerate.Count);
			//
			List<DocumentCollection> listDocumentCollections;
			if(this.DocumentCollectionsToGenerate == null)
				listDocumentCollections = new List<DocumentCollection>();
			else
				listDocumentCollections = this.DocumentCollectionsToGenerate;

			// Obtain the details of the Document Collections that need to be processed
			try
				{
				DocumentCollection.PopulateCollections(parSDDPdatacontext: parDataSet.SDDPdatacontext,
					parDocumentCollectionList: ref listDocumentCollections);
				}
			catch(GeneralException exc)
				{
				this.EmailBodyText = "Exception Error occurred: " + exc.Message + "\n HResult: " + exc.HResult + "\nInnerexception : " + exc.InnerException;
				Console.WriteLine(this.EmailBodyText);
				// Send the e-mail Technical Support
				SuccessfulSentEmail = eMail.SendEmail(
					parRecipient: Properties.AppResources.Email_Technical_Support,
					parSubject: "SDDP: Unexpected DocGenerator(s) Error occurred.)",
					parBody: EmailBodyText,
					parSendBcc: false);
				goto Procedure_Ends;
				}

//===g
			//+ Sequencially process each of the **DocumentCollections**
			//- =========================================
			//- Process each of the document collections.
			//- =========================================
			try
				{
				//- The Complete DataSet is in Memory, now process each Document Collection Entry
				foreach(DocumentCollection objDocCollection in DocumentCollectionsToGenerate)
					{
					Console.WriteLine("\r\nReady to generate Document Collection: {0} - {1}", objDocCollection.ID.ToString(),
						objDocCollection.Title);
					objDocCollection.UnexpectedErrors = false;

					//Prepare the E-mail Header that will be send to the user...
					EmailBodyText = "Good day,\n\nHerewith the generated document(s) that you requested from the Service Design and Delivery Portfolio "
						+ "as entry\n" + objDocCollection.ID + " - " + objDocCollection.Title + " in the Document Collections Library";

					//-- Check if any documents were specified to be generated, if nothing try to send an e-mail
					if(objDocCollection.Document_and_Workbook_objects == null
					|| objDocCollection.Document_and_Workbook_objects.Count() == 0)
						{
						//- Prepare and send an e-mail to the user...
						if(objDocCollection.NotificationEmail != null && objDocCollection.NotificationEmail != "None")
							{
							EmailBodyText = "\nYou submitted the Document Collection without specifing any document(s) to be generated.";
							SuccessfulSentEmail = eMail.SendEmail(
							parRecipient: objDocCollection.NotificationEmail,
							parSubject: "SDDP: Generated Document(s)",
							parBody: EmailBodyText);
							}
						//- Update the Document Collection Entry, else it will be continually processed, until the **Generation Status** is not blank or Pending.
						this.SuccessfullUpdatedDocCollection = objDocCollection.UpdateGenerateStatus(
								parGenerationStatus: enumGenerationStatus.Completed);

						if(this.SuccessfullUpdatedDocCollection)
							Console.WriteLine("Update Document Collection Status to 'Completed' was successful.");
						else
							Console.WriteLine("Update Document Collection Status to 'Completed' was unsuccessful.");
						}
					else
						{//+ Process each of the documents in the DocumentCollection
						foreach(dynamic objDocumentWorkbook in objDocCollection.Document_and_Workbook_objects)
							{
							Console.WriteLine("\r Generate ObjectType: {0}", objDocumentWorkbook.ToString());
							strDocWkbType = objDocumentWorkbook.ToString();
							strDocWkbType = strDocWkbType.Substring(strDocWkbType.IndexOf(".") + 1, 
								(strDocWkbType.Length - strDocWkbType.IndexOf(".") - 1));
							switch(strDocWkbType)
								{
//---g
								//+ Client_Requirements_Mapping_Workbook
								case ("Client_Requirements_Mapping_Workbook"):
									{
									//- Prepare to generate the Document
									Client_Requirements_Mapping_Workbook objCRMworkbook = objDocumentWorkbook;

									if(objCRMworkbook.ErrorMessages == null)
										objCRMworkbook.ErrorMessages = new List<string>();
									//- Execute the generation instruction
									objCRMworkbook.Generate(parDataSet: parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);

									// -Validate and finalise the document generation
									if(objCRMworkbook.DocumentStatus == enumDocumentStatusses.Done)
										{//+ Done - the document was generated and uploaded
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ");
										EmailBodyText += "\n       The workbook is stored at this url: " + objCRMworkbook.URLonSharePoint;
										//- if there were content errors, add those to the client message
										if(objCRMworkbook.ErrorMessages.Count() > 0)
											{//- include them in the message.
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objCRMworkbook.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following content issues occurred:";
											foreach(string errorEntry in objCRMworkbook.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											EmailBodyText += ", generated without any provisions.";
											}
										}
									else if(objCRMworkbook.DocumentStatus == enumDocumentStatusses.Error)
										{// there was an error that prevented the document's successful completion
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ") 
											+ " could NOT be generate for the following reason(s):";
										if(objCRMworkbook.ErrorMessages.Count() > 0)
											{//- include the erros in the message.
											foreach(string errorEntry in objCRMworkbook.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										}
									else if(objCRMworkbook.DocumentStatus == enumDocumentStatusses.FatalError)
										{//+ an Unexpected FATAL error occurred
										objDocCollection.UnexpectedErrors = true;
										objCRMworkbook.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objCRMworkbook.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									//- Place break between the different documents in the e-mail message
									EmailBodyText += "\n\n";
									break;
									}
								//+ Content_Status_Workbook
								case ("Content_Status_Workbook"):
									{
									//- Prepare to generate the Document
									Content_Status_Workbook objContentStatusWB = objDocumentWorkbook;

									if(objContentStatusWB.ErrorMessages == null)
										objContentStatusWB.ErrorMessages = new List<string>();
									objContentStatusWB.Generate(parDataSet: parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);

									// -Validate and finalise the document generation
									if(objContentStatusWB.DocumentStatus == enumDocumentStatusses.Done)
										{//+ Done - the document was generated and uploaded
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ");
										EmailBodyText += "\n       The workbook is stored at this url: " + objContentStatusWB.URLonSharePoint;
										//- if there were content errors, add those to the client message
										if(objContentStatusWB.ErrorMessages.Count() > 0)
											{//- include them in the message.
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objContentStatusWB.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following content issues occurred:";
											foreach(string errorEntry in objContentStatusWB.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											EmailBodyText += ", generated without any provisions.";
											}
										}
									else if(objContentStatusWB.DocumentStatus == enumDocumentStatusses.Error)
										{// there was an error that prevented the document's successful completion
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ") 
											+ " could NOT be generate for the following reason(s):";
										if(objContentStatusWB.ErrorMessages.Count() > 0)
											{//- include the erros in the message.
											foreach(string errorEntry in objContentStatusWB.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										}
									else if(objContentStatusWB.DocumentStatus == enumDocumentStatusses.FatalError)
										{//+ an Unexpected FATAL error occurred
										objDocCollection.UnexpectedErrors = true;
										objContentStatusWB.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objContentStatusWB.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									//- Place break between the different documents in the e-mail message
									EmailBodyText += "\n\n";
									break;
									}

//---g
								//+ Contract_SoW_Service_Description
								case ("Contract_SoW_Service_Description"):
									{
									// Prepare to generate the Document
									Contract_SoW_Service_Description objContractSoW = objDocumentWorkbook;

									if(objContractSoW.ErrorMessages == null)
										objContractSoW.ErrorMessages = new List<string>();

									objContractSoW.Generate(parDataSet: parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);

									//- Validate and finalise the document generation
									if(objContractSoW.DocumentStatus == enumDocumentStatusses.Done)
										{//+ Done - the document was generated and uploaded
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ");
										EmailBodyText += "\n       The document is stored at this url: " + objContractSoW.URLonSharePoint;
										//- if there were content errors, add those to the client message
										if(objContractSoW.ErrorMessages.Count() > 0)
											{//- include them in the message.
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objContractSoW.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following content issues occurred:";
											foreach(string errorEntry in objContractSoW.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											EmailBodyText += ", generated without any provisions.";
											}
										}

									else if(objContractSoW.DocumentStatus == enumDocumentStatusses.Error)
										{//+ there was an error that prevented the document's successful completion
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ") 
											+ " could NOT be generate for the following reason(s):";
										if(objContractSoW.ErrorMessages.Count() > 0)
											{//- include the erros in the message.
											foreach(string errorEntry in objContractSoW.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										}
									else if(objContractSoW.DocumentStatus == enumDocumentStatusses.FatalError)
										{//+ an Unexpected FATAL error occurred
										objDocCollection.UnexpectedErrors = true;
										objContractSoW.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objContractSoW.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									//- Place break between the different documents in the e-mail message
									EmailBodyText += "\n\n";
									break;

									}
//---g
								//+ CSD_based_on_ClientRequirementsMapping
								case ("CSD_based_on_ClientRequirementsMapping"):
									{
									//- Prepare to generate the Document
									CSD_based_on_ClientRequirementsMapping objCSDbasedCRM = objDocumentWorkbook;

									if(objCSDbasedCRM.ErrorMessages == null)
										objCSDbasedCRM.ErrorMessages = new List<string>();

									//- Generate the document...
									objCSDbasedCRM.Generate(parDataSet: parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);

									//- Validate and finalise the document generation
									if(objCSDbasedCRM.DocumentStatus == enumDocumentStatusses.Done)
										{//+ Done - the document was generated and uploaded
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ");
										EmailBodyText += "\n       The document is stored at this url: " + objCSDbasedCRM.URLonSharePoint;
										//- if there were content errors, add those to the client message
										if(objCSDbasedCRM.ErrorMessages.Count() > 0)
											{//- include them in the message.
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objCSDbasedCRM.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following content issues occurred:";
											foreach(string errorEntry in objCSDbasedCRM.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											EmailBodyText += ", generated without any provisions.";
											}
										}

									else if(objCSDbasedCRM.DocumentStatus == enumDocumentStatusses.Error)
										{//+ there was an error that prevented the document's successful completion
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ") 
											+ " could NOT be generate for the following reason(s):";
										if(objCSDbasedCRM.ErrorMessages.Count() > 0)
											{//- include the erros in the message.
											foreach(string errorEntry in objCSDbasedCRM.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										}
									else if(objCSDbasedCRM.DocumentStatus == enumDocumentStatusses.FatalError)
										{//+ an Unexpected FATAL error occurred
										objDocCollection.UnexpectedErrors = true;
										objCSDbasedCRM.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objCSDbasedCRM.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									//- Place break between the different documents in the e-mail message
									EmailBodyText += "\n\n";
									break;
									}
//---g
							//+ CSD_Document_DRM_Inline
							case ("CSD_Document_DRM_Inline"):
									{
									// Prepare to generate the Document
									CSD_Document_DRM_Inline objCSDdrmInline = objDocumentWorkbook;

									if(objCSDdrmInline.ErrorMessages == null)
										objCSDdrmInline.ErrorMessages = new List<string>();

									//- Generate the document...
									objCSDdrmInline.Generate(parDataSet: parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);

									//- Validate and finalise the document generation
									if(objCSDdrmInline.DocumentStatus == enumDocumentStatusses.Done)
										{//+ Done - the document was generated and uploaded
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ");
										EmailBodyText += "\n       The document is stored at this url: " + objCSDdrmInline.URLonSharePoint;
										//- if there were content errors, add those to the client message
										if(objCSDdrmInline.ErrorMessages.Count() > 0)
											{//- include them in the message.
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objCSDdrmInline.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following content issues occurred:";
											foreach(string errorEntry in objCSDdrmInline.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											EmailBodyText += ", generated without any provisions.";
											}
										}

									else if(objCSDdrmInline.DocumentStatus == enumDocumentStatusses.Error)
										{//+ there was an error that prevented the document's successful completion
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ") 
											+ " could NOT be generate for the following reason(s):";
										if(objCSDdrmInline.ErrorMessages.Count() > 0)
											{//- include the erros in the message.
											foreach(string errorEntry in objCSDdrmInline.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										}
									else if(objCSDdrmInline.DocumentStatus == enumDocumentStatusses.FatalError)
										{//+ an Unexpected FATAL error occurred
										objDocCollection.UnexpectedErrors = true;
										objCSDdrmInline.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objCSDdrmInline.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									//- Place break between the different documents in the e-mail message
									EmailBodyText += "\n\n";
									break;
									}
							//---g
							//+ CSD_Document_DRM_Sections
							case ("CSD_Document_DRM_Sections"):
									{
									// Prepare to generate the Document
									CSD_Document_DRM_Sections objCSDdrmSections = objDocumentWorkbook;

									if(objCSDdrmSections.ErrorMessages == null)
										objCSDdrmSections.ErrorMessages = new List<string>();

									//- Generate the document...
									objCSDdrmSections.Generate(parDataSet: parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);

									//- Validate and finalise the document generation
									if(objCSDdrmSections.DocumentStatus == enumDocumentStatusses.Done)
										{//+ Done - the document was generated and uploaded
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ");
										EmailBodyText += "\n       The document is stored at this url: " + objCSDdrmSections.URLonSharePoint;
										//- if there were content errors, add those to the client message
										if(objCSDdrmSections.ErrorMessages.Count() > 0)
											{//- include them in the message.
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objCSDdrmSections.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following content issues occurred:";
											foreach(string errorEntry in objCSDdrmSections.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											EmailBodyText += ", generated without any provisions.";
											}
										}

									else if(objCSDdrmSections.DocumentStatus == enumDocumentStatusses.Error)
										{//+ there was an error that prevented the document's successful completion
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ") + " could NOT be generate for the following reason(s):";
										if(objCSDdrmSections.ErrorMessages.Count() > 0)
											{//- include the erros in the message.
											foreach(string errorEntry in objCSDdrmSections.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										}
									else if(objCSDdrmSections.DocumentStatus == enumDocumentStatusses.FatalError)
										{//+ an Unexpected FATAL error occurred
										objDocCollection.UnexpectedErrors = true;
										objCSDdrmSections.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objCSDdrmSections.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									//- Place break between the different documents in the e-mail message
									EmailBodyText += "\n\n";
									break;

									}
							//---g
							//+ External_Technology_Coverage_Dashboard_Workbook
							case ("External_Technology_Coverage_Dashboard_Workbook"):
									{
									//- Prepare to generate the Document
									External_Technology_Coverage_Dashboard_Workbook objExtTechDashboard = objDocumentWorkbook;

									if(objExtTechDashboard.ErrorMessages == null)
										objExtTechDashboard.ErrorMessages = new List<string>();

									objExtTechDashboard.Generate(parDataSet: parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);

									//- Validate and finalise the document generation
									if(objExtTechDashboard.DocumentStatus == enumDocumentStatusses.Done)
										{//+ Done - the document was generated and uploaded
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ");
										EmailBodyText += "\n       The workbook is stored at this url: " + objExtTechDashboard.URLonSharePoint;
										//- if there were content errors, add those to the client message
										if(objExtTechDashboard.ErrorMessages.Count() > 0)
											{//- include them in the message.
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objExtTechDashboard.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following content issues occurred:";
											foreach(string errorEntry in objExtTechDashboard.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											EmailBodyText += ", generated without any provisions.";
											}
										}

									else if(objExtTechDashboard.DocumentStatus == enumDocumentStatusses.Error)
										{// there was an error that prevented the document's successful completion
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ") + " could NOT be generate for the following reason(s):";
										if(objExtTechDashboard.ErrorMessages.Count() > 0)
											{//- include the erros in the message.
											foreach(string errorEntry in objExtTechDashboard.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										}
									else if(objExtTechDashboard.DocumentStatus == enumDocumentStatusses.FatalError)
										{//+ an Unexpected FATAL error occurred
										objDocCollection.UnexpectedErrors = true;
										objExtTechDashboard.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objExtTechDashboard.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									//- Place break between the different documents in the e-mail message
									EmailBodyText += "\n\n";
									break;
									}
							//---g
							//+ Internal_Technology_Coverage_Dashboard_workbook
							case ("Internal_Technology_Coverage_Dashboard_Workbook"):
									{
									//- Prepare to generate the Document
									Internal_Technology_Coverage_Dashboard_Workbook objIntTechDashboard = objDocumentWorkbook;
									if(objIntTechDashboard.ErrorMessages == null)
										objIntTechDashboard.ErrorMessages = new List<string>();

									//- Generate the document...
									objIntTechDashboard.Generate(parDataSet: parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);

									//- Validate and finalise the document generation
									if(objIntTechDashboard.DocumentStatus == enumDocumentStatusses.Done)
										{//+ Done - the document was generated and uploaded
											//- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ");
										EmailBodyText += "\n       The document is stored at this url: " + objIntTechDashboard.URLonSharePoint;
										//- if there were content errors, add those to the client message
										if(objIntTechDashboard.ErrorMessages.Count() > 0)
											{//- include them in the message.
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objIntTechDashboard.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following content issues occurred:";
											foreach(string errorEntry in objIntTechDashboard.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											EmailBodyText += ", generated without any provisions.";
											}
										}

									else if(objIntTechDashboard.DocumentStatus == enumDocumentStatusses.Error)
										{//+ there was an error that prevented the document's successful completion
											//- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ") + " could NOT be generate for the following reason(s):";
										if(objIntTechDashboard.ErrorMessages.Count() > 0)
											{//- include the erros in the message.
											foreach(string errorEntry in objIntTechDashboard.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										}
									else if(objIntTechDashboard.DocumentStatus == enumDocumentStatusses.FatalError)
										{//+ an Unexpected FATAL error occurred
										objDocCollection.UnexpectedErrors = true;
										objIntTechDashboard.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objIntTechDashboard.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									//- Place break between the different documents in the e-mail message
									EmailBodyText += "\n\n";
									break;
									}
						//---g
							//+ ISD_Document_DRM_Inline
							case ("ISD_Document_DRM_Inline"):
									{
									//- Prepare to generate the Document
									ISD_Document_DRM_Inline objISDdrmInline = objDocumentWorkbook;
									//- Check and declare the List of Error Messages before generation begin...
									if(objISDdrmInline.ErrorMessages == null)
										objISDdrmInline.ErrorMessages = new List<string>();
									//- Generate the document...
									objISDdrmInline.Generate(parDataSet: parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);
									//- Validate and finalise the document generation
									if(objISDdrmInline.DocumentStatus == enumDocumentStatusses.Done)
										{//+ Done - the document was generated and uploaded
										//- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ");
										EmailBodyText += "\n       The document is stored at this url: " + objISDdrmInline.URLonSharePoint;
										 //- if there were content errors, add those to the client message
										if(objISDdrmInline.ErrorMessages.Count() > 0)
											{//- include them in the message.
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objISDdrmInline.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following content issues occurred:";
											foreach(string errorEntry in objISDdrmInline.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											EmailBodyText += ", generated without any provisions.";
											}
										}

									else if(objISDdrmInline.DocumentStatus == enumDocumentStatusses.Error)
										{//+ there was an error that prevented the document's successful completion
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ") + " could NOT be generate for the following reason(s):";
										if(objISDdrmInline.ErrorMessages.Count() > 0)
											{//- include the erros in the message.
											foreach(string errorEntry in objISDdrmInline.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										}
									else if(objISDdrmInline.DocumentStatus == enumDocumentStatusses.FatalError)
										{//+ an Unexpected FATAL error occurred
										objDocCollection.UnexpectedErrors = true;
										objISDdrmInline.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objISDdrmInline.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									//- Place break between the different documents in the e-mail message
									EmailBodyText += "\n\n";
									break;
									}
							//---g
							//+ ISD_Document_DRM_Sections
							case ("ISD_Document_DRM_Sections"):
									{
									//- Prepare to generate the Document
									ISD_Document_DRM_Sections objISDdrmSections = objDocumentWorkbook;

									if(objISDdrmSections.ErrorMessages == null)
										objISDdrmSections.ErrorMessages = new List<string>();

									//- Generate the document...
									objISDdrmSections.Generate(parDataSet: parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);
									//- Validate and finalise the document generation
									if(objISDdrmSections.DocumentStatus == enumDocumentStatusses.Done)
										{//+ Done - the document was generated and uploaded
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ");
										EmailBodyText += "\n       The document is stored at this url: " + objISDdrmSections.URLonSharePoint;
										//- if there were content errors, add those to the client message
										if(objISDdrmSections.ErrorMessages.Count() > 0)
											{//- include them in the message.
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objISDdrmSections.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following content issues occurred:";
											foreach(string errorEntry in objISDdrmSections.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											EmailBodyText += ", generated without any provisions.";
											}
										}

									else if(objISDdrmSections.DocumentStatus == enumDocumentStatusses.Error)
										{//+ there was an error that prevented the document's successful completion
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ") + " could NOT be generate for the following reason(s):";
										if(objISDdrmSections.ErrorMessages.Count() > 0)
											{//- include the erros in the message.
											foreach(string errorEntry in objISDdrmSections.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										}
									else if(objISDdrmSections.DocumentStatus == enumDocumentStatusses.FatalError)
										{//+ an Unexpected FATAL error occurred
										objDocCollection.UnexpectedErrors = true;
										objISDdrmSections.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objISDdrmSections.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									//- Place break between the different documents in the e-mail message
									EmailBodyText += "\n\n";
									break;
									}
//---g
							//+ Pricing_Addendum_Document
							case ("Pricing_Addendum_Document"):
									{
									// Prepare to generate the Document
									Pricing_Addendum_Document objPricingAddendum = objDocumentWorkbook;

									if(objPricingAddendum.ErrorMessages == null)
										objPricingAddendum.ErrorMessages = new List<string>();

									//Not currently implemented - Pricing is still WIP
									//- Generate the document...
									//objPricingAddendum.Generate(parDataSet: parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);
									//- Validate and finalise the document generation
									if(objPricingAddendum.DocumentStatus == enumDocumentStatusses.Done)
										{//+ Done - the document was generated and uploaded
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ");
										EmailBodyText += "\n       The document is stored at this url: " + objPricingAddendum.URLonSharePoint;
										//- if there were content errors, add those to the client message
										if(objPricingAddendum.ErrorMessages.Count() > 0)
											{//- include them in the message.
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objPricingAddendum.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following content issues occurred:";
											foreach(string errorEntry in objPricingAddendum.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											EmailBodyText += ", generated without any provisions.";
											}
										}

									else if(objPricingAddendum.DocumentStatus == enumDocumentStatusses.Error)
										{//+ there was an error that prevented the document's successful completion
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ") 
											+ " could NOT be generate for the following reason(s):";
										if(objPricingAddendum.ErrorMessages.Count() > 0)
											{//- include the erros in the message.
											foreach(string errorEntry in objPricingAddendum.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										}
									else if(objPricingAddendum.DocumentStatus == enumDocumentStatusses.FatalError)
										{//+ an Unexpected FATAL error occurred
										objDocCollection.UnexpectedErrors = true;
										objPricingAddendum.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objPricingAddendum.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									//- Place break between the different documents in the e-mail message
									EmailBodyText += "\n\n";
									break;
									}

//---g
							//+ RACI_Matrix_Workbook_per_Deliverable
							case ("RACI_Matrix_Workbook_per_Deliverable"):
									{
									// Prepare to generate the Document
									RACI_Matrix_Workbook_per_Deliverable objRACImatrix = objDocumentWorkbook;

									if(objRACImatrix.ErrorMessages == null)
										objRACImatrix.ErrorMessages = new List<string>();

									//- Generate the document...
									objRACImatrix.Generate(parDataSet: parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);
									//- Validate and finalise the document generation
									if(objRACImatrix.DocumentStatus == enumDocumentStatusses.Done)
										{//+ Done - the document was generated and uploaded
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ");
										EmailBodyText += "\n       The document is stored at this url: " + objRACImatrix.URLonSharePoint;
										//- if there were content errors, add those to the client message
										if(objRACImatrix.ErrorMessages.Count() > 0)
											{//- include them in the message.
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objRACImatrix.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following content issues occurred:";
											foreach(string errorEntry in objRACImatrix.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											EmailBodyText += ", generated without any provisions.";
											}
										}

									else if(objRACImatrix.DocumentStatus == enumDocumentStatusses.Error)
										{//+ there was an error that prevented the document's successful completion
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ") 
											+ " could NOT be generate for the following reason(s):";
										if(objRACImatrix.ErrorMessages.Count() > 0)
											{//- include the erros in the message.
											foreach(string errorEntry in objRACImatrix.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										}
									else if(objRACImatrix.DocumentStatus == enumDocumentStatusses.FatalError)
										{//+ an Unexpected FATAL error occurred
										objDocCollection.UnexpectedErrors = true;
										objRACImatrix.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objRACImatrix.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									//- Place break between the different documents in the e-mail message
									EmailBodyText += "\n\n";
									break;
									}

//---g
							//+ RACI_Workbook_per_Role
							case ("RACI_Workbook_per_Role"):
									{
									//- Prepare to generate the Document
									RACI_Workbook_per_Role objRACIperRole = objDocumentWorkbook;

									if(objRACIperRole.ErrorMessages == null)
										objRACIperRole.ErrorMessages = new List<string>();

									//- Generate the document...
									objRACIperRole.Generate(parDataSet: parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);
									//- Validate and finalise the document generation
									if(objRACIperRole.DocumentStatus == enumDocumentStatusses.Done)
										{//+ Done - the document was generated and uploaded
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ");
										EmailBodyText += "\n       The document is stored at this url: " + objRACIperRole.URLonSharePoint;
										//- if there were content errors, add those to the client message
										if(objRACIperRole.ErrorMessages.Count() > 0)
											{//- include them in the message.
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objRACIperRole.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following content issues occurred:";
											foreach(string errorEntry in objRACIperRole.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											EmailBodyText += ", generated without any provisions.";
											}
										}

									else if(objRACIperRole.DocumentStatus == enumDocumentStatusses.Error)
										{//+ there was an error that prevented the document's successful completion
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ") 
											+ " could NOT be generate for the following reason(s):";
										if(objRACIperRole.ErrorMessages.Count() > 0)
											{//- include the erros in the message.
											foreach(string errorEntry in objRACIperRole.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										}
									else if(objRACIperRole.DocumentStatus == enumDocumentStatusses.FatalError)
										{//+ an Unexpected FATAL error occurred
										objDocCollection.UnexpectedErrors = true;
										objRACIperRole.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objRACIperRole.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									//- Place break between the different documents in the e-mail message
									EmailBodyText += "\n\n";
									break;
									}

//---g
							//+ Service_Framework_document_DRM_Inline
							case ("Services_Framework_Document_DRM_Inline"):
									{
									//- Prepare to generate the Document
									Services_Framework_Document_DRM_Inline objSFdrmInline = objDocumentWorkbook;

									if(objSFdrmInline.ErrorMessages == null)
										objSFdrmInline.ErrorMessages = new List<string>();

									//- Generate the document...
									objSFdrmInline.Generate(parDataSet: parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);
									//- Validate and finalise the document generation
									if(objSFdrmInline.DocumentStatus == enumDocumentStatusses.Done)
										{//+ Done - the document was generated and uploaded
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ");
										EmailBodyText += "\n       The document is stored at this url: " + objSFdrmInline.URLonSharePoint;
										//- if there were content errors, add those to the client message
										if(objSFdrmInline.ErrorMessages.Count() > 0)
											{//- include them in the message.
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objSFdrmInline.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following content issues occurred:";
											foreach(string errorEntry in objSFdrmInline.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											EmailBodyText += ", generated without any provisions.";
											}
										}

									else if(objSFdrmInline.DocumentStatus == enumDocumentStatusses.Error)
										{//+ there was an error that prevented the document's successful completion
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ") 
											+ " could NOT be generate for the following reason(s):";
										if(objSFdrmInline.ErrorMessages.Count() > 0)
											{//- include the erros in the message.
											foreach(string errorEntry in objSFdrmInline.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										}
									else if(objSFdrmInline.DocumentStatus == enumDocumentStatusses.FatalError)
										{//+ an Unexpected FATAL error occurred
										objDocCollection.UnexpectedErrors = true;
										objSFdrmInline.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objSFdrmInline.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									//- Place break between the different documents in the e-mail message
									EmailBodyText += "\n\n";
									break;
									}

//---g
							//+ Services_Framework_Document_DRM_Sections
							case ("Services_Framework_Document_DRM_Sections"):
									{
									//- Prepare to generate the Document
									Services_Framework_Document_DRM_Sections objSFdrmSections = objDocumentWorkbook;

									if(objSFdrmSections.ErrorMessages == null)
										objSFdrmSections.ErrorMessages = new List<string>();

									//- Generate the document...
									objSFdrmSections.Generate(parDataSet: parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);
									//- Validate and finalise the document generation
									if(objSFdrmSections.DocumentStatus == enumDocumentStatusses.Done)
										{//+ Done - the document was generated and uploaded
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ");
										EmailBodyText += "\n       The document is stored at this url: " + objSFdrmSections.URLonSharePoint;
										//- if there were content errors, add those to the client message
										if(objSFdrmSections.ErrorMessages.Count() > 0)
											{//- include them in the message.
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objSFdrmSections.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following content issues occurred:";
											foreach(string errorEntry in objSFdrmSections.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											EmailBodyText += ", generated without any provisions.";
											}
										}
									else if(objSFdrmSections.DocumentStatus == enumDocumentStatusses.Error)
										{//+ there was an error that prevented the document's successful completion
										 //- compose the e-mail section for this document
										EmailBodyText += "\n     * " + strDocWkbType.Replace("_", " ") 
											+ " could NOT be generate for the following reason(s):";
										if(objSFdrmSections.ErrorMessages.Count() > 0)
											{//- include the erros in the message.
											foreach(string errorEntry in objSFdrmSections.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										}
									else if(objSFdrmSections.DocumentStatus == enumDocumentStatusses.FatalError)
										{//+ an Unexpected FATAL error occurred
										objDocCollection.UnexpectedErrors = true;
										objSFdrmSections.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objSFdrmSections.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									//- Place break between the different documents in the e-mail message
									EmailBodyText += "\n\n";
									break;
									}
								} //- switch (objectType)
							} //- foreach(dynamic objDocumentWorkbook in objDocCollection.Documen_and_Workbook_Objects...

						//---g
						//+ Process the Notification via E-mail
						//- Process the Notification via E-mail if the users selected to be notified.
						if(objDocCollection.NotifyMe && objDocCollection.NotificationEmail != null)
							{
							SuccessfulSentEmail = eMail.SendEmail(
							parRecipient: objDocCollection.NotificationEmail,
							parSubject: "SDDP: Generated Document(s)",
							parBody: EmailBodyText);

							if(SuccessfulSentEmail)
								Console.WriteLine("Sending e-mail successfully send to user!");
							else
								Console.WriteLine("*** ERROR *** \n Sending e-mail failed...\n");
							}
						//--------------------------------------------------------------------------------------------------------------------------------
						//- Check if there were unexpected errors and if there were, send an e-mail to the Technical Support team.
						if(objDocCollection.UnexpectedErrors)
							{
							this.SuccessfullUpdatedDocCollection = objDocCollection.UpdateGenerateStatus(
								parGenerationStatus: enumGenerationStatus.Failed);

							if(this.SuccessfullUpdatedDocCollection)
								Console.WriteLine("Update Document Collection Status to 'FAILED' was successful.");
							else
								Console.WriteLine("Update Document Collection Status to 'FAILED' was unsuccessful.");

							//- Prepare the e-mail
							SuccessfulSentEmail = eMail.SendEmail(
								parRecipient: Properties.AppResources.Email_Technical_Support,
								parSubject: "SDDP: Unexpected DocGenerator(s) Error occurred.",
								parBody: EmailBodyText,
								parSendBcc: false);

							if(SuccessfulSentEmail)
								Console.WriteLine("The error e-mail was successfully send to the technical team.");
							else
								Console.WriteLine("The error e-mail to the technical team FAILED!");
							}
						else // there was no UNEXPECTED errors
							{
							this.SuccessfullUpdatedDocCollection = objDocCollection.UpdateGenerateStatus(
								parGenerationStatus: enumGenerationStatus.Completed);

							if(this.SuccessfullUpdatedDocCollection)
								Console.WriteLine("Update Document Collection Status to 'Completed' was successful.");
							else
								Console.WriteLine("Update Document Collection Status to 'Completed' was unsuccessful.");
							}
						} // end if ...Count() > 0
					} // foreach(DocumentCollection objDocCollection in docCollectionsToGenerate)
				Console.WriteLine("\nDocuments for {0} Document Collection(s) were Generated.", DocumentCollectionsToGenerate.Count);
				}// end try

			catch(DataServiceTransportException exc)
				{
				if(exc.Message.Contains("timed out"))
					{
					EmailBodyText = "The data connection to SharePoint timed out - and the documents could not be generated..." +
						"The DocGenerator will retry to generate the document...";
					Console.WriteLine(EmailBodyText);
					}
				else
					{
					EmailBodyText = "Exception Error: " + exc.Message + "\n HResult: " + exc.HResult + "\nInnerexception : " + exc.InnerException;
					Console.WriteLine(EmailBodyText);
					}

				// Send the e-mail Technical Support
				SuccessfulSentEmail = eMail.SendEmail(
					parRecipient: Properties.AppResources.Email_Technical_Support,
					parSubject: "SDDP: Unexpected DocGenerator(s) Error occurred.)",
					parBody: EmailBodyText,
					parSendBcc: false);
				}
			catch(Exception exc)
				{
				EmailBodyText = EmailBodyText = "Exception Error: " + exc.Message + "\n HResult: " + exc.HResult + "\nInnerexception : " + exc.InnerException;
				Console.WriteLine(EmailBodyText);
				// Send the e-mail Technical Support
				SuccessfulSentEmail = eMail.SendEmail(
					parRecipient: Properties.AppResources.Email_Technical_Support,
					parSubject: "SDDP: Unexpected DocGenerator(s) Error occurred.)",
					parBody: EmailBodyText,
					parSendBcc: false);
				}

Procedure_Ends:
			Console.WriteLine("end of MainController in DocGeneratorCore.");
			return;
			} // end of method
		} // end of class
	} // end of Namespace