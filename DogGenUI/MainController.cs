using System;
using System.Collections.Generic;
using System.Data.Services.Client;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using DocGeneratorCore.SDDPServiceReference;
using DocGeneratorCore.Database.Classes;
using DocGeneratorCore.Database.Functions;

namespace DocGeneratorCore
	{
	#region Enumerations
	public enum enumPlatform
		{
		Development,
		QualityAssurance,
		Production
		}
	#endregion

	#region Classes

	//++ MainController Class
	/// <summary>
	/// This MainController controls the processing of the DocGenerator by means of the MainProcess.
	/// </summary>
	public class MainController
		{
#region Variables
		//- Object Variables
		//- CountdownEvent controller that is used for the Main Thread to wait the DataSet is build
		public static CountdownEvent mainThreadCountDown = new CountdownEvent(1);
		//- a string that is used to construct eroror message that are recorded and displayed
		private string strErrorMessage = String.Empty;
		#endregion

//===G
#region Properties
		public bool SuccessfulSentEmail{get; set;}
		public bool SuccessfullUpdatedDocCollection{get; set;}
		public List<DocumentCollection> DocumentCollectionsToGenerate{get; set;}

		#endregion

		//++MainProcess method
		public void MainProcess(ref CompleteDataSet parDataSet)
			{
			Console.WriteLine("Begin to execute the MainProcess in the DocGeneratorCore module - {0}", DateTime.UtcNow);
			Stopwatch objStopWatch1 = Stopwatch.StartNew();

			//-|Define the Email objects which is used to send confirmation and technical Emails
			eMail objTechnicalEmailgeneral = new eMail();
			//-|Set the currentHostname.
			Properties.Settings.Default.CurrentDatabaseHost = Dns.GetHostName();
			//-|Check if a dataset was passed into the app with **parDataset** parameter.
			//-|Keep in mind that the Dataset must change if the platform changed.
			if (parDataSet == null)
				{//- If the dataset is **Null** default to the **PRODUCTION** platform.
				parDataSet = new CompleteDataSet();
				parDataSet.LastRefreshedOn = new DateTime(2000, 1, 1, 0, 0, 0);
				parDataSet.RefreshingDateTimeStamp = DateTime.UtcNow;
				parDataSet.IsDataSetPopulated = false;
				}

			//-|Check if the required Platform correlates with the current platform for which DataSet is populated
			if (parDataSet.DatasetPlatform.ToString() != Properties.Settings.Default.CurrentPlatform)
				{
				switch (parDataSet.DatasetPlatform.ToString().ToUpper())
					{
				case "DEVELOPMENT":
					Properties.Settings.Default.CurrentPlatform = enumPlatform.Development.ToString();
					Properties.Settings.Default.CurrentDatabaseLocation = Properties.Settings.Default.DatabaseLocationDEV;
					Properties.Settings.Default.CurrentSDDPwebReference = Properties.Settings.Default.SDDPwebReferenceDEV;
					Properties.Settings.Default.CurrentURLSharePoint = Properties.Settings.Default.URLSharePointDEV;
					break;
				case "QUALITYASSURANCE":
					Properties.Settings.Default.CurrentPlatform = enumPlatform.QualityAssurance.ToString();
					Properties.Settings.Default.CurrentDatabaseLocation = Properties.Settings.Default.DatabaseLocationQA;
					Properties.Settings.Default.CurrentSDDPwebReference = Properties.Settings.Default.SDDPwebReferenceQA;
					Properties.Settings.Default.CurrentURLSharePoint = Properties.Settings.Default.URLSharePointQA;
					break;
				case "PRODUCTION":
					Properties.Settings.Default.CurrentPlatform = enumPlatform.Production.ToString();
					Properties.Settings.Default.CurrentDatabaseLocation = Properties.Settings.Default.DatabaseLocationPROD;
					Properties.Settings.Default.CurrentSDDPwebReference = Properties.Settings.Default.SDDPwebReferencePROD;
					Properties.Settings.Default.CurrentURLSharePoint = Properties.Settings.Default.URLSharePointPROD;

					break;
					}
				}
			Properties.Settings.Default.CurrentURLSharePointSitePortion = Properties.Settings.Default.URLSharePointSitePortion;
			Properties.Settings.Default.CurrentURLSharePoint = Properties.Settings.Default.CurrentURLSharePoint;
			

			//- Create a new DataContext if the *SDDPdatacontext* in **parDataSet** is null
			if(parDataSet.SDDPdatacontext == null)
				{
				parDataSet.SDDPdatacontext = new DesignAndDeliveryPortfolioDataContext(new
					Uri(Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion + Properties.AppResources.SharePointRESTuri));

				parDataSet.SDDPdatacontext.Credentials = new NetworkCredential(
						userName: Properties.AppResources.DocGenerator_AccountName,
						password: Properties.AppResources.DocGenerator_Account_Password,
						domain: Properties.AppResources.DocGenerator_AccountDomain);
				}
			parDataSet.SDDPdatacontext.MergeOption = MergeOption.NoTracking;

			//+Check and Populate the Dataset
			//- To ensure optimal Document Generation performance, the complete dataset is loaded into memory.
			//- Check if the **complete DataSet** is ready *_AND_* not older than **60 seconds** before beginning to generate the documents.
			try
				{
				//- If the DataSet is **incomplete**, rebuild the dataset from scratch...
				if(parDataSet.IsDataSetPopulated == false)
					{
					//- backdate the **LastRefreshedOn** property to a point in the past to ensure the complete dataset is loaded
					parDataSet.LastRefreshedOn = new DateTime(2000, 1, 1, 0, 0, 0);
					parDataSet.RefreshingDateTimeStamp = DateTime.UtcNow;
					parDataSet.IsDataSetPopulated = false;
					}
				else
					{
					//- Check if the current **Complete Dataset** is older than 3 minutes, if it is, refesh any changes in the dataset
					TimeSpan timeDifference = DateTime.UtcNow.Subtract(parDataSet.LastRefreshedOn);
					if(timeDifference.TotalSeconds > 180)
						{
						parDataSet.RefreshingDateTimeStamp = DateTime.UtcNow;
						parDataSet.IsDataSetPopulated = false;
						}
					}
				//- if the dataset is incomplete or outdated, build/and or refresh it
				if(parDataSet.IsDataSetPopulated == false)
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

					//- After populating the **objDataset**, check if is complete.
					if(objDataSet.IsDataSetPopulated == false)
						{//- Send an e-mail to Technical Support if the DataSet is not complete...
						strErrorMessage = "Please investigate, the DocGenerator was unable to successfully load the Complete DataSet from SharePoint.";
						Console.WriteLine("Error: ***" + strErrorMessage + "***");
						objTechnicalEmailgeneral.TechnicalEmailModel.MessageLines.Add(strErrorMessage);
						if(objTechnicalEmailgeneral.ComposeHTMLemail(parEmailType: enumEmailType.TechnicalSupport))
							{//-	 Only send the message if the HTML e-mail compiled successfully
							SuccessfulSentEmail = objTechnicalEmailgeneral.SendEmail(
							parDataSet: ref parDataSet,
							parReceipient: Properties.AppResources.Email_Technical_Support,
							parSubject: "SDDP: DocGenerator is experiencing and issue.)",
							parSendBcc: false);
							}
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
				parDataSet.IsDataSetPopulated = false;
				strErrorMessage = "The Following exception error occurred during the loading of the complete DataSet: ";
				Console.WriteLine(strErrorMessage + exc.Message + "\n HResult: " + exc.HResult + "\nInnerexception : " + exc.InnerException);
				// Send an e-mail to Technical Support
				objTechnicalEmailgeneral.TechnicalEmailModel.MessageLines.Add(strErrorMessage);
				objTechnicalEmailgeneral.TechnicalEmailModel.MessageLines.Add(exc.Message + "HResult: " + exc.HResult + "<br />Innerexception: " + exc.InnerException);
				if(objTechnicalEmailgeneral.ComposeHTMLemail(parEmailType: enumEmailType.TechnicalSupport))
					{//-	 Only send the message if the HTML e-mail compiled successfully
					SuccessfulSentEmail = objTechnicalEmailgeneral.SendEmail(
					parDataSet: ref parDataSet,
					parReceipient: Properties.AppResources.Email_Technical_Support,
					parSubject: "SDDP: DocGenerator Unexpected exception error occurred.)",
					parSendBcc: false);
					}
				goto Procedure_Ends;

				}
			objStopWatch1.Stop();
			Console.WriteLine("Time stamp Main controller: {0}", DateTime.UtcNow);
			Console.WriteLine("Time lapsed...............: {0})", objStopWatch1.Elapsed);

			//+ Obtain the details of the document Collections to be generated..
			string strDocWkbType = string.Empty;
			Console.WriteLine("{0} Document Collections to generate...", this.DocumentCollectionsToGenerate.Count);
			//
			List<DocumentCollection> listDocumentCollections;
			if(this.DocumentCollectionsToGenerate == null)
				listDocumentCollections = new List<DocumentCollection>();
			else
				listDocumentCollections = this.DocumentCollectionsToGenerate;

			// Obtain the details of the Document Collections that need to be processed, using the listDocumentCollection because you cannot pass the
			// this.Document CollectionsToGenerate as a referenced the object parameter.
			try
				{
				DocumentCollection.PopulateCollections(parDataSet: ref parDataSet, parDocumentCollectionList: ref listDocumentCollections);
				//- Once done set the this.DocumentCollectionsToGenerate property = to the listDocumentCollections object that now contains all the detail of the Document Collection
				this.DocumentCollectionsToGenerate = listDocumentCollections;
				}
			catch(GeneralException exc)
				{
				strErrorMessage = "The following exception error occurred while attempting to read the Data Collection Library: ";
				Console.WriteLine(strErrorMessage + exc.Message + "\n HResult: " + exc.HResult + "\nInnerexception : " + exc.InnerException);
				// Send an e-mail to Technical Support
				objTechnicalEmailgeneral.TechnicalEmailModel.MessageLines.Add(strErrorMessage);
				objTechnicalEmailgeneral.TechnicalEmailModel.MessageLines.Add(exc.Message + "HResult: " + exc.HResult + "<br />Innerexception: " + exc.InnerException);
				if(objTechnicalEmailgeneral.ComposeHTMLemail(parEmailType: enumEmailType.TechnicalSupport))
					{//-	 Only send the message if the HTML e-mail compiled successfully
					SuccessfulSentEmail = objTechnicalEmailgeneral.SendEmail(
						parDataSet: ref parDataSet,
						parReceipient: Properties.AppResources.Email_Technical_Support,
						parSubject: "SDDP: DocGenerator unexpected exception error occurred.)",
						parSendBcc: false);
					}
				goto Procedure_Ends;
				}

//===g
			//+ Sequencially process each of the **DocumentCollections**
			//- =========================================
			//- Process each of the document collections.
			//- =========================================
			eMail objConfirmationEmail = new eMail();
			eMail objTechnicalEmail = new eMail();

			try
				{
				//- The Complete DataSet is in Memory, now process each Document Collection Entry
				foreach(DocumentCollection objDocCollection in this.DocumentCollectionsToGenerate)
					{
					Console.WriteLine("\r\nReady to generate Document Collection: {0} - {1}", objDocCollection.ID.ToString(),
						objDocCollection.Title);
					objDocCollection.UnexpectedErrors = false;
					//- Reset all the Document Collection Specific variables and object variables
					objTechnicalEmail = new eMail();
					objTechnicalEmail.TechnicalEmailModel = new TechnicalSupportModel();
					objConfirmationEmail = new eMail();
					objConfirmationEmail.ConfirmationEmailModel = new EmailModel();
					//Prepare the E-mail Header that will be send to the user...

					objConfirmationEmail.ConfirmationEmailModel.CollectionID = objDocCollection.ID;
					objConfirmationEmail.ConfirmationEmailModel.CollectionTitle = objDocCollection.Title;
					objConfirmationEmail.ConfirmationEmailModel.CollectionURL = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion + Properties.AppResources.List_DocumentCollectionLibraryURI
						+ Properties.AppResources.EditFormURI + objDocCollection.ID;

					//-- Check if any documents were specified to be generated, if send an e-mail to the user stating that a no documents was sepecified.
					if(objDocCollection.Document_and_Workbook_objects == null
					|| objDocCollection.Document_and_Workbook_objects.Count() == 0)
						{
						//- Prepare and send an e-mail to the user...
						if(objDocCollection.NotificationEmail != null && objDocCollection.NotificationEmail != "None")
							{
							objConfirmationEmail.ConfirmationEmailModel.Failed = true;
							objConfirmationEmail.ConfirmationEmailModel.Error = "Unfortunatley, you submitted the Document Collection without specifing any document(s) to be generated."
								+ "<br /> Please specify any of the documents to be generated and then submit the Document Collection again.";
							if(objConfirmationEmail.ComposeHTMLemail(parEmailType: enumEmailType.UserErrorConfirmation))
								{//-	 Only send the message if the HTML e-mail compiled successfully
								SuccessfulSentEmail = objConfirmationEmail.SendEmail(
									parDataSet: ref parDataSet,
									parReceipient: objDocCollection.NotificationEmail,
									parSubject: "SDDP: Your generated document(s)");
								}
							}
						//- Update the Document Collection Entry, else it will be continually processed, until the **Generation Status** is not blank or Pending.
						this.SuccessfullUpdatedDocCollection = objDocCollection.UpdateGenerateStatus(
							parDataSet: ref parDataSet,
							parGenerationStatus: enumGenerationStatus.Completed);

						if(this.SuccessfullUpdatedDocCollection)
							Console.WriteLine("Update Document Collection Status to 'Completed' was successful.");
						else
							Console.WriteLine("Update Document Collection Status to 'Completed' was unsuccessful.");
						}
					else
						{//- The user soecified document - therefore process them....
						if(objConfirmationEmail.ConfirmationEmailModel.EmailGeneratedDocs == null)
							{
							objConfirmationEmail.ConfirmationEmailModel.EmailGeneratedDocs = new List<EmailGeneratedDocuments>();
							}
						foreach(dynamic objDocumentWorkbook in objDocCollection.Document_and_Workbook_objects)
							{
							Console.WriteLine("\r Generate ObjectType: {0}", objDocumentWorkbook.ToString());
							//- Declare the GeneratedDocument object that need to be added to the objConfirmationEmail.ConfirmationEmail.GeneratedDocs for inclusion in the e-mail
							EmailGeneratedDocuments objEmailGeneratedDocs = new EmailGeneratedDocuments();
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
									objCRMworkbook.Generate(parDataSet: ref parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);
									//- compose the e-mail section for this document
									objEmailGeneratedDocs.Title = "Client Requirements Mapping Workbook";
									objEmailGeneratedDocs.URL = objCRMworkbook.URLonSharePoint;

									// -Validate and finalise the document generation
									if(objCRMworkbook.DocumentStatus == enumDocumentStatusses.Done)
										{
										// Done - the document was generated and uploaded
										//- if there were content errors, add those to the client message
										if(objCRMworkbook.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objCRMworkbook.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											objEmailGeneratedDocs.IsSuccessful = true;
											}
										}
									else if(objCRMworkbook.DocumentStatus == enumDocumentStatusses.Error)
										{
										// there was an error that prevented the document's successful completion
										//- compose the e-mail section for this document
										//- if there were content errors, add those to the client message
										if(objCRMworkbook.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objCRMworkbook.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else if(objCRMworkbook.DocumentStatus == enumDocumentStatusses.FatalError)
											{// an Unexpected FATAL error occurred
											objDocCollection.UnexpectedErrors = true;
											objCRMworkbook.ErrorMessages.Add("Document Generation unexpectedly failed and the DocGenerator was "
												+ "unable to complete the generation of this document.");
											objCRMworkbook.ErrorMessages.Add("This message was also send to the SDDP Technical Team for "
												+ " further investigation. Once the issue is resolved the technical team will "
												+ "reschedule the generation of this document collection.");
											}
										}
									break;
									}
								//+ Content_Status_Workbook
								case ("Content_Status_Workbook"):
									{
									//- Prepare to generate the Document
									Content_Status_Workbook objContentStatusWB = objDocumentWorkbook;

									if(objContentStatusWB.ErrorMessages == null)
										objContentStatusWB.ErrorMessages = new List<string>();

									objContentStatusWB.Generate(parDataSet: ref parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);

									//- compose the e-mail section for this document
									objEmailGeneratedDocs.Title = "Content Status Workbook";
									objEmailGeneratedDocs.URL = objContentStatusWB.URLonSharePoint;

									// -Validate and finalise the document generation
									if(objContentStatusWB.DocumentStatus == enumDocumentStatusses.Done)
										{
										// Done - the document was generated and uploaded
										//- if there were content errors, add those to the client message
										if(objContentStatusWB.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objContentStatusWB.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											objEmailGeneratedDocs.IsSuccessful = true;
											}
										}
									else if(objContentStatusWB.DocumentStatus == enumDocumentStatusses.Error)
										{
										// there was an error that prevented the document's successful completion
										//- compose the e-mail section for this document
										//- if there were content errors, add those to the client message
										if(objContentStatusWB.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objContentStatusWB.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else if(objContentStatusWB.DocumentStatus == enumDocumentStatusses.FatalError)
											{// an Unexpected FATAL error occurred
											objDocCollection.UnexpectedErrors = true;
											objContentStatusWB.ErrorMessages.Add("Document Generation unexpectedly failed and the DocGenerator was "
												+ "unable to complete the generation of this document.");
											objContentStatusWB.ErrorMessages.Add("This message was also send to the SDDP Technical Team for "
												+ " further investigation. Once the issue is resolved the technical team will "
												+ "reschedule the generation of this document collection.");
											}
										}
									break;
									}

//---g
								//+ Contract_SoW_Service_Description
								case ("Contract_SoW_Service_Description"):
									{
									// Prepare to generate the Document
									Contract_SOW_Service_Description objContractSoW = objDocumentWorkbook;

									if(objContractSoW.ErrorMessages == null)
										objContractSoW.ErrorMessages = new List<string>();

									objContractSoW.Generate(parDataSet: ref parDataSet, 
										parRequestingUserID: objDocCollection.RequestingUserID,
										parClientName: objDocCollection.ClientName);

									//- compose the e-mail section for this document
									objEmailGeneratedDocs.Title = "Contract (SoW) Service Description Document";
									objEmailGeneratedDocs.URL = objContractSoW.URLonSharePoint;

									// -Validate and finalise the document generation
									if(objContractSoW.DocumentStatus == enumDocumentStatusses.Done)
										{
										// Done - the document was generated and uploaded
										//- if there were content errors, add those to the client message
										if(objContractSoW.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objContractSoW.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											objEmailGeneratedDocs.IsSuccessful = true;
											}
										}
									else if(objContractSoW.DocumentStatus == enumDocumentStatusses.Error)
										{
										// there was an error that prevented the document's successful completion
										//- compose the e-mail section for this document
										//- if there were content errors, add those to the client message
										if(objContractSoW.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objContractSoW.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else if(objContractSoW.DocumentStatus == enumDocumentStatusses.FatalError)
											{// an Unexpected FATAL error occurred
											objDocCollection.UnexpectedErrors = true;
											objContractSoW.ErrorMessages.Add("Document Generation unexpectedly failed and the DocGenerator was "
												+ "unable to complete the generation of this document.");
											objContractSoW.ErrorMessages.Add("This message was also send to the SDDP Technical Team for "
												+ " further investigation. Once the issue is resolved the technical team will "
												+ "reschedule the generation of this document collection.");
											}
										}
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
									objCSDbasedCRM.Generate(parDataSet: ref parDataSet, 
										parRequestingUserID: objDocCollection.RequestingUserID,
										parClientName: objDocCollection.ClientName);

									//- compose the e-mail section for this document
									objEmailGeneratedDocs.Title = "Client Service Description based on Requirements Mapping";
									objEmailGeneratedDocs.URL = objCSDbasedCRM.URLonSharePoint;

									// -Validate and finalise the document generation
									if(objCSDbasedCRM.DocumentStatus == enumDocumentStatusses.Done)
										{
										// Done - the document was generated and uploaded
										//- if there were content errors, add those to the client message
										if(objCSDbasedCRM.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objCSDbasedCRM.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											objEmailGeneratedDocs.IsSuccessful = true;
											}
										}
									else if(objCSDbasedCRM.DocumentStatus == enumDocumentStatusses.Error)
										{
										// there was an error that prevented the document's successful completion
										//- compose the e-mail section for this document
										//- if there were content errors, add those to the client message
										if(objCSDbasedCRM.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objCSDbasedCRM.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else if(objCSDbasedCRM.DocumentStatus == enumDocumentStatusses.FatalError)
											{// an Unexpected FATAL error occurred
											objDocCollection.UnexpectedErrors = true;
											objCSDbasedCRM.ErrorMessages.Add("Document Generation unexpectedly failed and the DocGenerator was "
												+ "unable to complete the generation of this document.");
											objCSDbasedCRM.ErrorMessages.Add("This message was also send to the SDDP Technical Team for "
												+ " further investigation. Once the issue is resolved the technical team will "
												+ "reschedule the generation of this document collection.");
											}
										}
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
									objCSDdrmInline.Generate(parDataSet: ref parDataSet, 
										parRequestingUserID: objDocCollection.RequestingUserID,
										parClientName: objDocCollection.ClientName);

									//- compose the e-mail section for this document
									objEmailGeneratedDocs.Title = "Client Service Description with inline Deliverables, Reports and Meetings";
									objEmailGeneratedDocs.URL = objCSDdrmInline.URLonSharePoint;

									// -Validate and finalise the document generation
									if(objCSDdrmInline.DocumentStatus == enumDocumentStatusses.Done)
										{
										// Done - the document was generated and uploaded
										//- if there were content errors, add those to the client message
										if(objCSDdrmInline.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objCSDdrmInline.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											objEmailGeneratedDocs.IsSuccessful = true;
											}
										}
									else if(objCSDdrmInline.DocumentStatus == enumDocumentStatusses.Error)
										{
										// there was an error that prevented the document's successful completion
										//- compose the e-mail section for this document
										//- if there were content errors, add those to the client message
										if(objCSDdrmInline.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objCSDdrmInline.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else if(objCSDdrmInline.DocumentStatus == enumDocumentStatusses.FatalError)
											{// an Unexpected FATAL error occurred
											objDocCollection.UnexpectedErrors = true;
											objCSDdrmInline.ErrorMessages.Add("Document Generation unexpectedly failed and the DocGenerator was "
												+ "unable to complete the generation of this document.");
											objCSDdrmInline.ErrorMessages.Add("This message was also send to the SDDP Technical Team for "
												+ " further investigation. Once the issue is resolved the technical team will "
												+ "reschedule the generation of this document collection.");
											}
										}
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
									objCSDdrmSections.Generate(parDataSet: ref parDataSet, 
										parRequestingUserID: objDocCollection.RequestingUserID,
										parClientName: objDocCollection.ClientName);

									//- compose the e-mail section for this document
									objEmailGeneratedDocs.Title = "Client Service Description with a Deliverables, Reports, Meetings "
										+ "Section Document";
									objEmailGeneratedDocs.URL = objCSDdrmSections.URLonSharePoint;

									// -Validate and finalise the document generation
									if(objCSDdrmSections.DocumentStatus == enumDocumentStatusses.Done)
										{
										// Done - the document was generated and uploaded
										//- if there were content errors, add those to the client message
										if(objCSDdrmSections.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objCSDdrmSections.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											objEmailGeneratedDocs.IsSuccessful = true;
											}
										}
									else if(objCSDdrmSections.DocumentStatus == enumDocumentStatusses.Error)
										{
										// there was an error that prevented the document's successful completion
										//- compose the e-mail section for this document
										//- if there were content errors, add those to the client message
										if(objCSDdrmSections.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objCSDdrmSections.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else if(objCSDdrmSections.DocumentStatus == enumDocumentStatusses.FatalError)
											{// an Unexpected FATAL error occurred
											objDocCollection.UnexpectedErrors = true;
											objCSDdrmSections.ErrorMessages.Add("Document Generation unexpectedly failed and the DocGenerator was "
												+ "unable to complete the generation of this document.");
											objCSDdrmSections.ErrorMessages.Add("This message was also send to the SDDP Technical Team for "
												+ " further investigation. Once the issue is resolved the technical team will "
												+ "reschedule the generation of this document collection.");
											}
										}
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

									objExtTechDashboard.Generate(parDataSet: ref parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);

									//- compose the e-mail section for this document
									objEmailGeneratedDocs.Title = "External Technology Coverage Dashboard Workbook";
									objEmailGeneratedDocs.URL = objExtTechDashboard.URLonSharePoint;

									//- Validate and finalise the document generation
									if(objExtTechDashboard.DocumentStatus == enumDocumentStatusses.Done)
										{
										// Done - the document was generated and uploaded
										//- if there were content errors, add those to the client message
										if(objExtTechDashboard.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objExtTechDashboard.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											objEmailGeneratedDocs.IsSuccessful = true;
											}
										}
									else if(objExtTechDashboard.DocumentStatus == enumDocumentStatusses.Error)
										{
										// there was an error that prevented the document's successful completion
										//- compose the e-mail section for this document
										//- if there were content errors, add those to the client message
										if(objExtTechDashboard.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objExtTechDashboard.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else if(objExtTechDashboard.DocumentStatus == enumDocumentStatusses.FatalError)
											{// an Unexpected FATAL error occurred
											objDocCollection.UnexpectedErrors = true;
											objExtTechDashboard.ErrorMessages.Add("Document Generation unexpectedly failed and the DocGenerator "
												+ "was unable to complete the generation of this document.");
											objExtTechDashboard.ErrorMessages.Add("This message was also send to the SDDP Technical Team for "
												+ " further investigation. Once the issue is resolved the technical team will "
												+ "reschedule the generation of this document collection.");
											}
										}
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
									objIntTechDashboard.Generate(parDataSet: ref parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);

									//- compose the e-mail section for this document
									objEmailGeneratedDocs.Title = "Internal Technology Coverage Dashboard Workbook";
									objEmailGeneratedDocs.URL = objIntTechDashboard.URLonSharePoint;

									// -Validate and finalise the document generation
									if(objIntTechDashboard.DocumentStatus == enumDocumentStatusses.Done)
										{
										// Done - the document was generated and uploaded
										//- if there were content errors, add those to the client message
										if(objIntTechDashboard.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objIntTechDashboard.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											objEmailGeneratedDocs.IsSuccessful = true;
											}
										}
									else if(objIntTechDashboard.DocumentStatus == enumDocumentStatusses.Error)
										{
										// there was an error that prevented the document's successful completion
										//- compose the e-mail section for this document
										//- if there were content errors, add those to the client message
										if(objIntTechDashboard.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objIntTechDashboard.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else if(objIntTechDashboard.DocumentStatus == enumDocumentStatusses.FatalError)
											{// an Unexpected FATAL error occurred
											objDocCollection.UnexpectedErrors = true;
											objIntTechDashboard.ErrorMessages.Add("Document Generation unexpectedly failed and the DocGenerator "
												+ "was unable to complete the generation of this document.");
											objIntTechDashboard.ErrorMessages.Add("This message was also send to the SDDP Technical Team for "
												+ " further investigation. Once the issue is resolved the technical team will "
												+ "reschedule the generation of this document collection.");
											}
										}
									break;
									}

								//+ Services_Model_Workbook
								case ("Services_Model_Workbook"):
									{
									//- Prepare to generate the Document
									Services_Model_Workbook objServicesModelWB = objDocumentWorkbook;
									if(objServicesModelWB.ErrorMessages == null)
										objServicesModelWB.ErrorMessages = new List<string>();

									//- Generate the document...
									objServicesModelWB.Generate(parDataSet: ref parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);

									//- compose the e-mail section for this document
									objEmailGeneratedDocs.Title = "Internal Services Mapping Workbook";
									objEmailGeneratedDocs.URL = objServicesModelWB.URLonSharePoint;

									// -Validate and finalise the document generation
									if(objServicesModelWB.DocumentStatus == enumDocumentStatusses.Done)
										{
										// Done - the document was generated and uploaded
										//- if there were content errors, add those to the client message
										if(objServicesModelWB.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objServicesModelWB.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											objEmailGeneratedDocs.IsSuccessful = true;
											}
										}
									else if(objServicesModelWB.DocumentStatus == enumDocumentStatusses.Error)
										{
										// there was an error that prevented the document's successful completion
										//- compose the e-mail section for this document
										//- if there were content errors, add those to the client message
										if(objServicesModelWB.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objServicesModelWB.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else if(objServicesModelWB.DocumentStatus == enumDocumentStatusses.FatalError)
											{// an Unexpected FATAL error occurred
											objDocCollection.UnexpectedErrors = true;
											objServicesModelWB.ErrorMessages.Add("Document Generation unexpectedly failed "
												+ "and the DocGenerator was unable to complete the generation of this document.");
											objServicesModelWB.ErrorMessages.Add("This message was also send to the SDDP "
												+ "Technical Team for further investigation. Once the issue is resolved the " 
												+ "technical team will investigate the issue and after fixing it, they will "
												+ "reschedule the generation of this document collection.");
											}
										}
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
									objISDdrmInline.Generate(parDataSet: ref parDataSet, 
										parRequestingUserID: objDocCollection.RequestingUserID,
										parClientName: objDocCollection.ClientName);

									//- compose the e-mail section for this document
									objEmailGeneratedDocs.Title = "Internal Service Definition with inline Deliverables, Reports, Meetings";
									objEmailGeneratedDocs.URL = objISDdrmInline.URLonSharePoint;

									// -Validate and finalise the document generation
									if(objISDdrmInline.DocumentStatus == enumDocumentStatusses.Done)
										{
										// Done - the document was generated and uploaded
										//- if there were content errors, add those to the client message
										if(objISDdrmInline.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objISDdrmInline.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											objEmailGeneratedDocs.IsSuccessful = true;
											}
										}
									else if(objISDdrmInline.DocumentStatus == enumDocumentStatusses.Error)
										{
										// there was an error that prevented the document's successful completion
										//- compose the e-mail section for this document
										//- if there were content errors, add those to the client message
										if(objISDdrmInline.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objISDdrmInline.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else if(objISDdrmInline.DocumentStatus == enumDocumentStatusses.FatalError)
											{// an Unexpected FATAL error occurred
											objDocCollection.UnexpectedErrors = true;
											objISDdrmInline.ErrorMessages.Add("Document Generation unexpectedly failed and the DocGenerator was "
												+ "unable to complete the generation of this document.");
											objISDdrmInline.ErrorMessages.Add("This message was also send to the SDDP Technical Team for "
												+ " further investigation. Once the issue is resolved the technical team will "
												+ "reschedule the generation of this document collection.");
											}
										}
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
									objISDdrmSections.Generate(parDataSet: ref parDataSet, 
										parRequestingUserID: objDocCollection.RequestingUserID,
										parClientName: objDocCollection.ClientName);
									//- compose the e-mail section for this document
									objEmailGeneratedDocs.Title = "Internal Service Definition with a Deliverables, Reports, Meetings Section";
									objEmailGeneratedDocs.URL = objISDdrmSections.URLonSharePoint;

									// -Validate and finalise the document generation
									if(objISDdrmSections.DocumentStatus == enumDocumentStatusses.Done)
										{
										// Done - the document was generated and uploaded
										//- if there were content errors, add those to the client message
										if(objISDdrmSections.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objISDdrmSections.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											objEmailGeneratedDocs.IsSuccessful = true;
											}
										}
									else if(objISDdrmSections.DocumentStatus == enumDocumentStatusses.Error)
										{
										// there was an error that prevented the document's successful completion
										//- compose the e-mail section for this document
										//- if there were content errors, add those to the client message
										if(objISDdrmSections.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objISDdrmSections.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else if(objISDdrmSections.DocumentStatus == enumDocumentStatusses.FatalError)
											{// an Unexpected FATAL error occurred
											objDocCollection.UnexpectedErrors = true;
											objISDdrmSections.ErrorMessages.Add("Document Generation unexpectedly failed and the DocGenerator was "
												+ "unable to complete the generation of this document.");
											objISDdrmSections.ErrorMessages.Add("This message was also send to the SDDP Technical Team for "
												+ " further investigation. Once the issue is resolved the technical team will "
												+ "reschedule the generation of this document collection.");
											}
										}
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
									//- compose the e-mail section for this document
									objEmailGeneratedDocs.Title = "Pricing Addendum Document";
									objEmailGeneratedDocs.URL = objPricingAddendum.URLonSharePoint;

									// -Validate and finalise the document generation
									if(objPricingAddendum.DocumentStatus == enumDocumentStatusses.Done)
										{
										// Done - the document was generated and uploaded
										//- if there were content errors, add those to the client message
										if(objPricingAddendum.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objPricingAddendum.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											objEmailGeneratedDocs.IsSuccessful = true;
											}
										}
									else if(objPricingAddendum.DocumentStatus == enumDocumentStatusses.Error)
										{
										// there was an error that prevented the document's successful completion
										//- compose the e-mail section for this document
										//- if there were content errors, add those to the client message
										if(objPricingAddendum.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objPricingAddendum.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else if(objPricingAddendum.DocumentStatus == enumDocumentStatusses.FatalError)
											{// an Unexpected FATAL error occurred
											objDocCollection.UnexpectedErrors = true;
											objPricingAddendum.ErrorMessages.Add("Document Generation unexpectedly failed and the DocGenerator was "
												+ "unable to complete the generation of this document.");
											objPricingAddendum.ErrorMessages.Add("This message was also send to the SDDP Technical Team for "
												+ " further investigation. Once the issue is resolved the technical team will "
												+ "reschedule the generation of this document collection.");
											}
										}
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
									objRACImatrix.Generate(parDataSet: ref parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);

									//- compose the e-mail section for this document
									objEmailGeneratedDocs.Title = "RACI Matrix per Deliverable Workbook";
									objEmailGeneratedDocs.URL = objRACImatrix.URLonSharePoint;

									// -Validate and finalise the document generation
									if(objRACImatrix.DocumentStatus == enumDocumentStatusses.Done)
										{
										// Done - the document was generated and uploaded
										//- if there were content errors, add those to the client message
										if(objRACImatrix.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objRACImatrix.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											objEmailGeneratedDocs.IsSuccessful = true;
											}
										}
									else if(objRACImatrix.DocumentStatus == enumDocumentStatusses.Error)
										{
										// there was an error that prevented the document's successful completion
										//- compose the e-mail section for this document
										//- if there were content errors, add those to the client message
										if(objRACImatrix.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objRACImatrix.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else if(objRACImatrix.DocumentStatus == enumDocumentStatusses.FatalError)
											{// an Unexpected FATAL error occurred
											objDocCollection.UnexpectedErrors = true;
											objRACImatrix.ErrorMessages.Add("Document Generation unexpectedly failed and the DocGenerator was "
												+ "unable to complete the generation of this document.");
											objRACImatrix.ErrorMessages.Add("This message was also send to the SDDP Technical Team for "
												+ " further investigation. Once the issue is resolved the technical team will "
												+ "reschedule the generation of this document collection.");
											}
										}
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
									objRACIperRole.Generate(parDataSet: ref parDataSet, parRequestingUserID: objDocCollection.RequestingUserID);

									//- compose the e-mail section for this document
									objEmailGeneratedDocs.Title = "RACI per Job Role Workbook";
									objEmailGeneratedDocs.URL = objRACIperRole.URLonSharePoint;

									// -Validate and finalise the document generation
									if(objRACIperRole.DocumentStatus == enumDocumentStatusses.Done)
										{
										// Done - the document was generated and uploaded
										//- if there were content errors, add those to the client message
										if(objRACIperRole.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objRACIperRole.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											objEmailGeneratedDocs.IsSuccessful = true;
											}
										}
									else if(objRACIperRole.DocumentStatus == enumDocumentStatusses.Error)
										{
										// there was an error that prevented the document's successful completion
										//- compose the e-mail section for this document
										//- if there were content errors, add those to the client message
										if(objRACIperRole.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objRACIperRole.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else if(objRACIperRole.DocumentStatus == enumDocumentStatusses.FatalError)
											{// an Unexpected FATAL error occurred
											objDocCollection.UnexpectedErrors = true;
											objRACIperRole.ErrorMessages.Add("Document Generation unexpectedly failed and the DocGenerator was "
												+ "unable to complete the generation of this document.");
											objRACIperRole.ErrorMessages.Add("This message was also send to the SDDP Technical Team for "
												+ " further investigation. Once the issue is resolved the technical team will "
												+ "reschedule the generation of this document collection.");
											}
										}
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
									objSFdrmInline.Generate(parDataSet: ref parDataSet, 
										parRequestingUserID: objDocCollection.RequestingUserID,
										parClientName: objDocCollection.ClientName);

									//- compose the e-mail section for this document
									objEmailGeneratedDocs.Title = "Services Framework with inline Deliverables, Reports, Meetings Document";
									objEmailGeneratedDocs.URL = objSFdrmInline.URLonSharePoint;

									// -Validate and finalise the document generation
									if(objSFdrmInline.DocumentStatus == enumDocumentStatusses.Done)
										{
										// Done - the document was generated and uploaded
										//- if there were content errors, add those to the client message
										if(objSFdrmInline.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objSFdrmInline.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											objEmailGeneratedDocs.IsSuccessful = true;
											}
										}
									else if(objSFdrmInline.DocumentStatus == enumDocumentStatusses.Error)
										{
										// there was an error that prevented the document's successful completion
										//- compose the e-mail section for this document
										//- if there were content errors, add those to the client message
										if(objSFdrmInline.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objSFdrmInline.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else if(objSFdrmInline.DocumentStatus == enumDocumentStatusses.FatalError)
											{// an Unexpected FATAL error occurred
											objDocCollection.UnexpectedErrors = true;
											objSFdrmInline.ErrorMessages.Add("Document Generation unexpectedly failed and the DocGenerator was "
												+ "unable to complete the generation of this document.");
											objSFdrmInline.ErrorMessages.Add("This message was also send to the SDDP Technical Team for "
												+ " further investigation. Once the issue is resolved the technical team will "
												+ "reschedule the generation of this document collection.");
											}
										}
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
									objSFdrmSections.Generate(parDataSet: ref parDataSet, 
										parRequestingUserID: objDocCollection.RequestingUserID,
										parClientName: objDocCollection.ClientName);

									//- compose the e-mail section for this document
									objEmailGeneratedDocs.Title = "Services Framework Document with a Deliverables, Report, Meetings Section";
									objEmailGeneratedDocs.URL = objSFdrmSections.URLonSharePoint;

									// -Validate and finalise the document generation
									if(objSFdrmSections.DocumentStatus == enumDocumentStatusses.Done)
										{
										// Done - the document was generated and uploaded
										//- if there were content errors, add those to the client message
										if(objSFdrmSections.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objSFdrmSections.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else
											{//- there were no content errors
											objEmailGeneratedDocs.IsSuccessful = true;
											}
										}
									else if(objSFdrmSections.DocumentStatus == enumDocumentStatusses.Error)
										{
										// there was an error that prevented the document's successful completion
										//- compose the e-mail section for this document
										//- if there were content errors, add those to the client message
										if(objSFdrmSections.ErrorMessages.Count() > 0)
											{//- include them in the message.
											objEmailGeneratedDocs.IsSuccessful = false;
											objEmailGeneratedDocs.Errors = new List<string>();
											foreach(string errorEntry in objSFdrmSections.ErrorMessages)
												{
												objEmailGeneratedDocs.Errors.Add(errorEntry);
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else if(objSFdrmSections.DocumentStatus == enumDocumentStatusses.FatalError)
											{// an Unexpected FATAL error occurred
											objDocCollection.UnexpectedErrors = true;
											objSFdrmSections.ErrorMessages.Add("Document Generation unexpectedly failed and the DocGenerator was "
												+ "unable to complete the generation of this document.");
											objSFdrmSections.ErrorMessages.Add("This message was also send to the SDDP Technical Team for "
												+ " further investigation. Once the issue is resolved the technical team will "
												+ "reschedule the generation of this document collection.");
											}
										}
									break;
									}
								} //- switch (objectType)
							//- Add the Generated document's e-mail content to the confirmation e-mail to ensure it appears in the generated document.
							objConfirmationEmail.ConfirmationEmailModel.EmailGeneratedDocs.Add(objEmailGeneratedDocs);
							} //- foreach(dynamic objDocumentWorkbook in objDocCollection.Documen_and_Workbook_Objects...

//---g
						//+ Process the User Notification E-mail
						//- Process the Notification via E-mail if the users selected to be notified.
						if(objDocCollection.NotifyMe && objDocCollection.NotificationEmail != null)
							{

							if(objConfirmationEmail.ComposeHTMLemail(parEmailType: enumEmailType.UserSuccessfulConfirmation))
								{
								SuccessfulSentEmail = objConfirmationEmail.SendEmail(
									parDataSet: ref parDataSet,
									parReceipient: objDocCollection.NotificationEmail,
									parSubject: "SDDP: your generated document(s)");

								if(SuccessfulSentEmail)
									Console.WriteLine("Sending e-mail successfully send to user!");
								else
									Console.WriteLine("*** ERROR *** \n Sending e-mail failed...\n");
								}
							else
								Console.WriteLine("Error composing the HTML email with Razor");
							}

						//+ Check if any **unexpected** errors occurred
						if(objDocCollection.UnexpectedErrors)
							{//- if there were unexpected errors, send an e-mail to the Technical Support team.
							this.SuccessfullUpdatedDocCollection = objDocCollection.UpdateGenerateStatus(
								parDataSet: ref parDataSet,
								parGenerationStatus: enumGenerationStatus.Failed);

							if(this.SuccessfullUpdatedDocCollection)
								Console.WriteLine("Update Document Collection Status to 'FAILED' was successful.");
							else
								Console.WriteLine("Update Document Collection Status to 'FAILED' was unsuccessful.");

							if(objTechnicalEmail.ComposeHTMLemail(parEmailType: enumEmailType.TechnicalSupport))
								{
								//- Prepare the e-mail Technical Support team's e-mail
								SuccessfulSentEmail = objTechnicalEmail.SendEmail(
									parDataSet: ref parDataSet,
									parReceipient: Properties.AppResources.Email_Technical_Support,
									parSubject: "SDDP: Unexpected Error occurred in the DocGenerator.");

								if(SuccessfulSentEmail)
									Console.WriteLine("The error e-mail was successfully send to the technical team.");
								else
									Console.WriteLine("The error e-mail to the technical team FAILED!");
								}
							}
						else
							{//- there was no UNEXPECTED errors
							this.SuccessfullUpdatedDocCollection = objDocCollection.UpdateGenerateStatus(
								parDataSet: ref parDataSet,
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
					strErrorMessage = "The data connection to SharePoint timed out - and documents could not be generated..." +
						"The DocGenerator will retry to generate the document. Keep an eye on any further e-mails and investigate it this error occur again shortly.";
					Console.WriteLine(strErrorMessage + exc.Message + "\n HResult: " + exc.HResult + "\nInnerexception : " + exc.InnerException);
					// Send an e-mail to Technical Support
					objTechnicalEmail.TechnicalEmailModel.MessageLines.Add(strErrorMessage);
					objTechnicalEmail.TechnicalEmailModel.MessageLines.Add(exc.Message + "HResult: " + exc.HResult + "<br />Innerexception: " 
						+ exc.InnerException);
					if(objTechnicalEmail.ComposeHTMLemail(parEmailType: enumEmailType.TechnicalSupport))
						{
						SuccessfulSentEmail = objTechnicalEmail.SendEmail(
							parDataSet: ref parDataSet,
							parReceipient: Properties.AppResources.Email_Technical_Support,
							parSubject: "SDDP: DocGenerator DataServiceTransportException (timeout) occurred.",
							parSendBcc: false);
						}
					}
				else
					{
					strErrorMessage = "Unexpected exception error: ";
					Console.WriteLine(strErrorMessage + exc.Message + "\n HResult: " + exc.HResult + "\nInnerexception : " + exc.InnerException);
					// Send an e-mail to Technical Support
					objTechnicalEmail.TechnicalEmailModel.MessageLines.Add(strErrorMessage);
					objTechnicalEmail.TechnicalEmailModel.MessageLines.Add(exc.Message + "HResult: " + exc.HResult + "<br />Innerexception: " 
						+ exc.InnerException);
					if(objTechnicalEmail.ComposeHTMLemail(parEmailType: enumEmailType.TechnicalSupport))
						{
						SuccessfulSentEmail = objTechnicalEmail.SendEmail(
							parDataSet: ref parDataSet,
							parReceipient: Properties.AppResources.Email_Technical_Support,
							parSubject: "SDDP: DocGenerator DataServicetransportException (unexpected) occurred.",
							parSendBcc: false);
						}
					}

				}
			catch(Exception exc)
				{
				strErrorMessage = "Unexpected exception error occurred";
				Console.WriteLine(strErrorMessage + exc.Message + "\n HResult: " + exc.HResult + "\nInnerexception : " + exc.InnerException);
				// Send an e-mail to Technical Support
				objTechnicalEmail.TechnicalEmailModel.MessageLines.Add(strErrorMessage);
				objTechnicalEmail.TechnicalEmailModel.MessageLines.Add(exc.Message + "HResult: " + exc.HResult + "<br />Innerexception: " + exc.InnerException);
				if(objTechnicalEmail.ComposeHTMLemail(parEmailType: enumEmailType.TechnicalSupport))
					{
					SuccessfulSentEmail = objTechnicalEmail.SendEmail(
						parDataSet: ref parDataSet,
						parReceipient: Properties.AppResources.Email_Technical_Support,
						parSubject: "SDDP: DocGenerator Unexpected Exception error occurred.",
						parSendBcc: false);
					}
				}

Procedure_Ends:
			Console.WriteLine("end of MainController in DocGeneratorCore.");
			return;
			} // end of method
		} // end of class
	#endregion
	} // end of Namespace