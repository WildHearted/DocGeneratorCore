using System;
using System.Collections.Generic;
using System.Data.Services.Client;
using System.IO;
using System.Linq;
using System.Net;
using DocGeneratorCore.SDDPServiceReference;

namespace DocGeneratorCore
	{
	public static class Globals
		{
		public static CompleteDataSet objCompleteDataSet;
		}

	public class MainController
		{
		public bool SuccessfulGeneratedDocument{get; set;}
		public bool SuccessfulPublishedDocument{get; set;}
		public bool SuccessfulSentEmail{get; set;}
		public bool SuccessfullUpdatedDocCollection{get; set;}
		public string EmailBodyText{get; set;}
		public string ReturnString{get; set;}
		public bool SuccessfulCompletion{get; set;}
		public List<DocumentCollection> DocumentCollectionsToGenerate{get; set;}
		public DesignAndDeliveryPortfolioDataContext SDDPdatacontext{get; set;}

		public void MainProcess()
			{
			string objectType = string.Empty;
			Console.WriteLine("Begin to execute the MainProcess in the DocGeneratorCore module");
			//Declare the SharePoint connection as a DataContext
			//DesignAndDeliveryPortfolioDataContext objSDDPdatacontext = new DesignAndDeliveryPortfolioDataContext(new
			//	Uri(Properties.AppResources.SharePointSiteURL + Properties.AppResources.SharePointRESTuri));

			//objSDDPdatacontext.Credentials = CredentialCache.DefaultCredentials;
			//objSDDPdatacontext.Credentials = new NetworkCredential(
			//	userName: Properties.AppResources.User_Credentials_UserName,
			//	password: Properties.AppResources.User_Credentials_Password,
			//	domain: Properties.AppResources.User_Credentials_Domain);
			//objSDDPdatacontext.MergeOption = MergeOption.NoTracking;
			this.SDDPdatacontext = new DesignAndDeliveryPortfolioDataContext(new
				Uri(Properties.AppResources.SharePointSiteURL + Properties.AppResources.SharePointRESTuri));

			this.SDDPdatacontext.Credentials = new NetworkCredential(
				userName: Properties.AppResources.User_Credentials_UserName,
				password: Properties.AppResources.User_Credentials_Password,
				domain: Properties.AppResources.User_Credentials_Domain);

			this.SDDPdatacontext.MergeOption = MergeOption.NoTracking;

			Console.WriteLine("Checking the Document Collection Library for any documents to generate...");
			this.ReturnString = String.Empty;
			this.SuccessfulGeneratedDocument = false;
			this.SuccessfulPublishedDocument = false;
			this.DocumentCollectionsToGenerate = new List<DocumentCollection>();

			// Obtain the Document Collections that need to be processed
			try
				{
				this.DocumentCollectionsToGenerate = DocumentCollection.GetCollectionsToGenerate(parSDDPdatacontext: this.SDDPdatacontext);

				if(this.DocumentCollectionsToGenerate == null
				|| this.DocumentCollectionsToGenerate.Count() < 1)
					{
					Console.WriteLine("Nothing to generate at this stage...");
					goto Procedure_Ends;
					}
				else
					{
					Console.WriteLine("{0} Document Collections to generate...", this.DocumentCollectionsToGenerate.Count);
					}
				}
			catch(GeneralException exc)
				{
				this.EmailBodyText = "Exception Error occurred: " + exc.Message + "\n HResult: " + exc.HResult + "\nInnerexception : " + exc.InnerException;
				Console.WriteLine(this.EmailBodyText);
				// Send the e-mail Technical Support
				SuccessfulSentEmail = eMail.SendEmail(
					parRecipient: Properties.AppResources.Email_Bcc_Address,
					parSubject: "SDDP: Unexpected DocGenerator(s) Error occurred.)",
					parBody: EmailBodyText,
					parSendBcc: false);
				goto Procedure_Ends;
				}

			//----------------------------------------------------
			// There are Document Collection entries to process...

			// To ensure optimal Document Generation performance:.
			// Load the complete DataSet before beginning to generate the documents.
			if(Globals.objCompleteDataSet == null)
				{
				//CompleteDataSet objDataSet = new CompleteDataSet();
				Globals.objCompleteDataSet = new CompleteDataSet();

				try
					{
					this.SuccessfulCompletion = Globals.objCompleteDataSet.PopulateBaseObjects(parDatacontexSDDP: this.SDDPdatacontext);
					if(!SuccessfulCompletion)
						{
						this.EmailBodyText = "DocGenerator was unable to successfully load the Complete DataSet from SharEPoint. Please investigate";
						Console.WriteLine(this.EmailBodyText);
						// Send the e-mail Technical Support
						SuccessfulSentEmail = eMail.SendEmail(
							parRecipient: Properties.AppResources.Email_Bcc_Address,
							parSubject: "SDDP: Unexpected DocGenerator Error occurred.)",
							parBody: EmailBodyText,
							parSendBcc: false);
						goto Procedure_Ends;
						}
					}
				catch(GeneralException exc)
					{
					this.EmailBodyText = "Exception Error occurred during the loading of the complete DataSet: " + exc.Message + "\n HResult: " + exc.HResult + "\nInnerexception : " + exc.InnerException;
					Console.WriteLine(this.EmailBodyText);
					// Send the e-mail Technical Support
					SuccessfulSentEmail = eMail.SendEmail(
						parRecipient: Properties.AppResources.Email_Bcc_Address,
						parSubject: "SDDP: Unexpected DocGenerator Error occurred.)",
						parBody: EmailBodyText,
						parSendBcc: false);
					goto Procedure_Ends;
					}
				}
			try
				{
				// Complete DataSet in Memory, now process each Document Collection Entry
				// Process each of the documents in the DocumentCollection
				foreach(DocumentCollection objDocCollection in DocumentCollectionsToGenerate)
					{
					Console.WriteLine("\r\nReady to generate Document Collection: {0} - {1}", objDocCollection.ID.ToString(),
						objDocCollection.Title);
					objDocCollection.UnexpectedErrors = false;

					//Prepare the  E-mail that will be send to the user...
					EmailBodyText = "Good day,\n\nHerewith the generated document(s) that you requested from the Service Design and Delivery Portfolio "
						+ "as entry\n" + objDocCollection.ID + " - " + objDocCollection.Title + " in the Document Collections Library";

					// Process each of the documents in the DocumentCollection
					if(objDocCollection.Document_and_Workbook_objects.Count() > 0)
						{
						//objDocCollection.Document_and_Workbook_objects.GetType();
						foreach(dynamic objDocumentWorkbook in objDocCollection.Document_and_Workbook_objects)
							{
							Console.WriteLine("\r Generate ObjectType: {0}", objDocumentWorkbook.ToString());
							objectType = objDocumentWorkbook.ToString();
							objectType = objectType.Substring(objectType.IndexOf(".") + 1, (objectType.Length - objectType.IndexOf(".") - 1));
							switch(objectType)
								{
							//--------------------------------------------
							case ("Client_Requirements_Mapping_Workbook"):
									{
									// Prepare to generate the Document
									SuccessfulGeneratedDocument = false;
									Client_Requirements_Mapping_Workbook objCRMworkbook = objDocumentWorkbook;

									if(objCRMworkbook.ErrorMessages == null)
										objCRMworkbook.ErrorMessages = new List<string>();

									SuccessfulGeneratedDocument = objCRMworkbook.Generate(
										parDataSet: ref Globals.objCompleteDataSet,
										parSDDPdatacontext: this.SDDPdatacontext);

									if(SuccessfulGeneratedDocument)
										{
										// set the Document status to Completed...
										objCRMworkbook.DocumentStatus = enumDocumentStatusses.Completed;
										// Prepare the inclusion of the text in the e-mail that the user will receive.
										EmailBodyText += "\n     * " + objDocumentWorkbook.DocumentType;
										// if there were errors, include them in the message.
										if(objCRMworkbook.ErrorMessages.Count() > 0)
											{
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objCRMworkbook.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following errors occurred:";
											foreach(string errorEntry in objCRMworkbook.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else // there were no generation errors.
											{
											Console.WriteLine("\t *** no errors occurred during the generation process.");
											EmailBodyText += ", which was generated without any errors.";
											}

										// begin to upload the document to SharePoint
										objCRMworkbook.DocumentStatus = enumDocumentStatusses.Uploading;
										Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

										// Upload the document to the Generated Documents Library
										SuccessfulPublishedDocument = objCRMworkbook.UploadDoc(
											parRequestingUserID: objDocCollection.RequestingUserID);
										// Check if the upload succeeded....
										if(SuccessfulPublishedDocument) //Upload Succeeded
											{
											Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
											// Insert the uploaded URL in the e-mail message body
											EmailBodyText += "\n       The document is stored at this url: " + objCRMworkbook.URLonSharePoint;
											objCRMworkbook.DocumentStatus = enumDocumentStatusses.Uploaded;
											// Delete the uploaded file from the Documents Directory
											if(File.Exists(path: objCRMworkbook.FileName))
												{
												File.Delete(path: objCRMworkbook.FileName);
												}
											}
										else // Upload failed Failed
											{
											Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
											objDocCollection.UnexpectedErrors = true;
											objCRMworkbook.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
											EmailBodyText += "\n       Unfortunately, a technical issue prevented the uploading of "
													+ "the generated document to the Generarated Documents Library on SharePoint.";
											}
										//Check if there were any Unhandled errors and flag the Document's collection
										if(objCRMworkbook.UnhandledError)
											{
											objDocCollection.UnexpectedErrors = true;
											}
										}
									else // The Document generation failed for some reason
										{
										Console.WriteLine("\t\t *** Unfortunately, the generation of the following document "
											+ "failed unexpectedly : {0}"
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)"
											, objDocumentWorkbook.GetType());
										objDocCollection.UnexpectedErrors = true;
										objCRMworkbook.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objCRMworkbook.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									EmailBodyText += "\n\n";
									break;
									}
							//---------------------------------------
							case ("Content_Status_Workbook"):
									{
									// Prepare to generate the Document
									SuccessfulGeneratedDocument = false;
									Content_Status_Workbook objcontentStatus = objDocumentWorkbook;

									if(objcontentStatus.ErrorMessages == null)
										objcontentStatus.ErrorMessages = new List<string>();
									SuccessfulGeneratedDocument = objcontentStatus.Generate(
										parDataSet: ref Globals.objCompleteDataSet);

									if(SuccessfulGeneratedDocument)
										{
										// set the Document status to Completed...
										objcontentStatus.DocumentStatus = enumDocumentStatusses.Completed;
										// Prepare the inclusion of the text in the e-mail that the user will receive.
										EmailBodyText += "\n     * " + objDocumentWorkbook.DocumentType;
										// if there were errors, include them in the message.
										if(objcontentStatus.ErrorMessages.Count() > 0)
											{
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objcontentStatus.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following errors occurred:";
											foreach(string errorEntry in objcontentStatus.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else // there were no generation errors.
											{
											Console.WriteLine("\t *** no errors occurred during the generation process.");
											EmailBodyText += ", which was generated without any errors.";
											}

										// begin to upload the document to SharePoint
										objcontentStatus.DocumentStatus = enumDocumentStatusses.Uploading;
										Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

										// Upload the document to the Generated Documents Library
										SuccessfulPublishedDocument = objcontentStatus.UploadDoc(
											parRequestingUserID: objDocCollection.RequestingUserID);
										// Check if the upload succeeded....
										if(SuccessfulPublishedDocument) //Upload Succeeded
											{
											Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
											// Insert the uploaded URL in the e-mail message body
											EmailBodyText += "\n       The document is stored at this url: " + objcontentStatus.URLonSharePoint;
											objcontentStatus.DocumentStatus = enumDocumentStatusses.Uploaded;
											// Delete the uploaded file from the Documents Directory
											if(File.Exists(path: objcontentStatus.FileName))
												{
												File.Delete(path: objcontentStatus.FileName);
												}
											}
										else // Upload failed Failed
											{
											Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
											objDocCollection.UnexpectedErrors = true;
											objcontentStatus.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
											EmailBodyText += "\n       Unfortunately, a technical issue prevented the uploading of "
													+ "the generated document to the Generarated Documents Library on SharePoint.";
											}
										//Check if there were any Unhandled errors and flag the Document's collection
										if(objcontentStatus.UnhandledError)
											{
											objDocCollection.UnexpectedErrors = true;
											}
										}
									else // The Document generation failed for some reason
										{
										Console.WriteLine("\t\t *** Unfortunately, the generation of the following document "
											+ "failed unexpectedly : {0}"
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)"
											, objDocumentWorkbook.GetType());
										objDocCollection.UnexpectedErrors = true;
										objcontentStatus.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objcontentStatus.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									EmailBodyText += "\n\n";
									break;
									}
							//--------------------------------------------
							case ("Contract_SoW_Service_Description"):
									{
									// Prepare to generate the Document
									SuccessfulGeneratedDocument = false;
									Contract_SoW_Service_Description objContractSoW = objDocumentWorkbook;

									if(objContractSoW.ErrorMessages == null)
										objContractSoW.ErrorMessages = new List<string>();

									SuccessfulGeneratedDocument = objContractSoW.Generate(
										parDataSet: ref Globals.objCompleteDataSet,
										parSDDPdatacontext: this.SDDPdatacontext);

									if(SuccessfulGeneratedDocument)
										{
										// set the Document status to Completed...
										objContractSoW.DocumentStatus = enumDocumentStatusses.Completed;
										// Prepare the inclusion of the text in the e-mail that the user will receive.
										EmailBodyText += "\n     * " + objDocumentWorkbook.DocumentType;
										// if there were errors, include them in the message.
										if(objContractSoW.ErrorMessages.Count() > 0)
											{
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objContractSoW.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following errors occurred:";
											foreach(string errorEntry in objContractSoW.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else // there were no generation errors.
											{
											Console.WriteLine("\t *** no errors occurred during the generation process.");
											EmailBodyText += ", which was generated without any errors.";
											}

										// begin to upload the document to SharePoint
										objContractSoW.DocumentStatus = enumDocumentStatusses.Uploading;
										Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

										// Upload the document to the Generated Documents Library
										SuccessfulPublishedDocument = objContractSoW.UploadDoc(
											parRequestingUserID: objDocCollection.RequestingUserID);
										// Check if the upload succeeded....
										if(SuccessfulPublishedDocument) //Upload Succeeded
											{
											Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
											// Insert the uploaded URL in the e-mail message body
											EmailBodyText += "\n       The document is stored at this url: " + objContractSoW.URLonSharePoint;
											objContractSoW.DocumentStatus = enumDocumentStatusses.Uploaded;
											// Delete the uploaded file from the Documents Directory
											if(File.Exists(path: objContractSoW.FileName))
												{
												File.Delete(path: objContractSoW.FileName);
												}
											}
										else // Upload failed Failed
											{
											Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
											objDocCollection.UnexpectedErrors = true;
											objContractSoW.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
											EmailBodyText += "\n       Unfortunately, a technical issue prevented the uploading of "
													+ "the generated document to the Generarated Documents Library on SharePoint.";
											}
										//Check if there were any Unhandled errors and flag the Document's collection
										if(objContractSoW.UnhandledError)
											{
											objDocCollection.UnexpectedErrors = true;
											}
										}
									else // The Document generation failed for some reason
										{
										Console.WriteLine("\t\t *** Unfortunately, the generation of the following document "
											+ "failed unexpectedly : {0}"
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)"
											, objDocumentWorkbook.GetType());
										objDocCollection.UnexpectedErrors = true;
										objContractSoW.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objContractSoW.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									EmailBodyText += "\n\n";
									break;
									}
							//----------------------------------------------
							case ("CSD_based_on_ClientRequirementsMapping"):
									{
									// Prepare to generate the Document
									SuccessfulGeneratedDocument = false;
									CSD_based_on_ClientRequirementsMapping objCSDbasedCRM = objDocumentWorkbook;

									if(objCSDbasedCRM.ErrorMessages == null)
										objCSDbasedCRM.ErrorMessages = new List<string>();

									SuccessfulGeneratedDocument = objCSDbasedCRM.Generate(
										parDataSet: ref Globals.objCompleteDataSet,
										parSDDPdatacontext: this.SDDPdatacontext);

									if(SuccessfulGeneratedDocument)
										{
										// set the Document status to Completed...
										objCSDbasedCRM.DocumentStatus = enumDocumentStatusses.Completed;
										// Prepare the inclusion of the text in the e-mail that the user will receive.
										EmailBodyText += "\n     * " + objDocumentWorkbook.DocumentType;
										// if there were errors, include them in the message.
										if(objCSDbasedCRM.ErrorMessages.Count() > 0)
											{
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objCSDbasedCRM.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following errors occurred:";
											foreach(string errorEntry in objCSDbasedCRM.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else // there were no generation errors.
											{
											Console.WriteLine("\t *** no errors occurred during the generation process.");
											EmailBodyText += ", which was generated without any errors.";
											}

										// begin to upload the document to SharePoint
										objCSDbasedCRM.DocumentStatus = enumDocumentStatusses.Uploading;
										Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

										// Upload the document to the Generated Documents Library
										SuccessfulPublishedDocument = objCSDbasedCRM.UploadDoc(
											parRequestingUserID: objDocCollection.RequestingUserID);
										// Check if the upload succeeded....
										if(SuccessfulPublishedDocument) //Upload Succeeded
											{
											Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
											// Insert the uploaded URL in the e-mail message body
											EmailBodyText += "\n       The document is stored at this url: " + objCSDbasedCRM.URLonSharePoint;
											objCSDbasedCRM.DocumentStatus = enumDocumentStatusses.Uploaded;
											// Delete the uploaded file from the Documents Directory
											if(File.Exists(path: objCSDbasedCRM.FileName))
												{
												File.Delete(path: objCSDbasedCRM.FileName);
												}
											}
										else // Upload failed Failed
											{
											Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
											objDocCollection.UnexpectedErrors = true;
											objCSDbasedCRM.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
											EmailBodyText += "\n       Unfortunately, a technical issue prevented the uploading of "
													+ "the generated document to the Generarated Documents Library on SharePoint.";
											}
										//Check if there were any Unhandled errors and flag the Document's collection
										if(objCSDbasedCRM.UnhandledError)
											{
											objDocCollection.UnexpectedErrors = true;
											}
										}
									else // The Document generation failed for some reason
										{
										Console.WriteLine("\t\t *** Unfortunately, the generation of the following document "
											+ "failed unexpectedly : {0}"
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)"
											, objDocumentWorkbook.GetType());
										objDocCollection.UnexpectedErrors = true;
										objCSDbasedCRM.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objCSDbasedCRM.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									EmailBodyText += "\n\n";
									break;
									}
							case ("CSD_Document_DRM_Inline"):
									{
									// Prepare to generate the Document
									SuccessfulGeneratedDocument = false;
									CSD_Document_DRM_Inline objCSDdrmInline = objDocumentWorkbook;

									if(objCSDdrmInline.ErrorMessages == null)
										objCSDdrmInline.ErrorMessages = new List<string>();

									SuccessfulGeneratedDocument = objCSDdrmInline.Generate(
										parDataSet: ref Globals.objCompleteDataSet,
										parSDDPdatacontext: this.SDDPdatacontext);

									if(SuccessfulGeneratedDocument)
										{
										// set the Document status to Completed...
										objCSDdrmInline.DocumentStatus = enumDocumentStatusses.Completed;
										// Prepare the inclusion of the text in the e-mail that the user will receive.
										EmailBodyText += "\n     * " + objDocumentWorkbook.DocumentType;
										// if there were errors, include them in the message.
										if(objCSDdrmInline.ErrorMessages.Count() > 0)
											{
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objCSDdrmInline.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following errors occurred:";
											foreach(string errorEntry in objCSDdrmInline.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else // there were no generation errors.
											{
											Console.WriteLine("\t *** no errors occurred during the generation process.");
											EmailBodyText += ", which was generated without any errors.";
											}

										// begin to upload the document to SharePoint
										objCSDdrmInline.DocumentStatus = enumDocumentStatusses.Uploading;
										Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

										// Upload the document to the Generated Documents Library
										SuccessfulPublishedDocument = objCSDdrmInline.UploadDoc(
											parRequestingUserID: objDocCollection.RequestingUserID);
										// Check if the upload succeeded....
										if(SuccessfulPublishedDocument) //Upload Succeeded
											{
											Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
											// Insert the uploaded URL in the e-mail message body
											EmailBodyText += "\n       The document is stored at this url: " + objCSDdrmInline.URLonSharePoint;
											objCSDdrmInline.DocumentStatus = enumDocumentStatusses.Uploaded;
											// Delete the uploaded file from the Documents Directory
											if(File.Exists(path: objCSDdrmInline.FileName))
												{
												File.Delete(path: objCSDdrmInline.FileName);
												}
											}
										else // Upload failed Failed
											{
											Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
											objDocCollection.UnexpectedErrors = true;
											objCSDdrmInline.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
											EmailBodyText += "\n       Unfortunately, a technical issue prevented the uploading of "
													+ "the generated document to the Generarated Documents Library on SharePoint.";
											}
										//Check if there were any Unhandled errors and flag the Document's collection
										if(objCSDdrmInline.UnhandledError)
											{
											objDocCollection.UnexpectedErrors = true;
											}
										}
									else // The Document generation failed for some reason
										{
										Console.WriteLine("\t\t *** Unfortunately, the generation of the following document "
											+ "failed unexpectedly : {0}"
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)"
											, objDocumentWorkbook.GetType());
										objDocCollection.UnexpectedErrors = true;
										objCSDdrmInline.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objCSDdrmInline.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									EmailBodyText += "\n\n";
									break;
									}
							//----------------------------------------
							case ("CSD_Document_DRM_Sections"):
									{
									// Prepare to generate the Document
									SuccessfulGeneratedDocument = false;
									CSD_Document_DRM_Sections objCSDdrmSections = objDocumentWorkbook;

									if(objCSDdrmSections.ErrorMessages == null)
										objCSDdrmSections.ErrorMessages = new List<string>();

									SuccessfulGeneratedDocument = objCSDdrmSections.Generate(
										parDataSet: ref Globals.objCompleteDataSet,
										parSDDPdatacontext: this.SDDPdatacontext);

									if(SuccessfulGeneratedDocument)
										{
										// set the Document status to Completed...
										objCSDdrmSections.DocumentStatus = enumDocumentStatusses.Completed;
										// Prepare the inclusion of the text in the e-mail that the user will receive.
										EmailBodyText += "\n     * " + objDocumentWorkbook.DocumentType;
										// if there were errors, include them in the message.
										if(objCSDdrmSections.ErrorMessages.Count() > 0)
											{
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objCSDdrmSections.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following errors occurred:";
											foreach(string errorEntry in objCSDdrmSections.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else // there were no generation errors.
											{
											Console.WriteLine("\t *** no errors occurred during the generation process.");
											EmailBodyText += ", which was generated without any errors.";
											}

										// begin to upload the document to SharePoint
										objCSDdrmSections.DocumentStatus = enumDocumentStatusses.Uploading;
										Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

										// Upload the document to the Generated Documents Library
										SuccessfulPublishedDocument = objCSDdrmSections.UploadDoc(
											parRequestingUserID: objDocCollection.RequestingUserID);
										// Check if the upload succeeded....
										if(SuccessfulPublishedDocument) //Upload Succeeded
											{
											Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
											// Insert the uploaded URL in the e-mail message body
											EmailBodyText += "\n       The document is stored at this url: " + objCSDdrmSections.URLonSharePoint;
											objCSDdrmSections.DocumentStatus = enumDocumentStatusses.Uploaded;
											// Delete the uploaded file from the Documents Directory
											if(File.Exists(path: objCSDdrmSections.FileName))
												{
												File.Delete(path: objCSDdrmSections.FileName);
												}
											}
										else // Upload failed Failed
											{
											Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
											objDocCollection.UnexpectedErrors = true;
											objCSDdrmSections.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
											EmailBodyText += "\n       Unfortunately, a technical issue prevented the uploading of "
													+ "the generated document to the Generarated Documents Library on SharePoint.";
											}
										//Check if there were any Unhandled errors and flag the Document's collection
										if(objCSDdrmSections.UnhandledError)
											{
											objDocCollection.UnexpectedErrors = true;
											}
										}
									else // The Document generation failed for some reason
										{
										Console.WriteLine("\t\t *** Unfortunately, the generation of the following document "
											+ "failed unexpectedly : {0}"
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)"
											, objDocumentWorkbook.GetType());
										objDocCollection.UnexpectedErrors = true;
										objCSDdrmSections.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objCSDdrmSections.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									EmailBodyText += "\n\n";
									break;
									}
							//-----------------------------------------------------
							case ("External_Technology_Coverage_Dashboard_Workbook"):
									{
									// Prepare to generate the Document
									SuccessfulGeneratedDocument = false;
									External_Technology_Coverage_Dashboard_Workbook objExtTechDashboard = objDocumentWorkbook;

									if(objExtTechDashboard.ErrorMessages == null)
										objExtTechDashboard.ErrorMessages = new List<string>();

									SuccessfulGeneratedDocument = objExtTechDashboard.Generate(
										parDataSet: ref Globals.objCompleteDataSet);

									if(SuccessfulGeneratedDocument)
										{
										// set the Document status to Completed...
										objExtTechDashboard.DocumentStatus = enumDocumentStatusses.Completed;
										// Prepare the inclusion of the text in the e-mail that the user will receive.
										EmailBodyText += "\n     * " + objDocumentWorkbook.DocumentType;
										// if there were errors, include them in the message.
										if(objExtTechDashboard.ErrorMessages.Count() > 0)
											{
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objExtTechDashboard.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following errors occurred:";
											foreach(string errorEntry in objExtTechDashboard.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else // there were no generation errors.
											{
											Console.WriteLine("\t *** no errors occurred during the generation process.");
											EmailBodyText += ", which was generated without any errors.";
											}

										// begin to upload the document to SharePoint
										objExtTechDashboard.DocumentStatus = enumDocumentStatusses.Uploading;
										Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

										// Upload the document to the Generated Documents Library
										SuccessfulPublishedDocument = objExtTechDashboard.UploadDoc(
											parRequestingUserID: objDocCollection.RequestingUserID);
										// Check if the upload succeeded....
										if(SuccessfulPublishedDocument) //Upload Succeeded
											{
											Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
											// Insert the uploaded URL in the e-mail message body
											EmailBodyText += "\n       The document is stored at this url: "
												+ objExtTechDashboard.URLonSharePoint;
											objExtTechDashboard.DocumentStatus = enumDocumentStatusses.Uploaded;
											// Delete the uploaded file from the Documents Directory
											if(File.Exists(path: objExtTechDashboard.FileName))
												{
												File.Delete(path: objExtTechDashboard.FileName);
												}
											}
										else // Upload failed Failed
											{
											Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
											objDocCollection.UnexpectedErrors = true;
											objExtTechDashboard.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
											EmailBodyText += "\n       Unfortunately, a technical issue prevented the uploading of "
													+ "the generated document to the Generarated Documents Library on SharePoint.";
											}
										//Check if there were any Unhandled errors and flag the Document's collection
										if(objExtTechDashboard.UnhandledError)
											{
											objDocCollection.UnexpectedErrors = true;
											}
										}
									else // The Document generation failed for some reason
										{
										Console.WriteLine("\t\t *** Unfortunately, the generation of the following document "
											+ "failed unexpectedly : {0}"
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)"
											, objDocumentWorkbook.GetType());
										objDocCollection.UnexpectedErrors = true;
										objExtTechDashboard.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objExtTechDashboard.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									EmailBodyText += "\n\n";
									break;
									}
							//---------------------------------------------------------
							case ("Internal_Technology_Coverage_Dashboard_Workbook"):
									{
									// Prepare to generate the Document
									SuccessfulGeneratedDocument = false;
									Internal_Technology_Coverage_Dashboard_Workbook objIntTechDashboard = objDocumentWorkbook;

									if(objIntTechDashboard.ErrorMessages == null)
										objIntTechDashboard.ErrorMessages = new List<string>();

									SuccessfulGeneratedDocument = objIntTechDashboard.Generate(
										parDataSet: ref Globals.objCompleteDataSet);

									if(SuccessfulGeneratedDocument)
										{
										// set the Document status to Completed...
										objIntTechDashboard.DocumentStatus = enumDocumentStatusses.Completed;
										// Prepare the inclusion of the text in the e-mail that the user will receive.
										EmailBodyText += "\n     * " + objDocumentWorkbook.DocumentType;
										// if there were errors, include them in the message.
										if(objIntTechDashboard.ErrorMessages.Count() > 0)
											{
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objIntTechDashboard.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following errors occurred:";
											foreach(string errorEntry in objIntTechDashboard.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else // there were no generation errors.
											{
											Console.WriteLine("\t *** no errors occurred during the generation process.");
											EmailBodyText += ", which was generated without any errors.";
											}

										// begin to upload the document to SharePoint
										objIntTechDashboard.DocumentStatus = enumDocumentStatusses.Uploading;
										Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

										// Upload the document to the Generated Documents Library
										SuccessfulPublishedDocument = objIntTechDashboard.UploadDoc(
											parRequestingUserID: objDocCollection.RequestingUserID);
										// Check if the upload succeeded....
										if(SuccessfulPublishedDocument) //Upload Succeeded
											{
											Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
											// Insert the uploaded URL in the e-mail message body
											EmailBodyText += "\n       The document is stored at this url: "
												+ objIntTechDashboard.URLonSharePoint;
											objIntTechDashboard.DocumentStatus = enumDocumentStatusses.Uploaded;
											// Delete the uploaded file from the Documents Directory
											if(File.Exists(path: objIntTechDashboard.FileName))
												{
												File.Delete(path: objIntTechDashboard.FileName);
												}
											}
										else // Upload failed Failed
											{
											Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
											objDocCollection.UnexpectedErrors = true;
											objIntTechDashboard.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
											EmailBodyText += "\n       Unfortunately, a technical issue prevented the uploading of "
													+ "the generated document to the Generarated Documents Library on SharePoint.";
											}
										//Check if there were any Unhandled errors and flag the Document's collection
										if(objIntTechDashboard.UnhandledError)
											{
											objDocCollection.UnexpectedErrors = true;
											}
										}
									else // The Document generation failed for some reason
										{
										Console.WriteLine("\t\t *** Unfortunately, the generation of the following document "
											+ "failed unexpectedly : {0}"
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)"
											, objDocumentWorkbook.GetType());
										objDocCollection.UnexpectedErrors = true;
										objIntTechDashboard.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objIntTechDashboard.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									EmailBodyText += "\n\n";
									break;
									}
							//-------------------------------------
							case ("ISD_Document_DRM_Inline"):
									{
									// Prepare to generate the Document
									SuccessfulGeneratedDocument = false;
									ISD_Document_DRM_Inline objISDdrmInline = objDocumentWorkbook;

									if(objISDdrmInline.ErrorMessages == null)
										objISDdrmInline.ErrorMessages = new List<string>();

									SuccessfulGeneratedDocument = objISDdrmInline.Generate(
										parDataSet: ref Globals.objCompleteDataSet,
										parSDDPdatacontext: this.SDDPdatacontext);

									if(SuccessfulGeneratedDocument)
										{
										// set the Document status to Completed...
										objISDdrmInline.DocumentStatus = enumDocumentStatusses.Completed;
										// Prepare the inclusion of the text in the e-mail that the user will receive.
										EmailBodyText += "\n     * " + objDocumentWorkbook.DocumentType;
										// if there were errors, include them in the message.
										if(objISDdrmInline.ErrorMessages.Count() > 0)
											{
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objISDdrmInline.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following errors occurred:";
											foreach(string errorEntry in objISDdrmInline.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else // there were no generation errors.
											{
											Console.WriteLine("\t *** no errors occurred during the generation process.");
											EmailBodyText += ", which was generated without any errors.";
											}

										// begin to upload the document to SharePoint
										objISDdrmInline.DocumentStatus = enumDocumentStatusses.Uploading;
										Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

										// Upload the document to the Generated Documents Library
										SuccessfulPublishedDocument = objISDdrmInline.UploadDoc(
											parRequestingUserID: objDocCollection.RequestingUserID);
										// Check if the upload succeeded....
										if(SuccessfulPublishedDocument) //Upload Succeeded
											{
											Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
											// Insert the uploaded URL in the e-mail message body
											EmailBodyText += "\n       The document is stored at this url: " + objISDdrmInline.URLonSharePoint;
											objISDdrmInline.DocumentStatus = enumDocumentStatusses.Uploaded;
											// Delete the uploaded file from the Documents Directory
											if(File.Exists(path: objISDdrmInline.FileName))
												{
												File.Delete(path: objISDdrmInline.FileName);
												}
											}
										else // Upload failed Failed
											{
											Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
											objDocCollection.UnexpectedErrors = true;
											objISDdrmInline.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
											EmailBodyText += "\n       Unfortunately, a technical issue prevented the uploading of "
													+ "the generated document to the Generarated Documents Library on SharePoint.";
											}
										//Check if there were any Unhandled errors and flag the Document's collection
										if(objISDdrmInline.UnhandledError)
											{
											objDocCollection.UnexpectedErrors = true;
											}
										}
									else // The Document generation failed for some reason
										{
										Console.WriteLine("\t\t *** Unfortunately, the generation of the following document "
											+ "failed unexpectedly : {0}"
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)"
											, objDocumentWorkbook.GetType());
										objDocCollection.UnexpectedErrors = true;
										objISDdrmInline.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objISDdrmInline.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									EmailBodyText += "\n\n";
									break;
									}
							//-------------------------------------
							case ("ISD_Document_DRM_Sections"):
									{
									// Prepare to generate the Document
									SuccessfulGeneratedDocument = false;
									ISD_Document_DRM_Sections objISDdrmSections = objDocumentWorkbook;

									if(objISDdrmSections.ErrorMessages == null)
										objISDdrmSections.ErrorMessages = new List<string>();

									SuccessfulGeneratedDocument = objISDdrmSections.Generate(
										parDataSet: ref Globals.objCompleteDataSet,
										parSDDPdatacontext: this.SDDPdatacontext);

									if(SuccessfulGeneratedDocument)
										{
										// set the Document status to Completed...
										objISDdrmSections.DocumentStatus = enumDocumentStatusses.Completed;
										// Prepare the inclusion of the text in the e-mail that the user will receive.
										EmailBodyText += "\n     * " + objDocumentWorkbook.DocumentType;
										// if there were errors, include them in the message.
										if(objISDdrmSections.ErrorMessages.Count() > 0)
											{
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objISDdrmSections.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following errors occurred:";
											foreach(string errorEntry in objISDdrmSections.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else // there were no generation errors.
											{
											Console.WriteLine("\t *** no errors occurred during the generation process.");
											EmailBodyText += ", which was generated without any errors.";
											}

										// begin to upload the document to SharePoint
										objISDdrmSections.DocumentStatus = enumDocumentStatusses.Uploading;
										Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

										// Upload the document to the Generated Documents Library
										SuccessfulPublishedDocument = objISDdrmSections.UploadDoc(
											parRequestingUserID: objDocCollection.RequestingUserID);
										// Check if the upload succeeded....
										if(SuccessfulPublishedDocument) //Upload Succeeded
											{
											Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
											// Insert the uploaded URL in the e-mail message body
											EmailBodyText += "\n       The document is stored at this url: " + objISDdrmSections.URLonSharePoint;
											objISDdrmSections.DocumentStatus = enumDocumentStatusses.Uploaded;
											// Delete the uploaded file from the Documents Directory
											if(File.Exists(path: objISDdrmSections.FileName))
												{
												File.Delete(path: objISDdrmSections.FileName);
												}
											}
										else // Upload failed Failed
											{
											Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
											objDocCollection.UnexpectedErrors = true;
											objISDdrmSections.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
											EmailBodyText += "\n       Unfortunately, a technical issue prevented the uploading of "
													+ "the generated document to the Generarated Documents Library on SharePoint.";
											}
										//Check if there were any Unhandled errors and flag the Document's collection
										if(objISDdrmSections.UnhandledError)
											{
											objDocCollection.UnexpectedErrors = true;
											}
										}
									else // The Document generation failed for some reason
										{
										Console.WriteLine("\t\t *** Unfortunately, the generation of the following document "
											+ "failed unexpectedly : {0}"
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)"
											, objDocumentWorkbook.GetType());
										objDocCollection.UnexpectedErrors = true;
										objISDdrmSections.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objISDdrmSections.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									EmailBodyText += "\n\n";
									break;
									}
							//------------------------------------------
							case ("Pricing_Addendum_Document"):
									{
									// Prepare to generate the Document
									SuccessfulGeneratedDocument = false;
									Pricing_Addendum_Document objPricingAddendum = objDocumentWorkbook;

									if(objPricingAddendum.ErrorMessages == null)
										objPricingAddendum.ErrorMessages = new List<string>();

									//bGenerateDocumentSuccessful = objPricingAddendum.Generate(
									//	parDataSet: ref Globals.objDataSet);

									if(SuccessfulGeneratedDocument)
										{
										// set the Document status to Completed...
										objPricingAddendum.DocumentStatus = enumDocumentStatusses.Completed;
										// Prepare the inclusion of the text in the e-mail that the user will receive.
										EmailBodyText += "\n     * " + objDocumentWorkbook.DocumentType;
										// if there were errors, include them in the message.
										if(objPricingAddendum.ErrorMessages.Count() > 0)
											{
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objPricingAddendum.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following errors occurred:";
											foreach(string errorEntry in objPricingAddendum.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else // there were no generation errors.
											{
											Console.WriteLine("\t *** no errors occurred during the generation process.");
											EmailBodyText += ", which was generated without any errors.";
											}

										// begin to upload the document to SharePoint
										objPricingAddendum.DocumentStatus = enumDocumentStatusses.Uploading;
										Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

										// Upload the document to the Generated Documents Library
										SuccessfulPublishedDocument = objPricingAddendum.UploadDoc(
											parRequestingUserID: objDocCollection.RequestingUserID);
										// Check if the upload succeeded....
										if(SuccessfulPublishedDocument) //Upload Succeeded
											{
											Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
											// Insert the uploaded URL in the e-mail message body
											EmailBodyText += "\n       The document is stored at this url: " + objPricingAddendum.URLonSharePoint;
											objPricingAddendum.DocumentStatus = enumDocumentStatusses.Uploaded;
											// Delete the uploaded file from the Documents Directory
											if(File.Exists(path: objPricingAddendum.FileName))
												{
												File.Delete(path: objPricingAddendum.FileName);
												}
											}
										else // Upload failed Failed
											{
											Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
											objDocCollection.UnexpectedErrors = true;
											objPricingAddendum.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
											EmailBodyText += "\n       Unfortunately, a technical issue prevented the uploading of "
													+ "the generated document to the Generarated Documents Library on SharePoint.";
											}
										//Check if there were any Unhandled errors and flag the Document's collection
										if(objPricingAddendum.UnhandledError)
											{
											objDocCollection.UnexpectedErrors = true;
											}
										}
									else // The Document generation failed for some reason
										{
										Console.WriteLine("\t\t *** Unfortunately, the {0} "
											+ "is not implemented at the moment because the Pricing Methodology is still Work in Progress."
											, objDocumentWorkbook.GetType());
										objDocCollection.UnexpectedErrors = false;
										//objPricingAddendum.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unfortunately, the " + objPricingAddendum.DocumentType
											+ " is not implemented at the moment because the Pricing Methodology is still Work in Progress.";
										}
									EmailBodyText += "\n\n";
									break;
									}
							//--------------------------------------
							case ("RACI_Matrix_Workbook_per_Deliverable"):
									{
									// Prepare to generate the Document
									SuccessfulGeneratedDocument = false;
									RACI_Matrix_Workbook_per_Deliverable objRACImatrix = objDocumentWorkbook;

									if(objRACImatrix.ErrorMessages == null)
										objRACImatrix.ErrorMessages = new List<string>();

									SuccessfulGeneratedDocument = objRACImatrix.Generate(
										parDataSet: ref Globals.objCompleteDataSet);

									if(SuccessfulGeneratedDocument)
										{
										// set the Document status to Completed...
										objRACImatrix.DocumentStatus = enumDocumentStatusses.Completed;
										// Prepare the inclusion of the text in the e-mail that the user will receive.
										EmailBodyText += "\n     * " + objDocumentWorkbook.DocumentType;
										// if there were errors, include them in the message.
										if(objRACImatrix.ErrorMessages.Count() > 0)
											{
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objRACImatrix.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following errors occurred:";
											foreach(string errorEntry in objRACImatrix.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else // there were no generation errors.
											{
											Console.WriteLine("\t *** no errors occurred during the generation process.");
											EmailBodyText += ", which was generated without any errors.";
											}

										// begin to upload the document to SharePoint
										objRACImatrix.DocumentStatus = enumDocumentStatusses.Uploading;
										Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

										// Upload the document to the Generated Documents Library
										SuccessfulPublishedDocument = objRACImatrix.UploadDoc(
											parRequestingUserID: objDocCollection.RequestingUserID);
										// Check if the upload succeeded....
										if(SuccessfulPublishedDocument) //Upload Succeeded
											{
											Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
											// Insert the uploaded URL in the e-mail message body
											EmailBodyText += "\n       The document is stored at this url: " + objRACImatrix.URLonSharePoint;
											objRACImatrix.DocumentStatus = enumDocumentStatusses.Uploaded;
											// Delete the uploaded file from the Documents Directory
											if(File.Exists(path: objRACImatrix.FileName))
												{
												File.Delete(path: objRACImatrix.FileName);
												}
											}
										else // Upload failed Failed
											{
											Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
											objDocCollection.UnexpectedErrors = true;
											objRACImatrix.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
											EmailBodyText += "\n       Unfortunately, a technical issue prevented the uploading of "
													+ "the generated document to the Generarated Documents Library on SharePoint.";
											}
										//Check if there were any Unhandled errors and flag the Document's collection
										if(objRACImatrix.UnhandledError)
											{
											objDocCollection.UnexpectedErrors = true;
											}
										}
									else // The Document generation failed for some reason
										{
										Console.WriteLine("\t\t *** Unfortunately, the generation of the following document "
											+ "failed unexpectedly : {0}"
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)"
											, objDocumentWorkbook.GetType());
										objDocCollection.UnexpectedErrors = true;
										objRACImatrix.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objRACImatrix.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									EmailBodyText += "\n\n";
									break;
									}
							//-----------------------------------
							case ("RACI_Workbook_per_Role"):
									{
									// Prepare to generate the Document
									SuccessfulGeneratedDocument = false;
									RACI_Workbook_per_Role objRACIperRole = objDocumentWorkbook;

									if(objRACIperRole.ErrorMessages == null)
										objRACIperRole.ErrorMessages = new List<string>();

									SuccessfulGeneratedDocument = objRACIperRole.Generate(
										parDataSet: ref Globals.objCompleteDataSet);

									if(SuccessfulGeneratedDocument)
										{
										// set the Document status to Completed...
										objRACIperRole.DocumentStatus = enumDocumentStatusses.Completed;
										// Prepare the inclusion of the text in the e-mail that the user will receive.
										EmailBodyText += "\n     * " + objDocumentWorkbook.DocumentType;
										// if there were errors, include them in the message.
										if(objRACIperRole.ErrorMessages.Count() > 0)
											{
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objRACIperRole.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following errors occurred:";
											foreach(string errorEntry in objRACIperRole.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else // there were no generation errors.
											{
											Console.WriteLine("\t *** no errors occurred during the generation process.");
											EmailBodyText += ", which was generated without any errors.";
											}

										// begin to upload the document to SharePoint
										objRACIperRole.DocumentStatus = enumDocumentStatusses.Uploading;
										Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

										// Upload the document to the Generated Documents Library
										SuccessfulPublishedDocument = objRACIperRole.UploadDoc(
											parRequestingUserID: objDocCollection.RequestingUserID);
										// Check if the upload succeeded....
										if(SuccessfulPublishedDocument) //Upload Succeeded
											{
											Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
											// Insert the uploaded URL in the e-mail message body
											EmailBodyText += "\n       The document is stored at this url: " + objRACIperRole.URLonSharePoint;
											objRACIperRole.DocumentStatus = enumDocumentStatusses.Uploaded;
											// Delete the uploaded file from the Documents Directory
											if(File.Exists(path: objRACIperRole.FileName))
												{
												File.Delete(path: objRACIperRole.FileName);
												}
											}
										else // Upload failed Failed
											{
											Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
											objDocCollection.UnexpectedErrors = true;
											objRACIperRole.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
											EmailBodyText += "\n       Unfortunately, a technical issue prevented the uploading of "
													+ "the generated document to the Generarated Documents Library on SharePoint.";
											}
										//Check if there were any Unhandled errors and flag the Document's collection
										if(objRACIperRole.UnhandledError)
											{
											objDocCollection.UnexpectedErrors = true;
											}
										}
									else // The Document generation failed for some reason
										{
										Console.WriteLine("\t\t *** Unfortunately, the generation of the following document "
											+ "failed unexpectedly : {0}"
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)"
											, objDocumentWorkbook.GetType());
										objDocCollection.UnexpectedErrors = true;
										objRACIperRole.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objRACIperRole.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									EmailBodyText += "\n\n";
									break;
									}
							//----------------------------------------
							case ("Services_Framework_Document_DRM_Inline"):
									{
									// Prepare to generate the Document
									SuccessfulGeneratedDocument = false;
									Services_Framework_Document_DRM_Inline objSFdrmInline = objDocumentWorkbook;

									if(objSFdrmInline.ErrorMessages == null)
										objSFdrmInline.ErrorMessages = new List<string>();

									SuccessfulGeneratedDocument = objSFdrmInline.Generate(
										parDataSet: ref Globals.objCompleteDataSet,
										parSDDPdatacontext: this.SDDPdatacontext);

									if(SuccessfulGeneratedDocument)
										{
										// set the Document status to Completed...
										objSFdrmInline.DocumentStatus = enumDocumentStatusses.Completed;
										// Prepare the inclusion of the text in the e-mail that the user will receive.
										EmailBodyText += "\n     * " + objDocumentWorkbook.DocumentType;
										// if there were errors, include them in the message.
										if(objSFdrmInline.ErrorMessages.Count() > 0)
											{
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objSFdrmInline.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following errors occurred:";
											foreach(string errorEntry in objSFdrmInline.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else // there were no generation errors.
											{
											Console.WriteLine("\t *** no errors occurred during the generation process.");
											EmailBodyText += ", which was generated without any errors.";
											}

										// begin to upload the document to SharePoint
										objSFdrmInline.DocumentStatus = enumDocumentStatusses.Uploading;
										Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

										// Upload the document to the Generated Documents Library
										SuccessfulPublishedDocument = objSFdrmInline.UploadDoc(
											parRequestingUserID: objDocCollection.RequestingUserID);
										// Check if the upload succeeded....
										if(SuccessfulPublishedDocument) //Upload Succeeded
											{
											Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
											// Insert the uploaded URL in the e-mail message body
											EmailBodyText += "\n       The document is stored at this url: " + objSFdrmInline.URLonSharePoint;
											objSFdrmInline.DocumentStatus = enumDocumentStatusses.Uploaded;
											// Delete the uploaded file from the Documents Directory
											if(File.Exists(path: objSFdrmInline.FileName))
												{
												File.Delete(path: objSFdrmInline.FileName);
												}
											}
										else // Upload failed Failed
											{
											Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
											objDocCollection.UnexpectedErrors = true;
											objSFdrmInline.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
											EmailBodyText += "\n       Unfortunately, a technical issue prevented the uploading of "
													+ "the generated document to the Generarated Documents Library on SharePoint.";
											}
										//Check if there were any Unhandled errors and flag the Document's collection
										if(objSFdrmInline.UnhandledError)
											{
											objDocCollection.UnexpectedErrors = true;
											}
										}
									else // The Document generation failed for some reason
										{
										Console.WriteLine("\t\t *** Unfortunately, the generation of the following document "
											+ "failed unexpectedly : {0}"
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)"
											, objDocumentWorkbook.GetType());
										objDocCollection.UnexpectedErrors = true;
										objSFdrmInline.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objSFdrmInline.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									EmailBodyText += "\n\n";
									break;
									}
							//---------------------------------------------
							case ("Services_Framework_Document_DRM_Sections"):
									{
									// Prepare to generate the Document
									SuccessfulGeneratedDocument = false;
									Services_Framework_Document_DRM_Sections objSFdrmSections = objDocumentWorkbook;

									if(objSFdrmSections.ErrorMessages == null)
										objSFdrmSections.ErrorMessages = new List<string>();

									SuccessfulGeneratedDocument = objSFdrmSections.Generate(
										parDataSet: ref Globals.objCompleteDataSet,
										parSDDPdatacontext: this.SDDPdatacontext);

									if(SuccessfulGeneratedDocument)
										{
										// set the Document status to Completed...
										objSFdrmSections.DocumentStatus = enumDocumentStatusses.Completed;
										// Prepare the inclusion of the text in the e-mail that the user will receive.
										EmailBodyText += "\n     * " + objDocumentWorkbook.DocumentType;
										// if there were errors, include them in the message.
										if(objSFdrmSections.ErrorMessages.Count() > 0)
											{
											Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
												objSFdrmSections.ErrorMessages.Count);
											EmailBodyText += ", which was generated but the following errors occurred:";
											foreach(string errorEntry in objSFdrmSections.ErrorMessages)
												{
												EmailBodyText += "\n          + " + errorEntry;
												Console.WriteLine("\t\t\t + {0}", errorEntry);
												}
											}
										else // there were no generation errors.
											{
											Console.WriteLine("\t *** no errors occurred during the generation process.");
											EmailBodyText += ", which was generated without any errors.";
											}

										// begin to upload the document to SharePoint
										objSFdrmSections.DocumentStatus = enumDocumentStatusses.Uploading;
										Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

										// Upload the document to the Generated Documents Library
										SuccessfulPublishedDocument = objSFdrmSections.UploadDoc(
											parRequestingUserID: objDocCollection.RequestingUserID);
										// Check if the upload succeeded....
										if(SuccessfulPublishedDocument) //Upload Succeeded
											{
											Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
											// Insert the uploaded URL in the e-mail message body
											EmailBodyText += "\n       The document is stored at this url: " + objSFdrmSections.URLonSharePoint;
											objSFdrmSections.DocumentStatus = enumDocumentStatusses.Uploaded;
											// Delete the uploaded file from the Documents Directory
											if(File.Exists(path: objSFdrmSections.FileName))
												{
												File.Delete(path: objSFdrmSections.FileName);
												}
											}
										else // Upload failed Failed
											{
											Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
											objDocCollection.UnexpectedErrors = true;
											objSFdrmSections.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
											EmailBodyText += "\n       Unfortunately, a technical issue prevented the uploading of "
													+ "the generated document to the Generarated Documents Library on SharePoint.";
											}
										//Check if there were any Unhandled errors and flag the Document's collection
										if(objSFdrmSections.UnhandledError)
											{
											objDocCollection.UnexpectedErrors = true;
											}
										}
									else // The Document generation failed for some reason
										{
										Console.WriteLine("\t\t *** Unfortunately, the generation of the following document "
											+ "failed unexpectedly : {0}"
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)"
											, objDocumentWorkbook.GetType());
										objDocCollection.UnexpectedErrors = true;
										objSFdrmSections.ErrorMessages.Add("Error: Document Generation unexpectedly failed...");
										EmailBodyText += "\n\t - Unable to complete the generation of document: "
											+ objSFdrmSections.DocumentType
											+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
										}
									EmailBodyText += "\n\n";
									break;
									}
								} // switch (objectType)
							} // foreach(dynamic objDocumentWorkbook in objDocCollection.Documen_and_Workbook_Objects...

						//--------------------------------------------------------------------------
						// Process the Notification via E-mail if the users selected to be notified.
						if(objDocCollection.NotifyMe && objDocCollection.NotificationEmail != null)
							{
							SuccessfulSentEmail = eMail.SendEmail(
							parRecipient: objDocCollection.NotificationEmail,
							parSubject: "SDDP: Generated Document(s)",
							parBody: EmailBodyText);

							if(SuccessfulSentEmail)
								{
								Console.WriteLine("Sending e-mail successfully send to user!");
								}
							else
								{
								Console.WriteLine("*** ERROR *** \n Sending e-mail failed...\n");
								}
							}
						//----------------------------------------------------------------------------
						// Check if there were unexpected errors and if there were, send an e-mail to the Technical Support team.
						if(objDocCollection.UnexpectedErrors)
							{
							this.SuccessfullUpdatedDocCollection = objDocCollection.UpdateGenerateStatus(
								parGenerationStatus: enumGenerationStatus.Failed);

							if(this.SuccessfullUpdatedDocCollection)
								Console.WriteLine("Update Document Collection Status to 'FAILED' was successful.");
							else
								Console.WriteLine("Update Document Collection Status to 'FAILED' was unsuccessful.");

							// Prepare the e-mail
							SuccessfulSentEmail = eMail.SendEmail(
								parRecipient: Properties.AppResources.Email_Bcc_Address,
								parSubject: "SDDP: Unexpected DocGenerator(s) Error occurred.)",
								parBody: EmailBodyText,
								parSendBcc: false);

							if(SuccessfulSentEmail)
								{
								Console.WriteLine("The error e-mail was successfully send to the technical team.");
								}
							else
								{
								Console.WriteLine("The error e-mail to the technical team FAILED!");
								}
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
					parRecipient: Properties.AppResources.Email_Bcc_Address,
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
					parRecipient: Properties.AppResources.Email_Bcc_Address,
					parSubject: "SDDP: Unexpected DocGenerator(s) Error occurred.)",
					parBody: EmailBodyText,
					parSendBcc: false);
				}

Procedure_Ends:
			Console.WriteLine("end of MainController in DocGeneratorCore.");
			this.SDDPdatacontext = null;
			return;
			} // end of method
		} // end of class
	} // end of Namespace