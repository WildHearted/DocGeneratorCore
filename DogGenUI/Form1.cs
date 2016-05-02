using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Data.Services.Client;
using System.Linq;
using System.Net;
using System.Windows.Forms;
using Microsoft.SharePoint.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using DocGenerator.ServiceReferenceSDDP;

namespace DocGenerator
	{
		
	public partial class Form1 : Form
		{

	
	public string ErrorLogMessage = "";

		public Form1()
			{
			InitializeComponent();
			}

		private void btnSDDP_Click(object sender, EventArgs e)
			{
			Cursor.Current = Cursors.WaitCursor;

			//	Declare the SharePoint connection as a DataContext
			DesignAndDeliveryPortfolioDataContext objSDDPdatacontext = new DesignAndDeliveryPortfolioDataContext(new
				Uri(Properties.AppResources.SharePointSiteURL + Properties.AppResources.SharePointRESTuri));
			//objSDDPdatacontext.Credentials = CredentialCache.DefaultCredentials;
			objSDDPdatacontext.Credentials = new NetworkCredential(
				userName: Properties.AppResources.User_Credentials_UserName,
				password: Properties.AppResources.User_Credentials_Password,
				domain: Properties.AppResources.User_Credentials_Domain);
			objSDDPdatacontext.MergeOption = MergeOption.NoTracking;
			//objSDDPdatacontext.ObjectTrackingEnabled = false;

			Console.WriteLine("Checking the Document Collection Library for any documents to generate...");

			string sReturnResult = "";
			string sEmailBody = "";
			bool bGenerateDocumentSuccessful = false;
			bool bPublishDocSuccessful = false;
			bool bSendEmailSuccessful = false;
			bool bUpdateDocColSuccessful = false;
			List<DocumentCollection> docCollectionsToGenerate = new List<DocumentCollection>();

			try
				{
				sReturnResult = DocumentCollection.GetCollectionsToGenerate(ref docCollectionsToGenerate, parSDDPdatacontext: objSDDPdatacontext);
				if (sReturnResult.Substring(0,4) == "Good")
					{
					Console.WriteLine("\r\nThere are {0} Document Collections to generate.", docCollectionsToGenerate.Count());
					}
				else if(sReturnResult.Substring(0,5) == "Error")
					{
					Console.WriteLine("\nERROR: There was an error, could not generate some documents for the Document Collections. \n{0}", sReturnResult);
					goto Procedure_Ends;
					}
				}
			catch(InvalidProgramException ex)
				{
				Console.WriteLine("\n\nException occurred [{0}] \n Inner Exception: {1}", ex.Message, ex.InnerException);
				goto Procedure_Ends;
				}
	
			string objectType = "";
			try
				{
				if(docCollectionsToGenerate.Count > 0)
					{
					// Continue here if there are any Document Collections to generate...
					// Load the complete DataSet before beginning to generate the documents - to ensure optimal Document Generation performance
					if(Globals.objDataSet == null)
						{
						//CompleteDataSet objDataSet = new CompleteDataSet();
						Globals.objDataSet = new CompleteDataSet();
						if(!Globals.objDataSet.PopulateBaseObjects(parDatacontexSDDP: objSDDPdatacontext))
							{
							MessageBox.Show("Unable to connect to SharePoint, please check the connection.", 
								"SharePoint not reachable", MessageBoxButtons.OK, MessageBoxIcon.Error);
							goto Procedure_Ends;
							}
						}

					foreach(DocumentCollection objDocCollection in docCollectionsToGenerate)
						{
						Console.WriteLine("\r\nReady to generate Document Collection: {0} - {1}", objDocCollection.ID.ToString(),
							objDocCollection.Title);
						objDocCollection.UnexpectedErrors = false;

						//Prepare the  E-mail that will be send to the user...
						sEmailBody = "Good day,\n\nHerewith the generated document(s) that you requested from the Service Design and Delivery Portfolio "
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
										bGenerateDocumentSuccessful = false;
										Client_Requirements_Mapping_Workbook objCRMworkbook = objDocumentWorkbook;

										if(objCRMworkbook.ErrorMessages == null)
											objCRMworkbook.ErrorMessages = new List<string>();

										bGenerateDocumentSuccessful = objCRMworkbook.Generate(
											parDataSet: ref Globals.objDataSet,
											parSDDPdatacontext: objSDDPdatacontext);

										if(bGenerateDocumentSuccessful)
											{
											// set the Document status to Completed...
											objCRMworkbook.DocumentStatus = enumDocumentStatusses.Completed;
											// Prepare the inclusion of the text in the e-mail that the user will receive.
											sEmailBody += "\n     * " + objDocumentWorkbook.DocumentType;
											// if there were errors, include them in the message.
											if(objCRMworkbook.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
													objCRMworkbook.ErrorMessages.Count);
												sEmailBody += ", which was generated but the following errors occurred:";
												foreach(string errorEntry in objCRMworkbook.ErrorMessages)
													{
													sEmailBody += "\n          + " + errorEntry;
													Console.WriteLine("\t\t\t + {0}", errorEntry);
													}
												}
											else // there were no generation errors.
												{
												Console.WriteLine("\t *** no errors occurred during the generation process.");
												sEmailBody += ", which was generated without any errors.";
												}

											// begin to upload the document to SharePoint
											objCRMworkbook.DocumentStatus = enumDocumentStatusses.Uploading;
											Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

											// Upload the document to the Generated Documents Library
											bPublishDocSuccessful = objCRMworkbook.UploadDoc(
												parRequestingUserID: objDocCollection.RequestingUserID);
											// Check if the upload succeeded....
											if(bPublishDocSuccessful) //Upload Succeeded
												{
												Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
												// Insert the uploaded URL in the e-mail message body
												sEmailBody += "\n       The document is stored at this url: " + objCRMworkbook.URLonSharePoint;
												objCRMworkbook.DocumentStatus = enumDocumentStatusses.Uploaded;
												}
											else // Upload failed Failed
												{
												Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
												objDocCollection.UnexpectedErrors = true;
												objCRMworkbook.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
												sEmailBody += "\n       Unfortunately, a technical issue prevented the uploading of "
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
											sEmailBody += "\n\t - Unable to complete the generation of document: "
												+ objCRMworkbook.DocumentType
												+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
											}
										sEmailBody += "\n\n";
										break;
										}
									//---------------------------------------
									case ("Content_Status_Workbook"):
										{
										// Prepare to generate the Document
										bGenerateDocumentSuccessful = false;
										Content_Status_Workbook objcontentStatus = objDocumentWorkbook;

										if(objcontentStatus.ErrorMessages == null)
											objcontentStatus.ErrorMessages = new List<string>();
										bGenerateDocumentSuccessful = objcontentStatus.Generate(
											parDataSet: ref Globals.objDataSet);

										if(bGenerateDocumentSuccessful)
											{
											// set the Document status to Completed...
											objcontentStatus.DocumentStatus = enumDocumentStatusses.Completed;
											// Prepare the inclusion of the text in the e-mail that the user will receive.
											sEmailBody += "\n     * " + objDocumentWorkbook.DocumentType;
											// if there were errors, include them in the message.
											if(objcontentStatus.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
													objcontentStatus.ErrorMessages.Count);
												sEmailBody += ", which was generated but the following errors occurred:";
												foreach(string errorEntry in objcontentStatus.ErrorMessages)
													{
													sEmailBody += "\n          + " + errorEntry;
													Console.WriteLine("\t\t\t + {0}", errorEntry);
													}
												}
											else // there were no generation errors.
												{
												Console.WriteLine("\t *** no errors occurred during the generation process.");
												sEmailBody += ", which was generated without any errors.";
												}

											// begin to upload the document to SharePoint
											objcontentStatus.DocumentStatus = enumDocumentStatusses.Uploading;
											Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

											// Upload the document to the Generated Documents Library
											bPublishDocSuccessful = objcontentStatus.UploadDoc(
												parRequestingUserID: objDocCollection.RequestingUserID);
											// Check if the upload succeeded....
											if(bPublishDocSuccessful) //Upload Succeeded
												{
												Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
												// Insert the uploaded URL in the e-mail message body
												sEmailBody += "\n       The document is stored at this url: " + objcontentStatus.URLonSharePoint;
												objcontentStatus.DocumentStatus = enumDocumentStatusses.Uploaded;
												}
											else // Upload failed Failed
												{
												Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
												objDocCollection.UnexpectedErrors = true;
												objcontentStatus.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
												sEmailBody += "\n       Unfortunately, a technical issue prevented the uploading of "
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
											sEmailBody += "\n\t - Unable to complete the generation of document: "
												+ objcontentStatus.DocumentType
												+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
											}
										sEmailBody += "\n\n";
										break;
										}
									//--------------------------------------------
									case ("Contract_SoW_Service_Description"):
										{
										// Prepare to generate the Document
										bGenerateDocumentSuccessful = false;
										Contract_SoW_Service_Description objContractSoW = objDocumentWorkbook;

										if(objContractSoW.ErrorMessages == null)
											objContractSoW.ErrorMessages = new List<string>();

										bGenerateDocumentSuccessful = objContractSoW.Generate(
											parDataSet: ref Globals.objDataSet,
											parSDDPdatacontext: objSDDPdatacontext);

										if(bGenerateDocumentSuccessful)
											{
											// set the Document status to Completed...
											objContractSoW.DocumentStatus = enumDocumentStatusses.Completed;
											// Prepare the inclusion of the text in the e-mail that the user will receive.
											sEmailBody += "\n     * " + objDocumentWorkbook.DocumentType;
											// if there were errors, include them in the message.
											if(objContractSoW.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
													objContractSoW.ErrorMessages.Count);
												sEmailBody += ", which was generated but the following errors occurred:";
												foreach(string errorEntry in objContractSoW.ErrorMessages)
													{
													sEmailBody += "\n          + " + errorEntry;
													Console.WriteLine("\t\t\t + {0}", errorEntry);
													}
												}
											else // there were no generation errors.
												{
												Console.WriteLine("\t *** no errors occurred during the generation process.");
												sEmailBody += ", which was generated without any errors.";
												}

											// begin to upload the document to SharePoint
											objContractSoW.DocumentStatus = enumDocumentStatusses.Uploading;
											Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

											// Upload the document to the Generated Documents Library
											bPublishDocSuccessful = objContractSoW.UploadDoc(
												parRequestingUserID: objDocCollection.RequestingUserID);
											// Check if the upload succeeded....
											if(bPublishDocSuccessful) //Upload Succeeded
												{
												Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
												// Insert the uploaded URL in the e-mail message body
												sEmailBody += "\n       The document is stored at this url: " + objContractSoW.URLonSharePoint;
												objContractSoW.DocumentStatus = enumDocumentStatusses.Uploaded;
												}
											else // Upload failed Failed
												{
												Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
												objDocCollection.UnexpectedErrors = true;
												objContractSoW.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
												sEmailBody += "\n       Unfortunately, a technical issue prevented the uploading of "
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
											sEmailBody += "\n\t - Unable to complete the generation of document: "
												+ objContractSoW.DocumentType
												+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
											}
										sEmailBody += "\n\n";
										break;
										}
									//----------------------------------------------
									case ("CSD_based_on_ClientRequirementsMapping"):
										{
										// Prepare to generate the Document
										bGenerateDocumentSuccessful = false;
										CSD_based_on_ClientRequirementsMapping objCSDbasedCRM = objDocumentWorkbook;

										if(objCSDbasedCRM.ErrorMessages == null)
											objCSDbasedCRM.ErrorMessages = new List<string>();

										bGenerateDocumentSuccessful = objCSDbasedCRM.Generate(
											parDataSet: ref Globals.objDataSet,
											parSDDPdatacontext: objSDDPdatacontext);

										if(bGenerateDocumentSuccessful)
											{
											// set the Document status to Completed...
											objCSDbasedCRM.DocumentStatus = enumDocumentStatusses.Completed;
											// Prepare the inclusion of the text in the e-mail that the user will receive.
											sEmailBody += "\n     * " + objDocumentWorkbook.DocumentType;
											// if there were errors, include them in the message.
											if(objCSDbasedCRM.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
													objCSDbasedCRM.ErrorMessages.Count);
												sEmailBody += ", which was generated but the following errors occurred:";
												foreach(string errorEntry in objCSDbasedCRM.ErrorMessages)
													{
													sEmailBody += "\n          + " + errorEntry;
													Console.WriteLine("\t\t\t + {0}", errorEntry);
													}
												}
											else // there were no generation errors.
												{
												Console.WriteLine("\t *** no errors occurred during the generation process.");
												sEmailBody += ", which was generated without any errors.";
												}

											// begin to upload the document to SharePoint
											objCSDbasedCRM.DocumentStatus = enumDocumentStatusses.Uploading;
											Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

											// Upload the document to the Generated Documents Library
											bPublishDocSuccessful = objCSDbasedCRM.UploadDoc(
												parRequestingUserID: objDocCollection.RequestingUserID);
											// Check if the upload succeeded....
											if(bPublishDocSuccessful) //Upload Succeeded
												{
												Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
												// Insert the uploaded URL in the e-mail message body
												sEmailBody += "\n       The document is stored at this url: " + objCSDbasedCRM.URLonSharePoint;
												objCSDbasedCRM.DocumentStatus = enumDocumentStatusses.Uploaded;
												}
											else // Upload failed Failed
												{
												Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
												objDocCollection.UnexpectedErrors = true;
												objCSDbasedCRM.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
												sEmailBody += "\n       Unfortunately, a technical issue prevented the uploading of "
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
											sEmailBody += "\n\t - Unable to complete the generation of document: "
												+ objCSDbasedCRM.DocumentType
												+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
											}
										sEmailBody += "\n\n";
										break;
										}
								case ("CSD_Document_DRM_Inline"):
										{
										// Prepare to generate the Document
										bGenerateDocumentSuccessful = false;
										CSD_Document_DRM_Inline objCSDdrmInline = objDocumentWorkbook;

										if(objCSDdrmInline.ErrorMessages == null)
											objCSDdrmInline.ErrorMessages = new List<string>();

										bGenerateDocumentSuccessful = objCSDdrmInline.Generate(
											parDataSet: ref Globals.objDataSet,
											parSDDPdatacontext: objSDDPdatacontext);

										if(bGenerateDocumentSuccessful)
											{
											// set the Document status to Completed...
											objCSDdrmInline.DocumentStatus = enumDocumentStatusses.Completed;
											// Prepare the inclusion of the text in the e-mail that the user will receive.
											sEmailBody += "\n     * " + objDocumentWorkbook.DocumentType;
											// if there were errors, include them in the message.
											if(objCSDdrmInline.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
													objCSDdrmInline.ErrorMessages.Count);
												sEmailBody += ", which was generated but the following errors occurred:";
												foreach(string errorEntry in objCSDdrmInline.ErrorMessages)
													{
													sEmailBody += "\n          + " + errorEntry;
													Console.WriteLine("\t\t\t + {0}", errorEntry);
													}
												}
											else // there were no generation errors.
												{
												Console.WriteLine("\t *** no errors occurred during the generation process.");
												sEmailBody += ", which was generated without any errors.";
												}

											// begin to upload the document to SharePoint
											objCSDdrmInline.DocumentStatus = enumDocumentStatusses.Uploading;
											Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

											// Upload the document to the Generated Documents Library
											bPublishDocSuccessful = objCSDdrmInline.UploadDoc(
												parRequestingUserID: objDocCollection.RequestingUserID);
											// Check if the upload succeeded....
											if(bPublishDocSuccessful) //Upload Succeeded
												{
												Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
												// Insert the uploaded URL in the e-mail message body
												sEmailBody += "\n       The document is stored at this url: " + objCSDdrmInline.URLonSharePoint;
												objCSDdrmInline.DocumentStatus = enumDocumentStatusses.Uploaded;
												}
											else // Upload failed Failed
												{
												Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
												objDocCollection.UnexpectedErrors = true;
												objCSDdrmInline.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
												sEmailBody += "\n       Unfortunately, a technical issue prevented the uploading of "
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
											sEmailBody += "\n\t - Unable to complete the generation of document: "
												+ objCSDdrmInline.DocumentType
												+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
											}
										sEmailBody += "\n\n";
										break;
										}
									//----------------------------------------
									case ("CSD_Document_DRM_Sections"):
										{
										// Prepare to generate the Document
										bGenerateDocumentSuccessful = false;
										CSD_Document_DRM_Sections objCSDdrmSections = objDocumentWorkbook;

										if(objCSDdrmSections.ErrorMessages == null)
											objCSDdrmSections.ErrorMessages = new List<string>();

										bGenerateDocumentSuccessful = objCSDdrmSections.Generate(
											parDataSet: ref Globals.objDataSet,
											parSDDPdatacontext: objSDDPdatacontext);

										if(bGenerateDocumentSuccessful)
											{
											// set the Document status to Completed...
											objCSDdrmSections.DocumentStatus = enumDocumentStatusses.Completed;
											// Prepare the inclusion of the text in the e-mail that the user will receive.
											sEmailBody += "\n     * " + objDocumentWorkbook.DocumentType;
											// if there were errors, include them in the message.
											if(objCSDdrmSections.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
													objCSDdrmSections.ErrorMessages.Count);
												sEmailBody += ", which was generated but the following errors occurred:";
												foreach(string errorEntry in objCSDdrmSections.ErrorMessages)
													{
													sEmailBody += "\n          + " + errorEntry;
													Console.WriteLine("\t\t\t + {0}", errorEntry);
													}
												}
											else // there were no generation errors.
												{
												Console.WriteLine("\t *** no errors occurred during the generation process.");
												sEmailBody += ", which was generated without any errors.";
												}

											// begin to upload the document to SharePoint
											objCSDdrmSections.DocumentStatus = enumDocumentStatusses.Uploading;
											Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

											// Upload the document to the Generated Documents Library
											bPublishDocSuccessful = objCSDdrmSections.UploadDoc(
												parRequestingUserID: objDocCollection.RequestingUserID);
											// Check if the upload succeeded....
											if(bPublishDocSuccessful) //Upload Succeeded
												{
												Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
												// Insert the uploaded URL in the e-mail message body
												sEmailBody += "\n       The document is stored at this url: " + objCSDdrmSections.URLonSharePoint;
												objCSDdrmSections.DocumentStatus = enumDocumentStatusses.Uploaded;
												}
											else // Upload failed Failed
												{
												Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
												objDocCollection.UnexpectedErrors = true;
												objCSDdrmSections.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
												sEmailBody += "\n       Unfortunately, a technical issue prevented the uploading of "
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
											sEmailBody += "\n\t - Unable to complete the generation of document: "
												+ objCSDdrmSections.DocumentType
												+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
											}
										sEmailBody += "\n\n";
										break;
										}
									//-----------------------------------------------------
									case ("External_Technology_Coverage_Dashboard_Workbook"):
										{
										// Prepare to generate the Document
										bGenerateDocumentSuccessful = false;
										External_Technology_Coverage_Dashboard_Workbook objExtTechDashboard = objDocumentWorkbook;

										if(objExtTechDashboard.ErrorMessages == null)
											objExtTechDashboard.ErrorMessages = new List<string>();

										bGenerateDocumentSuccessful = objExtTechDashboard.Generate(
											parDataSet: ref Globals.objDataSet);

										if(bGenerateDocumentSuccessful)
											{
											// set the Document status to Completed...
											objExtTechDashboard.DocumentStatus = enumDocumentStatusses.Completed;
											// Prepare the inclusion of the text in the e-mail that the user will receive.
											sEmailBody += "\n     * " + objDocumentWorkbook.DocumentType;
											// if there were errors, include them in the message.
											if(objExtTechDashboard.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
													objExtTechDashboard.ErrorMessages.Count);
												sEmailBody += ", which was generated but the following errors occurred:";
												foreach(string errorEntry in objExtTechDashboard.ErrorMessages)
													{
													sEmailBody += "\n          + " + errorEntry;
													Console.WriteLine("\t\t\t + {0}", errorEntry);
													}
												}
											else // there were no generation errors.
												{
												Console.WriteLine("\t *** no errors occurred during the generation process.");
												sEmailBody += ", which was generated without any errors.";
												}

											// begin to upload the document to SharePoint
											objExtTechDashboard.DocumentStatus = enumDocumentStatusses.Uploading;
											Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

											// Upload the document to the Generated Documents Library
											bPublishDocSuccessful = objExtTechDashboard.UploadDoc(
												parRequestingUserID: objDocCollection.RequestingUserID);
											// Check if the upload succeeded....
											if(bPublishDocSuccessful) //Upload Succeeded
												{
												Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
												// Insert the uploaded URL in the e-mail message body
												sEmailBody += "\n       The document is stored at this url: " 
													+ objExtTechDashboard.URLonSharePoint;
												objExtTechDashboard.DocumentStatus = enumDocumentStatusses.Uploaded;
												}
											else // Upload failed Failed
												{
												Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
												objDocCollection.UnexpectedErrors = true;
												objExtTechDashboard.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
												sEmailBody += "\n       Unfortunately, a technical issue prevented the uploading of "
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
											sEmailBody += "\n\t - Unable to complete the generation of document: "
												+ objExtTechDashboard.DocumentType
												+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
											}
										sEmailBody += "\n\n";
										break;
										}
									//---------------------------------------------------------
									case ("Internal_Technology_Coverage_Dashboard_Workbook"):
										{
										// Prepare to generate the Document
										bGenerateDocumentSuccessful = false;
										Internal_Technology_Coverage_Dashboard_Workbook objIntTechDashboard = objDocumentWorkbook;

										if(objIntTechDashboard.ErrorMessages == null)
											objIntTechDashboard.ErrorMessages = new List<string>();

										bGenerateDocumentSuccessful = objIntTechDashboard.Generate(
											parDataSet: ref Globals.objDataSet);

										if(bGenerateDocumentSuccessful)
											{
											// set the Document status to Completed...
											objIntTechDashboard.DocumentStatus = enumDocumentStatusses.Completed;
											// Prepare the inclusion of the text in the e-mail that the user will receive.
											sEmailBody += "\n     * " + objDocumentWorkbook.DocumentType;
											// if there were errors, include them in the message.
											if(objIntTechDashboard.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
													objIntTechDashboard.ErrorMessages.Count);
												sEmailBody += ", which was generated but the following errors occurred:";
												foreach(string errorEntry in objIntTechDashboard.ErrorMessages)
													{
													sEmailBody += "\n          + " + errorEntry;
													Console.WriteLine("\t\t\t + {0}", errorEntry);
													}
												}
											else // there were no generation errors.
												{
												Console.WriteLine("\t *** no errors occurred during the generation process.");
												sEmailBody += ", which was generated without any errors.";
												}

											// begin to upload the document to SharePoint
											objIntTechDashboard.DocumentStatus = enumDocumentStatusses.Uploading;
											Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

											// Upload the document to the Generated Documents Library
											bPublishDocSuccessful = objIntTechDashboard.UploadDoc(
												parRequestingUserID: objDocCollection.RequestingUserID);
											// Check if the upload succeeded....
											if(bPublishDocSuccessful) //Upload Succeeded
												{
												Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
												// Insert the uploaded URL in the e-mail message body
												sEmailBody += "\n       The document is stored at this url: " 
													+ objIntTechDashboard.URLonSharePoint;
												objIntTechDashboard.DocumentStatus = enumDocumentStatusses.Uploaded;
												}
											else // Upload failed Failed
												{
												Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
												objDocCollection.UnexpectedErrors = true;
												objIntTechDashboard.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
												sEmailBody += "\n       Unfortunately, a technical issue prevented the uploading of "
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
											sEmailBody += "\n\t - Unable to complete the generation of document: "
												+ objIntTechDashboard.DocumentType
												+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
											}
										sEmailBody += "\n\n";
										break;
										}
									//-------------------------------------
									case ("ISD_Document_DRM_Inline"):
										{
										// Prepare to generate the Document
										bGenerateDocumentSuccessful = false;
										ISD_Document_DRM_Inline objISDdrmInline = objDocumentWorkbook;

										if(objISDdrmInline.ErrorMessages == null)
											objISDdrmInline.ErrorMessages = new List<string>();

										bGenerateDocumentSuccessful = objISDdrmInline.Generate(
											parDataSet: ref Globals.objDataSet,
											parSDDPdatacontext: objSDDPdatacontext);

										if(bGenerateDocumentSuccessful)
											{
											// set the Document status to Completed...
											objISDdrmInline.DocumentStatus = enumDocumentStatusses.Completed;
											// Prepare the inclusion of the text in the e-mail that the user will receive.
											sEmailBody += "\n     * " + objDocumentWorkbook.DocumentType;
											// if there were errors, include them in the message.
											if(objISDdrmInline.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
													objISDdrmInline.ErrorMessages.Count);
												sEmailBody += ", which was generated but the following errors occurred:";
												foreach(string errorEntry in objISDdrmInline.ErrorMessages)
													{
													sEmailBody += "\n          + " + errorEntry;
													Console.WriteLine("\t\t\t + {0}", errorEntry);
													}
												}
											else // there were no generation errors.
												{
												Console.WriteLine("\t *** no errors occurred during the generation process.");
												sEmailBody += ", which was generated without any errors.";
												}

											// begin to upload the document to SharePoint
											objISDdrmInline.DocumentStatus = enumDocumentStatusses.Uploading;
											Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

											// Upload the document to the Generated Documents Library
											bPublishDocSuccessful = objISDdrmInline.UploadDoc(
												parRequestingUserID: objDocCollection.RequestingUserID);
											// Check if the upload succeeded....
											if(bPublishDocSuccessful) //Upload Succeeded
												{
												Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
												// Insert the uploaded URL in the e-mail message body
												sEmailBody += "\n       The document is stored at this url: " + objISDdrmInline.URLonSharePoint;
												objISDdrmInline.DocumentStatus = enumDocumentStatusses.Uploaded;
												}
											else // Upload failed Failed
												{
												Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
												objDocCollection.UnexpectedErrors = true;
												objISDdrmInline.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
												sEmailBody += "\n       Unfortunately, a technical issue prevented the uploading of "
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
											sEmailBody += "\n\t - Unable to complete the generation of document: "
												+ objISDdrmInline.DocumentType
												+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
											}
										sEmailBody += "\n\n";
										break;
										}
									//-------------------------------------
									case ("ISD_Document_DRM_Sections"):
										{
										// Prepare to generate the Document
										bGenerateDocumentSuccessful = false;
										ISD_Document_DRM_Sections objISDdrmSections = objDocumentWorkbook;

										if(objISDdrmSections.ErrorMessages == null)
											objISDdrmSections.ErrorMessages = new List<string>();

										bGenerateDocumentSuccessful = objISDdrmSections.Generate(
											parDataSet: ref Globals.objDataSet,
											parSDDPdatacontext: objSDDPdatacontext);

										if(bGenerateDocumentSuccessful)
											{
											// set the Document status to Completed...
											objISDdrmSections.DocumentStatus = enumDocumentStatusses.Completed;
											// Prepare the inclusion of the text in the e-mail that the user will receive.
											sEmailBody += "\n     * " + objDocumentWorkbook.DocumentType;
											// if there were errors, include them in the message.
											if(objISDdrmSections.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
													objISDdrmSections.ErrorMessages.Count);
												sEmailBody += ", which was generated but the following errors occurred:";
												foreach(string errorEntry in objISDdrmSections.ErrorMessages)
													{
													sEmailBody += "\n          + " + errorEntry;
													Console.WriteLine("\t\t\t + {0}", errorEntry);
													}
												}
											else // there were no generation errors.
												{
												Console.WriteLine("\t *** no errors occurred during the generation process.");
												sEmailBody += ", which was generated without any errors.";
												}

											// begin to upload the document to SharePoint
											objISDdrmSections.DocumentStatus = enumDocumentStatusses.Uploading;
											Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

											// Upload the document to the Generated Documents Library
											bPublishDocSuccessful = objISDdrmSections.UploadDoc(
												parRequestingUserID: objDocCollection.RequestingUserID);
											// Check if the upload succeeded....
											if(bPublishDocSuccessful) //Upload Succeeded
												{
												Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
												// Insert the uploaded URL in the e-mail message body
												sEmailBody += "\n       The document is stored at this url: " + objISDdrmSections.URLonSharePoint;
												objISDdrmSections.DocumentStatus = enumDocumentStatusses.Uploaded;
												}
											else // Upload failed Failed
												{
												Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
												objDocCollection.UnexpectedErrors = true;
												objISDdrmSections.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
												sEmailBody += "\n       Unfortunately, a technical issue prevented the uploading of "
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
											sEmailBody += "\n\t - Unable to complete the generation of document: "
												+ objISDdrmSections.DocumentType
												+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
											}
										sEmailBody += "\n\n";
										break;
										}
									//------------------------------------------
									case ("Pricing_Addendum_Document"):
										{
										// Prepare to generate the Document
										bGenerateDocumentSuccessful = false;
										Pricing_Addendum_Document objPricingAddendum = objDocumentWorkbook;

										if(objPricingAddendum.ErrorMessages == null)
											objPricingAddendum.ErrorMessages = new List<string>();

										//bGenerateDocumentSuccessful = objPricingAddendum.Generate(
										//	parDataSet: ref Globals.objDataSet);

										if(bGenerateDocumentSuccessful)
											{
											// set the Document status to Completed...
											objPricingAddendum.DocumentStatus = enumDocumentStatusses.Completed;
											// Prepare the inclusion of the text in the e-mail that the user will receive.
											sEmailBody += "\n     * " + objDocumentWorkbook.DocumentType;
											// if there were errors, include them in the message.
											if(objPricingAddendum.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
													objPricingAddendum.ErrorMessages.Count);
												sEmailBody += ", which was generated but the following errors occurred:";
												foreach(string errorEntry in objPricingAddendum.ErrorMessages)
													{
													sEmailBody += "\n          + " + errorEntry;
													Console.WriteLine("\t\t\t + {0}", errorEntry);
													}
												}
											else // there were no generation errors.
												{
												Console.WriteLine("\t *** no errors occurred during the generation process.");
												sEmailBody += ", which was generated without any errors.";
												}

											// begin to upload the document to SharePoint
											objPricingAddendum.DocumentStatus = enumDocumentStatusses.Uploading;
											Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

											// Upload the document to the Generated Documents Library
											bPublishDocSuccessful = objPricingAddendum.UploadDoc(
												parRequestingUserID: objDocCollection.RequestingUserID);
											// Check if the upload succeeded....
											if(bPublishDocSuccessful) //Upload Succeeded
												{
												Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
												// Insert the uploaded URL in the e-mail message body
												sEmailBody += "\n       The document is stored at this url: " + objPricingAddendum.URLonSharePoint;
												objPricingAddendum.DocumentStatus = enumDocumentStatusses.Uploaded;
												}
											else // Upload failed Failed
												{
												Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
												objDocCollection.UnexpectedErrors = true;
												objPricingAddendum.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
												sEmailBody += "\n       Unfortunately, a technical issue prevented the uploading of "
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
											sEmailBody += "\n\t - Unfortunately, the "+ objPricingAddendum.DocumentType
												+ " is not implemented at the moment because the Pricing Methodology is still Work in Progress.";
											}
										sEmailBody += "\n\n";
										break;
										}
									//--------------------------------------
									case ("RACI_Matrix_Workbook_per_Deliverable"):
										{
										// Prepare to generate the Document
										bGenerateDocumentSuccessful = false;
										RACI_Matrix_Workbook_per_Deliverable objRACImatrix = objDocumentWorkbook;

										if(objRACImatrix.ErrorMessages == null)
											objRACImatrix.ErrorMessages = new List<string>();

										bGenerateDocumentSuccessful = objRACImatrix.Generate(
											parDataSet: ref Globals.objDataSet);

										if(bGenerateDocumentSuccessful)
											{
											// set the Document status to Completed...
											objRACImatrix.DocumentStatus = enumDocumentStatusses.Completed;
											// Prepare the inclusion of the text in the e-mail that the user will receive.
											sEmailBody += "\n     * " + objDocumentWorkbook.DocumentType;
											// if there were errors, include them in the message.
											if(objRACImatrix.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
													objRACImatrix.ErrorMessages.Count);
												sEmailBody += ", which was generated but the following errors occurred:";
												foreach(string errorEntry in objRACImatrix.ErrorMessages)
													{
													sEmailBody += "\n          + " + errorEntry;
													Console.WriteLine("\t\t\t + {0}", errorEntry);
													}
												}
											else // there were no generation errors.
												{
												Console.WriteLine("\t *** no errors occurred during the generation process.");
												sEmailBody += ", which was generated without any errors.";
												}

											// begin to upload the document to SharePoint
											objRACImatrix.DocumentStatus = enumDocumentStatusses.Uploading;
											Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

											// Upload the document to the Generated Documents Library
											bPublishDocSuccessful = objRACImatrix.UploadDoc(
												parRequestingUserID: objDocCollection.RequestingUserID);
											// Check if the upload succeeded....
											if(bPublishDocSuccessful) //Upload Succeeded
												{
												Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
												// Insert the uploaded URL in the e-mail message body
												sEmailBody += "\n       The document is stored at this url: " + objRACImatrix.URLonSharePoint;
												objRACImatrix.DocumentStatus = enumDocumentStatusses.Uploaded;
												}
											else // Upload failed Failed
												{
												Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
												objDocCollection.UnexpectedErrors = true;
												objRACImatrix.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
												sEmailBody += "\n       Unfortunately, a technical issue prevented the uploading of "
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
											sEmailBody += "\n\t - Unable to complete the generation of document: "
												+ objRACImatrix.DocumentType
												+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
											}
										sEmailBody += "\n\n";
										break;
										}
									//-----------------------------------
									case ("RACI_Workbook_per_Role"):
										{
										// Prepare to generate the Document
										bGenerateDocumentSuccessful = false;
										RACI_Workbook_per_Role objRACIperRole = objDocumentWorkbook;

										if(objRACIperRole.ErrorMessages == null)
											objRACIperRole.ErrorMessages = new List<string>();

										bGenerateDocumentSuccessful = objRACIperRole.Generate(
											parDataSet: ref Globals.objDataSet);

										if(bGenerateDocumentSuccessful)
											{
											// set the Document status to Completed...
											objRACIperRole.DocumentStatus = enumDocumentStatusses.Completed;
											// Prepare the inclusion of the text in the e-mail that the user will receive.
											sEmailBody += "\n     * " + objDocumentWorkbook.DocumentType;
											// if there were errors, include them in the message.
											if(objRACIperRole.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
													objRACIperRole.ErrorMessages.Count);
												sEmailBody += ", which was generated but the following errors occurred:";
												foreach(string errorEntry in objRACIperRole.ErrorMessages)
													{
													sEmailBody += "\n          + " + errorEntry;
													Console.WriteLine("\t\t\t + {0}", errorEntry);
													}
												}
											else // there were no generation errors.
												{
												Console.WriteLine("\t *** no errors occurred during the generation process.");
												sEmailBody += ", which was generated without any errors.";
												}

											// begin to upload the document to SharePoint
											objRACIperRole.DocumentStatus = enumDocumentStatusses.Uploading;
											Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

											// Upload the document to the Generated Documents Library
											bPublishDocSuccessful = objRACIperRole.UploadDoc(
												parRequestingUserID: objDocCollection.RequestingUserID);
											// Check if the upload succeeded....
											if(bPublishDocSuccessful) //Upload Succeeded
												{
												Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
												// Insert the uploaded URL in the e-mail message body
												sEmailBody += "\n       The document is stored at this url: " + objRACIperRole.URLonSharePoint;
												objRACIperRole.DocumentStatus = enumDocumentStatusses.Uploaded;
												}
											else // Upload failed Failed
												{
												Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
												objDocCollection.UnexpectedErrors = true;
												objRACIperRole.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
												sEmailBody += "\n       Unfortunately, a technical issue prevented the uploading of "
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
											sEmailBody += "\n\t - Unable to complete the generation of document: "
												+ objRACIperRole.DocumentType
												+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
											}
										sEmailBody += "\n\n";
										break;
										}
									//----------------------------------------
									case ("Services_Framework_Document_DRM_Inline"):
										{
										// Prepare to generate the Document
										bGenerateDocumentSuccessful = false;
										Services_Framework_Document_DRM_Inline objSFdrmInline = objDocumentWorkbook;

										if(objSFdrmInline.ErrorMessages == null)
											objSFdrmInline.ErrorMessages = new List<string>();

										bGenerateDocumentSuccessful = objSFdrmInline.Generate(
											parDataSet: ref Globals.objDataSet,
											parSDDPdatacontext: objSDDPdatacontext);

										if(bGenerateDocumentSuccessful)
											{
											// set the Document status to Completed...
											objSFdrmInline.DocumentStatus = enumDocumentStatusses.Completed;
											// Prepare the inclusion of the text in the e-mail that the user will receive.
											sEmailBody += "\n     * " + objDocumentWorkbook.DocumentType;
											// if there were errors, include them in the message.
											if(objSFdrmInline.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
													objSFdrmInline.ErrorMessages.Count);
												sEmailBody += ", which was generated but the following errors occurred:";
												foreach(string errorEntry in objSFdrmInline.ErrorMessages)
													{
													sEmailBody += "\n          + " + errorEntry;
													Console.WriteLine("\t\t\t + {0}", errorEntry);
													}
												}
											else // there were no generation errors.
												{
												Console.WriteLine("\t *** no errors occurred during the generation process.");
												sEmailBody += ", which was generated without any errors.";
												}

											// begin to upload the document to SharePoint
											objSFdrmInline.DocumentStatus = enumDocumentStatusses.Uploading;
											Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

											// Upload the document to the Generated Documents Library
											bPublishDocSuccessful = objSFdrmInline.UploadDoc(
												parRequestingUserID: objDocCollection.RequestingUserID);
											// Check if the upload succeeded....
											if(bPublishDocSuccessful) //Upload Succeeded
												{
												Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
												// Insert the uploaded URL in the e-mail message body
												sEmailBody += "\n       The document is stored at this url: " + objSFdrmInline.URLonSharePoint;
												objSFdrmInline.DocumentStatus = enumDocumentStatusses.Uploaded;
												}
											else // Upload failed Failed
												{
												Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
												objDocCollection.UnexpectedErrors = true;
												objSFdrmInline.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
												sEmailBody += "\n       Unfortunately, a technical issue prevented the uploading of "
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
											sEmailBody += "\n\t - Unable to complete the generation of document: "
												+ objSFdrmInline.DocumentType
												+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
											}
											sEmailBody += "\n\n";
										break;
										}
									//---------------------------------------------
									case ("Services_Framework_Document_DRM_Sections"):
										{
										// Prepare to generate the Document
										bGenerateDocumentSuccessful = false;
										Services_Framework_Document_DRM_Sections objSFdrmSections = objDocumentWorkbook;

										if(objSFdrmSections.ErrorMessages == null)
											objSFdrmSections.ErrorMessages = new List<string>();

										bGenerateDocumentSuccessful = objSFdrmSections.Generate(
											parDataSet: ref Globals.objDataSet,
											parSDDPdatacontext: objSDDPdatacontext);

										if(bGenerateDocumentSuccessful)
											{
											// set the Document status to Completed...
											objSFdrmSections.DocumentStatus = enumDocumentStatusses.Completed;
											// Prepare the inclusion of the text in the e-mail that the user will receive.
											sEmailBody += "\n     * " + objDocumentWorkbook.DocumentType;
											// if there were errors, include them in the message.
											if(objSFdrmSections.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.",
													objSFdrmSections.ErrorMessages.Count);
												sEmailBody += ", which was generated but the following errors occurred:";
												foreach(string errorEntry in objSFdrmSections.ErrorMessages)
													{
													sEmailBody += "\n          + " + errorEntry;
													Console.WriteLine("\t\t\t + {0}", errorEntry);
													}
												}
											else // there were no generation errors.
												{
												Console.WriteLine("\t *** no errors occurred during the generation process.");
												sEmailBody += ", which was generated without any errors.";
												}

											// begin to upload the document to SharePoint
											objSFdrmSections.DocumentStatus = enumDocumentStatusses.Uploading;
											Console.WriteLine("\t Uploading Document to SharePoint's Generated Documents Library");

											// Upload the document to the Generated Documents Library
											bPublishDocSuccessful = objSFdrmSections.UploadDoc(
												parRequestingUserID: objDocCollection.RequestingUserID);
											// Check if the upload succeeded....
											if(bPublishDocSuccessful) //Upload Succeeded
												{
												Console.WriteLine("+ {0}, was Successfully Uploaded.", objDocumentWorkbook.DocumentType);
												// Insert the uploaded URL in the e-mail message body
												sEmailBody += "\n       The document is stored at this url: " + objSFdrmSections.URLonSharePoint;
												objSFdrmSections.DocumentStatus = enumDocumentStatusses.Uploaded;
												}
											else // Upload failed Failed
												{
												Console.WriteLine("*** Uploading of {0} FAILED.", objDocumentWorkbook.DocumentType);
												objDocCollection.UnexpectedErrors = true;
												objSFdrmSections.ErrorMessages.Add("Error: Unable to upload the document to SharePoint");
												sEmailBody += "\n       Unfortunately, a technical issue prevented the uploading of "
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
											sEmailBody += "\n\t - Unable to complete the generation of document: "
												+ objSFdrmSections.DocumentType
												+ "\n (This message was also send to the SDDP Technical Team for further investigation.)";
											}
										sEmailBody += "\n\n";
										break;
										}
									} // switch (objectType)
								} // foreach(dynamic objDocumentWorkbook in objDocCollection.Documen_and_Workbook_Objects...

							//--------------------------------------------------------------------------
							// Process the Notification via E-mail if the users selected to be notified.
							if(objDocCollection.NotifyMe && objDocCollection.NotificationEmail != null)
								{
								bSendEmailSuccessful = eMail.SendEmail(
								parRecipient: objDocCollection.NotificationEmail,
								parSubject: "SDDP: Generated Document(s)",
								parBody: sEmailBody);

								if(bSendEmailSuccessful)
									Console.WriteLine("Sending e-mail successfully send to user!");
								else
									Console.WriteLine("*** ERROR *** \n Sending e-mail failed...\n");

								}
							//----------------------------------------------------------------------------
							// Check if there were unexpected errors and if there were, send an e-mail to the Technical Support team.
							if(objDocCollection.UnexpectedErrors)
								{
								bUpdateDocColSuccessful = objDocCollection.UpdateGenerateStatus(
									parGenerationStatus: enumGenerationStatus.Failed);

								if(bUpdateDocColSuccessful)
									Console.WriteLine("Update Document Collection Status to 'FAILED' was successful.");
								else
									Console.WriteLine("Update Document Collection Status to 'FAILED' was unsuccessful.");

								// Prepare the e-mail
								bSendEmailSuccessful = eMail.SendEmail(
									parRecipient: Properties.AppResources.Email_Bcc_Address,
									parSubject: "SDDP: Unexpected DocGenerator(s) Error occurred.)",
									parBody: sEmailBody,
									parSendBcc: false);

								if(bSendEmailSuccessful)
									Console.WriteLine("The error e-mail was successfully send to the technical team.");
								else
									Console.WriteLine("The error e-mail to the technical team FAILED!");

								}
							else // there was no UNEXPECTED errors
								{
								bUpdateDocColSuccessful = objDocCollection.UpdateGenerateStatus(
									parGenerationStatus: enumGenerationStatus.Completed);

								if(bUpdateDocColSuccessful)
									Console.WriteLine("Update Document Collection Status to 'Completed' was successful.");
								else
									Console.WriteLine("Update Document Collection Status to 'Completed' was unsuccessful.");
								}
							} // end if ...Count() > 0

						} // foreach(DocumentCollection objDocCollection in docCollectionsToGenerate)
					Console.WriteLine("\nDocuments for {0} Document Collection(s) were Generated.", docCollectionsToGenerate.Count);
					} //if(docCollectionsToGenerate.Count > 0)
				else
					{
					Console.WriteLine("Sorry, nothing to generate at this stage.");
					}

				objSDDPdatacontext = null;
				} // end try
			catch(DataServiceTransportException exc)
				{
				if(exc.Message.Contains("timed out"))
					{
					Console.WriteLine("The data connection to SharePoint timed out - and the documents could not be generated..." +
						"The {0} will retry to generate the document(2)", "DocGenerator");
					}
				else
					{
					Console.WriteLine("\n\nException Error: {0} occurred and means {1}", exc.HResult, exc.Message);
					}
				Console.WriteLine("\n\nException Error: {0} occurred and means {1}", exc.HResult, exc.Message);
				}
			catch(OpenXmlPackageException exc)
				{
				Console.WriteLine("\n\nException Error: {0} occurred and means {1}", exc.HResult, exc.Message);
				}
			catch(Exception exc) // if the List is empty - nothing to generate
				{
				Console.WriteLine("\n\nException Error: {0} occurred and means {1}", exc.HResult, exc.Message);
				}
Procedure_Ends:
			Console.WriteLine("Everything done, waiting for next cycle...");
			Cursor.Current = Cursors.Default;
		}

	private void btnOpenMSwordDocument(object sender, EventArgs e)
		{
		string parTemplateURL = "https://teams.dimensiondata.com/sites/ServiceCatalogue/DocumentTemplates/ServicesFrameworkDocumentTemplate.dotx";
		enumDocumentTypes parDocumentType = enumDocumentTypes.Service_Framework_Document_DRM_sections;
		int tableCaptionCounter = 0;
		int imageCaptionCounter = 0;
		int iPictureNo = 49;
		int hyperlinkCounter = 1;
			
		// define a new objOpenXMLdocument
		oxmlDocument objOXMLdocument = new oxmlDocument();
		// use CreateDocumentFromTemplate method to create a new MS Word Document based on the relevant template
		if(objOXMLdocument.CreateDocWbkFromTemplate(
			parDocumentOrWorkbook: enumDocumentOrWorkbook.Document,
			parTemplateURL: parTemplateURL, 
			parDocumentType: parDocumentType))
			{
			Console.WriteLine("objOXMLdocument:\n" +
			"                + LocalDocumentPath: {0}\n" +
			"                + DocumentFileName.: {1}\n" +
			"                + DocumentURI......: {2}", objOXMLdocument.LocalPath, objOXMLdocument.Filename, objOXMLdocument.LocalURI);
			}
		else
			{
			// if the creation failed.
			Console.WriteLine("An ERROR occurred and the new MS Word Document could not be created due to above stated ERROR conditions.");
			return;
			}
		try
			{
			// Open the MS Word document in Edit mode
			WordprocessingDocument objWPdocument = WordprocessingDocument.Open(path: objOXMLdocument.LocalURI, isEditable: true);

			// Define all open XML objects to use for building the document
			MainDocumentPart objMainDocumentPart = objWPdocument.MainDocumentPart;
			Body objBody = objWPdocument.MainDocumentPart.Document.Body;
			Paragraph objParagraph = new Paragraph();    // Define the objParagraph	
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();

			// Now begin to write the content to the document

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

			// Just some text to validate all routines
			objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
			objRun = oxmlDocument.Construct_RunText(
				parText2Write: Properties.AppResources.Document_IntruductorySection_HeadingText,
				parIsNewSection: true);
			objParagraph.Append(objRun);
			objBody.Append(objParagraph);
			objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
			objRun = oxmlDocument.Construct_RunText(Properties.AppResources.Document_Introduction_HeadingText);
			objParagraph.Append(objRun);
			objBody.Append(objParagraph);
			objParagraph = oxmlDocument.Construct_Paragraph(1);
			objRun = oxmlDocument.Construct_RunText("This is a run of Text with ");
			objParagraph.Append(objRun);
			objRun = oxmlDocument.Construct_RunText(" Bold, ", parBold: true);
			objParagraph.Append(objRun);
			objRun = oxmlDocument.Construct_RunText("Bold Underline, ", parBold: true, parUnderline: true);
			objParagraph.Append(objRun);
			objRun = oxmlDocument.Construct_RunText(" Bold Italic, ", parBold: true, parItalic: true);
			objParagraph.Append(objRun);
			objRun = oxmlDocument.Construct_RunText(" Italic, ", parItalic: true);
			objParagraph.Append(objRun);
			objRun = oxmlDocument.Construct_RunText("Underline,", parUnderline: true);
			objParagraph.Append(objRun);
			objRun = oxmlDocument.Construct_RunText(" and ");
			objParagraph.Append(objRun);
			objRun = oxmlDocument.Construct_RunText("Italic Underline", parItalic: true, parUnderline: true);
			objParagraph.Append(objRun);
			objRun = oxmlDocument.Construct_RunText(" properties.");
			objParagraph.Append(objRun);
			objBody.Append(objParagraph);

			objParagraph = oxmlDocument.Construct_Paragraph(1);
			objRun = oxmlDocument.Construct_RunText("Another paragraph with just normal text.");
			objParagraph.Append(objRun);
			objBody.Append(objParagraph);

			objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
			objRun = oxmlDocument.Construct_RunText(parText2Write: "A send level heading (error)", parIsError: true);
			objParagraph.Append(objRun);
			objBody.Append(objParagraph);
			objParagraph = oxmlDocument.Construct_Paragraph(2);
			objRun = oxmlDocument.Construct_RunText("Below is an image of my favourite car. ");
			objParagraph.Append(objRun);
			objBody.Append(objParagraph);

			// Determine the Page Size for the current Body object.
			SectionProperties objSectionProperties = new SectionProperties();
			UInt32 pageWidth = Convert.ToUInt32(Properties.AppResources.DefaultPageWidth);
			UInt32 pageHeight = Convert.ToUInt32(Properties.AppResources.DefaultPageHeight);
			
			if(objBody.GetFirstChild<SectionProperties>() != null)
				{
				objSectionProperties = objBody.GetFirstChild<SectionProperties>();
				PageSize objPageSize = objSectionProperties.GetFirstChild<PageSize>();
				PageMargin objPageMargin = objSectionProperties.GetFirstChild<PageMargin>();
				if(objPageSize != null)
					{
					pageWidth = objPageSize.Width;
					Console.WriteLine("Page Width.: {0}", objPageSize.Width);
					pageHeight = objPageSize.Height;
					Console.WriteLine("Page Height: {0}", objPageSize.Height);
					}
				if(objPageMargin != null)
					{
					if(objPageMargin.Left != null)
						{
						pageWidth -= objPageMargin.Left;
						Console.WriteLine("Left Margin: {0}", objPageMargin.Right);
						}
					if(objPageMargin.Right != null)
						{
						pageWidth -= objPageMargin.Right;
						Console.WriteLine("Right Margin: {0}", objPageMargin.Right);
						}
					if(objPageMargin.Top != null)
						{
						string tempTop = objPageMargin.Top.ToString();
						Console.WriteLine("Top Margin: {0}", tempTop);
						pageHeight -= Convert.ToUInt32(tempTop);
						}
					if(objPageMargin.Bottom != null)
						{
						string tempBottom = objPageMargin.Bottom.ToString();
						Console.WriteLine("Bottom Margin: {0}", tempBottom);
						pageHeight -= Convert.ToUInt32(tempBottom);
						}
					}
	               }
			Console.WriteLine("Effective pageWidth.: {0}twips", pageWidth);
			Console.WriteLine("Effective pageHeight: {0}twips", pageHeight);

			// Insert and image in the document
			objParagraph = oxmlDocument.Construct_Paragraph(2);
			objRun = oxmlDocument.InsertImage(
				parMainDocumentPart: ref objMainDocumentPart,
				parEffectivePageTWIPSheight: pageHeight,
				parEffectivePageTWIPSwidth: pageWidth,
				parParagraphLevel: 2,
				parPictureSeqNo: 1,
				parImageURL: @Properties.AppResources.TestData_Location + "RS5.jpg");
			if(objRun != null)
				{
				objParagraph.Append(objRun);
				objBody.AppendChild<Paragraph>(objParagraph);
				}
			else
				{
				objRun = oxmlDocument.Construct_RunText("ERROR: Unable to insert the image - an error occurred");
				objBody.Append(objParagraph);
				}
			// Insert the Image Caption
				// First increment the Image Caption Counter with 1
			imageCaptionCounter += 1;
			objParagraph = oxmlDocument.Construct_Caption(
				parCaptionType: "Image", 
				parCaptionText: Properties.AppResources.Document_Caption_Image_Text + imageCaptionCounter + ": " + "An awesome machine.");
			objBody.Append(objParagraph);

			// Insert a Heading for the Table section.
			objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
			objRun = oxmlDocument.Construct_RunText(
				parText2Write: "Tables",
				parIsNewSection: true);
			objParagraph.Append(objRun);
			objBody.Append(objParagraph);
			// Insert a paragraph of text
			objParagraph = oxmlDocument.Construct_Paragraph(2);
			objRun = oxmlDocument.Construct_RunText("This demonstrates how tables are handled by the DocGenerator application.", parBold: true);
			objParagraph.Append(objRun);
			objBody.Append(objParagraph);

			//Table Construction code
			// Construct a Table object instance
			DocumentFormat.OpenXml.Wordprocessing.Table objTable = new DocumentFormat.OpenXml.Wordprocessing.Table();
			objTable = oxmlDocument.ConstructTable(
				parPageWidth: pageWidth,
				parFirstRow: true, 
				parFirstColumn: true, 
				parLastColumn: true, 
				parLastRow: true, 
				parNoVerticalBand: true, 
				parNoHorizontalBand: false);
			// Create the Table Row and append it to the Table object
			TableRow objTableRow = new TableRow();
			TableCell objTableCell = new TableCell();
			bool IsFirstRow = false;
			bool IsLastRow = false;
			bool IsFirstColumn = false;
			bool IsLastColumn = false;
			int numberOfRows = 6;
			int numberOfColumns = 4;
			string tableText = "";
			UInt32 columnWidth = pageWidth / Convert.ToUInt32(numberOfColumns);
			// Construct a TableGrid object instance
			TableGrid objTableGrid = new TableGrid();
			List<UInt32> lstTableColumns = new List<UInt32>();
			for(int i = 0; i < numberOfColumns; i++)
				{
				lstTableColumns.Add(columnWidth);
				}
			objTableGrid = oxmlDocument.ConstructTableGrid(lstTableColumns);
			// Append the TableGrid object instance to the Table object instance
			objTable.Append(objTableGrid);
				
			// Create a TableRow object instance
			for(int r = 1; r < numberOfRows+1; r++)
				{
				// Construct a TableRow
				if(r == 1) // the Hear row
					IsFirstRow = true;
				else
					IsFirstRow = false;

				if(r == numberOfRows)
					IsLastRow = true;
				else
					IsLastRow = false;

				objTableRow = oxmlDocument.ConstructTableRow(parIsFirstRow: IsFirstRow, parIsLastRow: IsLastRow);
				// Create the TableCells for each Column
				for(int c = 1; c < numberOfColumns+1; c++)
					{
					objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1, parIsTableParagraph: true);
					if(c == 1)
						IsFirstColumn = true;
					else
						IsFirstColumn = false;

					if(c == numberOfColumns)
						IsLastColumn = true;
					else
						IsLastColumn = false;

					objTableCell = oxmlDocument.ConstructTableCell(
						lstTableColumns[c-1],
						parIsFirstRow: IsFirstRow,
						parIsLastRow: IsLastRow,
						parIsFirstColumn: IsFirstColumn,
						parIsLastColumn: IsLastColumn);

					// Create a Pargaraph for the text to go into the TableCell
					objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
					tableText = "Row " + r + ", Column " + c + " Text";
					objRun = oxmlDocument.Construct_RunText(tableText);
					objParagraph.Append(objRun);
					objTableCell.Append(objParagraph);
					objTableRow.Append(objTableCell);
					} //end For numberOfColumns loop
				objTable.Append(objTableRow);
				} // end For numberOfRows loop
			// Insert the Table object into the document Body
			objBody.Append(objTable);

			// Insert the Table Caption
			// increment the table Caption Counter with 1
			tableCaptionCounter += 1;
			objParagraph = oxmlDocument.Construct_Caption(
				parCaptionType: "Table",
				parCaptionText: Properties.AppResources.Document_Caption_Table_Text + tableCaptionCounter + ": " + "A table generated by the app.");
			objBody.Append(objParagraph);
				
			// Insert a new XML Table based on an HTML table input from a local file.
			objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 3);
			objRun = oxmlDocument.Construct_RunText(
				parText2Write: "How HTML content is handled",
				parIsNewSection: true);
			objParagraph.Append(objRun);
			objBody.Append(objParagraph);

			HTMLdecoder objHTMLdecoder = new HTMLdecoder();
			objHTMLdecoder.WPbody = objBody;
			string sCurrentDirectory = Directory.GetCurrentDirectory();
			Console.WriteLine("Current Directory is {0}", sCurrentDirectory);
			string sFile = @Properties.AppResources.TestData_Location + "IntoFromSharePoint.txt";
			string sContent = System.IO.File.ReadAllText(sFile);
			objHTMLdecoder.DecodeHTML(
				parMainDocumentPart: ref objMainDocumentPart,
				parDocumentLevel: 3,
				parPageWidthTwips: pageWidth,
				parPageHeightTwips: pageHeight,
				parHTML2Decode: sContent,
				parTableCaptionCounter: ref tableCaptionCounter,
				parImageCaptionCounter: ref imageCaptionCounter,
				parPictureNo: ref iPictureNo,
				parHyperlinkID: ref hyperlinkCounter);
			hyperlinkCounter += 1;

			// Close the document
			objParagraph = oxmlDocument.Construct_Paragraph(1);
			objRun = oxmlDocument.Construct_RunText("--- end of the document --- ");
			objParagraph.Append(objRun);
			objBody.Append(objParagraph);

			Console.WriteLine("\t\t Document generated, now saving and closing the document.");

			// Save and close the Document
			objWPdocument.Close();

			Console.WriteLine("Document saved and closed!!!");
			} // end Try

		catch(OpenXmlPackageException exc)
			{
			//TODO: add code to catch exception.
			}
		catch(ArgumentNullException exc)
			{
			//TODO: add code to catch exception.
			}
			
		}

	private void Form1_Load(object sender, EventArgs e)
		{
			
          }

		private void buttonTestSpeed_Click(object sender, EventArgs e)
			{
			Console.WriteLine("\n\nButton clicked to begin Access speed comparisson - {0}", DateTime.Now);

			List<Int32> listPortfolios = new List<Int32>() {1, 2, 3};
			DateTime timeStarted = DateTime.Now;
			DateTime timeLap;
			//Initialize the Data access to SharePoint
			DesignAndDeliveryPortfolioDataContext objSDDPdatacontext = new DesignAndDeliveryPortfolioDataContext(new
			Uri(Properties.AppResources.SharePointSiteURL + Properties.AppResources.SharePointRESTuri)); //"/_vti_bin/listdata.svc"));
			//datacontexSDDP.Credentials = CredentialCache.DefaultCredentials;
			objSDDPdatacontext.Credentials = new NetworkCredential(
					userName: Properties.AppResources.User_Credentials_UserName,
					password: Properties.AppResources.User_Credentials_Password,
					domain: Properties.AppResources.User_Credentials_Domain);
			objSDDPdatacontext.MergeOption = MergeOption.NoTracking;
		// https://msdn.microsoft.com/en-us/library/ff798478.aspx

		//var rsDeliverables1 = from deliverableEntry in datacontexSDDP.Deliverables select deliverableEntry;
		timeStarted = DateTime.Now;
		try
			{
			// Individual record accesss and object populate
			Console.WriteLine("\nRead {1} individual Portfolios started at: {0}", timeStarted, listPortfolios.Count);
			foreach(var specificID in listPortfolios)
				{
				//timeLap = DateTime.Now;
				//ServicePortfolio objPortfolio = new ServicePortfolio();
				//objPortfolio.PopulateObject(datacontexSDDP, specificID);
				//if(objPortfolio != null)
				//	Console.WriteLine("\t + {0} - {1} ({2})", objPortfolio.ID, objPortfolio.Title, DateTime.Now - timeLap);
				}
			Console.WriteLine("Total time: {0}sec", DateTime.Now - timeStarted);

			// Read ALL entries up front and then find individuals entries
			timeStarted = DateTime.Now;
			Console.WriteLine("\n\nRead {1} Porfolios started at: {0}", timeStarted, listPortfolios.Count);

			CompleteDataSet objDataSet = new CompleteDataSet();
			objDataSet.PopulateBaseObjects(objSDDPdatacontext);
			if(objDataSet.dsPortfolios != null && objDataSet.dsPortfolios.Count > 0)
				{
				foreach(var recPortfolio in objDataSet.dsPortfolios.Where(por => listPortfolios.Contains(por.Key)))
					{
					timeLap = DateTime.Now;
					Console.WriteLine("\t + {0} - {1}", recPortfolio.Key, recPortfolio.Value.Title);

					foreach(var recFamily in objDataSet.dsFamilies.Where(fam => fam.Value.ServicePortfolioID == recPortfolio.Key))
						{
						Console.WriteLine("\t\t + {0} - {1})", recFamily.Key, recFamily.Value.Title);
						}
					}
				}

			Console.WriteLine("Total time: {0}sec", DateTime.Now - timeStarted);
			}
		catch(DataServiceQueryException exc)
			{
			Console.WriteLine("Unable to access SharePoint Error: " + exc.HResult + "\n - " + exc.Message);
			}
		catch(DataServiceClientException exc)
			{
			Console.WriteLine("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
			}
		}

		///%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
		/// <summary>
		/// 
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void Button_GenerateExcel_Click(object sender, EventArgs e)
			{
			string parTemplateURL = "https://teams.dimensiondata.com/sites/ServiceCatalogue/DocumentTemplates/ClientRequirementsMappingOutputTemplate.xltx";
			enumDocumentTypes parDocumentType = enumDocumentTypes.Client_Requirement_Mapping_Workbook;
			int intHyperlinkCounter = 1;
			string strCommentText;
			int intSharedStringIndex;
			Dictionary<string, string> dictionaryOfComments = new Dictionary<string, string>();

			// define a new objOpenXMLdocument
			Console.WriteLine("Begin to Generate an MS Excel Workbook");
			oxmlWorkbook objOXMLworkbook = new oxmlWorkbook();
			// use CreateDocumentFromTemplate method to create a new MS Excel Workbook Document based on the relevant template
			Console.WriteLine("Creating the new Worksheet from an existing Template");
			if(objOXMLworkbook.CreateDocWbkFromTemplate(
				parDocumentOrWorkbook: enumDocumentOrWorkbook.Workbook,
				parTemplateURL: parTemplateURL,
				parDocumentType: parDocumentType))
				{
				Console.WriteLine("\t objOXMLworkbook:\n" +
				"\t\t + LocalDocumentPath: {0}\n" +
				"\t\t + DocumentFileName.: {1}\n" +
				"\t\t + DocumentURI......: {2}", objOXMLworkbook.LocalPath, objOXMLworkbook.Filename, objOXMLworkbook.LocalURI);
				}
			else
				{
				// if the creation failed.
				Console.WriteLine("An ERROR occurred and the new MS Word Document could not be created due to above stated ERROR conditions.");
				return;
				}
			try
				{
				// Open the MS Excel document in Edit mode
				// https://msdn.microsoft.com/en-us/library/office/hh298534.aspx
				SpreadsheetDocument objSpreadsheetDocument = SpreadsheetDocument.Open(path: objOXMLworkbook.LocalURI, isEditable: true);
				// Obtain the WorkBookPart from the spreadsheet.
				if(objSpreadsheetDocument.WorkbookPart == null)
					{
					throw new ArgumentException(objOXMLworkbook.LocalURI +
						" does not contain a WorkbookPart.");
					}
				WorkbookPart objWorkbookPart = objSpreadsheetDocument.WorkbookPart;

				//Obtain the SharedStringTablePart - If it doesn't exisit, create new one.
				SharedStringTablePart objSharedStringTablePart;
				if(objSpreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
					{
					objSharedStringTablePart = objSpreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
					}
				else
					{
					objSharedStringTablePart = objSpreadsheetDocument.AddNewPart<SharedStringTablePart>();
					}

				// Get the Matrix worksheet in the Workbook.
				Sheet objMatrixWorksheet = objWorkbookPart.Workbook.Descendants<Sheet>().Where(sht => sht.Name == Properties.AppResources.Workbook_CRM_Worksheet_Name_Matrix).FirstOrDefault();
				if(objMatrixWorksheet == null)
					{
					throw new ArgumentException("The " + Properties.AppResources.Workbook_CRM_Worksheet_Name_Matrix +
						" worksheet could not be loacated in the workbook.");
					}
				// obtain the WorksheetPart of the objMatrixWorksheet
				WorksheetPart objMatrixWorksheetPart = (WorksheetPart)(objWorkbookPart.GetPartById(objMatrixWorksheet.Id));

				// construct the CommentsList to which Comments objects can be appended
				CommentList objMatrixCommentList = new CommentList();

				//WorksheetCommentsPart objMatrixCommentsPart = (WorksheetCommentsPart)(objMatrixWorksheetPart.GetPartsOfType<WorksheetCommentsPart>().First());

				// Copy the Formats from Row 8 into the List of Formats from where it can be applied to every Row at a later stage
				Client_Requirements_Mapping_Workbook objCRMworkbook = new Client_Requirements_Mapping_Workbook();
				List<UInt32Value> listColumnStyles = new List<UInt32Value>();
				int intLastColumn = 15;
				int intStyleSourceRow = 7;
				string strCellAddress = "";
                    for(int i = 0; i < intLastColumn - 1; i++)
					{
					strCellAddress = aWorkbook.GetColumnLetter(i) + intStyleSourceRow;
					Cell objSourceCell = objMatrixWorksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == strCellAddress).FirstOrDefault();
					if(objSourceCell != null)
						{
						listColumnStyles.Add(objSourceCell.StyleIndex);
						}
					else
						listColumnStyles.Add(0U);
                         } // loop

				// Display the List of Formats that was collected.
				Console.WriteLine("Cell Styles..:");
				int col = 0;
				foreach(UInt32Value item in listColumnStyles)
					{
					Console.WriteLine("\t {0} ({1}) = {2} ", col, aWorkbook.GetColumnLetter(col), item.Value);
					col += 1;
					}

				UInt16 intRowNumber = 7;
				Cell objCell = new Cell();

				//--- Column A -------------------------------
				// Write a normal String
				oxmlWorkbook.PopulateCell(
					parWorksheetPart: objMatrixWorksheetPart,
					parColumnLetter: "A",
					parRowNumber: intRowNumber,
					parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("A"))),
					parCellDatatype: CellValues.String,
					parCellcontents: "Service Tower 1");
				//--- Column B --------------------------------
				// Copy only the formatting no content
				oxmlWorkbook.PopulateCell(
					parWorksheetPart: objMatrixWorksheetPart,
					parColumnLetter: "B",
					parRowNumber: intRowNumber,
					parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("B"))),
					parCellDatatype: CellValues.String);
				//--- Column C --------------------------------
				oxmlWorkbook.PopulateCell(
					parWorksheetPart: objMatrixWorksheetPart,
					parColumnLetter: "C",
					parRowNumber: intRowNumber,
					parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("C"))),
					parCellDatatype: CellValues.String);
				//--- Column D --------------------------------
				oxmlWorkbook.PopulateCell(
					parWorksheetPart: objMatrixWorksheetPart,
					parColumnLetter: "D",
					parRowNumber: intRowNumber,
					parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("D"))),
					parCellDatatype: CellValues.String);
				//--- Column E --------------------------------
				oxmlWorkbook.PopulateCell(
					parWorksheetPart: objMatrixWorksheetPart,
					parColumnLetter: "E",
					parRowNumber: intRowNumber,
					parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("E"))),
					parCellDatatype: CellValues.String);
				//--- Column F --------------------------------
				// Write the NEW as a Shared String value to the workbook
				intSharedStringIndex = oxmlWorkbook.InsertSharedStringItem(
					parText2Insert: Properties.AppResources.Workbook_CRM_Matrix_NewColumn_Text, parShareStringPart: objSharedStringTablePart);
				oxmlWorkbook.PopulateCell(
					parWorksheetPart: objMatrixWorksheetPart,
					parColumnLetter: "F",
					parRowNumber: intRowNumber,
					parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("F"))),
					parCellDatatype: CellValues.SharedString,
					parCellcontents: intSharedStringIndex.ToString());

				//add the commments to the dictionaryOfComments - the comments will be added as a last step.
				strCommentText = "The comments in this column contains...";
                    dictionaryOfComments.Add(key: "F" + "|" + intRowNumber , value: strCommentText);

				strCommentText = "This is just a test comment...\nIt will be replaced with actual data.)";
				dictionaryOfComments.Add(key: "G" + "|" + intRowNumber, value: strCommentText);

				//--- Column G --------------------------------
				intSharedStringIndex = oxmlWorkbook.InsertSharedStringItem(
					parText2Insert: Properties.AppResources.Workbook_CRM_Matrix_RowType_TowerOfService, parShareStringPart: objSharedStringTablePart);
				oxmlWorkbook.PopulateCell(
					parWorksheetPart: objMatrixWorksheetPart,
					parColumnLetter: "G",
					parRowNumber: intRowNumber,
					parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("G"))),
					parCellDatatype: CellValues.SharedString,
					parCellcontents: intSharedStringIndex.ToString());
				//--- Column H --------------------------------
				intSharedStringIndex = oxmlWorkbook.InsertSharedStringItem(
					parText2Insert: "Fully Compliant", 
					parShareStringPart: objSharedStringTablePart);
				oxmlWorkbook.PopulateCell(
					parWorksheetPart: objMatrixWorksheetPart,
					parColumnLetter: "H",
					parRowNumber: intRowNumber,
					parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("H"))),
					parCellDatatype: CellValues.SharedString,
					parCellcontents: intSharedStringIndex.ToString());
				//--- Column I --------------------------------
				oxmlWorkbook.PopulateCell(
					parWorksheetPart: objMatrixWorksheetPart,
					parColumnLetter: "I",
					parRowNumber: intRowNumber,
					parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("I"))),
					parCellDatatype: CellValues.String);
				//--- Column J --------------------------------
				oxmlWorkbook.PopulateCell(
					parWorksheetPart: objMatrixWorksheetPart,
					parColumnLetter: "J",
					parRowNumber: intRowNumber,
					parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("J"))),
					parCellDatatype: CellValues.String,
					parCellcontents: "RFP Page " + intRowNumber + " Paragraph: Kii");
				//--- Column K --------------------------------
				oxmlWorkbook.PopulateCell(
					parWorksheetPart: objMatrixWorksheetPart,
					parColumnLetter: "K",
					parRowNumber: intRowNumber,
					parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("K"))),
					parCellDatatype: CellValues.String);
				//--- Column L --------------------------------
				intHyperlinkCounter += 1;
				oxmlWorkbook.PopulateCell(
					parWorksheetPart: objMatrixWorksheetPart,
					parColumnLetter: "L",
					parRowNumber: intRowNumber,
					parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("L"))),
					parCellDatatype: CellValues.Number,
					parCellcontents: "101",
					parHyperlinkCounter: intHyperlinkCounter,
					parHyperlinkURL: Properties.AppResources.SharePointURL +
						Properties.AppResources.List_MappingServiceTowers +
						Properties.AppResources.EditFormURI + 101 );
				//--- Column M --------------------------------
				intHyperlinkCounter += 1;
				oxmlWorkbook.PopulateCell(
					parWorksheetPart: objMatrixWorksheetPart,
					parColumnLetter: "M",
					parRowNumber: intRowNumber,
					parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("M"))),
					parCellDatatype: CellValues.Number,
					parCellcontents: "2",
					parHyperlinkCounter: intHyperlinkCounter,
					parHyperlinkURL: Properties.AppResources.SharePointURL +
						Properties.AppResources.List_DeliverablesURI +
						Properties.AppResources.EditFormURI + 2);
				//--- Column N --------------------------------
				intHyperlinkCounter += 1;
				oxmlWorkbook.PopulateCell(
					parWorksheetPart: objMatrixWorksheetPart,
					parColumnLetter: "N",
					parRowNumber: intRowNumber,
					parStyleId: (UInt32Value)(listColumnStyles.ElementAt(aWorkbook.GetColumnNumber("N"))),
					parCellDatatype: CellValues.Number,
					parCellcontents: "1",
					parHyperlinkCounter: intHyperlinkCounter,
					parHyperlinkURL: Properties.AppResources.SharePointURL +
						Properties.AppResources.List_ServiceLevelsURI +
						Properties.AppResources.EditFormURI + 1);

				// Now insert all the Comments
				aWorkbook.InsertWorksheetComments(
					parWorksheetPart: objMatrixWorksheetPart,
					parDictionaryOfComments: dictionaryOfComments);
				//oxmlWorkbook.InsertComments(
				//	parWorksheetPart: objMatrixWorksheetPart,
				//	parDictCommentsToAdd: dictionaryOfComments);

				// Close the document
				//Validate the document with OpenXML validator
				OpenXmlValidator objOXMLvalidator = new OpenXmlValidator(fileFormat: FileFormatVersions.Office2010);
				int errorCount = 0;
				Console.WriteLine("\n\rValidating document....");
				foreach(ValidationErrorInfo validationError in objOXMLvalidator.Validate(objSpreadsheetDocument))
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

				Console.WriteLine("Workbook generation completed, saving and closing the document.");
				// Save and close the Document
				objSpreadsheetDocument.Close();

				Console.WriteLine("Workbook saved and closed!!!");
				} // end Try

			catch(OpenXmlPackageException exc)
				{
				//TODO: add code to catch exception.
				}
			catch(ArgumentNullException exc)
				{
				//TODO: add code to catch exception.
				}
			}

		
		private void buttonSQLiteTest_Click(object sender, EventArgs e)
			{
				
			//string strDBfilePath = Path.GetFullPath("\\") + Properties.AppResources.LocalDatabasePath;
			//string strDBfileName = Properties.AppResources.LocalDatabaseName;
			//string strDB = Path.Combine(strDBfilePath, strDBfileName);
			Console.Write("\nTest Starts...");

			string str1 = "ABC";
			int int1 = 123;
			string strResult = str1 + int1.ToString();

			Console.Write("\n\t - Concatenated value of {0} and {1} = {2}", str1, int1, strResult);
			//if(File.Exists(strDB))
			//	{
			//	Console.Write("Yes\n");
			//	}
			//else
			//	{
			//	Console.Write("No, needs to CREATE the Database;");
			//	bSuccess = ObjectDatabase.CreateSQLDataBase(parDBFilePath: strDBfilePath, parDBFileName: strDBfileName);
			//	}

			// Initialise the connection to the SQLiteDatabase...
			//var objSQLiteConnection = new SQLiteConnection(strDB);
			//var objLINQcontext = new DataContext(objSQLiteConnection);


			Console.WriteLine("Test Completed...");
			}
		}

	public static class Globals
		{
		public static CompleteDataSet objDataSet;
		}

	}
