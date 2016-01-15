using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Services.Client;
using System.Drawing;
using System.Dynamic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.SharePoint;
using DogGenUI.SDDPServiceReference;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
namespace DogGenUI
	{

	public partial class Form1 : Form
		{
		/// <summary>
		///	Declare the SharePoint connection as a DataContext
		/// </summary>
		private static string websiteURL = "https://teams.dimensiondata.com/sites/ServiceCatalogue";
		DesignAndDeliveryPortfolioDataContext datacontexSDDP = new DesignAndDeliveryPortfolioDataContext(new Uri(websiteURL + "/_vti_bin/listdata.svc"));
		public string ErrorLogMessage = "";

		public Form1()
			{
			InitializeComponent();
			datacontexSDDP.Credentials = CredentialCache.DefaultCredentials;
			}

		private void btnSDDP_Click(object sender, EventArgs e)
			{
			Cursor.Current = Cursors.WaitCursor;
			string returnResult = "";
			List<DocumentCollection> docCollectionsToGenerate = new List<DocumentCollection>();
			try
				{
				returnResult = DocumentCollection.GetCollectionsToGenerate(ref docCollectionsToGenerate);
				if (returnResult.Substring(0,4) == "Good")
					{
					Console.WriteLine("\r\nThere are {0} Document Collections to generate.", docCollectionsToGenerate.Count());
					lblConnect.Text = "There are " + docCollectionsToGenerate.Count + " Document Collections to generate...";
					}
				else if(returnResult.Substring(0,5) == "Error")
					{
					Console.WriteLine("\nERROR: There was an error accessing the Document Collections. \n{0}", returnResult);
					lblConnect.Text = "ERROR: There was an error processing the Document Collections.";
					}
				}
			catch(InvalidProgramException ex)
				{
				Console.WriteLine("Exception occurred [{0}] \n Inner Exception: {1}", ex.Message, ex.InnerException);
				}
			// Continue here if there are any Document Collections to generate...
			lblConnect.Refresh();

			string objectType = "";
			try
				{
				if(docCollectionsToGenerate.Count > 0)
					{
					foreach(DocumentCollection objDocCollection in docCollectionsToGenerate)
						{
						Console.WriteLine("\r\nReady to generate Document Collection: {0} - {1}", objDocCollection.ID.ToString(), objDocCollection.Title);
						lblConnect.Text = "Generating " + objDocCollection.ID.ToString() + " - " + objDocCollection.Title + "...";
						lblConnect.Refresh();

						// Process each of the documents in the DocumentCollection
						if(objDocCollection.Document_and_Workbook_objects.Count() > 0)
							{
							objDocCollection.Document_and_Workbook_objects.GetType();
							foreach(dynamic objDocumentWorkbook in objDocCollection.Document_and_Workbook_objects)
								{
								Console.WriteLine("\r\t Generate ObjectType: {0}", objDocumentWorkbook.ToString());
								objectType = objDocumentWorkbook.ToString();
								objectType = objectType.Substring(objectType.IndexOf(".")+1,(objectType.Length - objectType.IndexOf(".")-1));
								switch(objectType)
									{
									case ("Client_Requirements_Mapping_Workbook"):
										{
										Client_Requirements_Mapping_Workbook objCRMworkbook = objDocumentWorkbook;
										if(objCRMworkbook.Generate())
											{
											if(objCRMworkbook.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.", objCRMworkbook.ErrorMessages.Count);
												Utilities.WriteErrorsToConsole(objCRMworkbook.ErrorMessages);
												}

											Console.WriteLine("\t Completed generation of {0}", objDocumentWorkbook.GetType());
											}
										break;
										}
									case ("Content_Status_Workbook"):
										{
										Content_Status_Workbook objcontentStatus = objDocumentWorkbook;
										if(objcontentStatus.Generate())
											{
											if(objcontentStatus.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t {0} error(s) occurred during the generation process.", objcontentStatus.ErrorMessages.Count);
												Utilities.WriteErrorsToConsole(objcontentStatus.ErrorMessages);
												}

											Console.WriteLine("\t Completed generation of {0}", objDocumentWorkbook.GetType());
											}
										break;
										}
									case ("Contract_SoW_Service_Description"):
										{
										Contract_SoW_Service_Description objContractSoW = objDocumentWorkbook;
										if(objContractSoW.Generate())
											{
											if(objContractSoW.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t {0} error(s) occurred during the generation process.", objContractSoW.ErrorMessages.Count);
												Utilities.WriteErrorsToConsole(objContractSoW.ErrorMessages);
												}

											Console.WriteLine("\t Completed generation of {0}", objDocumentWorkbook.GetType());
											}
										break;
										}
									case ("CSD_based_on_ClientRequirementsMapping"):
										{
										CSD_based_on_ClientRequirementsMapping objCSDbasedCRM = objDocumentWorkbook;
										if(objCSDbasedCRM.Generate())
											{
											if(objCSDbasedCRM.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t {0} error(s) occurred during the generation process.", objCSDbasedCRM.ErrorMessages.Count);
												Utilities.WriteErrorsToConsole(objCSDbasedCRM.ErrorMessages);
												}

											Console.WriteLine("\t Completed generation of {0}", objDocumentWorkbook.GetType());
											}
										break;
										}
									case ("CSD_Document_DRM_Inline"):
										{
										CSD_Document_DRM_Inline objCSDdrmInline = objDocumentWorkbook;
										if(objCSDdrmInline.Generate())
											{
											if(objCSDdrmInline.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t {0} error(s) occurred during the generation process.", objCSDdrmInline.ErrorMessages.Count);
												Utilities.WriteErrorsToConsole(objCSDdrmInline.ErrorMessages);
												}

											Console.WriteLine("\t Completed generation of {0}", objDocumentWorkbook.GetType());
											}
										break;
										}
									case ("CSD_Document_DRM_Sections"):
										{
										CSD_Document_DRM_Sections objCSDdrmSections = objDocumentWorkbook;
										if(objCSDdrmSections.Generate())
											{
											if(objCSDdrmSections.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t {0} error(s) occurred during the generation process.", objCSDdrmSections.ErrorMessages.Count);
												Utilities.WriteErrorsToConsole(objCSDdrmSections.ErrorMessages);
												}

											Console.WriteLine("\t Completed generation of {0}", objDocumentWorkbook.GetType());
											}
										break;
										}
									case ("External_Technology_Coverage_Dashboard_Workbook"):
										{
										External_Technology_Coverage_Dashboard_Workbook objExtTechDashboard = objDocumentWorkbook;
										if(objExtTechDashboard.Generate())
											{
											if(objExtTechDashboard.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.", objExtTechDashboard.ErrorMessages.Count);
												Utilities.WriteErrorsToConsole(objExtTechDashboard.ErrorMessages);
												}

											Console.WriteLine("\t Completed generation of {0}", objDocumentWorkbook.GetType());
											}
										break;
										}
									case ("Internal_Technology_Coverage_Dashboard_Workbook"):
										{
										Internal_Technology_Coverage_Dashboard_Workbook objIntTechDashboard = objDocumentWorkbook;
										if(objIntTechDashboard.Generate())
											{
											if(objIntTechDashboard.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.", objIntTechDashboard.ErrorMessages.Count);
												Utilities.WriteErrorsToConsole(objIntTechDashboard.ErrorMessages);
												}

											Console.WriteLine("\t Completed generation of {0}", objDocumentWorkbook.GetType());
											}
										break;
										}
									case ("ISD_Document_DRM_Inline"):
										{
										ISD_Document_DRM_Inline objISDdrmInline = objDocumentWorkbook;
										if(objISDdrmInline.Generate())
											{
											if(objISDdrmInline.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.", objISDdrmInline.ErrorMessages.Count);
												Utilities.WriteErrorsToConsole(objISDdrmInline.ErrorMessages);
												}

											Console.WriteLine("\t Completed generation of {0}", objDocumentWorkbook.GetType());
											}
										break;
										}
									case ("ISD_Document_DRM_Sections"):
										{
										ISD_Document_DRM_Sections objISDdrmSections = objDocumentWorkbook;
										if(objISDdrmSections.Generate())
											{
											if(objISDdrmSections.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.", objISDdrmSections.ErrorMessages.Count);
												Utilities.WriteErrorsToConsole(objISDdrmSections.ErrorMessages);
												}

											Console.WriteLine("\t Completed generation of {0}", objDocumentWorkbook.GetType());
											}
										break;
										}
									case ("Pricing_Addendum_Document"):
										{
										Pricing_Addendum_Document objPricingAddendum = objDocumentWorkbook;
										if(objPricingAddendum.Generate())
											{
											if(objPricingAddendum.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.", objPricingAddendum.ErrorMessages.Count);
												Utilities.WriteErrorsToConsole(objPricingAddendum.ErrorMessages);
												}

											Console.WriteLine("\t Completed generation of {0}", objDocumentWorkbook.GetType());
											}
										break;
										}
									case ("RACI_Matrix_Workbook_per_Deliverable"):
										{
										RACI_Matrix_Workbook_per_Deliverable objRACImatrix = objDocumentWorkbook;
										if(objRACImatrix.Generate())
											{
											if(objRACImatrix.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.", objRACImatrix.ErrorMessages.Count);
												Utilities.WriteErrorsToConsole(objRACImatrix.ErrorMessages);
												}

											Console.WriteLine("\t Completed generation of {0}", objDocumentWorkbook.GetType());
											}
										break;
										}
									case ("RACI_Workbook_per_Role"):
										{
										RACI_Workbook_per_Role objRACIperRole = objDocumentWorkbook;
										if(objRACIperRole.Generate())
											{
											if(objRACIperRole.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.", objRACIperRole.ErrorMessages.Count);
												Utilities.WriteErrorsToConsole(objRACIperRole.ErrorMessages);
												}

											Console.WriteLine("\t Completed generation of {0}", objDocumentWorkbook.GetType());
											}
										break;
										}
									case ("Services_Framework_Document_DRM_Inline"):
										{
										Services_Framework_Document_DRM_Inline objSFdrmInline = objDocumentWorkbook;
										if(objSFdrmInline.Generate())
											{
											if(objSFdrmInline.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.", objSFdrmInline.ErrorMessages.Count);
												Utilities.WriteErrorsToConsole(objSFdrmInline.ErrorMessages);
												}

											Console.WriteLine("\t Completed generation of {0}", objDocumentWorkbook.GetType());
											}
										break;
										}
									case ("Services_Framework_Document_DRM_Sections"):
										{
										Services_Framework_Document_DRM_Sections objSFdrmSections = objDocumentWorkbook;
										if(objSFdrmSections.Generate())
											{
											if(objSFdrmSections.ErrorMessages.Count() > 0)
												{
												Console.WriteLine("\t *** {0} error(s) occurred during the generation process.", objSFdrmSections.ErrorMessages.Count);
												Utilities.WriteErrorsToConsole(objSFdrmSections.ErrorMessages);
												}

											Console.WriteLine("\t Completed generation of {0}", objDocumentWorkbook.GetType());
											}
										break;
										}
									default:
										break;
									}
								}
							} // end if ...Count() > 0

						}

					Console.WriteLine("\nDocuments for {0} Document Collection(s) were Generated.", docCollectionsToGenerate.Count);
					lblConnect.Text = "Document Generation completed for " + docCollectionsToGenerate.Count + " document collections.";
					lblConnect.Refresh();
					}
				else
					{
					Console.WriteLine("Sorry, nothing to generate at this stage.");
					lblConnect.Text = "Sorry, nothing to generate at this stage.";
					lblConnect.Refresh();
					}
				}
			catch(Exception ex)	// if the List is empty - nothing to generate
				{
				Console.WriteLine("Exception Error: {0} occurred and means {1}", ex.Source, ex.Message);
				lblConnect.Text = "Exception error" + ex.HResult + " - " + ex.Message;
				lblConnect.Refresh();
				}
			finally
				{
				Cursor.Current = Cursors.Default;
				}
			}

		private void btnOpenMSwordDocument(object sender, EventArgs e)
			{
			string parTemplateURL = "https://teams.dimensiondata.com/sites/ServiceCatalogue/DocumentTemplates/InternalServiceDefinitionTemplate.dotx";
			enumDocumentTypes parDocumentType = enumDocumentTypes.ISD_Document_DRM_Inline;

			// define a new objOpenXMLdocument
			oxmlDocument objOXMLdocument = new oxmlDocument();
			// use CreateDocumentFromTemplate method to create a new MS Word Document based on the relevant template
			if (objOXMLdocument.CreateDocumentFromTemplate(parTemplateURL: parTemplateURL, parDocumentType: parDocumentType))
				{
				Console.WriteLine("objOXMLdocument:\n" +
				"                + LocalDocumentPath: {0}\n" +
				"                + DocumentFileName.: {1}\n" +
				"                + DocumentURI......: {2}", objOXMLdocument.LocalDocumentPath, objOXMLdocument.DocumentFilename, objOXMLdocument.LocalDocumentURI);
				}
			else
				{
				// if the creation failed.
				Console.WriteLine("An ERROR occurred and the new MS Word Document could not be created due to above stated ERROR conditions.");
				return;
				}

			

			string newText = "";

			newText = "This is a Heading 1 - added with oXML";
			// Open the MS Word document in Edit mode
			WordprocessingDocument objDocument = WordprocessingDocument.Open(path: objOXMLdocument.LocalDocumentURI, isEditable: true);
			// Define the objBody of the document
			Body objBody = objDocument.MainDocumentPart.Document.Body;
			// Insert a new Paragraph to the end of the Body of the objDocument 
			Paragraph objParagraph = objBody.AppendChild(new Paragraph());
			// Insert a new Run object in the new objParagraph
			Run objRun = objParagraph.AppendChild(new Run());
			// Insert the text in the objRun of the objParagraph
			objRun.AppendChild(new Text(newText));
			// Check if the paragraph has any paragraph properties, if not add ParagraphProperties to it.
			if(objParagraph.Elements<ParagraphProperties>().Count() == 0)
			 	{
			 	objParagraph.PrependChild<ParagraphProperties>(new ParagraphProperties());
				}
			// Get the first PropertiesElement for the paragraph.
			ParagraphProperties objParagraphProperties = objParagraph.Elements<ParagraphProperties>().First();
			// Set the value of the ParagraphStyleId to "Heading1"
			objParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = "Heading1" };

			// Insert a new Paragraph to the end of the Body of the objDocument 
			objParagraph = objBody.AppendChild(new Paragraph());
			// Insert a new Run object in the new objParagraph
			objRun = objParagraph.AppendChild(new Run());
			// Insert the text in the objRun of the objParagraph
			objRun.AppendChild(new Text("This is a Heading 2 - added with oXML"));
			// Check if the paragraph has any paragraph properties, if not add ParagraphProperties to it.
			if(objParagraph.Elements<ParagraphProperties>().Count() == 0)
				{
				objParagraph.PrependChild<ParagraphProperties>(new ParagraphProperties());
				}
			// Get the first PropertiesElement for the paragraph.
			objParagraphProperties = objParagraph.Elements<ParagraphProperties>().First();
			// Set the value of the ParagraphStyleId to "Heading2"
			objParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = "Heading2" };



			Console.WriteLine("Paragraph updated, now saving and closing the document.");
			// Save and close the Document
			objDocument.Close();
			
			}

		private void Form1_Load(object sender, EventArgs e)
			{
			textBoxFileName.Text = "https://teams.dimensiondata.com/sites/ServiceCatalogue/DocumentTemplates/InternalServiceDefinitionTemplate.dotx";
               }
		}
	}
