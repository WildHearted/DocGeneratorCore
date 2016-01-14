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
using DogGenUI;
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
					Console.WriteLine("\n\n ERROR: There was an error accessing the Document Collections. \n{0}", returnResult);
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
						Console.WriteLine("\r\nReady to generate entry: {0} - {1}", objDocCollection.ID.ToString(), objDocCollection.Title);
						lblConnect.Text = "Generating " + objDocCollection.ID.ToString() + " - " + objDocCollection.Title + "...";
						lblConnect.Refresh();

						// Process each of the documents in the DocumentCollection
						if(objDocCollection.Document_and_Workbook_objects.Count() > 0)
							{
							objDocCollection.Document_and_Workbook_objects.GetType();
							foreach(dynamic objDocumentWorkbook in objDocCollection.Document_and_Workbook_objects)
								{
								Console.WriteLine("\t\t ObjectType: {0}", objDocumentWorkbook.GetType());
								objectType = objDocumentWorkbook.GetType();
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
													Console.WriteLine("");
													}
												else
													{
													Console.WriteLine("");
													}
												Console.WriteLine("");
												}
											break;
										}
									case ("Content_Status_Workbook"):
										{
										break;
										}
									case ("Contract_SoW_Service_Description"):
										{
										break;
										}
									case ("CSD_based_on_ClientRequirementsMapping"):
										{
										break;
										}
									case ("CSD_Document_DRM_Inline"):
										{
										break;
										}
									case ("CSD_Document_DRM_Sections"):
										{
										break;
										}
									case ("External_Technology_Coverage_Dashboard_Workbook"):
										{
										break;
										}
									case ("Internal_Technology_Coverage_Dashboard_Workbook"):
										{
										break;
										}
									case ("ISD_Document_DRM_Inline"):
										{
										break;
										}
									case ("ISD_Document_DRM_Sections"):
										{
										break;
										}
									case ("Pricing_Addendum_Document"):
										{
										break;
										}
									case ("RACI_Matrix_Workbook_per_Deliverable"):
										{
										break;
										}
									case ("RACI_Workbook_per_Role"):
										{
										break;
										}
									case ("Services_Framework_Document_DRM_Inline"):
										{
										break;
										}
									case ("Services_Framework_Document_DRM_Sections"):
										{
										break;
										}
									default:
										break;
									}
								}
							} // end if ...Count() > 0

						}

					Console.WriteLine("\n\n{0} Document Collection(s) were Generated.", docCollectionsToGenerate.Count);
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
