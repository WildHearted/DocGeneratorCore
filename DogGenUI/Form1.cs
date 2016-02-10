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
					}
				else if(returnResult.Substring(0,5) == "Error")
					{
					Console.WriteLine("\nERROR: There was an error accessing the Document Collections. \n{0}", returnResult);
					}
				}
			catch(InvalidProgramException ex)
				{
				Console.WriteLine("Exception occurred [{0}] \n Inner Exception: {1}", ex.Message, ex.InnerException);
				}
			// Continue here if there are any Document Collections to generate...

			string objectType = "";
			try
				{
				if(docCollectionsToGenerate.Count > 0)
					{
					foreach(DocumentCollection objDocCollection in docCollectionsToGenerate)
						{
						Console.WriteLine("\r\nReady to generate Document Collection: {0} - {1}", objDocCollection.ID.ToString(), objDocCollection.Title);

						// Process each of the documents in the DocumentCollection
						if(objDocCollection.Document_and_Workbook_objects.Count() > 0)
							{
							//objDocCollection.Document_and_Workbook_objects.GetType();
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
					}
				else
					{
					Console.WriteLine("Sorry, nothing to generate at this stage.");
					}
				}
			catch(Exception ex)	// if the List is empty - nothing to generate
				{
				Console.WriteLine("Exception Error: {0} occurred and means {1}", ex.Source, ex.Message);
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
			int TableCaptionCounter = 0;
			int ImageCaptionCounter = 0;
			
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
			try
				{
				// Open the MS Word document in Edit mode
				WordprocessingDocument objWPdocument = WordprocessingDocument.Open(path: objOXMLdocument.LocalDocumentURI, isEditable: true);
				// Define all open XML objects to use for building the document
				MainDocumentPart objMainDocumentPart = objWPdocument.MainDocumentPart;
				Body objBody = objWPdocument.MainDocumentPart.Document.Body;
				Paragraph objParagraph = new Paragraph();    // Define the objParagraph	
				Run objRun = new Run();
				// Now begin to write the content to the document
				objParagraph = oxmlDocument.Insert_Section("Introductory");
				objBody.Append(objParagraph);
				objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 1, parText2Write: "Introduction", parRestartNumbering: true);
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
				objRun = oxmlDocument.Construct_RunText("Another paragrpah with just normal text.");
				objParagraph.Append(objRun);
				objBody.Append(objParagraph);

				objParagraph = oxmlDocument.Insert_Heading(2, "Executive Summary", false);
				objBody.Append(objParagraph);
				objParagraph = oxmlDocument.Construct_Paragraph(2);
				objRun = oxmlDocument.Construct_RunText("Below is an image of my favourite car. ");
				objParagraph.Append(objRun);
				objBody.Append(objParagraph);

				// Determine the Page Size for the current Body object.
				SectionProperties objSectionProperties = new SectionProperties();
				UInt32 pageWith = 11900U;
				int pixelsPerInch;
				using(System.Drawing.Graphics drawGraphics = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
					{pixelsPerInch = (int)drawGraphics.DpiX;}
				int twipsPerPixel = 1440 / pixelsPerInch;
				Console.WriteLine("TwipsPerPixel: {0}", twipsPerPixel);

				if(objBody.GetFirstChild<SectionProperties>() != null)
					{
					Console.WriteLine("SectionProperties: {0}", objBody.GetFirstChild<SectionProperties>());
					objSectionProperties = objBody.GetFirstChild<SectionProperties>();
					PageSize objPageSize = objSectionProperties.GetFirstChild<PageSize>();
					PageMargin objPageMargin = objSectionProperties.GetFirstChild<PageMargin>();
					if(objPageSize != null)
						pageWith = objPageSize.Width;
					if(objPageMargin != null)
						{
						if(objPageMargin.Left != null)
							pageWith -= objPageMargin.Left;
						if(objPageMargin.Right != null)
							pageWith -= objPageMargin.Right;
						}
	                    }
				Console.WriteLine("The pageWidth in Pixels: {0}px", pageWith);
	
				// Insert and image in the document
				objParagraph = oxmlDocument.Construct_Paragraph(2);
				objRun = oxmlDocument.InsertImage(
					parMainDocumentPart: ref objMainDocumentPart,
					parParagraphLevel: 2,
					parPictureSeqNo: 1,
					parImageURL: "C:\\Users\\ben.vandenberg\\Desktop\\2015-10-05 22.31.26.jpg");
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
				ImageCaptionCounter += 1;
				objParagraph = oxmlDocument.Construct_Caption(parCaptionType: "Image", parCaptionSequence: ImageCaptionCounter, parCaptionText: "An awesome machine.");
				objBody.Append(objParagraph);


				// Insert a Section and Heading for the Table section.
				//objParagraph = oxmlDocument.Insert_Section(parText2Write: "Tables");
				objParagraph = oxmlDocument.Insert_Heading(2,"Tables");
				objBody.Append(objParagraph);
				objParagraph = oxmlDocument.Construct_Paragraph(2);
				objRun = oxmlDocument.Construct_RunText("This demonstrates how tables are handled by the DocGenerator application.", parBold: true);
				objParagraph.Append(objRun);
				objBody.Append(objParagraph);

				//Table Construction code

				// Construct a Table object instance
				DocumentFormat.OpenXml.Wordprocessing.Table objTable = new DocumentFormat.OpenXml.Wordprocessing.Table();
				objTable = oxmlDocument.ConstructTable(parTableWidth: pageWith, parFirstRow: true, parFirstColumn: true, parLastColumn: true, parLastRow: true, parNoVerticalBand: true, parNoHorizontalBand: false);
				DocumentFormat.OpenXml.Wordprocessing.TableRow objTableRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
				DocumentFormat.OpenXml.Wordprocessing.TableCell objTableCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
				bool IsFirstRow = false;
				bool IsLastRow = false;
				bool IsFirstColumn = false;
				bool IsLastColumn = false;
				int numberOfRows = 6;
				int numberOfColumns = 4;
				string tableText = "";
				UInt32 columnWidth = pageWith / Convert.ToUInt32(numberOfColumns);
				// Construct a TableGrid object instance
				DocumentFormat.OpenXml.Wordprocessing.TableGrid objTableGrid = new DocumentFormat.OpenXml.Wordprocessing.TableGrid();
				List<UInt32> lstTableColumns = new List<UInt32>();
				for(int i = 0; i < numberOfColumns; i++)
					{
					lstTableColumns.Add(columnWidth);
					}
				objTableGrid = oxmlDocument.ConstructTableGrid(lstTableColumns, "px", pageWith);
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
				objBody.Append(objTable);
				objParagraph = oxmlDocument.Construct_Caption(parCaptionType: "Table", parCaptionSequence: ImageCaptionCounter, parCaptionText: "A table generated by the app.");
				objBody.Append(objParagraph);
				
				// Insert a new XML Table based on an HTML table input from a local file.
				objParagraph = oxmlDocument.Insert_Section(parText2Write: "HTML Content Test" );
				objBody.Append(objParagraph);
				objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 1, parText2Write: "First part of HTML Content", parRestartNumbering: true);
				objBody.Append(objParagraph);
				HTMLdecoder objHTMLdecoder = new HTMLdecoder();
				objHTMLdecoder.WPbody = objBody;

				string sCurrentDirectory = Directory.GetCurrentDirectory();
				Console.WriteLine("Current Directory is {0}", sCurrentDirectory);
				//string sFile = @"C:\Users\ben.vandenberg\Desktop\HTMLtest\IntroSimple.txt";
				//string sFile = @"C:\Users\ben.vandenberg\Desktop\HTMLtest\IntroComplex.txt";
				string sFile = @"C:\Users\ben.vandenberg\Desktop\HTMLtest\IntoFromSharePoint.txt";
				string sContent = System.IO.File.ReadAllText(sFile);
				objHTMLdecoder.DecodeHTML(
					parMainDocumentPart: ref objMainDocumentPart,
					parDocumentLevel: 1,
					parPageWidth: pageWith,
					parHTML2Decode: sContent,
					parTableCaptionCounter: ref TableCaptionCounter,
					parImageCaptionCounter: ref ImageCaptionCounter);

				// Close the document
				objParagraph = oxmlDocument.Construct_Paragraph(1);
				objRun = oxmlDocument.Construct_RunText("-- - end of the document --- ");
				objParagraph.Append(objRun);
				objBody.Append(objParagraph);

				Console.WriteLine("\t\t Document generated, now saving and closing the document.");

				Console.WriteLine("Paragraph updated, now saving and closing the document.");
				// Save and close the Document
				objWPdocument.Close();
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
			textBoxFileName.Text = "https://teams.dimensiondata.com/sites/ServiceCatalogue/DocumentTemplates/InternalServiceDefinitionTemplate.dotx";
               }
		}
	}
