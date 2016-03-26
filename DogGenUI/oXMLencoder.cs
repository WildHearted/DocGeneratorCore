using System;
using System.IO;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DrwWp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DrwWp2010 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using Drw =DocumentFormat.OpenXml.Drawing;
using Drw2010 = DocumentFormat.OpenXml.Office2010.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Spreadsheet;

// Reference sources:
// https://msdn.microsoft.com/en-us/library/office/ff478255.aspx (Baic Open XML Documents)
// https://msdn.microsoft.com/en-us/library/dd469465%28v=office.12%29.aspx (Examples with merging and Presentations)
// http://blogs.msdn.com/b/vsod/archive/2012/02/18/how-to-create-a-document-from-a-template-dotx-dotm-and-attach-to-it-using-open-xml-sdk.aspx (Example of creating a new document based on a .dotx template.)
// (Example to Replace text in a document) http://www.codeproject.com/Tips/666751/Use-OpenXML-to-Create-a-Word-Document-from-an-Exis
// (Structure of an oXML document) https://msdn.microsoft.com/en-us/library/office/gg278308.aspx
namespace DocGenerator
	{

	public enum enumDocumentOrWorkbook
		{
		Document = 1,
		Workbook = 2
		}

	public class oxmlDocumentWorkbook
		{
		// Object Properties
		private string _localPath = "";
		public string LocalPath
			{
			get{return this._localPath;}
			private set{this._localPath = value;}
			}

		private string _fileName = "";
		public string Filename
			{
			get{return this._fileName;}
			private set{this._fileName = value;}
			}

		private string _localURI = "";
		public string LocalURI
			{
			get{return this._localURI;}
			private set{this._localURI = value;}
			}

		private enumDocumentOrWorkbook _documentOrWorkbook;
		public enumDocumentOrWorkbook DocumentOrWorkbook
			{
			get{return this._documentOrWorkbook;}
			set{this._documentOrWorkbook = value;}
			}

		//----------------------------------
		//--- CreateDocumentFromTemplate ---
		/// <summary>
		/// Use this method to create the new document object with which to work.
		/// It will create the new document based on the specified Tempate and Document Type. Upon creation, the LocalDocument
		/// </summary>
		/// <param name="parTemplateURL">
		/// This value must be the web URI of the template residing in the Document Templates List in SharePoint</param>
		/// <param name="parDocumentOrWorkbook">
		/// This value is the enumerated Document Type</param>
		/// <returns>
		/// Returns a bool with true if the creatin of the oxmlDoument object was successful and false if it failed.
		/// Validate that the bool is TRUE on return of the method.
		/// </returns>
		public bool CreateDocWbkFromTemplate(
			enumDocumentOrWorkbook parDocumentOrWorkbook,
			string parTemplateURL, 
			enumDocumentTypes parDocumentType)
			{
			string ErrorLogMessage = "";
			this.DocumentOrWorkbook = parDocumentOrWorkbook;

			//Derive the file name of the template document
			//			Console.WriteLine(" Template URL: [{0}] \r\n" +
			//"         1         2         3         4         5         6         7         8         9        11        12        13        14        15\r\n" +
			//"12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890 \r\n" ,parTemplateURL);

			string templateFileName = parTemplateURL.Substring(parTemplateURL.LastIndexOf("/") + 1, (parTemplateURL.Length - parTemplateURL.LastIndexOf("/")) - 1);

			// Check if the DocGenerator Template Directory Exist and that it is accessable
			// Configure and validate for the relevant Template
			string templateDirectory = Path.GetFullPath("\\") + Properties.AppResources.LocalTemplatePath;
			try
				{
				if(Directory.Exists(@templateDirectory))
					{
					Console.WriteLine("\t\t\t The templateDirectory [" + templateDirectory + "] exist and are ready to be used.");
					}
				else
					{
					DirectoryInfo templateDirInfo = Directory.CreateDirectory(@templateDirectory);
					Console.WriteLine("\t\t\t The templateDirectory [" + templateDirectory + "] was created and are ready to be used.");
					}
				}
			catch(UnauthorizedAccessException exc)
				{
				ErrorLogMessage = "The current user: [" + System.Security.Principal.WindowsIdentity.GetCurrent().Name +
				"] does not have the required security permissions to access the template directory at: " + templateDirectory +
				"\r\n " + exc.Message + " in " + exc.Source;
				Console.WriteLine(ErrorLogMessage);
				return false;
				}
			catch(NotSupportedException exc)
				{
				ErrorLogMessage = "The path of template directory [" + templateDirectory + 
					"] contains invalid characters. Ensure that the path is valid and  contains legible path characters only. \r\n " + 
					exc.Message + " in " + exc.Source;
				Console.WriteLine(ErrorLogMessage);
				return false;
				}
			catch(DirectoryNotFoundException exc)
				{
				ErrorLogMessage = "The path of template directory [" + templateDirectory + 
					"] is invalid. Check that the drive is mapped and exist /r/n " + exc.Message + " in " + exc.Source;
				Console.WriteLine(ErrorLogMessage);
				return false;
				}
			
			// Check if the template file exist in the template directory
			if(File.Exists(templateDirectory + templateFileName))
				{
				// If the the template exist just proceed...
				Console.WriteLine("\t\t\t The template already exist and are ready for use: " + templateDirectory + templateFileName);
				}
			else
				{
				// Download the relevant template from SharePoint
				WebClient objWebClient = new WebClient();
				objWebClient.UseDefaultCredentials = true;
				//objWebClient.Credentials = CredentialCache.DefaultCredentials;
				try
					{
					objWebClient.DownloadFile(parTemplateURL, templateDirectory + "\\" + templateFileName);
					}
				catch(WebException exc)
					{
					ErrorLogMessage = "The template file could not be downloaded from SharePoint List [" + parTemplateURL + "]. " +
						"\n - Check that the template exist in SharePoint \n - that it is accessible \n - " +
						"and that the network connection is working. \n " + exc.Message + " in " + exc.Source;
					Console.WriteLine(ErrorLogMessage);
					return false;
					}
				}
			Console.WriteLine("\t\t\t Template: {0} exist in directory: {1}? {2}", 
				templateFileName, templateDirectory, File.Exists(templateDirectory + templateFileName));

			// Check if the DocGenerator\Documents Directory exist and that it is accessable
			string documentDirectory = Path.GetFullPath("\\") + Properties.AppResources.LocalDocumentPath;
			if(!Directory.Exists(documentDirectory))
				{
				try
					{
					Directory.CreateDirectory(@documentDirectory);
					}
				catch(UnauthorizedAccessException exc)
					{
					ErrorLogMessage = "The current user: [" + System.Security.Principal.WindowsIdentity.GetCurrent().Name +
						"] does not have the required security permissions to access the Document Directory at: " + documentDirectory +
						"\r\n " + exc.Message + " in " + exc.Source;
					Console.WriteLine(ErrorLogMessage);
					return false;
					}
				catch(NotSupportedException exc)
					{
					ErrorLogMessage = "The path of Document Directory [" + documentDirectory + "] contains invalid characters." +
						" Ensure that the path is valid and consist of legible path characters only. \r\n " + exc.Message + " in " + exc.Source;
					Console.WriteLine(ErrorLogMessage);
					return false;
					}
				catch(DirectoryNotFoundException exc)
					{
					ErrorLogMessage = "The path of Document Directory [" + 
						documentDirectory + "] is invalid. Check that the drive is mapped and exist \r\n " + exc.Message + " in " + exc.Source;
					Console.WriteLine(ErrorLogMessage);
					return false;
					}
				}
			Console.WriteLine("\t\t\t The documentDirectory [" + documentDirectory + "] exist and is ready to be used.");
			// Set the object's LocalDocumentPath property
			this.LocalPath = documentDirectory;

			// Construct a name for the New Document/Workbook
			string docwbkFilename = DateTime.Now.ToShortDateString();
			docwbkFilename = docwbkFilename.Replace("/", "-") + "_" + DateTime.Now.ToLongTimeString();
			//Console.WriteLine("filename: [{0}]", documentFilename);
			docwbkFilename = docwbkFilename.Replace(":", "-");
			docwbkFilename = docwbkFilename.Replace(" ", "_");

			if(parDocumentOrWorkbook == enumDocumentOrWorkbook.Document)
				docwbkFilename = parDocumentType + "_" + docwbkFilename + ".docx";
			else
				docwbkFilename = parDocumentType + "_" + docwbkFilename + ".xlsx";

			Console.WriteLine("\t\t\t Document Or Workbook filename: [{0}]", docwbkFilename);
			// Set the object's Filename property
			this.Filename = docwbkFilename;

			// Create a new file based on a template.
			try
				{
				File.Copy(sourceFileName: templateDirectory + templateFileName, destFileName: documentDirectory + docwbkFilename, overwrite: true);
				}
			catch(FileNotFoundException exc)
				{
				ErrorLogMessage = "The template file: [" + templateDirectory + "\\" + templateFileName + "] does not exist. \r\n " + exc.Message + " in " + exc.Source;
				Console.WriteLine(ErrorLogMessage);
				return false;
				}
			catch(DirectoryNotFoundException exc)
				{
				ErrorLogMessage = "Either template or document directory could not be found. \r\n - Template Dir: [" + templateDirectory + "] " +
					"\r\n - Document Dir: [" + documentDirectory + "] \r\n " + exc.Message + " in " + exc.Source;
				Console.WriteLine(ErrorLogMessage);
				return false;
				}
			catch(UnauthorizedAccessException exc)
				{
				ErrorLogMessage = "The DocGenerator process doesn't have the required permissions to access the template or to create the new document. " +
					"\r\n " + exc.Message + " in " + exc.Source;
				Console.WriteLine(ErrorLogMessage);
				return false;
				}
			catch(IOException exc)
				{
				ErrorLogMessage = "An IO error occurred while attempting to copy the Template file for the new Document. \r\n " + exc.Message + " in " + exc.Source;
				Console.WriteLine(ErrorLogMessage);
				return false;
				}

			if(this.DocumentOrWorkbook == enumDocumentOrWorkbook.Document)
				{
				// Open the new Word document which is still in .dotx format to save it as a .docx file
				try
					{
					WordprocessingDocument objDocument = WordprocessingDocument.Open(path: documentDirectory + docwbkFilename, isEditable: true);
					// Change the document Type from .dotx to docx format.
					objDocument.ChangeDocumentType(newType: WordprocessingDocumentType.Document);
					objDocument.Close();
					}
				catch(OpenXmlPackageException exc)
					{
					ErrorLogMessage = "Unable to open new Document: [" + documentDirectory + "\\" + docwbkFilename + "] \r\n " + exc.Message + " in " + exc.Source;
					Console.WriteLine(ErrorLogMessage);
					return false;
					}
				}

			if(this.DocumentOrWorkbook == enumDocumentOrWorkbook.Workbook)
				{
				// Open the new Word document which is still in .xltx format to save it as a .xlsx file
				try
					{
					SpreadsheetDocument objWorksbook = SpreadsheetDocument.Open(path: documentDirectory + docwbkFilename, isEditable: true);
					// Change the document Type from .xltx to .xlsx format.
					objWorksbook.ChangeDocumentType(newType: SpreadsheetDocumentType.Workbook);
					objWorksbook.Close();
					}
				catch(OpenXmlPackageException exc)
					{
					ErrorLogMessage = "Unable to open new Workbook: [" + documentDirectory + "\\" + docwbkFilename + "] \r\n " + exc.Message + " in " + exc.Source;
					Console.WriteLine(ErrorLogMessage);
					return false;
					}
				}

			Console.WriteLine("\t\t\t Successfully created the new document: {0}", documentDirectory + docwbkFilename);
			// Set the object's DocumentURI property
			this.LocalURI = documentDirectory + docwbkFilename;
			return true;
			}
		} // end of oxmlDocumentWorkbook class

	public class oxmlDocument : oxmlDocumentWorkbook
		{
		// ----------------------
		//---Construct_Heading ---
		// ----------------------
		/// <summary>
		/// This method constructs a new Heading Paragraph which can be inserted into the Body object of the oXML document
		/// </summary>
		/// <param name="parHeadingLevel">
		/// Pass an integer between 1 and 9 depending of the level of the Heading that need to be inserted.
		/// </param>
		/// <param name="parBookMark">
		/// Optional Parameter. When a Bookmark must be created for the heading, pass the BookMark label 
		/// (without any spaces or odd characters) that need to be inserted as a string. By default the value is Null, 
		/// which means the heading will not contain a Bookmark. If a value is passed a Bookmark will be inserted.
		/// </param>
		public static Paragraph Construct_Heading(
			int parHeadingLevel,
			string parBookMark = null,
			bool parNoNumberedHeading = false)
			{
			if(parHeadingLevel < 1)
				parHeadingLevel = 1;
			else if(parHeadingLevel > 9)
				parHeadingLevel = 9;
			
			Paragraph objParagraph = new Paragraph();
			ParagraphProperties objParagraphProperties = new ParagraphProperties();
			ParagraphStyleId objParagraphStyleID = new ParagraphStyleId();
			if(parNoNumberedHeading)
				{
				objParagraphStyleID.Val = "DDHeadingNoNumber";
				}
			else
				{
				objParagraphStyleID.Val = "DDHeading" + parHeadingLevel.ToString();
				}
			objParagraphProperties.Append(objParagraphStyleID);
			objParagraph.Append(objParagraphProperties);
			if(parBookMark != null)
				{
				BookmarkStart objBookmarkStart = new BookmarkStart();
				objBookmarkStart.Name = parBookMark;
				string bookMarkID = parBookMark.Substring(parBookMark.IndexOf("_", 0) + 1, parBookMark.Length - parBookMark.IndexOf("_", 0) -1);
				objBookmarkStart.Id = bookMarkID;
				objParagraph.Append(objBookmarkStart);

				BookmarkEnd objBookmarkEnd = new BookmarkEnd();
				objBookmarkEnd.Id = bookMarkID;
				objParagraph.Append(objBookmarkEnd);
				}

			return objParagraph;
			}

		// -------------------------
		//--- Construct_Paragraph ---
		// -------------------------
		/// <summary>
		/// Use this method to create a new Paragraph object
		/// </summary>
		/// <param name="parBodyTextLevel">
		/// An optional parameter, default is 0, this parameter is used to determine the Style of the Paragraph
		/// </param>
		/// <param name="parIsTableParagraph">
		/// Pass boolean value of TRUE if the paragraph is going to be a table paragraph, else leave it blank because the default value is FALSE.
		/// </param>
		/// <returns>
		/// The paragraph object that can be inserted into a document.
		/// </returns>
		public static Paragraph Construct_Paragraph(
			int parBodyTextLevel = 0, 
			bool parIsTableParagraph = false)
			{

			if(parBodyTextLevel > 9)
				parBodyTextLevel = 9;

			//Construct a Paragraph instance.
			Paragraph objParagraph = new Paragraph();
			//Construct a ParagraphProperties object instance for the paragraph.
			ParagraphProperties objParagraphProperties = new ParagraphProperties();
			//Construct the ParagraphStyle to be used
			ParagraphStyleId objParagraphStyleID = new ParagraphStyleId();
			if(parIsTableParagraph)
				{
				objParagraphStyleID.Val = "DDTableBodyText";
				}
			else
				{
				objParagraphStyleID.Val = "DDBodyText" + parBodyTextLevel.ToString();
				}
			objParagraphProperties.Append(objParagraphStyleID);
			objParagraph.Append(objParagraphProperties);
			return objParagraph;
			}

		// -------------------------------------
		//--- Construct_BulletNumberParagraph ---
		// -------------------------------------
		/// <summary>
		/// Use this method to insert a new Bullet Text Paragraph
		/// </summary>
		/// <param name="parBulletLevel">
		/// Pass an integer between 0 and 9 depending of the level of the body text level that need to be inserted.
		/// </param>
		/// <param name="parIsTableBullet">
		///  Pass boolean value of TRUE if the paragraph is for a Table else leave blank because the default value is FALSE.
		/// </param>
		/// <param name="parIsBullet">
		/// Pass boolean FALSE if NOT a normal paragraph bullet - Default value is TRUE
		/// </param>
		/// <returns> Paragraph object</returns>
		public static Paragraph Construct_BulletNumberParagraph(
			int parBulletLevel, 
			bool parIsBullet = true, 
			bool parIsTableBullet = false)
			{
			if(parBulletLevel > 9)
				parBulletLevel = 9;

			//Create a new Paragraph instance
			Paragraph objParagraph = new Paragraph();
			ParagraphProperties objParagraphProperties = new ParagraphProperties();
			ParagraphStyleId objParagraphStyleID = new ParagraphStyleId();
			if(parIsTableBullet)
				{
				objParagraphStyleID.Val = "DDTableBullet";
				}
			else
				{
				if(parIsBullet == true)
					{
					objParagraphStyleID.Val = "DDBullet" + parBulletLevel.ToString();
					}
				else
					{
					objParagraphStyleID.Val = "DDNumber" + parBulletLevel.ToString();
					}
				}
			objParagraphProperties.Append(objParagraphStyleID);
			objParagraph.Append(objParagraphProperties);
			return objParagraph;
			}

		//--------------------------
		//---Construct_Error ---
		//--------------------------
		/// <summary>
		/// Use this method to insert a new Body Text Paragraph and highlights it in RED text 
		/// to indicate an error in the SharePoint Enahanced Rich Text.
		/// </summary>
		/// <param name="parText">
		/// Pass the text as a string which need to be inserted into the document. 
		/// </param>
		/// <returns>
		/// The paragraph object that is inserted into the Body object will be returned as a Paragraph object.
		/// </returns>
		public static Paragraph Construct_Error(string parText)
			{
			//Create a Paragraph instance.
			Paragraph objParagraph = new Paragraph();
			//Create a ParagraphProperties object instance for the paragraph.
			ParagraphProperties objParagraphProperties = new ParagraphProperties();
			ParagraphStyleId objParagraphStyleID = new ParagraphStyleId();
			objParagraphStyleID.Val = "DDContentError";
			objParagraphProperties.Append(objParagraphStyleID);
			objParagraph.Append(objParagraphProperties);
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
			objRun = oxmlDocument.Construct_RunText(parText);
			objParagraph.Append(objRun);
			return objParagraph;
			}

		//--------------------------
		//---Construct Caption   ---
		//--------------------------
		/// <summary>
		/// Use this method to insert a new Caption into the document for an Image or a Table.
		/// </summary>
		/// <param name="parCaptionType">
		/// Pass one of the following values: "Image" or "Table" to indicate whether an image or a table caption must be inserted.
		/// </param>
		/// <param name="parCaptionSequence">
		/// Pass an integer that will be number of the Image of Table to be inserted.
		/// </param>
		/// <param name="parCaptionText">
		/// Pass the text string that must be inserted as the Caption text.
		/// </param>
		/// <returns>
		/// The paragraph object that can be inserted into the Body object will be returned.
		/// </returns>
		public static Paragraph Construct_Caption(
			string parCaptionType,
			string parCaptionText)
			{
			//Create a Paragraph instance.
			Paragraph objParagraph = new Paragraph();
			// Create the Paragraph Properties instance.
			ParagraphProperties objParagraphProperties = new ParagraphProperties();
			ParagraphStyleId objParagraphStyleID = new ParagraphStyleId();
			if(parCaptionType == "Table")
				{ objParagraphStyleID.Val = "DDCaptionTable"; }
			else
				{ objParagraphStyleID.Val = "DDCaptionImage"; }
			objParagraphProperties.Append(objParagraphStyleID);
			//Append the ParagraphProerties to the Paragraph
			objParagraph.Append(objParagraphProperties);

			// Create the Caption Run Object
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
			DocumentFormat.OpenXml.Wordprocessing.Text objText = new DocumentFormat.OpenXml.Wordprocessing.Text();
			objText.Space = SpaceProcessingModeValues.Preserve;
			objText.Text = parCaptionText;
			objRun.Append(objText);
			objParagraph.Append(objRun);

			return objParagraph;
			}

		//------------------------
		//--- Construct_RunText ---
		//------------------------
		public static DocumentFormat.OpenXml.Wordprocessing.Run Construct_RunText(
				string parText2Write,
				bool parIsError = false,
				bool parIsNewSection = false,
				String parContentLayer = "None",
				bool parBold = false,
				bool parItalic = false,
				bool parUnderline = false,
				bool parSubscript = false,
				bool parSuperscript = false)
			{
			// Create a new Run object in the objParagraph
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();

			if(parIsNewSection)
				{
				LastRenderedPageBreak objLastRenderedPageBreak = new LastRenderedPageBreak();
				objRun.Append(objLastRenderedPageBreak);
				}
			else // if(!parIsNewSection)
				{
				// Create a Run Properties instance.

				DocumentFormat.OpenXml.Wordprocessing.RunProperties objRunProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
				// Insert the colour coding for Content Layering if applicable.
				if(parContentLayer != "None")
					{
					if(parContentLayer == "Layer1")
						{
						DocumentFormat.OpenXml.Wordprocessing.Color objLayer1Color = new DocumentFormat.OpenXml.Wordprocessing.Color();
						objLayer1Color.Val = Properties.AppResources.Layer1Color;
						objRunProperties.Append(objLayer1Color);
						}
					else if(parContentLayer == "Layer2")
						{
						DocumentFormat.OpenXml.Wordprocessing.Color objLayer2Color = new DocumentFormat.OpenXml.Wordprocessing.Color();
						objLayer2Color.Val = Properties.AppResources.Layer2Color;
						objRunProperties.Append(objLayer2Color);
						}
					else if(parContentLayer == "Layer3")
						{
						DocumentFormat.OpenXml.Wordprocessing.Color objLayer3Color = new DocumentFormat.OpenXml.Wordprocessing.Color();
						objLayer3Color.Val = Properties.AppResources.Layer3Color;
						objRunProperties.Append(objLayer3Color);
						}
					}

				if(parBold || parItalic || parUnderline || parSubscript || parSuperscript)
					{
					// Set the properties for the Run
					if(parBold)
						objRunProperties.Bold = new DocumentFormat.OpenXml.Wordprocessing.Bold();
					if(parItalic)
						objRunProperties.Italic = new DocumentFormat.OpenXml.Wordprocessing.Italic();
					if(parUnderline)
						objRunProperties.Underline = new DocumentFormat.OpenXml.Wordprocessing.Underline() { Val = DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single };
					if(parSubscript)
						{
						DocumentFormat.OpenXml.Wordprocessing.VerticalTextAlignment objVerticalTextAlignment = new DocumentFormat.OpenXml.Wordprocessing.VerticalTextAlignment();
						objVerticalTextAlignment.Val = VerticalPositionValues.Subscript;
						objRunProperties.Append(objVerticalTextAlignment);
						}
					if(parSuperscript)
						{
						DocumentFormat.OpenXml.Wordprocessing.VerticalTextAlignment objVerticalTextAlignment = new DocumentFormat.OpenXml.Wordprocessing.VerticalTextAlignment();
						objVerticalTextAlignment.Val = VerticalPositionValues.Superscript;
						objRunProperties.Append(objVerticalTextAlignment);
						}
					}
				if(parIsError)
					{
					DocumentFormat.OpenXml.Wordprocessing.Color objColorRed = new DocumentFormat.OpenXml.Wordprocessing.Color();
					objColorRed.Val = Properties.AppResources.ErrorTextColor;
					DocumentFormat.OpenXml.Wordprocessing.Underline objUnderline = new DocumentFormat.OpenXml.Wordprocessing.Underline();
					objUnderline.Val = DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Wave;
					objRunProperties.Append(objColorRed);
					objRunProperties.Append(objUnderline);
					}
				
				// Append the Run Properties to the Run object
				objRun.AppendChild(objRunProperties);
			} // if(parIsNewSection)
			// Insert the text in the objRun
			DocumentFormat.OpenXml.Wordprocessing.Text objText = new DocumentFormat.OpenXml.Wordprocessing.Text();
			objText.Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
			objText.Text = parText2Write;
			//Console.WriteLine("**** Text ****: {0} \tBold:{1} Italic:{2} Underline:{3}", objText.Text,parBold, parItalic, parUnderline);
			
			objRun.AppendChild(objText);
			return objRun;
			}

		//-------------------
		//--- InsertImage ---
		//-------------------
		/// <summary>
		/// 
		/// </summary>
		/// <param name="parMainDocumentPart"></param>
		/// <param name="parEffectivePageTWIPSwidth"></param>
		/// <param name="parEffectivePageTWIPSheight"></param>
		/// <param name="parParagraphLevel"></param>
		/// <param name="parPictureSeqNo"></param>
		/// <param name="parImageURL">
		/// Required String prameter. The location of the URL to the specific image.
		/// </param>
		/// <returns></returns>
		public static DocumentFormat.OpenXml.Wordprocessing.Run InsertImage(
			ref MainDocumentPart parMainDocumentPart, 
			UInt32 parEffectivePageTWIPSwidth,
			UInt32 parEffectivePageTWIPSheight,
			int parParagraphLevel, 
			int parPictureSeqNo, 
			string parImageURL)
			{
			if(parParagraphLevel < 1)
				parParagraphLevel = 1;
			else if(parParagraphLevel > 9)
				parParagraphLevel = 9;

			string ErrorLogMessage = "";
			string imageType = "";
			string relationshipID = "";
			string imageFileName = "";
			string imageDirectory = System.IO.Path.GetFullPath("\\") + DocGenerator.Properties.AppResources.LocalImagePath;
			try
				{
				
				// Download the image from SharePoint if it is a http:// based image
				imageType = parImageURL.Substring(parImageURL.LastIndexOf(".") + 1, (parImageURL.Length - parImageURL.LastIndexOf(".") - 1));
				if(parImageURL.IndexOf("\\") < 0)
					{
					ErrorLogMessage = "";
					//Derive the file name of the image file
					//Console.WriteLine(
					//"         1         2         3         4         5         6         7         " + 
					//"8         9        11        12        13        14        15\r\n" +
					//"1234567890123456789012345678901234567890123456789012345678901234567890123456789" +
					//"0123456789012345678901234567890123456789012345678901234567890 \r{0}", parImageURL);
					imageFileName = parImageURL.Substring(parImageURL.LastIndexOf("/") + 1, (parImageURL.Length - parImageURL.LastIndexOf("/")) - 1);
					// Construct the local name for the New Image file
					imageFileName = imageFileName.Replace("%20", "_");
					imageFileName = imageFileName.Replace(" ", "-");
					Console.WriteLine("\t\t\t local imageFileName: [{0}]", imageFileName);
					// Check if the DocGenerator Image Directory Exist and that it is accessable

					try
						{
						if(Directory.Exists(@imageDirectory))
							{
							Console.WriteLine("\t\t\t The imageDirectory [" + imageDirectory + "] exist and are ready to be used.");
							}
						else
							{
							DirectoryInfo templateDirInfo = Directory.CreateDirectory(@imageDirectory);
							Console.WriteLine("\t\t\t The imageDirectory [" + imageDirectory + "] was created and are ready to be used.");
							}
						}
					catch(UnauthorizedAccessException exc)
						{
						ErrorLogMessage = "The current user: [" + System.Security.Principal.WindowsIdentity.GetCurrent().Name +
						"] does not have the required security permissions to access the template directory at: " + imageDirectory +
						"\r\n " + exc.Message + " in " + exc.Source;
						Console.WriteLine("\t\t\t" + ErrorLogMessage);
						//TODO: insert code to write an error line in the document
						return null;
						}
					catch(NotSupportedException exc)
						{
						ErrorLogMessage = "The path of template directory [" + imageDirectory + "] contains invalid characters. Ensure that the path is valid and  contains legible path characters only. \r\n " + exc.Message + " in " + exc.Source;
						Console.WriteLine("\t\t\t" + ErrorLogMessage);
						//TODO: insert code to write an error line in the document
						return null;
						}
					catch(DirectoryNotFoundException exc)
						{
						ErrorLogMessage = "The path of template directory [" + imageDirectory + "] is invalid. Check that the drive is mapped and exist /r/n " + exc.Message + " in " + exc.Source;
						Console.WriteLine("\t\t\t" + ErrorLogMessage);
						//TODO: insert code to write an error line in the document
						return null;
						}

					// Check if the Image file already exist in the local Image directory
					if(File.Exists(imageDirectory + "\\" + imageFileName))
						{
						// If the the image file exist just proceed...
						Console.WriteLine("\t\t\t The image to already exist, just use it:" + imageDirectory + "\\" + imageFileName);
						}
					else // If the image doesn't exist already, then download it...
						{
						// Download the relevant image from SharePoint
						WebClient objWebClient = new WebClient();
						objWebClient.UseDefaultCredentials = true;
						//objWebClient.Credentials = CredentialCache.DefaultCredentials;
						try
							{
							objWebClient.DownloadFile(parImageURL, imageDirectory + "\\" + imageFileName);
							}
						catch(WebException exc)
							{
							ErrorLogMessage = "The template file could not be downloaded from SharePoint List [" + parImageURL + "]. " +
								"\n - Check that the template exist in SharePoint \n - that it is accessible \n - " +
								"and that the network connection is working. \n " + exc.Message + " in " + exc.Source;
							Console.WriteLine("\t\t\t" + ErrorLogMessage);
							return null;
							}
						}

					Console.WriteLine("\t\t\t {2} this Image:[{0}] exist in this directory:[{1}]", imageFileName, imageDirectory, File.Exists(imageDirectory + "\\" + imageFileName));
					parImageURL = imageDirectory + imageFileName;
					}
				else //if(parImageURL.IndexOf("/") > 0) // if it is a local file (not an URL...)
					{
					imageFileName = parImageURL.Substring(parImageURL.LastIndexOf("\\") + 1, (parImageURL.Length - parImageURL.LastIndexOf("\\") - 1));
					}

				var img = System.Drawing.Image.FromFile(parImageURL);
				//https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
				int imagePIXELheight = img.Height;
				int imagePIXELwidth = img.Width;

				Console.WriteLine("Image dimensions (H x W): {0} x {1} pixels per Inch", imagePIXELheight, imagePIXELwidth);
				Console.WriteLine("Horizontal Resolution...: {0} pixels per inch", img.HorizontalResolution);

				img.Dispose(); 
				img = null;

				// Load the image into the Media section of the Document and store the relaionshipID in the variable.
				// Insert the image into the MainDocumentPartdocument 
				switch(imageType)
					{
					case "JPG":
					case "jpg":
							{
							ImagePart objImagePart = parMainDocumentPart.AddImagePart(ImagePartType.Jpeg);
							using(FileStream objFileStream = new FileStream(path: parImageURL, mode: FileMode.Open))
								{
								objImagePart.FeedData(objFileStream);
								}
							relationshipID = parMainDocumentPart.GetIdOfPart(part: objImagePart);
							break;
							}
					case "GIF":
					case "gif":
							{
							ImagePart objImagePart = parMainDocumentPart.AddImagePart(ImagePartType.Gif);
							using(FileStream objFileStream = new FileStream(path: parImageURL, mode: FileMode.Open))
								{
								objImagePart.FeedData(objFileStream);
								}
							relationshipID = parMainDocumentPart.GetIdOfPart(part: objImagePart);
							break;
							}
					case "BMP":
					case "bmp":
							{
							ImagePart objImagePart = parMainDocumentPart.AddImagePart(ImagePartType.Bmp);
							using(FileStream objFileStream = new FileStream(path: parImageURL, mode: FileMode.Open))
								{
								objImagePart.FeedData(objFileStream);
								}
							relationshipID = parMainDocumentPart.GetIdOfPart(part: objImagePart);
							break;
							}
					case "PNG":
					case "png":
							{
							ImagePart objImagePart = parMainDocumentPart.AddImagePart(ImagePartType.Png);
							using(FileStream objFileStream = new FileStream(path: parImageURL, mode: FileMode.Open))
								{
								objImagePart.FeedData(objFileStream);
								}
							relationshipID = parMainDocumentPart.GetIdOfPart(part: objImagePart);
							break;
							}
					case "TIFF":
					case "tiff":
							{
							ImagePart objImagePart = parMainDocumentPart.AddImagePart(ImagePartType.Tiff);
							using(FileStream objFileStream = new FileStream(path: parImageURL, mode: FileMode.Open))
								{
								objImagePart.FeedData(objFileStream);
								}
							relationshipID = parMainDocumentPart.GetIdOfPart(part: objImagePart);
							break;
							}
					default:
							{
							break;
							}
					}

				// Define the Drawing Object instance
				DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
				// Define the Anchor object
				DrwWp.Anchor objAnchor = new DrwWp.Anchor();
				objAnchor.DistanceFromTop = (UInt32Value)57150U;
				objAnchor.DistanceFromBottom = (UInt32Value)57150U;
				objAnchor.DistanceFromLeft = (UInt32Value) 0U;
				objAnchor.DistanceFromRight = (UInt32Value) 0U;
				objAnchor.RelativeHeight = (UInt32Value) 0U; 
				objAnchor.SimplePos = false;
				objAnchor.BehindDoc = true;
				objAnchor.Locked = true;
				objAnchor.LayoutInCell = false;
				objAnchor.AllowOverlap = false;

				// Define the Simple Position of the image.
				DrwWp.SimplePosition objSimplePosition = new DrwWp.SimplePosition();
				objSimplePosition.X = 0L;
				objSimplePosition.Y = 0L;
				objAnchor.Append(objSimplePosition);

				//Define the Horizontal Position
				DrwWp.HorizontalPosition objHorizontalPosition = new DrwWp.HorizontalPosition();
				objHorizontalPosition.RelativeFrom = DrwWp.HorizontalRelativePositionValues.Margin;
				// for flush Left Margin alignment
				//DrwWp.HorizontalAlignment objHorizontalAlignment = new DrwWp.HorizontalAlignment();
				//objHorizontalAlignment.Text = "left";
				//objHorizontalPosition.Append(objHorizontalAlignment);
				// for Left indentation
				DrwWp.PositionOffset objHorizontalPositionOffset = new DrwWp.PositionOffset();
				objHorizontalPositionOffset.Text = Properties.AppResources.Document_Image_Left_Indent;
				objHorizontalPosition.Append(objHorizontalPositionOffset);
				objAnchor.Append(objHorizontalPosition);

				// Define the Vertical Position
				DrwWp.VerticalPosition objVerticalPosition = new DrwWp.VerticalPosition();
				objVerticalPosition.RelativeFrom = DrwWp.VerticalRelativePositionValues.Paragraph;
				DrwWp.PositionOffset objVerticalPositionOffset = new DrwWp.PositionOffset();
				objVerticalPositionOffset.Text = "0";
				objVerticalPosition.Append(objVerticalPositionOffset);
				objAnchor.Append(objVerticalPosition);

				// Define the Extent for the image (Canvas)
				//If the image is wider than the Effective Width of the page
				double imageDXAwidth = 0;
				double imageDXAheight = 0;
				if((imagePIXELwidth * 20) > parEffectivePageTWIPSwidth)
					{
					imageDXAwidth = (((parEffectivePageTWIPSwidth / (imagePIXELwidth * 20D)) * imagePIXELwidth) * 20D) * 635D;
					imageDXAheight = (((parEffectivePageTWIPSwidth / (imagePIXELwidth * 20D)) * imagePIXELheight) * 20D) * 635D;
					}
				else if((imageDXAheight * 20) > parEffectivePageTWIPSheight)
					{
					imageDXAwidth = (((parEffectivePageTWIPSheight / (imagePIXELheight * 20D)) * imagePIXELwidth) * 20D) * 635D;
					imageDXAheight = (((parEffectivePageTWIPSheight / (imagePIXELheight * 20D)) * imagePIXELheight) * 20D) * 635D;
					}
				else
					{
					imageDXAwidth = imagePIXELwidth * 635D;
					imageDXAheight = imagePIXELheight * 635D;
					}
				Console.WriteLine("imageDXAwidth: {0}", imageDXAwidth);
				Console.Write(" imageDXAheight: {0}", imageDXAheight);

				DrwWp.Extent objExtent = new DrwWp.Extent();
				objExtent.Cx = Convert.ToInt64(imageDXAwidth);
				objExtent.Cy = Convert.ToInt64(imageDXAheight);
				objAnchor.Append(objExtent);

				// Define Extent Effects
				DrwWp.EffectExtent objEffectExtent = new DrwWp.EffectExtent();
				objEffectExtent.LeftEdge = 0L;
				objEffectExtent.TopEdge = 0L;
				objEffectExtent.RightEdge = 9525L;
				objEffectExtent.BottomEdge = 9525L;
				objAnchor.Append(objEffectExtent);

				// Define how text is wrapped around the image
				DrwWp.WrapTopBottom objWrapTopBottom = new DrwWp.WrapTopBottom();
				objAnchor.Append(objWrapTopBottom);

				// Define the Document Properties by linking the image to identifier of the imaged where it was inserted in the MainDocumentPart.
				DrwWp.DocProperties objDocProperties = new DrwWp.DocProperties();
				objDocProperties.Id = Convert.ToUInt32(parPictureSeqNo);
				objDocProperties.Name = "Picture " + parPictureSeqNo.ToString();
				objAnchor.Append(objDocProperties);

				// Define the Graphic Frame for the image
				DrwWp.NonVisualGraphicFrameDrawingProperties objNonVisualGraphicFrameDrawingProperties = new DrwWp.NonVisualGraphicFrameDrawingProperties();
				objAnchor.Append(objNonVisualGraphicFrameDrawingProperties);
				Drw.GraphicFrameLocks objGraphicFrameLocks = new Drw.GraphicFrameLocks();
				objGraphicFrameLocks.NoChangeAspect = true;
				objGraphicFrameLocks.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
				objNonVisualGraphicFrameDrawingProperties.Append(objGraphicFrameLocks);

				// Configure the graphic
				Drw.Graphic objGraphic = new Drw.Graphic();
				objGraphic.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
				Drw.GraphicData objGraphicData = new Drw.GraphicData();
				objGraphicData.Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture";

				// Define the Picture
				Pic.Picture objPicture = new Pic.Picture();
				objPicture.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
				// Define the Picture's NonVisual Properties
				Pic.NonVisualPictureProperties objNonVisualPictureProperties = new Pic.NonVisualPictureProperties();
				// Define the NonVisual Drawing Properties
				Pic.NonVisualDrawingProperties objNonVisualDrawingProperties = new Pic.NonVisualDrawingProperties();
				objNonVisualDrawingProperties.Id = Convert.ToUInt32(parPictureSeqNo);
				objNonVisualDrawingProperties.Name = imageFileName;
				// Define the Picture's NonVisual Picture Drawing Properties
				Pic.NonVisualPictureDrawingProperties objNonVisualPictureDrawingProperties = new Pic.NonVisualPictureDrawingProperties();
				objNonVisualPictureProperties.Append(objNonVisualDrawingProperties);
				objNonVisualPictureProperties.Append(objNonVisualPictureDrawingProperties);

				// Define the Blib
				Drw.Blip objBlip = new Drw.Blip();
				objBlip.Embed = relationshipID;
				Drw.BlipExtensionList objBlipExtensionList = new Drw.BlipExtensionList();
				Drw.BlipExtension objBlipExtension = new Drw.BlipExtension();
				objBlipExtension.Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}";
				Drw2010.UseLocalDpi objUseLocalDpi = new Drw2010.UseLocalDpi();
				objUseLocalDpi.Val = false;
				objUseLocalDpi.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
				objBlipExtension.Append(objUseLocalDpi);
				objBlipExtensionList.Append(objBlipExtension);
				objBlip.Append(objBlipExtensionList);

				// Define how the image is filled
				Drw.Stretch objStretch = new Drw.Stretch();
				Drw.FillRectangle objFillRectangle = new Drw.FillRectangle();
				objStretch.Append(objFillRectangle);
				Pic.BlipFill objBlipFill = new Pic.BlipFill();
				objBlipFill.Append(objBlip);
				objBlipFill.Append(objStretch);

				// Define the Picture's Shape Properties
				Pic.ShapeProperties objShapeProperties = new Pic.ShapeProperties();
				Drw.Transform2D objTransform2D = new Drw.Transform2D();
				Drw.Offset objOffset = new Drw.Offset();
				objOffset.X = 0L;
				objOffset.Y = 0L;
				Drw.Extents objExtents = new Drw.Extents();
				objExtents.Cx = Convert.ToInt64(imageDXAwidth);
				objExtents.Cy = Convert.ToInt64(imageDXAheight);
				objTransform2D.Append(objOffset);
				objTransform2D.Append(objExtents);

				// Define the Preset Geometry
				Drw.PresetGeometry objPresetGeometry = new Drw.PresetGeometry();
				objPresetGeometry.Preset = Drw.ShapeTypeValues.Rectangle;
				Drw.AdjustValueList objAdjustValueList = new Drw.AdjustValueList();
				objPresetGeometry.Append(objAdjustValueList);
				objShapeProperties.Append(objTransform2D);
				objShapeProperties.Append(objPresetGeometry);

				// Append the Definitions to the Picture Object Instance...
				objPicture.Append(objNonVisualPictureProperties);
				objPicture.Append(objBlipFill);
				objPicture.Append(objShapeProperties);

				// Append the the picture object to the Graphic object Instance
				objGraphicData.Append(objPicture);
				objGraphic.Append(objGraphicData);
				objAnchor.Append(objGraphic);

				// Define the drawings relative width
				DrwWp2010.RelativeWidth objRelativeWidth = new DrwWp2010.RelativeWidth();
				objRelativeWidth.ObjectId = DrwWp2010.SizeRelativeHorizontallyValues.InsideMargin;
				DrwWp2010.PercentageWidth objPercentageWidth = new DrwWp2010.PercentageWidth();
				objPercentageWidth.Text = "0";
				objRelativeWidth.Append(objPercentageWidth);
				objAnchor.Append(objRelativeWidth);

				// Define the drawings relative Height
				DrwWp2010.RelativeHeight objRelativeHeight = new DrwWp2010.RelativeHeight();
				objRelativeHeight.RelativeFrom = DrwWp2010.SizeRelativeVerticallyValues.InsideMargin;
				DrwWp2010.PercentageHeight objPercentageHeight = new DrwWp2010.PercentageHeight();
				objPercentageHeight.Text = "0";
				objRelativeHeight.Append(objPercentageHeight);
				objAnchor.Append(objRelativeHeight);
				
				// Append the Anchor object to the Drawing object...
				objDrawing.Append(objAnchor);

				// Define the Run object and append the Drawing object to it...
				DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
				objRun.Append(objDrawing);
				// Return the Run object which now contains the complete Image to be added to a Paragraph in the document.
				return objRun;
				}
			catch(Exception exc)
				{
				ErrorLogMessage = "The image file: [" + parImageURL + "] couldn't be located and was not inserted. \r\n " + exc.Message + " in " + exc.Source;
				Console.WriteLine(ErrorLogMessage);
				return null;
				}
			}


		/// <summary>
		/// This method inserts the image (defined) in the resource file, into the MainDocumentPart and returnes the RelationshipID.
		/// </summary>
		/// <param name="parMainDocumentPart">
		/// Pass the MainDocumentPart as an object by reference.
		/// </param>
		/// <returns>
		/// The actual Relationship ID where the image was inserted in the MainDocumentPart, is returned as a string. 
		/// If the image could not be added to the MaindocumentPart, an "ERROR:..." is returned instead of the Relationship ID.
		/// </returns>
		//----------------------------
		//--- InsertHyperlinkImage ---
		//----------------------------
		public static string InsertHyperlinkImage(
			ref MainDocumentPart parMainDocumentPart)
			{
			string ErrorLogMessage = "";
			string relationshipID = "";

			try
				{
				// Insert the image into the MainDocumentPart 
				Assembly objAssembly = Assembly.GetExecutingAssembly();
				Console.WriteLine("Assembly.Location: {0}", objAssembly.Location);
				Console.WriteLine("Directory: {0}", objAssembly.Location.Substring(
					startIndex: 0,length: objAssembly.Location.LastIndexOf("\\")+1) + Properties.AppResources.ClickLinkImageURL);
				ImagePart objImagePart = parMainDocumentPart.AddImagePart(ImagePartType.Png);
				string hyperlinkImageURL = objAssembly.Location.Substring(
					startIndex: 0, length: objAssembly.Location.LastIndexOf("\\") + 1) + Properties.AppResources.ClickLinkImageURL;

				using(FileStream objFileStream = new FileStream(path: hyperlinkImageURL, mode: FileMode.Open))
					{
					objImagePart.FeedData(objFileStream);
					}
				relationshipID = parMainDocumentPart.GetIdOfPart(part: objImagePart);

				return relationshipID;
				}
			catch(Exception exc)
				{
				ErrorLogMessage = "The image file: [" + Properties.AppResources.ClickLinkImageURL + "] couldn't be located and was not inserted. \r\n " + exc.Message + " in " + exc.Source;
				Console.WriteLine(ErrorLogMessage);
				return null;
				}
			}

		//------------------------------------
		// --- ConstructClickLinkHyperlink ---
		// -----------------------------------
		public static DocumentFormat.OpenXml.Wordprocessing.Drawing ConstructClickLinkHyperlink(
			ref MainDocumentPart parMainDocumentPart,
			string parImageRelationshipId,
			string parClickLinkURL,
			int parHyperlinkID)
			{

			Uri objUri = new Uri(parClickLinkURL);
			string hyperlinkID = "";
			// Check if the hyperlink already exist in the document
			foreach(HyperlinkRelationship hyperRelationship in parMainDocumentPart.HyperlinkRelationships)
				{
				if(hyperRelationship.Uri == objUri)
					{
					//urlExist = true;
					hyperlinkID = hyperRelationship.Id;
					}
				}
			// If no matching hyperlikID was found, add a new Hyperlink to the MainDocumentPart.
			if(hyperlinkID == "")
				{
				HyperlinkRelationship objHyperlinkRelationship = parMainDocumentPart.AddHyperlinkRelationship(hyperlinkUri: objUri, isExternal: true);
				hyperlinkID = objHyperlinkRelationship.Id;
				}
			
			// Define a Drawing Object instance
			DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
			// Define the Anchor object
			DrwWp.Anchor objAnchor = new DrwWp.Anchor();
			objAnchor.DistanceFromTop = (UInt32Value) 0U;
			objAnchor.DistanceFromBottom = (UInt32Value) 0U;
			objAnchor.DistanceFromLeft = (UInt32Value) 114300U;
			objAnchor.DistanceFromRight = (UInt32Value) 114300U;
			objAnchor.RelativeHeight = (UInt32Value) 251659264U;
			objAnchor.SimplePos = false;
			objAnchor.BehindDoc = false;
			objAnchor.Locked = true;
			objAnchor.LayoutInCell = true;
			objAnchor.AllowOverlap = true;

			// Define the Simple Position of the image.
			DrwWp.SimplePosition objSimplePosition = new DrwWp.SimplePosition();
			objSimplePosition.X = 0L;
			objSimplePosition.Y = 0L;
			objAnchor.Append(objSimplePosition);

			//Define the Horizontal Position
			DrwWp.HorizontalPosition objHorizontalPosition = new DrwWp.HorizontalPosition();
			objHorizontalPosition.RelativeFrom = DrwWp.HorizontalRelativePositionValues.LeftMargin;
			DrwWp.PositionOffset objHorizontalPositionOffSet = new DrwWp.PositionOffset();
			objHorizontalPositionOffSet.Text = "238125";
			objHorizontalPosition.Append(objHorizontalPositionOffSet);
			objAnchor.Append(objHorizontalPosition);

			// Define the Vertical Position
			DrwWp.VerticalPosition objVerticalPosition = new DrwWp.VerticalPosition();
			objVerticalPosition.RelativeFrom = DrwWp.VerticalRelativePositionValues.Line;
			DrwWp.PositionOffset objVerticalPositionOffset = new DrwWp.PositionOffset();
			objVerticalPositionOffset.Text = "114300";
			objVerticalPosition.Append(objVerticalPositionOffset);
			objAnchor.Append(objVerticalPosition);

			// Define the Canvas or Extent in which the the image will be placed
			DrwWp.Extent objExtent = new DrwWp.Extent();
			objExtent.Cx = 180975L;
			objExtent.Cy = 123825L;
			objAnchor.Append(objExtent);
			// Define Extent Effects
			DrwWp.EffectExtent objEffectExtent = new DrwWp.EffectExtent();
			objEffectExtent.LeftEdge = 0L;
			objEffectExtent.TopEdge = 0L;
			objEffectExtent.RightEdge = 9525L;
			objEffectExtent.BottomEdge = 9525L;
			objAnchor.Append(objEffectExtent);

			// Define how text is wrapped around the image
			DrwWp.WrapNone objWrapNone = new DrwWp.WrapNone();
			objAnchor.Append(objWrapNone);

			// Define the Document Properties by linking the image to identifier of the image where it was inserted in the MainDocumentPart.
			DrwWp.DocProperties objDocProperties = new DrwWp.DocProperties();
			objDocProperties.Id = Convert.ToUInt32(parHyperlinkID);
			objDocProperties.Name = "ClickLink " + parHyperlinkID;
			
			// Define the Hyperlink to be added
			Drw.HyperlinkOnClick objHyperlinkOnClick = new Drw.HyperlinkOnClick();
			objHyperlinkOnClick.Id = hyperlinkID;
			objHyperlinkOnClick.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
			objDocProperties.Append(objHyperlinkOnClick);
			objAnchor.Append(objDocProperties);

			// Define the Graphic Frame for the image
			DrwWp.NonVisualGraphicFrameDrawingProperties objNonVisualGraphicFrameDrawingProperties = new DrwWp.NonVisualGraphicFrameDrawingProperties();
			objAnchor.Append(objNonVisualGraphicFrameDrawingProperties);
			Drw.GraphicFrameLocks objGraphicFrameLocks = new Drw.GraphicFrameLocks();
			objGraphicFrameLocks.NoChangeAspect = true;
			objGraphicFrameLocks.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
			objNonVisualGraphicFrameDrawingProperties.Append(objGraphicFrameLocks);

			// Configure the graphic
			Drw.Graphic objGraphic = new Drw.Graphic();
			objGraphic.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
			Drw.GraphicData objGraphicData = new Drw.GraphicData();
			objGraphicData.Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture";

			// Define the Picture
			Pic.Picture objPicture = new Pic.Picture();
			objPicture.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
			// Define the Picture's NonVisual Properties
			Pic.NonVisualPictureProperties objNonVisualPictureProperties = new Pic.NonVisualPictureProperties();
			// Define the NonVisual Drawing Properties
			Pic.NonVisualDrawingProperties objNonVisualDrawingProperties = new Pic.NonVisualDrawingProperties();
			objNonVisualDrawingProperties.Id = Convert.ToUInt32(0);
			objNonVisualDrawingProperties.Name = Properties.AppResources.ClickLinkFileName;
			// Define the Picture's NonVisual Picture Drawing Properties
			Pic.NonVisualPictureDrawingProperties objNonVisualPictureDrawingProperties = new Pic.NonVisualPictureDrawingProperties();
			objNonVisualPictureProperties.Append(objNonVisualDrawingProperties);
			objNonVisualPictureProperties.Append(objNonVisualPictureDrawingProperties);

			// Define the Blib
			Drw.Blip objBlip = new Drw.Blip();
			objBlip.Embed = parImageRelationshipId;
			Drw.BlipExtensionList objBlipExtensionList = new Drw.BlipExtensionList();
			Drw.BlipExtension objBlipExtension = new Drw.BlipExtension();
			objBlipExtension.Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}";
			Drw2010.UseLocalDpi objUseLocalDpi = new Drw2010.UseLocalDpi();
			objUseLocalDpi.Val = false;
			objUseLocalDpi.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
			objBlipExtension.Append(objUseLocalDpi);
			objBlipExtensionList.Append(objBlipExtension);
			objBlip.Append(objBlipExtensionList);

			// Define how the image is filled
			Drw.Stretch objStretch = new Drw.Stretch();
			Drw.FillRectangle objFillRectangle = new Drw.FillRectangle();
			objStretch.Append(objFillRectangle);

			Pic.BlipFill objBlipFill = new Pic.BlipFill();
			objBlipFill.Append(objBlip);
			objBlipFill.Append(objStretch);

			// Define the Picture's Shape Properties
			Pic.ShapeProperties objShapeProperties = new Pic.ShapeProperties();
			Drw.Transform2D objTransform2D = new Drw.Transform2D();
			Drw.Offset objOffset = new Drw.Offset();
			objOffset.X = 0L;
			objOffset.Y = 0L;
			Drw.Extents objExtents = new Drw.Extents();
			objExtents.Cx = 180975L;
			objExtents.Cy = 123825L;
			objTransform2D.Append(objOffset);
			objTransform2D.Append(objExtents);
			// Define the Preset Geometry
			Drw.PresetGeometry objPresetGeometry = new Drw.PresetGeometry();
			objPresetGeometry.Preset = Drw.ShapeTypeValues.Rectangle;
			Drw.AdjustValueList objAdjustValueList = new Drw.AdjustValueList();
			objPresetGeometry.Append(objAdjustValueList);
			objShapeProperties.Append(objTransform2D);
			objShapeProperties.Append(objPresetGeometry);

			// Append the Definitions to the Picture Object Instance...
			objPicture.Append(objNonVisualPictureProperties);
			objPicture.Append(objBlipFill);
			objPicture.Append(objShapeProperties);

			// Append the the picture object to the Graphic object Instance
			objGraphicData.Append(objPicture);
			objGraphic.Append(objGraphicData);
			objAnchor.Append(objGraphic);

			// Define the drawings relative width
			DrwWp2010.RelativeWidth objRelativeWidth = new DrwWp2010.RelativeWidth();
			objRelativeWidth.ObjectId = DrwWp2010.SizeRelativeHorizontallyValues.Margin;
			DrwWp2010.PercentageWidth objPercentageWidth = new DrwWp2010.PercentageWidth();
			objPercentageWidth.Text = "0";
			objRelativeWidth.Append(objPercentageWidth);
			objAnchor.Append(objRelativeWidth);

			// Define the drawings relative Height
			DrwWp2010.RelativeHeight objRelativeHeight = new DrwWp2010.RelativeHeight();
			objRelativeHeight.RelativeFrom = DrwWp2010.SizeRelativeVerticallyValues.Margin;
			DrwWp2010.PercentageHeight objPercentageHeight = new DrwWp2010.PercentageHeight();
			objPercentageHeight.Text = "0";
			objRelativeHeight.Append(objPercentageHeight);
			objAnchor.Append(objRelativeHeight);

			// Append the Anchor object to the Drawing object...
			objDrawing.Append(objAnchor);

			// Return the Run object which now contains the complete Image to be added to a Paragraph in the document.
			return objDrawing;

			}

		//------------------------------------
		// --- Construct_BookmarkHyperlink ---
		// -----------------------------------
		public static DocumentFormat.OpenXml.Wordprocessing.Paragraph Construct_BookmarkHyperlink(
			int parBodyTextLevel,
			string parBookmarkValue)
			{
			// Create the object instances for the ParagraphProperties
			DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties objParagraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties();
			DocumentFormat.OpenXml.Wordprocessing.ParagraphStyleId objParagraPhStyleID = new DocumentFormat.OpenXml.Wordprocessing.ParagraphStyleId();
			objParagraPhStyleID.Val = "DDBodyText" + parBodyTextLevel;
			objParagraphProperties.Append(objParagraPhStyleID);

			// Create the object instances for the Hyperlink.
			DocumentFormat.OpenXml.Wordprocessing.Hyperlink objHyperlink = new DocumentFormat.OpenXml.Wordprocessing.Hyperlink();
			objHyperlink.History = true;
			objHyperlink.Anchor = parBookmarkValue; // use the Bookmark Parameter as the Anchor for the hyperlink.
			// Create object instances for the RunProperties
			DocumentFormat.OpenXml.Wordprocessing.RunProperties objRunProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
			DocumentFormat.OpenXml.Wordprocessing.RunStyle objRunStyle = new DocumentFormat.OpenXml.Wordprocessing.RunStyle();
			objRunStyle.Val = "Hyperlink";
			Spacing objSpacing = new Spacing();
			objSpacing.Val = 14;
			objRunProperties.Append(objRunStyle);
			objRunProperties.Append(objSpacing);

			// Create the object instances for the Text in the Hyperlink.
			DocumentFormat.OpenXml.Wordprocessing.Text objText = new DocumentFormat.OpenXml.Wordprocessing.Text();
			objText.Text = Properties.AppResources.Document_DRM_ClickHere;
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
			objRun.Append(objRunProperties);
			objRun.Append(objText);
			objHyperlink.Append(objRun);

			DocumentFormat.OpenXml.Wordprocessing.Run objRun2 = new DocumentFormat.OpenXml.Wordprocessing.Run();
			DocumentFormat.OpenXml.Wordprocessing.Text objText2 = new DocumentFormat.OpenXml.Wordprocessing.Text();
			objText2.Space = SpaceProcessingModeValues.Preserve;
			objText2.Text = Properties.AppResources.Document_DRM_Navigate_To_Detail;
			objRun2.Append(objText2);

			// Construct the Paragraph
			DocumentFormat.OpenXml.Wordprocessing.Paragraph objParagraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
			objParagraph.Append(objParagraphProperties);
			objParagraph.Append(objHyperlink);
			objParagraph.Append(objRun2);

			// Return the Paragraph object which now contains the complete Hyperlink text.
			return objParagraph;

			}

		//----------------------
		//--- ConstructTable ---
		//----------------------
		/// <summary>
		/// 
		/// </summary>
		/// <param name="parPageWidth">
		/// parameter value is the percentage of the available page width. If greater than 100 it will be set to 100% if less than 10 it will be set to 10%
		/// </param>
		/// <param name="parFirstColumn"></param>
		/// <param name="parLastColumn"></param>
		/// <param name="parFirstRow"></param>
		/// <param name="parLastRow"></param>
		/// <param name="parNoVerticalBand"></param>
		/// <param name="parNoHorizontalBand"></param>
		/// <returns></returns>
		public static DocumentFormat.OpenXml.Wordprocessing.Table ConstructTable(
			UInt32  parPageWidth,
			bool parFirstColumn = false, 
			bool parLastColumn = false,  
			bool parFirstRow = false, 
			bool parLastRow = false,
			bool parNoVerticalBand = true,
			bool parNoHorizontalBand = false)
			{
			
			// Creates a Table instance
			DocumentFormat.OpenXml.Wordprocessing.Table objTable = new DocumentFormat.OpenXml.Wordprocessing.Table();
			// Create and set the Table Properties instance
			TableProperties objTableProperties = new TableProperties();
			// Create and add the Table Style
			DocumentFormat.OpenXml.Wordprocessing.TableStyle objTableStyle = new DocumentFormat.OpenXml.Wordprocessing.TableStyle();
			objTableStyle.Val = "DDGreenHeaderTable";
			objTableProperties.Append(objTableStyle);
			// Define and add the table width
			TableWidth objTableWidth = new TableWidth();
			if(parPageWidth == 0)
				{
				objTableWidth.Width = "0";
				objTableWidth.Type = TableWidthUnitValues.Auto;
				}
			else
				{
				// Subtract the static Left Indent value from the page width
				objTableWidth.Width = parPageWidth.ToString();
				objTableWidth.Type = TableWidthUnitValues.Dxa;
				}
			objTableProperties.Append(objTableWidth);
			
			// Define the Table Indentation
			TableIndentation objTableIndentation = new TableIndentation();
			objTableIndentation.Width = Convert.ToInt32(Properties.AppResources.Document_Table_Left_Indent);
			objTableIndentation.Type = TableWidthUnitValues.Dxa;
               objTableProperties.Append(objTableIndentation);
			
			// Define the Table Layout
			TableLayout objTableLayout = new TableLayout();
			objTableLayout.Type = TableLayoutValues.Fixed;
			objTableProperties.Append(objTableLayout);

			// Define the TableCalMargins
			TableCellMarginDefault objTableCellMarginDefault = new TableCellMarginDefault();
			TopMargin objTopMargin = new TopMargin();
			objTopMargin.Width = "15";
			objTopMargin.Type = TableWidthUnitValues.Dxa;
			BottomMargin objBottomMargin = new BottomMargin();
			objBottomMargin.Width = "15";
			objBottomMargin.Type = TableWidthUnitValues.Dxa;
			TableCellLeftMargin objTableCellLeftMargin = new TableCellLeftMargin();
			objTableCellLeftMargin.Width = 60;
			objTableCellLeftMargin.Type = TableWidthValues.Dxa;
			TableCellRightMargin objTableCellRightMargin = new TableCellRightMargin();
			objTableCellRightMargin.Width = 60;
			objTableCellRightMargin.Type = TableWidthValues.Dxa;
			objTableCellMarginDefault.Append(objTopMargin);
			objTableCellMarginDefault.Append(objTableCellLeftMargin);
			objTableCellMarginDefault.Append(objBottomMargin);
			objTableCellMarginDefault.Append(objTableCellRightMargin);
			objTableProperties.Append(objTableCellMarginDefault);

			// Define and add the Table Justification
			//TableJustification objTableJustification = new TableJustification();
			//objTableJustification.Val = TableRowAlignmentValues.Left;
			//objTableProperties.Append(objTableJustification);

			// Define and add the Table Look
			TableLook objTableLook = new TableLook()
				{Val = "0600",
                    FirstColumn = parFirstColumn,
				FirstRow = parFirstRow,
				LastColumn = parLastColumn,
				LastRow = parLastRow,
				NoVerticalBand = parNoVerticalBand,
				NoHorizontalBand = parNoHorizontalBand
				};
			objTableProperties.Append(objTableLook);
			// Append the TableProperties instance to the Table instance
			objTable.Append(objTableProperties);

			return objTable;

			}

		//--------------------------
		//--- ConstructTableGrid ---
		//--------------------------
		/// <summary>
		/// Constructs a TableGrid which can then be appended to a Table object.
		/// </summary>
		/// <param name="parColumnWidthList">
		/// Pass a List of integers which contains the width of each table column in points per inch (Pix)
		/// </param>
		/// <returns></returns>
		public static DocumentFormat.OpenXml.Wordprocessing.TableGrid ConstructTableGrid (
			List<UInt32> parColumnWidthList)
			{
			// Create the TableGrid instance
			TableGrid objTableGrid = new TableGrid();
			// Process the columns as defined in the parColumnWidthList
               foreach (UInt32 columnItem in parColumnWidthList)
				{
				GridColumn objGridColumn = new GridColumn();
				objGridColumn.Width = columnItem.ToString();
				objTableGrid.Append(objGridColumn);
				};
			return objTableGrid;
			}

		//-------------------------
		//--- ConstructTableRow ---
		//-------------------------
		/// <summary>
		/// 
		/// </summary>
		/// <param name="parIsFirstRow"></param>
		/// <param name="parIsLastRow"></param>
		/// <returns></returns>
		public static TableRow ConstructTableRow(
			bool parIsFirstRow = false,
			bool parIsLastRow = false,
			bool parIsFirstColumn = false,
			bool parIsLastColumn = false,
			bool parIsOddHorizontalBand = false,
			bool parIsEvenHorizontalBand = false,
			bool parHasCondinalStyle = true) 
			{
			// Create a TableRow object
			TableRow objTableRow = new TableRow();
			objTableRow.RsidTableRowAddition = "005C4C4F";
			objTableRow.RsidTableRowProperties = "005C4C4F";
			// Create the TableRowProperties object
			TableRowProperties objTableRowProperties = new TableRowProperties();
			
			//if required, create and add the Conditional Format Style
			if(parHasCondinalStyle || parIsFirstRow)
				{
				if(parHasCondinalStyle)
					{
					// Construct a ConditionalFormatStyle instance
					ConditionalFormatStyle objConditionalFormatStyle = new ConditionalFormatStyle()
						{
						Val = "100000000000",
						FirstRow = parIsFirstRow,
						LastRow = parIsLastRow,
						FirstColumn = parIsFirstColumn,
						LastColumn = parIsLastColumn,
						OddVerticalBand = false,
						EvenVerticalBand = false,
						OddHorizontalBand = parIsOddHorizontalBand,
						EvenHorizontalBand = parIsEvenHorizontalBand
						};
					objTableRowProperties.Append(objConditionalFormatStyle);
					}
				if(parIsFirstRow)
					{
					TableHeader objTableHeader = new TableHeader();
					objTableRowProperties.Append(objTableHeader);
					}
				objTableRow.Append(objTableRowProperties);
				}
			return objTableRow;
			}

		//-------------------------
		//---ConstructTableCell ---
		//-------------------------
		/// <summary>
		/// 
		/// </summary>
		/// <param name="parCellWidth">width of the cell in Dxa (20ths of a Pixel per inch)</param>
		/// <param name="parHasCondtionalFormatting">OPTIONAL, default value = FALSE, determinse whater a Conditional formatting instance will be inserted for the table cell</param>
		/// <param name="parIsFirstRow">OPTIONAL; default = FALSE</param>
		/// <param name="parIsLastRow">OPTIONAL; default = FALSE</param>
		/// <param name="parIsFirstColumn">OPTIONAL; default = FALSE</param>
		/// <param name="parIsLastColumn">OPTIONAL; default = FALSE</param>
		/// <param name="parFirstRowFirstColumn">OPTIONAL; default = FALSE</param>
		/// <param name="parLastRowFirstColumn">OPTIONAL; default = FALSE</param>
		/// <param name="parFirstRowLastColumn">OPTIONAL; default = FALSE</param>
		/// <param name="parLastRowLastColumn">OPTIONAL; default = FALSE</param>
		/// <param name="parEvenHorizontalBand">OPTIONAL; default = FALSE</param>
		/// <param name="parOddHorizontalBand">OPTIONAL; default = FALSE</param>
		/// <returns>returns a suitably consructed TableCell object</returns>
		public static TableCell ConstructTableCell(
			UInt32Value parCellWidth,
			bool parHasCondtionalFormatting = false,
			bool parIsFirstRow = false,
			bool parIsLastRow = false,
			bool parIsFirstColumn = false,
			bool parIsLastColumn = false,
			bool parFirstRowFirstColumn = false,
			bool parLastRowFirstColumn = false,
			bool parFirstRowLastColumn = false,
			bool parLastRowLastColumn = false,
			bool parEvenHorizontalBand = false,
			bool parOddHorizontalBand = false)
			{

			// Create new TableCell instance that will be returned to the calling instruction.
			TableCell objTableCell = new TableCell();
			// Create a new TableCellProperty object
			TableCellProperties objTableCellProperties = new TableCellProperties();
			// Construct the TableWidth object
			TableCellWidth objTableCellWidth = new TableCellWidth();
			objTableCellWidth.Width = parCellWidth.ToString();
			objTableCellWidth.Type = TableWidthUnitValues.Dxa;
			objTableCellProperties.Append(objTableCellWidth);

			// Construct the Cell Alignment
			TableCellVerticalAlignment objTableCellVerticalAlignment = new TableCellVerticalAlignment();
			if (parIsFirstRow)
				objTableCellVerticalAlignment.Val = TableVerticalAlignmentValues.Center;
			else
				objTableCellVerticalAlignment.Val = TableVerticalAlignmentValues.Top;
			objTableCellProperties.Append(objTableCellVerticalAlignment);

			if(parHasCondtionalFormatting)
				{
				// Create new ConditionalFormatStyle instance
				DocumentFormat.OpenXml.Wordprocessing.ConditionalFormatStyle objConditionalFormatStyle = new ConditionalFormatStyle()
					{
					//Val = "001000000100",
					FirstRow = parIsFirstRow,
					LastRow = parIsLastRow,
					FirstColumn = parIsFirstColumn,
					LastColumn = parIsLastColumn,
					OddVerticalBand = false,
					EvenVerticalBand = false,
					OddHorizontalBand = parOddHorizontalBand,
					EvenHorizontalBand = parEvenHorizontalBand,
					FirstRowFirstColumn = parFirstRowFirstColumn,
					FirstRowLastColumn = parFirstRowLastColumn,
					LastRowFirstColumn = parLastRowFirstColumn,
					LastRowLastColumn = parLastRowLastColumn
					};
				// Append the ConditionalFormatStyle object to the TableCellProperties object.
				objTableCellProperties.Append(objConditionalFormatStyle);
				}
			
			// Append the TableCallProperties object to the TableCell object.
			objTableCell.Append(objTableCellProperties);
			return objTableCell;
			} // end of ConstructTableCell

		} //End of oxmlDocument Class


	class oxmlWorkbook : oxmlDocumentWorkbook
		{

		//==============================
		//=== InsertSharedStringItem ===
		//==============================
		/// <summary>
		/// Creates a SharedStringItem with the specified text parameter and inserts it into the SharedStringTablePart. 
		/// If the item already exists, returns its index
		/// </summary>
		/// <param name="text"></param>
		/// <param name="parShareStringPart"></param>
		/// <returns></returns>
		public static int InsertSharedStringItem(
			string parText2Insert, 
			SharedStringTablePart parShareStringPart)
			{
			int i = 0;

			// Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
			foreach(SharedStringItem item in parShareStringPart.SharedStringTable.Elements<SharedStringItem>())
				{
				if(item.InnerText == parText2Insert)
					{
					return i;
					}
				i++;
				}

			// The text does not exist in the part. Create the SharedStringItem and return its index.
			parShareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(parText2Insert)));
			parShareStringPart.SharedStringTable.Save();

			return i;
			} // InsertSharedStringItem method

		//=============================
		//--- InsertCellInWorksheet ---
		//=============================
		// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
		// If the cell already exists, returns it. 

		/// <summary>
		/// Insert a Cell into a WorksSheet, given the Column Name, Row Index and the WorksheetPart.
		/// If the cell already exists, return it
		/// </summary>
		/// <param name="parWorksheetPart"></param>
		/// <param name="parColumnName">The column letter at which to insert the cell</param>
		/// <param name="parRowNumber">The row at which to insert the cell</param>
		/// <returns>an Inserted Cell object</returns>
		public static Cell InsertCellInWorksheet(
			WorksheetPart parWorksheetPart,
               string parColumnName,
			UInt16 parRowNumber
			)
			{
			Worksheet objWorksheet = parWorksheetPart.Worksheet;
			SheetData objSheetData = objWorksheet.GetFirstChild<SheetData>();
			string strCellReference = parColumnName + parRowNumber;

			// If the worksheet does not contain a row with the specified row index, insert one.
			Row objRow;
			if(objSheetData.Elements<Row>().Where(r => r.RowIndex == parRowNumber).Count() != 0)
				{
				objRow = objSheetData.Elements<Row>().Where(r => r.RowIndex == parRowNumber).First();
				}
			else
				{
				objRow = new Row() { RowIndex = parRowNumber };
				objSheetData.Append(objRow);
				}

			// If there is not a cell with the specified column name, insert one.  
			if(objRow.Elements<Cell>().Where(c => c.CellReference.Value == parColumnName + parRowNumber).Count() > 0)
				{
				return objRow.Elements<Cell>().Where(c => c.CellReference.Value == strCellReference).First();
				}
			else
				{
				// Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
				Cell objReferenceCell = null;
				foreach(Cell objCell in objRow.Elements<Cell>())
					{
					if(string.Compare(objCell.CellReference.Value, strCellReference, true) > 0)
						{
						objReferenceCell = objCell;
						break;
						}
					}

				Cell objNewCell = new Cell();
				objNewCell.CellReference = strCellReference;
				objRow.InsertBefore(newChild: objNewCell, refChild: objReferenceCell);

				objWorksheet.Save();
				return objNewCell;
				}

			} // end of InsertCellIntoWorksheet method

		public static void InsertHyperlink(
			WorksheetPart parWorksheetPart,
			string parCellReference,
			string parHyperlinkURL,
			string parHyperLinkID
               )
			{
			Hyperlinks objHyperlinks = new Hyperlinks();
			// Check if any the Hyperlinks already exist.
			if(parWorksheetPart.Worksheet.GetFirstChild<Hyperlinks>() == null)
				{
				// Get the PageMargins, in order to insert the Hyperlink BEFORE the PageMargins
				PageMargins objPageMargins = parWorksheetPart.Worksheet.Descendants<PageMargins>().First();
				parWorksheetPart.Worksheet.InsertBefore<Hyperlinks>(newChild: objHyperlinks, refChild: objPageMargins);
				//parWorksheetPart.Worksheet.Save();
				//parWorksheetPart.Worksheet.Append(objHyperlinks);
				}
			else
				objHyperlinks = parWorksheetPart.Worksheet.Descendants<Hyperlinks>().First();
			
			//Construnct the new Hyperlink
			DocumentFormat.OpenXml.Spreadsheet.Hyperlink objHyperlink = new DocumentFormat.OpenXml.Spreadsheet.Hyperlink();
			objHyperlink.Reference = parCellReference;
			objHyperlink.Id = parHyperLinkID.ToString();
			// Append the new Hyperlink to the Hyperlinks Object
			objHyperlinks.Append(objHyperlink);
			parWorksheetPart.Worksheet.Save();

			// Insert the HyperlinkRelationship
			parWorksheetPart.AddHyperlinkRelationship(new System.Uri(uriString: parHyperlinkURL, uriKind: UriKind.Absolute), 
				isExternal: true, 
				id: parHyperLinkID.ToString());
			
			} // end of InsertHyperlink procedure


		/// <summary>
		/// Insert a comment into a Worksheet, provide the CellReference and the Text to insert as the 
		/// </summary>
		/// <param name="parCellReference"></param>
		/// <param name="parText2Add"></param>
		/// <returns></returns>
		public static DocumentFormat.OpenXml.Spreadsheet.Comment InsertComment(
			string parCellReference,
			string parText2Add
			)
			{
			// Compose the Comment Object containing the comment reference to the Cell
			DocumentFormat.OpenXml.Spreadsheet.Comment objComment = new DocumentFormat.OpenXml.Spreadsheet.Comment();
			objComment.Reference = parCellReference;
			objComment.ShapeId = 0U;
			objComment.AuthorId = 0U;

			// Construct the CommentText Object.
			CommentText objCommentText = new CommentText();

			// Construct the Run object and RunProperties object
			DocumentFormat.OpenXml.Spreadsheet.Run objRun = new DocumentFormat.OpenXml.Spreadsheet.Run();
			DocumentFormat.OpenXml.Spreadsheet.RunProperties objRunProperties = new DocumentFormat.OpenXml.Spreadsheet.RunProperties();
			// Construct the Font Family
			DocumentFormat.OpenXml.Spreadsheet.FontFamily objFontFamily = new DocumentFormat.OpenXml.Spreadsheet.FontFamily();
			objFontFamily.Val = Convert.ToInt32(Properties.AppResources.Workbooks_Comments_FontFamily);
			// Construct the Run Font
			RunFont objRunFont = new RunFont();
			objRunFont.Val = Properties.AppResources.Workbook_Comments_RunFont;
			// Construct the Font Size
			DocumentFormat.OpenXml.Spreadsheet.FontSize objFontSize = new DocumentFormat.OpenXml.Spreadsheet.FontSize();
			objFontSize.Val = Convert.ToDouble(Properties.AppResources.Workbooks_Comments_FontSize);
			// Construct the Text Color
			DocumentFormat.OpenXml.Spreadsheet.Color objColor = new DocumentFormat.OpenXml.Spreadsheet.Color();
			objColor.Indexed = Convert.ToUInt32(Properties.AppResources.Workbook_Comments_FontColor);
			//Build the RunProperties
			objRunProperties.Append(objFontSize);
			objRunProperties.Append(objColor);
			objRunProperties.Append(objRunFont);
			objRunProperties.Append(objFontFamily);
			// Construct the Text Object
			DocumentFormat.OpenXml.Spreadsheet.Text objText = new DocumentFormat.OpenXml.Spreadsheet.Text();
			objText.Text = parText2Add;

			objRun.Append(objRunProperties);
			objRun.Append(objText);

			objCommentText.Append(objRun);
			objComment.Append(objCommentText);

			return objComment;
			} // end of InsertComment procedure

		public static void PopulateCell(
			WorksheetPart parWorksheetPart,
			string parColumnLetter,
			ushort parRowNumber,
			UInt32Value parStyleId,
			CellValues parCellDatatype,
			String parCellcontents = null,
			int parHyperlinkCounter = 0,
			string parHyperlinkURL = null
			)
			{
			string strCellReference = parColumnLetter + parRowNumber;
			SheetData objSheetData = parWorksheetPart.Worksheet.GetFirstChild<SheetData>();
			
			//Populate the Cell
			Cell objCell = new Cell();
			
			// Populate the Hyperlink if required
			if(parHyperlinkURL != null)
				{
				oxmlWorkbook.InsertHyperlink(
					parWorksheetPart: parWorksheetPart,
					parCellReference: strCellReference,
					parHyperLinkID: "Hyp" + parHyperlinkCounter,
					parHyperlinkURL: parHyperlinkURL);
				}
			
			// Now determine the position where the objCell must be inserted.
			Row objRow;
			// If the worksheet does not contain a row with the specified row index, insert one.
			if(objSheetData.Elements<Row>().Where(r => r.RowIndex == parRowNumber).Count() != 0)
				{
				objRow = objSheetData.Elements<Row>().Where(r => r.RowIndex == parRowNumber).First();
				}
			else
				{
				objRow = new Row() { RowIndex = parRowNumber };
				objSheetData.Append(objRow);
				}

			// Check if the required cell exist in the row, exist, assign the objCell to it (overwriting it,
			// If the cell does not exist in the row, then insert one. 
			if(objRow.Elements<Cell>().Where(c => c.CellReference.Value == strCellReference).Count() > 0)
				{
				Cell objExistingCell = objRow.Elements<Cell>().Where(c => c.CellReference.Value == strCellReference).First();
				objCell = objExistingCell;
				}
			else // The cell doesn't exist...
				{
				// Cells must be in sequential order according to CellReference :. Determine where to insert the new cell.
				Cell objReferenceCell = null;
				foreach(Cell itemCell in objRow.Elements<Cell>())
					{
					if(string.Compare(itemCell.CellReference.Value, strCellReference, true) > 0)
						{
						objReferenceCell = itemCell;
						break;
						}
					}
				Cell objNewCell = new Cell();
				objNewCell.CellReference = strCellReference;
				objRow.InsertBefore(newChild: objNewCell, refChild: objReferenceCell);
				objCell = objNewCell;
				}

			if(parCellcontents != null)
				{
				objCell.DataType = new EnumValue<CellValues>(parCellDatatype);
				objCell.CellValue = new CellValue(parCellcontents);
				}
			objCell.StyleIndex = parStyleId;

			parWorksheetPart.Worksheet.Save();
			} // end PopulateCell procedure
		
		} //End of oxmlWorkbook class
	} // End of Namespace
