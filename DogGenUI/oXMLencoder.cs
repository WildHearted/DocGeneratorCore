using System;
using System.IO;
using System.Net;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Resources;
using System.Text;
using System.Threading.Tasks;
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
namespace DogGenUI
	{
	public class oxmlDocument
		{
		// Object Variables
		private const string localTemplatePath = @"DocGenerator\Templates";
		private const string localDocumentPath = @"DocGenerator\Documents";
		private const string localImagePath = @"DocGenerator\Images";
		// Object Properties
		private string _localDocumentPath = "";
		public string LocalDocumentPath
			{
			get{return this._localDocumentPath;}
			private set{this._localDocumentPath = value;}
			}
		private string _documentFileName = "";
		public string DocumentFilename
			{
			get{return this._documentFileName;}
			private set{this._documentFileName = value;}
			}

		private string _localDocumentURI = "";
		public string LocalDocumentURI
			{
			get{return this._localDocumentURI;}
			private set{this._localDocumentURI = value;}
			}
//--- CreateDocumentFromTemplate ---
		/// <summary>
		/// Use this method to create the new document object with which to work.
		/// It will create the new document based on the specified Tempate and Document Type. Upon creation, the LocalDocument
		/// </summary>
		/// <param name="parTemplateURL">
		/// This value must be the web URI of the template residing in the Document Templates List in SharePoint</param>
		/// <param name="parDocumentType">
		/// This value is the enumerated Document Type</param>
		/// <returns>
		/// Returns a bool with true if the creatin of the oxmlDoument object was successful and false if it failed.
		/// Validate that the bool is TRUE on return of the method.
		/// </returns>
		public bool CreateDocumentFromTemplate(string parTemplateURL, enumDocumentTypes parDocumentType)
			{
			string ErrorLogMessage = "";
			//Derive the file name of the template document
			//			Console.WriteLine(" Template URL: [{0}] \r\n" +
			//"         1         2         3         4         5         6         7         8         9        11        12        13        14        15\r\n" +
			//"12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890 \r\n" ,parTemplateURL);

			string templateFileName = parTemplateURL.Substring(parTemplateURL.LastIndexOf("/") + 1, (parTemplateURL.Length - parTemplateURL.LastIndexOf("/")) - 1);

			// Check if the DocGenerator Template Directory Exist and that it is accessable
			// Configure and validate for the relevant Template
			string templateDirectory = System.IO.Path.GetFullPath("\\") + localTemplatePath;
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
				ErrorLogMessage = "The path of template directory [" + templateDirectory + "] contains invalid characters. Ensure that the path is valid and  contains legible path characters only. \r\n " + exc.Message + " in " + exc.Source;
				Console.WriteLine(ErrorLogMessage);
				return false;
				}
			catch(DirectoryNotFoundException exc)
				{
				ErrorLogMessage = "The path of template directory [" + templateDirectory + "] is invalid. Check that the drive is mapped and exist /r/n " + exc.Message + " in " + exc.Source;
				Console.WriteLine(ErrorLogMessage);
				return false;
				}
			// Check if the template file exist in the template directory
			if(File.Exists(templateDirectory + "\\" + templateFileName))
				{
				// If the the template exist just proceed...
				Console.WriteLine("The template to use:" + templateDirectory + "\\" + templateFileName);
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
			Console.WriteLine("\t\t\t Template: {0} exist in directory: {1}? {2}", templateFileName, templateDirectory, File.Exists(templateDirectory + "\\" + templateFileName));

			// Check if the DocGenerator\Documents Directory exist and that it is accessable
			string documentDirectory = System.IO.Path.GetFullPath("\\") + localDocumentPath;
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
					ErrorLogMessage = "The path of Document Directory [" + documentDirectory + "] is invalid. Check that the drive is mapped and exist \r\n " + exc.Message + " in " + exc.Source;
					Console.WriteLine(ErrorLogMessage);
					return false;
					}
				}
			Console.WriteLine("\t\t\t The documentDirectory [" + documentDirectory + "] exist and are ready to be used.");
			// Set the object's LocalDocumentPath property
			this.LocalDocumentPath = documentDirectory;

			// Construct a name for the New Document
			string documentFilename = DateTime.Now.ToShortDateString();
			documentFilename = documentFilename.Replace("/", "-") + "_" + DateTime.Now.ToLongTimeString();
			//Console.WriteLine("filename: [{0}]", documentFilename);
			documentFilename = documentFilename.Replace(":", "-");
			documentFilename = documentFilename.Replace(" ", "_");
			documentFilename = parDocumentType + "_" + documentFilename + ".docx";
			Console.WriteLine("\t\t\t Document filename: [{0}]", documentFilename);
			// Set the object's Filename property
			this.DocumentFilename = documentFilename;

			// Create the file based on a template.
			try
				{
				File.Copy(sourceFileName: templateDirectory + "\\" + templateFileName, destFileName: documentDirectory + "\\" + documentFilename, overwrite: true);
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

			// Open the new document which is still in .dotx format to save it as a docx file
			try
				{
				WordprocessingDocument objDocument = WordprocessingDocument.Open(path: documentDirectory + "\\" + documentFilename, isEditable: true);
				// Change the document Type from .dotx to docx format.
				objDocument.ChangeDocumentType(newType: DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
				objDocument.Close();
				}
			catch(OpenXmlPackageException exc)
				{
				ErrorLogMessage = "Unable to open new Document: [" + documentDirectory + "\\" + documentFilename + "] \r\n " + exc.Message + " in " + exc.Source;
				Console.WriteLine(ErrorLogMessage);
				return false;
				}

			Console.WriteLine("\t\t\t Successfully created the new document: {0}", documentDirectory + "\\" + documentFilename);
			// Set the object's DocumentURI property
			this.LocalDocumentURI = documentDirectory + "\\" + documentFilename;
			return true;
			}
//---------------------
//---Insert Section ---
//---------------------
		/// <summary>
		/// 
		/// </summary>
		/// <param name="parBody"></param>
		/// <param name="parText2Write"></param>
		public static Paragraph Insert_Section(string parText2Write)
			{
			Paragraph objParagraph = new Paragraph();
			ParagraphProperties objParagraphProperties = new ParagraphProperties();
			ParagraphStyleId objParagraphStyleId = new ParagraphStyleId();
			objParagraphStyleId.Val = "DDSection";
			objParagraphProperties.Append(objParagraphStyleId);
			objParagraph.Append(objParagraphProperties);
			// Define the Run object instance which will containt the Text of the Section
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
			LastRenderedPageBreak objLastRenderedPageBreak = new LastRenderedPageBreak();
               objRun.Append(objLastRenderedPageBreak);
			DocumentFormat.OpenXml.Wordprocessing.Text objText = new DocumentFormat.OpenXml.Wordprocessing.Text();
			objText.Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
			objText.Text = parText2Write;
			objRun.Append(objText);
			objParagraph.Append(objRun);
			return objParagraph;
			}

//---------------------
//---Insert Heading ---
//---------------------
		/// <summary>
		/// This method inserts a new Heading Paragraph into the Body object of an oXML document
		/// </summary>
		/// <param name="parHeadingLevel">
		/// Pass an integer between 1 and 9 depending of the level of the Heading that need to be inserted.
		/// </param>
		/// <param name="parText2Write">
		/// Pass the text as astring, it will be inserted as the heading text.
		/// </param>
		public static Paragraph Insert_Heading(int parHeadingLevel, string parText2Write, bool parRestartNumbering = false)
			{
			if(parHeadingLevel < 1)
				parHeadingLevel = 1;
			else if(parHeadingLevel > 9)
				parHeadingLevel = 9;

			Paragraph objParagraph = new Paragraph();
			ParagraphProperties objParagraphProperties = new ParagraphProperties();
			ParagraphStyleId objParagraphStyleID = new ParagraphStyleId();
			objParagraphStyleID.Val = "Heading" + parHeadingLevel.ToString();
			objParagraphProperties.Append(objParagraphStyleID);
			if(parRestartNumbering)
				{
				//NumberingProperties objNumberingProperties = new NumberingProperties();
				//NumberingLevelReference objNumberingLevelReference = new NumberingLevelReference();
				//objNumberingLevelReference.Val = 0;
				//NumberingId objNumberingID = new NumberingId();
				//objNumberingID.Val = 30;
				//objNumberingProperties.Append(objNumberingLevelReference);
				//objNumberingProperties.Append(objNumberingID);
				//objParagraphProperties.Append(objNumberingProperties);
				}
			objParagraph.Append(objParagraphProperties);
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
			DocumentFormat.OpenXml.Wordprocessing.Text objText = new DocumentFormat.OpenXml.Wordprocessing.Text();
			objText.Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
			objText.Text = parText2Write;
			objRun.Append(objText);
			objParagraph.Append(objRun);
			return objParagraph;
			}

		//--------------------------
		//---Construct Paragraph ---
		//--------------------------
		/// <summary>
		/// Use this method to insert a new Body Text Paragraph
		/// </summary>
		/// <param name="parBody">
		/// Pass a refrence to a Body object
		/// </param>
		/// <param name="parIsTableParagraph">
		/// Pass boolean value of TRUE if the paragraph is for a Table else leave blank because the default value is FALSE.
		/// </param>
		/// <returns>
		/// The paragraph object that is inserted into the Body object will be returned as a Paragraph object.
		/// </returns>
		public static Paragraph Construct_Paragraph(int parBodyTextLevel, bool parIsTableParagraph = false)
			{
			if(parBodyTextLevel > 9)
				parBodyTextLevel = 9;

			//Create a Paragraph instance.
			Paragraph objParagraph = new Paragraph();
			//Create a ParagraphProperties object instance for the paragraph.
			ParagraphProperties objParagraphProperties = new ParagraphProperties();
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

		//-----------------------------
		//--- ConstructBulletParagraph ---
		//-----------------------------
		/// <summary>
		/// Use this method to insert a new Bullet Text Paragraph
		/// </summary>
		/// <param name="parBulletLevel">
		/// Pass an integer between 0 and 9 depending of the level of the body text level that need to be inserted.
		/// </param>
		/// <param name="parIsTableBullet">
		///  Pass boolean value of TRUE if the paragraph is for a Table else leave blank because the default value is FALSE.
		/// </param>
		public static Paragraph Construct_BulletNumberParagraph(int parBulletLevel, bool parIsBullet = true, bool parIsTableBullet = false)
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
		//---Construct Paragraph ---
		//--------------------------
		/// <summary>
		/// Use this method to insert a new Body Text Paragraph
		/// </summary>
		/// <param name="parBody">
		/// Pass a refrence to a Body object
		/// </param>
		/// <param name="parIsTableParagraph">
		/// Pass boolean value of TRUE if the paragraph is for a Table else leave blank because the default value is FALSE.
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
		/// Use this method to insert a new Caption into the document
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
			int parCaptionSequence,
			string parCaptionText)
			{

			//Create a Paragraph instance.
			Paragraph objParagraph = new Paragraph();
			// Create the Paragraph Properties instance.
			ParagraphProperties objParagraphProperties = new ParagraphProperties();
			ParagraphStyleId objParagraphStyleID = new ParagraphStyleId();
			if(parCaptionType == "Table")
				objParagraphStyleID.Val = "DDCaptionTable";
			else
				objParagraphStyleID.Val = "DDCaptionImage";
			objParagraphProperties.Append(objParagraphStyleID);
			//Append the ParagraphProerties to the Paragraph
			objParagraph.Append(objParagraphProperties);

			BookmarkStart objBookmarkStart = new BookmarkStart();
			objBookmarkStart.Name = "_" + parCaptionType + parCaptionSequence.ToString();
			objBookmarkStart.Id = parCaptionType + "_" + parCaptionSequence.ToString();
			objParagraph.Append(objBookmarkStart);

			// Create the Caption Run Object
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
			DocumentFormat.OpenXml.Wordprocessing.Text objText = new DocumentFormat.OpenXml.Wordprocessing.Text();
			objText.Space = SpaceProcessingModeValues.Preserve;
			objText.Text = parCaptionType + " ";
			objRun.Append(objText);
			objParagraph.Append(objRun);

			// Create and append the FieldCharacter "Begin"
			DocumentFormat.OpenXml.Wordprocessing.Run objRunFieldCharBegin = new DocumentFormat.OpenXml.Wordprocessing.Run();
			FieldChar objFieldCharacterBegin = new FieldChar();
			objFieldCharacterBegin.FieldCharType = FieldCharValues.Begin;
			objRunFieldCharBegin.Append(objFieldCharacterBegin);
			objParagraph.Append(objRunFieldCharBegin);

			DocumentFormat.OpenXml.Wordprocessing.Run objRunFieldCode = new DocumentFormat.OpenXml.Wordprocessing.Run();
			FieldCode objFieldCode = new FieldCode();
			objFieldCode.Space = SpaceProcessingModeValues.Preserve;
			if(parCaptionType == "Table")
				objFieldCode.Text = " SEQ Table \\* ARABIC ";
			else
				objFieldCode.Text = " SEQ Image \\* ARABIC ";
			objRunFieldCode.Append(objFieldCode);
			objParagraph.Append(objRunFieldCode);

			DocumentFormat.OpenXml.Wordprocessing.Run objRunFieldCharSeparate = new DocumentFormat.OpenXml.Wordprocessing.Run();
			FieldChar objFieldCharacterSeparate = new FieldChar();
			objFieldCharacterSeparate.FieldCharType = FieldCharValues.Separate;
			objRunFieldCharSeparate.Append(objFieldCharacterSeparate);
			objParagraph.Append(objRunFieldCharSeparate);

			DocumentFormat.OpenXml.Wordprocessing.Run objRunCaptionSequence = new DocumentFormat.OpenXml.Wordprocessing.Run();
			DocumentFormat.OpenXml.Wordprocessing.RunProperties objRunPropertyCaptionSeq = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
			DocumentFormat.OpenXml.Wordprocessing.NoProof objNoProof = new DocumentFormat.OpenXml.Wordprocessing.NoProof();
			objRunPropertyCaptionSeq.Append(objNoProof);
			objRunCaptionSequence.AppendChild(objRunPropertyCaptionSeq);
               DocumentFormat.OpenXml.Wordprocessing.Text objText_CaptionSequence = new DocumentFormat.OpenXml.Wordprocessing.Text();
			objText_CaptionSequence.Text = parCaptionSequence.ToString();
			objRunCaptionSequence.Append(objText_CaptionSequence);
			objParagraph.Append(objRunCaptionSequence);

			// Create and append the FieldCharacter "End"
			DocumentFormat.OpenXml.Wordprocessing.Run objRunFieldCharEnd = new DocumentFormat.OpenXml.Wordprocessing.Run();
			FieldChar objFieldCharacterEnd = new FieldChar();
			objFieldCharacterEnd.FieldCharType = FieldCharValues.End;
			objRunFieldCharEnd.Append(objFieldCharacterEnd);
			objParagraph.Append(objRunFieldCharEnd);

			// Create and append the Cation text
			DocumentFormat.OpenXml.Wordprocessing.Run objRunCaptionText = new DocumentFormat.OpenXml.Wordprocessing.Run();
			DocumentFormat.OpenXml.Wordprocessing.Text objTextCaptiontext = new DocumentFormat.OpenXml.Wordprocessing.Text();
			objTextCaptiontext.Text = ": " + parCaptionText;
			objRunCaptionText.Append(objTextCaptiontext);
			objParagraph.Append(objRunCaptionText);

			BookmarkEnd objBookmarkEnd = new BookmarkEnd();
			objBookmarkEnd.Id = parCaptionType + "_" + parCaptionSequence.ToString();
			objParagraph.Append(objBookmarkEnd);

			return objParagraph;
			}

		//------------------------
		//--- Construct_RunText ---
		//------------------------
		public static DocumentFormat.OpenXml.Wordprocessing.Run Construct_RunText(
				string parText2Write,
				bool parBold = false,
				bool parItalic = false,
				bool parUnderline = false,
				bool parSubscript = false,
				bool parSuperscript = false)
			{
			// Create a new Run object in the objParagraph
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
			// Create a Run Properties instance.
			DocumentFormat.OpenXml.Wordprocessing.RunProperties objRunProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
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
			// Append the Run Properties to the Run object
			objRun.Append(objRunProperties);
			// Insert the text in the objRun
			DocumentFormat.OpenXml.Wordprocessing.Text objText = new DocumentFormat.OpenXml.Wordprocessing.Text();
			objText.Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
			objText.Text = parText2Write;
			Console.WriteLine("\t\t**** Text writtent to document: {0} \tBold:{1} Italic:{2} Underline:{3}", objText.Text,parBold, parItalic, parUnderline);

			objRun.AppendChild(objText);
			return objRun;
			}

		//-------------------
		//--- InsertImage ---
		//-------------------
		public static DocumentFormat.OpenXml.Wordprocessing.Run InsertImage(
			ref MainDocumentPart parMainDocumentPart, 
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

			try
				{
				// Load the image into the Media section of the Document and store the relaionshipID in the variable.

				// Download the image from SharePoint if it is a http:// based image
				imageType = parImageURL.Substring(parImageURL.LastIndexOf(".") + 1, (parImageURL.Length - parImageURL.LastIndexOf(".") - 1));
				if(parImageURL.IndexOf("\\") < 0)
					{
					ErrorLogMessage = "";
					//Derive the file name of the image file
					Console.WriteLine(
					"         1         2         3         4         5         6         7         8         9        11        12        13        14        15\r\n" +
					"12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890 \r{0}" ,parImageURL);
					imageFileName = parImageURL.Substring(parImageURL.LastIndexOf("/") + 1, (parImageURL.Length - parImageURL.LastIndexOf("/")) - 1);
					// Construct the local name for the New Image file
					imageFileName = imageFileName.Replace("%20", "_");
					imageFileName = imageFileName.Replace(" ", "-");
					Console.WriteLine("\t\t\t local imageFileName: [{0}]", imageFileName);
					// Check if the DocGenerator Image Directory Exist and that it is accessable
					string imageDirectory = System.IO.Path.GetFullPath("\\") + localImagePath;
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
					parImageURL = imageDirectory + "\\" + imageFileName;
                         }
				else //if(parImageURL.IndexOf("/") > 0) // if it is a local file (not an URL...)
					imageFileName = parImageURL.Substring(parImageURL.LastIndexOf("/") + 1, (parImageURL.Length - parImageURL.LastIndexOf("/") - 1));

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
				objAnchor.DistanceFromTop = (UInt32Value) 0U;
				objAnchor.DistanceFromBottom = (UInt32Value) 0U;
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
				DrwWp.HorizontalAlignment objHorizontalAlignment = new DrwWp.HorizontalAlignment();
				objHorizontalAlignment.Text = "left";
				objHorizontalPosition.Append(objHorizontalAlignment);
				objAnchor.Append(objHorizontalPosition);

				// Define the Vertical Position
				DrwWp.VerticalPosition objVerticalPosition = new DrwWp.VerticalPosition();
				//objVerticalPosition.RelativeFrom = DrwWp.VerticalRelativePositionValues.Paragraph;
				objVerticalPosition.RelativeFrom = DrwWp.VerticalRelativePositionValues.Line;
				DrwWp.PositionOffset objVerticalPositionOffset = new DrwWp.PositionOffset();
				objVerticalPositionOffset.Text = "76200";
				objVerticalPosition.Append(objVerticalPositionOffset);
				objAnchor.Append(objVerticalPosition);

				// Define the Extent for the image
				DrwWp.Extent objExtent = new DrwWp.Extent(); // { Cx = 6010275L, Cy = 6010275L };
				objExtent.Cx = 6619875L;
				objExtent.Cy = 1457325L;
				objAnchor.Append(objExtent);
				// Define Extent Effects
				DrwWp.EffectExtent objEffectExtent = new DrwWp.EffectExtent(); // { LeftEdge = 0L, TopEdge = 0L, RightEdge = 9525L, BottomEdge = 9525L };
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
				objExtents.Cx = 6619875L;
				objExtents.Cy = 1457325L;
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

//----------------------
//--- ConstructTable ---
//----------------------
		/// <summary>
		/// 
		/// </summary>
		/// <param name="parTableWidth">
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
			UInt32  parTableWidth,
			bool parFirstColumn = false, 
			bool parLastColumn = false,  
			bool parFirstRow = false, 
			bool parLastRow = false,
			bool parNoVerticalBand = true,
			bool parNoHorizontalBand = false)
			{

			//To get the parTableWith value to 50ths of a percentage
			//Multiply by 50 to get it in 50ths of a percentage.
			//parTableWidth *= 50;
			//if(parTableWidth > 100 * 50)
			//	parTableWidth = 100 * 50;
			//else if(parTableWidth < 10 * 50)
			//	parTableWidth = 10 * 50;
			
			DocumentFormat.OpenXml.OnOffValue FirstColumnValue = parFirstColumn;
			DocumentFormat.OpenXml.OnOffValue LastColumnValue = parLastColumn;
			DocumentFormat.OpenXml.OnOffValue FirstRowValue = parFirstRow;
			DocumentFormat.OpenXml.OnOffValue LastRowValue = parLastRow;
			DocumentFormat.OpenXml.OnOffValue NoVerticalBandValue = parNoVerticalBand;
			DocumentFormat.OpenXml.OnOffValue NoHorizontalBandValue = parNoHorizontalBand;
			
			// Creates a Table instance
			DocumentFormat.OpenXml.Wordprocessing.Table objTable = new DocumentFormat.OpenXml.Wordprocessing.Table();
			// Create and set the Table Properties instance
			DocumentFormat.OpenXml.Wordprocessing.TableProperties objTableProperties = new DocumentFormat.OpenXml.Wordprocessing.TableProperties();
			DocumentFormat.OpenXml.Wordprocessing.TableStyle objTableStyle = new DocumentFormat.OpenXml.Wordprocessing.TableStyle() { Val = "DDGreenHeaderTable" };
			DocumentFormat.OpenXml.Wordprocessing.TableWidth objTableWidth = new TableWidth()
				{ Width = Convert.ToString(parTableWidth), Type = TableWidthUnitValues.Dxa };
			DocumentFormat.OpenXml.Wordprocessing.TableJustification objTableJustification = new TableJustification();
			objTableJustification.Val = TableRowAlignmentValues.Left;
			DocumentFormat.OpenXml.Wordprocessing.TableLook objTableLook = new DocumentFormat.OpenXml.Wordprocessing.TableLook()
				{Val = "04A0",
                    FirstColumn = FirstColumnValue,
				FirstRow = FirstRowValue,
				LastColumn = LastColumnValue,
				LastRow = LastRowValue,
				NoVerticalBand = NoVerticalBandValue,
				NoHorizontalBand = NoHorizontalBandValue};

			objTableProperties.Append(objTableStyle);
			objTableProperties.Append(objTableWidth);
			objTableProperties.Append(objTableJustification);
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
/// Pass a List of integers which contains the width of each table column in points)
/// </param>
/// <returns></returns>
		public static DocumentFormat.OpenXml.Wordprocessing.TableGrid ConstructTableGrid (
			List<UInt32> parColumnWidthList,
			string parTableColumnUnit,
			UInt32 parTableWidth)
			{
			// Create the TableGrid instance
			DocumentFormat.OpenXml.Wordprocessing.TableGrid objTableGrid = new DocumentFormat.OpenXml.Wordprocessing.TableGrid();
			// Process the columns as defined in the parColumnWidthList
               foreach (UInt32 columnItem in parColumnWidthList)
				{
				GridColumn objGridColumn = new GridColumn();
				// the 
				if(parTableColumnUnit == "%")
					{
					if (columnItem > 100)
						objGridColumn.Width = (columnItem / parColumnWidthList.Count).ToString();
					else
						objGridColumn.Width = columnItem.ToString();
					}
				else
					{
					objGridColumn.Width = columnItem.ToString();
					}
				
				objTableGrid.Append(objGridColumn);
				};
			return objTableGrid;
			}
		//--------------------------
		//--- ConstructTableRow ---
		//-------------------------
		/// <summary>
		/// 
		/// </summary>
		/// <param name="parIsFirstRow"></param>
		/// <param name="parIsLastRow"></param>
		/// <returns></returns>
		public static DocumentFormat.OpenXml.Wordprocessing.TableRow ConstructTableRow(
			bool parIsFirstRow = false,
			bool parIsLastRow = false,
			bool parIsFirstColumn = false,
			bool parIsLastColumn = false,
			bool parIsOddHorizontalBand = false,
			bool parIsEvenHorizontalBand = false) 
			{
			// Create a TableRow instance
			TableRow objTableRow = new TableRow() { };
			// Create a TableRowProperties instance
			TableRowProperties objTableRowProperties = new TableRowProperties();
			// Create a ConditionalFormatStyle instance
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
			if(parIsFirstRow)
				{
				TableHeader objTableHeader = new TableHeader();
				objTableRowProperties.Append(objTableHeader);
				}
			objTableRow.Append(objTableRowProperties);
			return objTableRow;
			}

		//-------------------------
		//---ConstructTableCell ---
		//-------------------------
		/// <summary>
		/// 
		/// </summary>
		/// <param name="parColumnWidthPercentage"></param>
		/// <param name="parIsFirstRowCell"></param>
		/// <param name="parIsLastRowCell"></param>
		/// <param name="parIsFirstColumnCell"></param>
		/// <param name="parIsLastColumnCell"></param>
		/// <returns></returns>
		public static DocumentFormat.OpenXml.Wordprocessing.TableCell ConstructTableCell(
			//int parColumnWidthPercentage,
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

			// Create new TableCell instance
			DocumentFormat.OpenXml.Wordprocessing.TableCell objTableCell = new TableCell();
			// Create a new TableCellProperty Instance
			DocumentFormat.OpenXml.Wordprocessing.TableCellProperties objTableCellProperties = new TableCellProperties();

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

			TableCellWidth objTableCellWidth = new TableCellWidth()
				{
				//Width = (parColumnWidthPercentage *= 50).ToString(),
				Type = TableWidthUnitValues.Auto
				};
			// Append the ConditionalFormatStyle object and TableCellWidth object to the TableCellProperties object.
			objTableCellProperties.Append(objConditionalFormatStyle);
			objTableCellProperties.Append(objTableCellWidth);
			// Append the TableCallProperties object to the TableCell object.
			objTableCell.Append(objTableCellProperties);
			return objTableCell;
			}

	} //End of oxmlDocument Class

	class oxmlWorkbook
		{
		}
		
	} //End of oxmlWorkbook class
