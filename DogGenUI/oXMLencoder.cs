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
		// Object Properties
		private string _localDocumentPath = "";
		public string LocalDocumentPath
			{
			get
				{
				return this._localDocumentPath;
				}
			private set
				{
				this._localDocumentPath = value;
				}
			}
		private string _documentFileName = "";
		public string DocumentFilename
			{
			get
				{
				return this._documentFileName;
				}
			private set
				{
				this._documentFileName = value;
				}
			}
		private string _localDocumentURI = "";
		public string LocalDocumentURI
			{
			get
				{
				return this._localDocumentURI;
				}
			private set
				{
				this._localDocumentURI = value;
				}
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
			//string localTemplatePath = "";
			//string localDocumentPath = "";
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
		/// <param name="parBodyTextLevel">
		/// Pass an integer between 0 and 9 depending of the level of the body text level that need to be inserted.
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
			// Get the first PropertiesElement for the paragraph.
			ParagraphProperties objParagraphProperties = new ParagraphProperties();
			ParagraphStyleId objParagraphStyleID = new ParagraphStyleId();
			if(parIsTableParagraph)
				{
				//objParagraphStyleID.Val = "DDTableBodyText";
				//objParagraphProperties.Append(objParagraphStyleID);
				}
			else
				{
				objParagraphStyleID.Val = "DDBodyText" + parBodyTextLevel.ToString();
				objParagraphProperties.Append(objParagraphStyleID);
				}
			
			objParagraph.Append(objParagraphProperties);
			return objParagraph;
			}

//-----------------------------
//--- InsertBulletParagraph ---
//-----------------------------
/// <summary>
		/// Use this method to insert a new Bullet Text Paragraph
		/// </summary>
		/// <param name="parBody">
		/// Pass a refrence to a Body object
		/// </param>
		/// <param name="parBulletLevel">
		/// Pass an integer between 0 and 9 depending of the level of the body text level that need to be inserted.
		/// </param>
		/// <param name="parText2Write">
		/// Pass the text as astring, it will be inserted as the heading text.
		/// </param>
		public static Paragraph Insert_BulletParagraph(ref Body parBody, int parBulletLevel, string parText2Write)
			{
			if(parBulletLevel > 9)
				parBulletLevel = 9;
			//Insert a new Paragraph to the end of the Body of the objDocument
			Paragraph objParagraph = parBody.AppendChild(new Paragraph());
			
			// Get the first PropertiesElement for the paragraph.
			if(objParagraph.Elements<ParagraphProperties>().Count() == 0)
				objParagraph.PrependChild<ParagraphProperties>(new ParagraphProperties());

			ParagraphProperties objParagraphProperties = objParagraph.Elements<ParagraphProperties>().First();
			objParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = "DDBulletText" + parBulletLevel.ToString() };
			// Check if the run object has any Run Properties, if not add RunProperties to it.
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = objParagraph.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run());
			if(objRun.Elements<DocumentFormat.OpenXml.Wordprocessing.RunProperties>().Count() == 0)
				objRun.PrependChild<DocumentFormat.OpenXml.Wordprocessing.RunProperties>(new DocumentFormat.OpenXml.Wordprocessing.RunProperties());
			DocumentFormat.OpenXml.Wordprocessing.Text objText = objRun.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text());
			objText.Space = SpaceProcessingModeValues.Preserve;
			objRun.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(parText2Write));
			objParagraph.Append(objParagraphProperties);
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
			objText.Text = parCaptionType;
			objRun.Append(objText);
			objParagraph.Append(objRun);

			// Create and append the FieldCharacter "Begin"
			DocumentFormat.OpenXml.Wordprocessing.Run objRunFieldCharBegin = new DocumentFormat.OpenXml.Wordprocessing.Run();
			FieldChar objFieldCharacterBegin = new FieldChar();
			objFieldCharacterBegin.FieldCharType = FieldCharValues.Begin;
			objRunFieldCharBegin.Append(objFieldCharacterBegin);
			objParagraph.Append(objRunFieldCharBegin);

			DocumentFormat.OpenXml.Wordprocessing.Run objRunFielCodePreserve = new DocumentFormat.OpenXml.Wordprocessing.Run();
			FieldCode objFieldCode = new FieldCode();
			objFieldCode.Space = SpaceProcessingModeValues.Preserve;
			if(parCaptionType == "Table")
				objFieldCode.Text = " SEQ Table \\ ARABIC ";
			else
				objFieldCode.Text = " SEQ Image \\ ARABIC ";
			objRunFielCodePreserve.Append(objFieldCode);
			objParagraph.Append(objRunFielCodePreserve);

			DocumentFormat.OpenXml.Wordprocessing.Run objRunFieldCharSeparate = new DocumentFormat.OpenXml.Wordprocessing.Run();
			FieldChar objFieldCharacterSeparate = new FieldChar();
			objFieldCharacterSeparate.FieldCharType = FieldCharValues.Separate;
			objRunFieldCharSeparate.Append(objFieldCharacterSeparate);
			objParagraph.Append(objRunFieldCharSeparate);

			DocumentFormat.OpenXml.Wordprocessing.Run objRunCaptionSequence = new DocumentFormat.OpenXml.Wordprocessing.Run();
			//DocumentFormat.OpenXml.Wordprocessing.RunProperties objRunPropertyCaptionSeq = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
			//DocumentFormat.OpenXml.Wordprocessing.NoProof objNoProof = new NoProof();
			//objRunPropertyCaptionSeq.Append(objNoProof);
			//objRunCaptionSequence.AppendChild(objRunPropertyCaptionSeq);
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
			//DocumentFormat.OpenXml.Wordprocessing.RunProperties objRunPropertyCaptionText = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
			//objRunPropertyCaptionText.Append(objNoProof);
			//objRunCaptionText.Append(objRunPropertyCaptionText);
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
		//--- ConstructrunText ---
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
			WordprocessingDocument parWPdocument, 
			int parParagraphLevel, 
			int parPictureSeqNo, 
			string parImageURL)
			{
			if(parParagraphLevel < 1)
				parParagraphLevel = 1;
			else if(parParagraphLevel > 9)
				parParagraphLevel = 9;

			string imgFileName = "";
			string ErrorLogMessage = "";
			string imgType = "";
			string relationshipID = "";

			try
				{
				// Load the image into the Media section of the Document and store the relaionshipID in the variable.
				MainDocumentPart objMainDocumentPart = parWPdocument.MainDocumentPart;

				imgType = parImageURL.Substring(parImageURL.LastIndexOf(".") + 1, (parImageURL.Length - parImageURL.LastIndexOf(".") - 1));
				if(parImageURL.IndexOf("\\") > 0)
					imgFileName = parImageURL.Substring(parImageURL.LastIndexOf("\\") + 1, (parImageURL.Length - parImageURL.LastIndexOf("\\") - 1));
				else if(parImageURL.IndexOf("/") > 0)
					imgFileName = parImageURL.Substring(parImageURL.LastIndexOf("/") + 1, (parImageURL.Length - parImageURL.LastIndexOf("/") - 1));
				switch(imgType)
					{
					case "jpg":
							{
							ImagePart objImagePart = objMainDocumentPart.AddImagePart(ImagePartType.Jpeg);
							using(FileStream objFileStream = new FileStream(path: parImageURL, mode: FileMode.Open))
								{
								objImagePart.FeedData(objFileStream);
								}
							relationshipID = objMainDocumentPart.GetIdOfPart(part: objImagePart);
							break;
							}
					case "gif":
							{
							ImagePart objImagePart = objMainDocumentPart.AddImagePart(ImagePartType.Gif);
							using(FileStream objFileStream = new FileStream(path: parImageURL, mode: FileMode.Open))
								{
								objImagePart.FeedData(objFileStream);
								}
							relationshipID = objMainDocumentPart.GetIdOfPart(part: objImagePart);
							break;
							}
					case "bmp":
							{
							ImagePart objImagePart = objMainDocumentPart.AddImagePart(ImagePartType.Bmp);
							using(FileStream objFileStream = new FileStream(path: parImageURL, mode: FileMode.Open))
								{
								objImagePart.FeedData(objFileStream);
								}
							relationshipID = objMainDocumentPart.GetIdOfPart(part: objImagePart);
							break;
							}
					case "png":
							{
							ImagePart objImagePart = objMainDocumentPart.AddImagePart(ImagePartType.Png);
							using(FileStream objFileStream = new FileStream(path: parImageURL, mode: FileMode.Open))
								{
								objImagePart.FeedData(objFileStream);
								}
							relationshipID = objMainDocumentPart.GetIdOfPart(part: objImagePart);
							break;
							}
					case "tiff":
							{
							ImagePart objImagePart = objMainDocumentPart.AddImagePart(ImagePartType.Tiff);
							using(FileStream objFileStream = new FileStream(path: parImageURL, mode: FileMode.Open))
								{
								objImagePart.FeedData(objFileStream);
								}
							relationshipID = objMainDocumentPart.GetIdOfPart(part: objImagePart);
							break;
							}
					default:
							{
							break;
							}
					}
				DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
				DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
				// Prepare the Anchor object
				DrwWp.Anchor objAnchor = new DrwWp.Anchor() { DistanceFromTop = (UInt32Value) 0U, DistanceFromBottom = (UInt32Value) 0U, DistanceFromLeft = (UInt32Value) 114300U, DistanceFromRight = (UInt32Value) 114300U, SimplePos = false, RelativeHeight = (UInt32Value) 251658240U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true, EditId = "09096F23", AnchorId = "411CCDA1" };
				DrwWp.SimplePosition objSimplePosition = new DrwWp.SimplePosition() { X = 0L, Y = 0L };

				DrwWp.HorizontalPosition objHorizontalPosition = new DrwWp.HorizontalPosition() { RelativeFrom = DrwWp.HorizontalRelativePositionValues.Column };
				DrwWp.PositionOffset objHorizontalPositionOffset = new DrwWp.PositionOffset();
				objHorizontalPositionOffset.Text = "393065";
				objHorizontalPosition.Append(objHorizontalPositionOffset);

				DrwWp.VerticalPosition ObjVerticalPosition = new DrwWp.VerticalPosition() { RelativeFrom = DrwWp.VerticalRelativePositionValues.Paragraph };
				DrwWp.PositionOffset objVerticalPositionOffset = new DrwWp.PositionOffset();
				objVerticalPositionOffset.Text = "377190";
				ObjVerticalPosition.Append(objVerticalPositionOffset);

				DrwWp.Extent objExtent = new DrwWp.Extent() { Cx = 6010275L, Cy = 6010275L };
				DrwWp.EffectExtent objEffectExtent = new DrwWp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 9525L, BottomEdge = 9525L };
				DrwWp.WrapTopBottom objWrapTopBottom = new DrwWp.WrapTopBottom();

				DrwWp.DocProperties objDocProperties = new DrwWp.DocProperties() { Id = Convert.ToUInt32(parPictureSeqNo), Name = "Picture " + parPictureSeqNo.ToString() };

				DrwWp.NonVisualGraphicFrameDrawingProperties objNonVisualGraphicFrameDrawingProperties = new DrwWp.NonVisualGraphicFrameDrawingProperties();

				Drw.GraphicFrameLocks objGraphicFrameLocks = new Drw.GraphicFrameLocks() { NoChangeAspect = true };
				objGraphicFrameLocks.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

				objNonVisualGraphicFrameDrawingProperties.Append(objGraphicFrameLocks);

				Drw.Graphic objGgraphic = new Drw.Graphic();
				objGgraphic.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

				Drw.GraphicData objGraphicData = new Drw.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

				Pic.Picture objPicture = new Pic.Picture();
				objPicture.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

				Pic.NonVisualPictureProperties objNonVisualPictureProperties = new Pic.NonVisualPictureProperties();
				Pic.NonVisualDrawingProperties objNonVisualDrawingProperties = new Pic.NonVisualDrawingProperties()
					{
					Id = Convert.ToUInt32(parPictureSeqNo),
					Name = imgFileName
					};

				Pic.NonVisualPictureDrawingProperties objNonVisualPictureDrawingProperties = new Pic.NonVisualPictureDrawingProperties();

				objNonVisualPictureProperties.Append(objNonVisualDrawingProperties);
				objNonVisualPictureProperties.Append(objNonVisualPictureDrawingProperties);

				Pic.BlipFill objBlipFill = new Pic.BlipFill();

				Drw.Blip objBlip = new Drw.Blip() { Embed = relationshipID };
				Drw.BlipExtensionList objBlipExtensionList = new Drw.BlipExtensionList();

				Drw.BlipExtension objBlipExtension = new Drw.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };
				Drw2010.UseLocalDpi objUseLocalDpi = new Drw2010.UseLocalDpi() { Val = false };
				objUseLocalDpi.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
				objBlipExtension.Append(objUseLocalDpi);
				objBlipExtensionList.Append(objBlipExtension);
				objBlip.Append(objBlipExtensionList);

				Drw.Stretch objStretch = new Drw.Stretch();
				Drw.FillRectangle objFillRectangle = new Drw.FillRectangle();
				objStretch.Append(objFillRectangle);
				objBlipFill.Append(objBlip);
				objBlipFill.Append(objStretch);

				Pic.ShapeProperties objShapeProperties = new Pic.ShapeProperties();

				Drw.Transform2D objTransform2D = new Drw.Transform2D();
				Drw.Offset objOffset = new Drw.Offset() { X = 0L, Y = 0L };
				Drw.Extents objExtents = new Drw.Extents() { Cx = 6010275L, Cy = 6010275L };

				objTransform2D.Append(objOffset);
				objTransform2D.Append(objExtents);

				Drw.PresetGeometry objPresetGeometry = new Drw.PresetGeometry() { Preset = Drw.ShapeTypeValues.Rectangle };
				Drw.AdjustValueList objAdjustValueList = new Drw.AdjustValueList();

				objPresetGeometry.Append(objAdjustValueList);

				objShapeProperties.Append(objTransform2D);
				objShapeProperties.Append(objPresetGeometry);

				objPicture.Append(objNonVisualPictureProperties);
				objPicture.Append(objBlipFill);
				objPicture.Append(objShapeProperties);

				objGraphicData.Append(objPicture);

				objGgraphic.Append(objGraphicData);

				DrwWp2010.RelativeWidth objRelativeWidth = new DrwWp2010.RelativeWidth() { ObjectId = DrwWp2010.SizeRelativeHorizontallyValues.Page };
				DrwWp2010.PercentageWidth objPercentageWidth = new DrwWp2010.PercentageWidth();
				objPercentageWidth.Text = "0";
				objRelativeWidth.Append(objPercentageWidth);

				DrwWp2010.RelativeHeight objRelativeHeight = new DrwWp2010.RelativeHeight() { RelativeFrom = DrwWp2010.SizeRelativeVerticallyValues.Page };
				DrwWp2010.PercentageHeight objPercentageHeight = new DrwWp2010.PercentageHeight();
				objPercentageHeight.Text = "0";
				objRelativeHeight.Append(objPercentageHeight);

				objAnchor.Append(objSimplePosition);
				objAnchor.Append(objHorizontalPosition);
				objAnchor.Append(ObjVerticalPosition);
				objAnchor.Append(objExtent);
				objAnchor.Append(objEffectExtent);
				objAnchor.Append(objWrapTopBottom);
				objAnchor.Append(objDocProperties);
				objAnchor.Append(objNonVisualGraphicFrameDrawingProperties);
				objAnchor.Append(objGgraphic);
				objAnchor.Append(objRelativeWidth);
				objAnchor.Append(objRelativeHeight);

				objDrawing.Append(objAnchor);

				objRun.Append(objDrawing);
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
			objTableJustification.Val = TableRowAlignmentValues.Right;
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
