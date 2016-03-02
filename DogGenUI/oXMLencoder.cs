using System;
using System.IO;
using System.Net;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Resources;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DrwWp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DrwWp2010 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using Drw =DocumentFormat.OpenXml.Drawing;
using Drw2010 = DocumentFormat.OpenXml.Office2010.Drawing;
using Drw2013 = DocumentFormat.OpenXml.Office2013.Drawing;
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
	public class oxmlDocument
		{
		// Object Variables

		// Object Properties
		private string _localDocumentPath = "";
		public string LocalDocumentPath
			{
			get { return this._localDocumentPath; }
			private set { this._localDocumentPath = value; }
			}
		private string _documentFileName = "";
		public string DocumentFilename
			{
			get { return this._documentFileName; }
			private set { this._documentFileName = value; }
			}

		private string _localDocumentURI = "";
		public string LocalDocumentURI
			{
			get { return this._localDocumentURI; }
			private set { this._localDocumentURI = value; }
			}

		//----------------------------------
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
			string templateDirectory = System.IO.Path.GetFullPath("\\") + DocGenerator.Properties.AppResources.LocalTemplatePath;
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
			string documentDirectory = System.IO.Path.GetFullPath("\\") + DocGenerator.Properties.AppResources.LocalDocumentPath;
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
		/// This method inserts a complete new Section in the document.
		/// </summary>
		/// <param name="parText2Write">
		/// Provide the string that will be inserted as the text of the new Section.
		/// </param>
		/// <param name="parIsError">
		/// Optional parameter: True if an error occurred and the Section text will be emphasised in Red with waved underline.
		/// Defaults to False.
		/// </param>
		/// <param name="parHyperlinkRelationshipID">
		/// Optional String parameter defaults to a "" (blank string). If a Hyperlink must be inserted, this parameter must contain a value which represents the 
		/// image relationship ID of the ClickLink image as inserted in the Main Document Body
		/// </param>
		/// <param name="parHyperlinkURL">
		/// Optional String prameter which defaults to a "" (blank string). If a Hyperlink must be inserted, the complete URL to the specific entry in SharePoint
		/// must be provided in this parameter. This complete string will be added to the image provided in the parHyperlinkRelationshipID image as the URL to be
		/// inserted in the document.
		/// </param>
		/// <returns></returns>
		public static Paragraph Insert_Section()
			{
			Paragraph objParagraph = new Paragraph();
			ParagraphProperties objParagraphProperties = new ParagraphProperties();
			ParagraphStyleId objParagraphStyleId = new ParagraphStyleId();
			objParagraphStyleId.Val = "DDSection";
			objParagraphProperties.Append(objParagraphStyleId);

			objParagraph.Append(objParagraphProperties);

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
		/// <param name="parBookMark">
		/// Optional Parameter. When a Bookmark must be created for the heading, pass the BookMark label 
		/// (without any spaces or odd characters) that need to be inserted as a string. By default the value is Null, 
		/// which means the heading will not contain a Bookmark. If a value is passed a Bookmark will be inserted.
		/// </param>
		public static Paragraph Insert_Heading(
			int parHeadingLevel,
			string parBookMark = null)
			{
			if(parHeadingLevel < 1)
				parHeadingLevel = 1;
			else if(parHeadingLevel > 9)
				parHeadingLevel = 9;
			
			Paragraph objParagraph = new Paragraph();
			ParagraphProperties objParagraphProperties = new ParagraphProperties();
			ParagraphStyleId objParagraphStyleID = new ParagraphStyleId();
			objParagraphStyleID.Val = "DDHeading" + parHeadingLevel.ToString();
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

		//---------------------------
		//--- Construct_Paragraph ---
		//---------------------------
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

		//-----------------------------
		//--- Construct_BulletNumberParagraph ---
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
		//---Construct_Error ---
		//--------------------------
		/// <summary>
		/// Use this method to insert a new Body Text Paragraph and highlights it in RED text 
		/// to indicate an error in the SharePoint Enahanced Rich Text.
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

			//BookmarkStart objBookmarkStart = new BookmarkStart();
			//objBookmarkStart.Name = "_" + parCaptionType + parCaptionSequence.ToString();
			//objBookmarkStart.Id = parCaptionType + "_" + parCaptionSequence.ToString();
			//objParagraph.Append(objBookmarkStart);

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

			//BookmarkEnd objBookmarkEnd = new BookmarkEnd();
			//objBookmarkEnd.Id = parCaptionType + "_" + parCaptionSequence.ToString();
			//objParagraph.Append(objBookmarkEnd);

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
				// Append the Run Properties to the Run object
				objRun.AppendChild(objRunProperties);
			} // if(parIsNewSection)
			// Insert the text in the objRun
			DocumentFormat.OpenXml.Wordprocessing.Text objText = new DocumentFormat.OpenXml.Wordprocessing.Text();
			objText.Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
			objText.Text = parText2Write;
			Console.WriteLine("**** Text ****: {0} \tBold:{1} Italic:{2} Underline:{3}", objText.Text,parBold, parItalic, parUnderline);
			
			objRun.AppendChild(objText);
			return objRun;
			}

		//-------------------
		//--- InsertImage ---
		//-------------------
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
				// Load the image into the Media section of the Document and store the relaionshipID in the variable.

				// Download the image from SharePoint if it is a http:// based image
				imageType = parImageURL.Substring(parImageURL.LastIndexOf(".") + 1, (parImageURL.Length - parImageURL.LastIndexOf(".") - 1));
				if(parImageURL.IndexOf("\\") < 0)
					{
					ErrorLogMessage = "";
					//Derive the file name of the image file
					Console.WriteLine(
					"         1         2         3         4         5         6         7         8         9        11        12        13        14        15\r\n" +
					"12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890 \r{0}", parImageURL);
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
			//DocumentFormat.OpenXml.OnOffValue FirstColumnValue = parFirstColumn;
			//DocumentFormat.OpenXml.OnOffValue LastColumnValue = parLastColumn;
			//DocumentFormat.OpenXml.OnOffValue FirstRowValue = parFirstRow;
			//DocumentFormat.OpenXml.OnOffValue LastRowValue = parLastRow;
			//DocumentFormat.OpenXml.OnOffValue NoVerticalBandValue = parNoVerticalBand;
			//DocumentFormat.OpenXml.OnOffValue NoHorizontalBandValue = parNoHorizontalBand;
			
			// Creates a Table instance
			DocumentFormat.OpenXml.Wordprocessing.Table objTable = new DocumentFormat.OpenXml.Wordprocessing.Table();
			// Create and set the Table Properties instance
			DocumentFormat.OpenXml.Wordprocessing.TableProperties objTableProperties = new DocumentFormat.OpenXml.Wordprocessing.TableProperties();
			DocumentFormat.OpenXml.Wordprocessing.TableStyle objTableStyle = new DocumentFormat.OpenXml.Wordprocessing.TableStyle() { Val = "DDGreenHeaderTable" };
			DocumentFormat.OpenXml.Wordprocessing.TableWidth objTableWidth = new TableWidth()
				{ Width = "0", Type = TableWidthUnitValues.Auto };
			DocumentFormat.OpenXml.Wordprocessing.TableJustification objTableJustification = new TableJustification();
			objTableJustification.Val = TableRowAlignmentValues.Left;
			DocumentFormat.OpenXml.Wordprocessing.TableLook objTableLook = new DocumentFormat.OpenXml.Wordprocessing.TableLook()
				{Val = "0600",
                    FirstColumn = parFirstColumn,
				FirstRow = parFirstRow,
				LastColumn = parLastColumn,
				LastRow = parLastRow,
				NoVerticalBand = parNoVerticalBand,
				NoHorizontalBand = parNoHorizontalBand
				};

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
		/// Pass a List of integers which contains the width of each table column in points per inch (Pix)
		/// </param>
		/// <returns></returns>
		public static DocumentFormat.OpenXml.Wordprocessing.TableGrid ConstructTableGrid (
			List<UInt32> parColumnWidthList)
			{
			// Create the TableGrid instance
			DocumentFormat.OpenXml.Wordprocessing.TableGrid objTableGrid = new DocumentFormat.OpenXml.Wordprocessing.TableGrid();
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
		public static DocumentFormat.OpenXml.Wordprocessing.TableRow ConstructTableRow(
			bool parIsFirstRow = false,
			bool parIsLastRow = false,
			bool parIsFirstColumn = false,
			bool parIsLastColumn = false,
			bool parIsOddHorizontalBand = false,
			bool parIsEvenHorizontalBand = false,
			bool parHasCondinalStyle = true) 
			{
			// Create a TableRow instance
			TableRow objTableRow = new TableRow();
			objTableRow.RsidTableRowAddition = "005C4C4F";
			objTableRow.RsidTableRowProperties = "005C4C4F";

			if(parHasCondinalStyle || parIsFirstRow)
				{
				TableRowProperties objTableRowProperties = new TableRowProperties();
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
		public static DocumentFormat.OpenXml.Wordprocessing.TableCell ConstructTableCell(
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
			DocumentFormat.OpenXml.Wordprocessing.TableCell objTableCell = new TableCell();
			// Create a new TableCellProperty object
			DocumentFormat.OpenXml.Wordprocessing.TableCellProperties objTableCellProperties = new TableCellProperties();
			// Construct the TableWidth object
			TableCellWidth objTableCellWidth = new TableCellWidth();
			objTableCellWidth.Width = parCellWidth.ToString();
			objTableCellWidth.Type = TableWidthUnitValues.Dxa;
			// Append the TableCellWidth object to the TableCellProperties object.
			objTableCellProperties.Append(objTableCellWidth);

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

	class oxmlWorkbook
		{
		}
		
	} //End of oxmlWorkbook class
