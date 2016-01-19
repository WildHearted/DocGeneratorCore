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
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using A =DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
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
		/// <summary>
		/// 
		/// </summary>
		/// <param name="parBody"></param>
		/// <param name="parText2Write"></param>
		public static void Insert_Section(ref Body parBody, string parText2Write)
			{
			//Insert a new Paragraph to the end of the Body of the objDocument
			DocumentFormat.OpenXml.Wordprocessing.Paragraph objParagraph = parBody.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
			// Get the first ParagraphProperties.Element for the paragraph.
			if(objParagraph.Elements<DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties>().Count() == 0)
				objParagraph.PrependChild<DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties>(new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties());
			DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties objParagraphProperties = objParagraph.Elements<DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties>().First();
			objParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = "DDSection" };

			DocumentFormat.OpenXml.Wordprocessing.Run objRun = objParagraph.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run());
			// Check if the run object has any Run Properties, if not add RunProperties to it.
			//if(objRun.Elements<DocumentFormat.OpenXml.Wordprocessing.RunProperties>().Count() == 0)
			//	objRun.PrependChild(new DocumentFormat.OpenXml.Wordprocessing.RunProperties());
			// Get the first Run Properties Element for the run.
			//DocumentFormat.OpenXml.Wordprocessing.RunProperties objRunProperties = objRun.Elements<DocumentFormat.OpenXml.Wordprocessing.RunProperties>().First();
			//objRunProperties.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.LastRenderedPageBreak());
			objRun.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.LastRenderedPageBreak());
			DocumentFormat.OpenXml.Wordprocessing.Text objText = objRun.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text() );
			objText.Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
			objRun.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(parText2Write) );
			}


		/// <summary>
		/// This method inserts a new Heading Paragraph into the Body object of an oXML document
		/// </summary>
		/// <param name="parBody"></param>
		/// Pass a refrence to a Body object
		/// <param name="parHeadingLevel">
		/// Pass an integer between 1 and 9 depending of the level of the Heading that need to be inserted.
		/// </param>
		/// <param name="parText2Write">
		/// Pass the text as astring, it will be inserted as the heading text.
		/// </param>
		public static void Insert_Heading(ref Body parBody, int parHeadingLevel, string parText2Write)
			{
			if(parHeadingLevel < 1)
				parHeadingLevel = 1;
			else if(parHeadingLevel > 9)
				parHeadingLevel = 9;
			//Insert a new Paragraph to the end of the Body of the objDocument
			DocumentFormat.OpenXml.Wordprocessing.Paragraph objParagraph = parBody.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
			// Get the first PropertiesElement for the paragraph.
			if(objParagraph.Elements<DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties>().Count() == 0)
				objParagraph.PrependChild<DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties>(new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties());
			DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties objParagraphProperties = objParagraph.Elements<DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties>().First();
			objParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = "Heading" + parHeadingLevel.ToString() };
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = objParagraph.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run());
			DocumentFormat.OpenXml.Wordprocessing.Text objText = objRun.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text());
			objText.Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
			objRun.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(parText2Write));
			}

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
		public static Paragraph Insert_BodyTextParagraph(ref Body parBody, int parBodyTextLevel)
			{
			if(parBodyTextLevel < 1)
				parBodyTextLevel = 1;
			else if(parBodyTextLevel > 9)
				parBodyTextLevel = 9;

			//Insert a new Paragraph to the end of the Body of the objDocument
			Paragraph objParagraph = parBody.AppendChild(new Paragraph());
			
			// Get the first PropertiesElement for the paragraph.
			if(objParagraph.Elements<ParagraphProperties>().Count() == 0)
				objParagraph.PrependChild(new ParagraphProperties());
			ParagraphProperties objParagraphProperties = objParagraph.Elements<ParagraphProperties>().First();
			objParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = "DDBodyText" + parBodyTextLevel.ToString() };
			return objParagraph;
			}

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
		public static void Insert_BulletParagraph(ref Body parBody, int parBulletLevel, string parText2Write)
			{
			if(parBulletLevel < 1)
				parBulletLevel = 1;
			else if(parBulletLevel > 9)
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
			}


		public static void Insert_Run_Text(
			Paragraph parParagraphObj,
				string parText2Write,
				bool parBold = false,
				bool parItalic = false,
				bool parUnderline = false)
			{
			string ErrorLogMessage = "";
			// Insert a new Run object in the objParagraph
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = parParagraphObj.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run());
			
			//objRun.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(parText2Write));
			
			if(parBold || parItalic || parUnderline)
				{
				// Check if the run object has any Run Properties, if not add RunProperties to it.
				if(objRun.Elements<DocumentFormat.OpenXml.Wordprocessing.RunProperties>().Count() == 0)
					objRun.PrependChild<DocumentFormat.OpenXml.Wordprocessing.RunProperties>(new DocumentFormat.OpenXml.Wordprocessing.RunProperties());
				// Get the first Run Properties Element for the run.
				DocumentFormat.OpenXml.Wordprocessing.RunProperties objRunProperties = objRun.Elements<DocumentFormat.OpenXml.Wordprocessing.RunProperties>().First();
				// Set the properties for the run
				if(parBold)
					objRunProperties.Bold = new DocumentFormat.OpenXml.Wordprocessing.Bold();
				if(parItalic)
					objRunProperties.Italic = new DocumentFormat.OpenXml.Wordprocessing.Italic();
				if(parUnderline)
					objRunProperties.Underline = new DocumentFormat.OpenXml.Wordprocessing.Underline() { Val = DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single };
				}

			// Insert the text in the objRun of the objParagraph
			DocumentFormat.OpenXml.Wordprocessing.Text objText = objRun.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text());
			objText.Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
			objText.Text = parText2Write;
			//objRun.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(parText2Write));
			}


		public static void InsertImage(WordprocessingDocument parWPdocument, int parParagraphLevel, int parPictureSeqNo, string parImageURL)
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
				else if (parImageURL.IndexOf("/") > 0)
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

				// Define the objBody of the document
				Body objBody = objMainDocumentPart.Document.Body;
				Paragraph objParargraph = oxmlDocument.Insert_BodyTextParagraph(ref objBody, parParagraphLevel);
				DocumentFormat.OpenXml.Wordprocessing.Run objRun = objParargraph.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run()); 
				DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
				// Prepare the Anchor object
				Wp.Anchor objAnchor = new Wp.Anchor() { DistanceFromTop = (UInt32Value) 0U, DistanceFromBottom = (UInt32Value) 0U, DistanceFromLeft = (UInt32Value) 114300U, DistanceFromRight = (UInt32Value) 114300U, SimplePos = false, RelativeHeight = (UInt32Value) 251658240U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true, EditId = "09096F23", AnchorId = "411CCDA1" };
				Wp.SimplePosition objSimplePosition = new Wp.SimplePosition() { X = 0L, Y = 0L };

				Wp.HorizontalPosition objHorizontalPosition = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Column };
				Wp.PositionOffset objHorizontalPositionOffset = new Wp.PositionOffset();
				objHorizontalPositionOffset.Text = "393065";
				objHorizontalPosition.Append(objHorizontalPositionOffset);

				Wp.VerticalPosition ObjVerticalPosition = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph };
				Wp.PositionOffset objVerticalPositionOffset = new Wp.PositionOffset();
				objVerticalPositionOffset.Text = "377190";
				ObjVerticalPosition.Append(objVerticalPositionOffset);

				Wp.Extent objExtent = new Wp.Extent() { Cx = 6010275L, Cy = 6010275L };
				Wp.EffectExtent objEffectExtent = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 9525L, BottomEdge = 9525L };
				Wp.WrapTopBottom objWrapTopBottom = new Wp.WrapTopBottom();

				Wp.DocProperties objDocProperties = new Wp.DocProperties() { Id = Convert.ToUInt32(parPictureSeqNo), Name = "Picture " + parPictureSeqNo.ToString() };

				Wp.NonVisualGraphicFrameDrawingProperties objNonVisualGraphicFrameDrawingProperties = new Wp.NonVisualGraphicFrameDrawingProperties();

				A.GraphicFrameLocks objGraphicFrameLocks = new A.GraphicFrameLocks() { NoChangeAspect = true };
				objGraphicFrameLocks.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

				objNonVisualGraphicFrameDrawingProperties.Append(objGraphicFrameLocks);

				A.Graphic objGgraphic = new A.Graphic();
				objGgraphic.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

				A.GraphicData objGraphicData = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

				Pic.Picture objPicture = new Pic.Picture();
				objPicture.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

				Pic.NonVisualPictureProperties objNonVisualPictureProperties = new Pic.NonVisualPictureProperties();
				Pic.NonVisualDrawingProperties objNonVisualDrawingProperties = new Pic.NonVisualDrawingProperties()
					{ Id = Convert.ToUInt32(parPictureSeqNo), Name = imgFileName };

				Pic.NonVisualPictureDrawingProperties objNonVisualPictureDrawingProperties = new Pic.NonVisualPictureDrawingProperties();

				objNonVisualPictureProperties.Append(objNonVisualDrawingProperties);
				objNonVisualPictureProperties.Append(objNonVisualPictureDrawingProperties);

				Pic.BlipFill objBlipFill = new Pic.BlipFill();

				A.Blip objBlip = new A.Blip() { Embed = relationshipID };
				A.BlipExtensionList objBlipExtensionList = new A.BlipExtensionList();

				A.BlipExtension objBlipExtension = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };
				A14.UseLocalDpi objUseLocalDpi = new A14.UseLocalDpi() { Val = false };
				objUseLocalDpi.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
				objBlipExtension.Append(objUseLocalDpi);
				objBlipExtensionList.Append(objBlipExtension);
				objBlip.Append(objBlipExtensionList);

				A.Stretch objStretch = new A.Stretch();
				A.FillRectangle objFillRectangle = new A.FillRectangle();
				objStretch.Append(objFillRectangle);
				objBlipFill.Append(objBlip);
				objBlipFill.Append(objStretch);

				Pic.ShapeProperties objShapeProperties = new Pic.ShapeProperties();

				A.Transform2D objTransform2D = new A.Transform2D();
				A.Offset objOffset = new A.Offset() { X = 0L, Y = 0L };
				A.Extents objExtents = new A.Extents() { Cx = 6010275L, Cy = 6010275L };

				objTransform2D.Append(objOffset);
				objTransform2D.Append(objExtents);

				A.PresetGeometry objPresetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
				A.AdjustValueList objAdjustValueList = new A.AdjustValueList();

				objPresetGeometry.Append(objAdjustValueList);

				objShapeProperties.Append(objTransform2D);
				objShapeProperties.Append(objPresetGeometry);

				objPicture.Append(objNonVisualPictureProperties);
				objPicture.Append(objBlipFill);
				objPicture.Append(objShapeProperties);

				objGraphicData.Append(objPicture);

				objGgraphic.Append(objGraphicData);

				Wp14.RelativeWidth objRelativeWidth = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Page };
				Wp14.PercentageWidth objPercentageWidth = new Wp14.PercentageWidth();
				objPercentageWidth.Text = "0";
				objRelativeWidth.Append(objPercentageWidth);

				Wp14.RelativeHeight objRelativeHeight = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Page };
				Wp14.PercentageHeight objPercentageHeight = new Wp14.PercentageHeight();
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

			}
			catch(Exception exc)
				{
                    ErrorLogMessage = "The image file: [" + parImageURL + "] couldn't be located and was not inserted. \r\n " + exc.Message + " in " + exc.Source;
				Console.WriteLine(ErrorLogMessage);
				return;
				}
	}

		class oxmlWorkbook
			{
			}
		}
	}
