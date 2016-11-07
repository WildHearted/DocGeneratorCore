using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
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
using DocGeneratorCore.Database.Classes;

// Reference sources:
// https://msdn.microsoft.com/en-us/library/office/ff478255.aspx (Baic Open XML Documents)
// https://msdn.microsoft.com/en-us/library/dd469465%28v=office.12%29.aspx (Examples with merging and Presentations)
// http://blogs.msdn.com/b/vsod/archive/2012/02/18/how-to-create-a-document-from-a-template-dotx-dotm-and-attach-to-it-using-open-xml-sdk.aspx (Example of creating a new document based on a .dotx template.)
// (Example to Replace text in a document) http://www.codeproject.com/Tips/666751/Use-OpenXML-to-Create-a-Word-Document-from-an-Exis
// (Structure of an oXML document) https://msdn.microsoft.com/en-us/library/office/gg278308.aspx
namespace DocGeneratorCore
	{
	public enum enumTableRowMergeType
		{
		Restart,
		Continue,
		None
		}

	public enum enumDocumentOrWorkbook
		{
		Document,
		Workbook
		}

	public class oxmlDocumentWorkbook
		{
		// Object Properties
		private string _localPath = "";
		public string LocalPath
			{
			get { return this._localPath; }
			private set { this._localPath = value; }
			}

		private string _fileName = "";
		public string Filename
			{
			get { return this._fileName; }
			private set { this._fileName = value; }
			}

		private string _localURI = "";
		public string LocalURI
			{
			get { return this._localURI; }
			private set { this._localURI = value; }
			}

		private enumDocumentOrWorkbook _documentOrWorkbook;
		public enumDocumentOrWorkbook DocumentOrWorkbook
			{
			get { return this._documentOrWorkbook; }
			set { this._documentOrWorkbook = value; }
			}

		//===g
		//++ CreateDocWbkFromTemplate
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
			enumDocumentTypes parDocumentType,
			ref CompleteDataSet parDataSet)
			{
			string ErrorLogMessage = "";
			this.DocumentOrWorkbook = parDocumentOrWorkbook;

			//- Derive the file name of the template document

			string templateFileName = parTemplateURL.Substring(parTemplateURL.LastIndexOf("/") + 1,
				(parTemplateURL.Length - parTemplateURL.LastIndexOf("/")) - 1);

			//- Check if the DocGenerator Template Directory Exist and that it is accessable Configure and validate for the relevant Template
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

			//- Check if the required **template file** exist in the template directory
			if(File.Exists(templateDirectory + templateFileName))
				{
				// If the the template exist just proceed...
				Console.WriteLine("\t\t\t The template already exist and are ready for use: " + templateDirectory + templateFileName);
				}
			else
				{
				// Download the relevant template from SharePoint
				WebClient objWebClient = new WebClient();

				objWebClient.Credentials = parDataSet.SDDPdatacontext.Credentials;
				//objWebClient.Credentials = new NetworkCredential(
				//	userName: Properties.AppResources.DocGenerator_AccountName,
				//	password: Properties.AppResources.DocGenerator_Account_Password,
				//	domain: Properties.AppResources.DocGenerator_AccountDomain);
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
				// Open the new Excel workbook which is still in .xltx format to save it as a .xlsx file
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

	public class oxmlDocument:oxmlDocumentWorkbook
		{
		//===g
		//++ Construct_Heading

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

				objParagraphStyleID.Val = Properties.AppResources.Document_StyleName_Heading_Unnumbered;
			else
				objParagraphStyleID.Val = Properties.AppResources.Document_StyleName_Headings + parHeadingLevel.ToString();

			objParagraphProperties.Append(objParagraphStyleID);
			objParagraph.Append(objParagraphProperties);
			if(parBookMark != null)
				{
				BookmarkStart objBookmarkStart = new BookmarkStart();
				objBookmarkStart.Name = parBookMark;
				string bookMarkID = parBookMark.Substring(parBookMark.IndexOf("_", 0) + 1, parBookMark.Length - parBookMark.IndexOf("_", 0) - 1);
				objBookmarkStart.Id = bookMarkID;
				objParagraph.Append(objBookmarkStart);

				BookmarkEnd objBookmarkEnd = new BookmarkEnd();
				objBookmarkEnd.Id = bookMarkID;
				objParagraph.Append(objBookmarkEnd);
				}

			return objParagraph;
			}

		//===g
		//++ Construct_Paragraph
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
			bool parIsTableParagraph = false,
			bool parIsTableHeader = false,
			bool parFirstRow = false,
			bool parLastRow = false,
			bool parFirstColumn = false,
			bool parLastColumn = false)
			{

			if(parBodyTextLevel > 9)
				parBodyTextLevel = 9;

			//Construct a Paragraph instance.
			Paragraph objParagraph = new Paragraph();
			objParagraph.RsidParagraphAddition = "00CE4AAA";
			objParagraph.RsidRunAdditionDefault = "00551B06";

			//Construct a ParagraphProperties object instance for the paragraph.
			ParagraphProperties objParagraphProperties = new ParagraphProperties();
			//Construct the ParagraphStyle to be used
			ParagraphStyleId objParagraphStyleID = new ParagraphStyleId();
			if(parIsTableParagraph)
				{
				if(parIsTableHeader)
					objParagraphStyleID.Val = Properties.AppResources.Document_StyleName_TableHeaderText;
				else
					objParagraphStyleID.Val = Properties.AppResources.Document_StyleName_TableBodyText;

				//ConditionalFormatStyle objConditionalFormatStyle = new ConditionalFormatStyle();
				//objConditionalFormatStyle.Val = "000000000000";
				//objConditionalFormatStyle.FirstRow = parFirstRow;
				//objConditionalFormatStyle.LastRow = parLastRow;
				//objConditionalFormatStyle.FirstColumn = parFirstColumn;
				//objConditionalFormatStyle.LastColumn = parLastColumn;
				//if (parFirstRow && parFirstColumn)
				//	objConditionalFormatStyle.FirstRowFirstColumn = true;
				//else
				//	objConditionalFormatStyle.FirstRowFirstColumn = false;
				//if (parFirstRow && parLastColumn)
				//	objConditionalFormatStyle.FirstRowLastColumn = true;
				//else
				//	objConditionalFormatStyle.FirstRowLastColumn = false;
				//if (parLastRow && parFirstColumn)
				//	objConditionalFormatStyle.LastRowFirstColumn = true;
				//else
				//	objConditionalFormatStyle.LastRowFirstColumn = false;
				//if (parLastRow && parLastColumn)
				//	objConditionalFormatStyle.LastRowLastColumn = true;
				//else
				//	objConditionalFormatStyle.LastRowLastColumn = false;
				//objConditionalFormatStyle.OddHorizontalBand = false;
				//objConditionalFormatStyle.EvenHorizontalBand = false;
				//objConditionalFormatStyle.OddVerticalBand = false;
				//objConditionalFormatStyle.EvenVerticalBand = false;
				//objParagraphProperties.Append(objConditionalFormatStyle);
				}
			else
				{
				//-The change required that only one level of DD Body Text is used instead 9 levels, therefore
				//-it was removed the levels from code.
				objParagraphStyleID.Val = Properties.AppResources.Document_StyleName_BodyText;
				}
			objParagraphProperties.Append(objParagraphStyleID);
			objParagraph.Append(objParagraphProperties);
			return objParagraph;
			}

		//===g
		//++ Construct_BulletNumberParagraph

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
		public static Paragraph Construct_BulletParagraph(
			int parBulletLevel,
			bool parIsTableBullet = false)
			{
			string bulletStyle = string.Empty;
			Paragraph objParagraph = new Paragraph();
			try
				{
				//-Flatten the cascade when it exceeds 3 levels
				int bulletLevel = 1;
				if (parBulletLevel < 1)
					bulletLevel = 1;
				else if (parBulletLevel > 3)
					throw new InvalidContentFormatException("The bullet list exceeds 3 levels. The Dimension Data standard template "
						+ "limits cascaded bullets to three levels. Please review the content and correct the error.");
				else
					bulletLevel = parBulletLevel;

				//Configure the **Bullet Style** 

				if (parIsTableBullet)
					{
					//-Insert a **Table BULLET** (not a normal paragraph) bullet
					bulletStyle = Properties.AppResources.Document_StyleName_TableBullet + bulletLevel.ToString();
					}
				else
					{//-Insert a **normal bullet style** (not a table) bullet
					bulletStyle = Properties.AppResources.Document_StyleName_Bullet + bulletLevel.ToString();
					}

				//-Create a new Paragraph instance
				ParagraphProperties objParagraphProperties = new ParagraphProperties();
				ParagraphStyleId objParagraphStyleID = new ParagraphStyleId();
				objParagraphStyleID.Val = bulletStyle;
				objParagraphProperties.Append(objParagraphStyleID);

				objParagraph.Append(objParagraphProperties);
				}
			catch (InvalidContentFormatException exc)
				{
				throw new InvalidContentFormatException(exc.Message);
				}

			return objParagraph;
			}

		//++ Construct_NumberParagraph
		/// <summary>
		/// Use this method to insert a new Bullet Text Paragraph
		/// </summary>
		/// <param name="parNumberLevel">
		/// Pass an integer between 0 and 9 depending of the level of the body text level that need to be inserted.
		/// </param>
		/// <param name="parRestartNumbering">
		/// Indicates whether the numbering for numbered lists need to be restarted.
		/// </param>
		/// <param name="parNumberingId">
		/// The NumberingID is Unique and is required to enable the creation of Restarting of the numbering list.</param>
		/// <param name="parIsTableNumber">
		///  Pass boolean value of TRUE if the paragraph is for a Table else leave blank because the default value is FALSE.
		/// <returns> Paragraph object</returns>
		public static Paragraph Construct_NumberParagraph(
			ref MainDocumentPart  parMainDocumentPart,
			int parNumberLevel,
			bool parRestartNumbering = false,
			int parNumberingId = 0,
			bool parIsTableNumber = false)
			{
			int numberLevel = 1;
			string numberStyle = string.Empty;
			Paragraph objParagraph = new Paragraph();
			try
				{
				//+Report a Content Error numbering exceeds 3 levels
				if (parNumberLevel > 3)
					{
					throw new InvalidContentFormatException("The numbering list exceeds 3 levels. The Dimension Data standard template "
						+ "limits cascaded numbering to three levels. Please review the content and correct the error.");
					}
				else if (parNumberLevel < 1)
					numberLevel = 1;
				else
					numberLevel = parNumberLevel;

				//+Configure the **Number Style** 
				if (parIsTableNumber)
					{//- Insert a **Table BULLET** (not a normal paragraph) bullet
					numberStyle = Properties.AppResources.Document_StyleName_TableNumber + numberLevel.ToString();
					}
				else
					{//-Insert a **normal bullet style** (not a table) bullet
					numberStyle = Properties.AppResources.Document_StyleName_Number + numberLevel.ToString();
					}

				ParagraphProperties objParagraphProperties = new ParagraphProperties();
				ParagraphStyleId objParagraphStyleID = new ParagraphStyleId();
				objParagraphStyleID.Val = numberStyle;
				objParagraphProperties.Append(objParagraphStyleID);

				//-Add **numbering restart** if renumbering must be restart.
				if ( parRestartNumbering)
					{
					//-|When the numbering needs to restart a new **NumberingInstance** must be created in the *NumberingDefinitionsPart*
					NumberingDefinitionsPart objNumberingDefinitionPart = parMainDocumentPart.NumberingDefinitionsPart;
					Numbering objNumbering = objNumberingDefinitionPart.Numbering;
					Int32Value abstractNumeringId = 24;
					if (parIsTableNumber)
						abstractNumeringId = 16;

					//-| Create the new **NumberingInstance** for the restart of the numbering
					NumberingInstance objNumberingInstance = new NumberingInstance();
					objNumberingInstance.NumberID = Convert.ToInt32(parNumberingId);
					//-|Create the **AbstractNumberingId** to link the NumberingInstance to the *AbstractNumber* 
					AbstractNumId objAbstractNumID = new AbstractNumId();
					objAbstractNumID.Val = abstractNumeringId;
					objNumberingInstance.Append(objAbstractNumID);
					//!Create all the **LevelOverrides** to be applied to the *Numbering Instance*.
					//-|Add the **NumberingInstance** to the Numbering
					objNumbering.Append(objNumberingInstance);
					//-|Override Level 0
					LevelOverride objLevelOveride0 = new LevelOverride() { LevelIndex = 0 };
					StartOverrideNumberingValue objStartOverrideNumberingValue0 = new StartOverrideNumberingValue() { Val = 1 };
					objLevelOveride0.Append(objStartOverrideNumberingValue0);
					objNumberingInstance.Append(objLevelOveride0);

					//-|Override Level 1
					LevelOverride objLevelOveride1 = new LevelOverride() { LevelIndex = 1 };
					StartOverrideNumberingValue objStartOverrideNumberingValue1 = new StartOverrideNumberingValue() { Val = 1 };
					objLevelOveride1.Append(objStartOverrideNumberingValue1);
					objNumberingInstance.Append(objLevelOveride1);

					//-|Override Level 2
					LevelOverride objLevelOveride2 = new LevelOverride() { LevelIndex = 2 };
					StartOverrideNumberingValue objStartOverrideNumberingValue2 = new StartOverrideNumberingValue() { Val = 1 };
					objLevelOveride2.Append(objStartOverrideNumberingValue2);
					objNumberingInstance.Append(objLevelOveride2);

					//-|Override Level 3
					LevelOverride objLevelOveride3 = new LevelOverride() { LevelIndex = 3 };
					StartOverrideNumberingValue objStartOverrideNumberingValue3 = new StartOverrideNumberingValue() { Val = 1 };
					objLevelOveride3.Append(objStartOverrideNumberingValue3);
					objNumberingInstance.Append(objLevelOveride3);

					//-|Override Level 4
					LevelOverride objLevelOveride4 = new LevelOverride() { LevelIndex = 4 };
					StartOverrideNumberingValue objStartOverrideNumberingValue4 = new StartOverrideNumberingValue() { Val = 1 };
					objLevelOveride4.Append(objStartOverrideNumberingValue4);
					objNumberingInstance.Append(objLevelOveride4);

					//-|Override Level 5
					LevelOverride objLevelOveride5 = new LevelOverride() { LevelIndex = 5 };
					StartOverrideNumberingValue objStartOverrideNumberingValue5 = new StartOverrideNumberingValue() { Val = 1 };
					objLevelOveride5.Append(objStartOverrideNumberingValue5);
					objNumberingInstance.Append(objLevelOveride5);

					//-|Override Level 6
					LevelOverride objLevelOveride6 = new LevelOverride() { LevelIndex = 6 };
					StartOverrideNumberingValue objStartOverrideNumberingValue6 = new StartOverrideNumberingValue() { Val = 1 };
					objLevelOveride6.Append(objStartOverrideNumberingValue6);
					objNumberingInstance.Append(objLevelOveride6);

					//-|Override Level 7
					LevelOverride objLevelOveride7 = new LevelOverride() { LevelIndex = 7 };
					StartOverrideNumberingValue objStartOverrideNumberingValue7 = new StartOverrideNumberingValue() { Val = 1 };
					objLevelOveride7.Append(objStartOverrideNumberingValue7);
					objNumberingInstance.Append(objLevelOveride7);

					//-|Override Level 8
					LevelOverride objLevelOveride8 = new LevelOverride() { LevelIndex = 8 };
					StartOverrideNumberingValue objStartOverrideNumberingValue8 = new StartOverrideNumberingValue() { Val = 1 };
					objLevelOveride8.Append(objStartOverrideNumberingValue8);
					objNumberingInstance.Append(objLevelOveride8);

					//-|Define the Paragraph properties referencing the *NumberingInstance*
					//-| Define a **NumberingLevelReference**
					NumberingLevelReference objNumberingLevelReference = new NumberingLevelReference();
					objNumberingLevelReference.Val = 0;
					//-Define the **NumberingId** which is the same number as the *NumberingInstance* creating the refenece to the NI.
					NumberingId objNumberingId = new NumberingId();
					objNumberingId.Val = parNumberingId;
					//-|Define a new **NumberingProperty** and add/append the *NumberingLevelReference* and *NumberId* objects
					NumberingProperties objNumberingProperties = new NumberingProperties();
					objNumberingProperties.Append(objNumberingLevelReference);
					objNumberingProperties.Append(objNumberingId);
					//-|Append the **NumberingProperties** to the *ParagraphProperties*
					objParagraphProperties.Append(objNumberingProperties);
					}
				//-|Append the **ParagrpahProperties** to the *paragraph*
				objParagraph.Append(objParagraphProperties);
				}
			catch (InvalidContentFormatException exc)
				{
				throw new InvalidContentFormatException(exc.Message);
				}

			return objParagraph;
			}

		//===R
		//++Construct_Error

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
			objParagraphStyleID.Val = Properties.AppResources.Document_StyleName_ContentErrorText;
			objParagraphProperties.Append(objParagraphStyleID);
			objParagraph.Append(objParagraphProperties);
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
			objRun = oxmlDocument.Construct_RunText(parText);
			objParagraph.Append(objRun);
			return objParagraph;
			}


		//===g
		//++ Construct Caption
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
			string parCaptionText,
			DocumentFormat.OpenXml.Wordprocessing.Run parImageRun = null
			)
			{
			//Create a Paragraph instance.
			Paragraph objParagraph = new Paragraph();
			// Create the Paragraph Properties instance.
			ParagraphProperties objParagraphProperties = new ParagraphProperties();
			ParagraphStyleId objParagraphStyleID = new ParagraphStyleId();
			if(parCaptionType == "Table")
				{ objParagraphStyleID.Val = Properties.AppResources.Document_StyleName_Caption_Table; }
			else
				{ objParagraphStyleID.Val = Properties.AppResources.Document_StyleName_Caption_Figure; }
			objParagraphProperties.Append(objParagraphStyleID);

			//Append the ParagraphProperties to the Paragraph
			objParagraph.Append(objParagraphProperties);
			if(parCaptionType == "Image"
			&& parImageRun != null)
				{
				objParagraph.Append(parImageRun);
				}

			// Create the Caption Run Object
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
			DocumentFormat.OpenXml.Wordprocessing.Text objText = new DocumentFormat.OpenXml.Wordprocessing.Text();
			objText.Space = SpaceProcessingModeValues.Preserve;
			objText.Text = parCaptionText;
			objRun.Append(objText);
			objParagraph.Append(objRun);

			return objParagraph;
			}


		//===g
		//++ Construct_RunText

		public static DocumentFormat.OpenXml.Wordprocessing.Run Construct_RunText(
				string parText2Write,
				bool parIsError = false,
				bool parIsNewSection = false,
				String parContentLayer = "None",
				bool parBold = false,
				bool parItalic = false,
				bool parUnderline = false,
				bool parSubscript = false,
				bool parSuperscript = false,
				bool parStrikeTrough = false)
			{
			// Create a new Run object in the objParagraph
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();

			if(parIsNewSection)
				{
				LastRenderedPageBreak objLastRenderedPageBreak = new LastRenderedPageBreak();
				objRun.Append(objLastRenderedPageBreak);
				}
			else //- if(!parIsNewSection)
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


		//===g
		//++ InsertImage
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
		public static DocumentFormat.OpenXml.Wordprocessing.Run Insert_Image(
			ref MainDocumentPart parMainDocumentPart,
			uint parEffectivePageWidthDxa,
			uint parEffectivePageHeightDxa,
			int parParagraphLevel,
			int parPictureSeqNo,
			string parImageURL,
			int parImageHeight,
			enumWidthHeightType parImageHeightType,
			int parImageWidth,
			enumWidthHeightType parImageWidthType)

			{

			string ErrorLogMessage = "";
			string imageType = "";
			string relationshipID = "";
			string imageFileName = "";
			string imageDirectory = Path.GetFullPath("\\") + DocGeneratorCore.Properties.AppResources.LocalImagePath;
			try
				{
				//-|Download the image from SharePoint if it is a http:// based image
				imageType = parImageURL.Substring(parImageURL.LastIndexOf(".") + 1, (parImageURL.Length - parImageURL.LastIndexOf(".") - 1));
				if(parImageURL.IndexOf("\\") < 0)
					{
					ErrorLogMessage = "";
					//-|Derive the file name of the image file
					imageFileName = parImageURL.Substring(parImageURL.LastIndexOf("/") + 1, (parImageURL.Length - parImageURL.LastIndexOf("/")) - 1);
					//-|Construct the local name for the New Image file
					imageFileName = imageFileName.Replace("%20", "_");
					imageFileName = imageFileName.Replace(" ", "-");
					Console.WriteLine("\t\t\t local imageFileName: [{0}]", imageFileName);
					//-|Check if the DocGenerator Image Directory Exist and that it is accessable
					try
						{
						if(Directory.Exists(@imageDirectory))
							{
							Console.WriteLine("\t\t\t\t The imageDirectory [" + imageDirectory + "] exist and are ready to be used.");
							}
						else
							{
							DirectoryInfo templateDirInfo = Directory.CreateDirectory(@imageDirectory);
							Console.WriteLine("\t\t\t\t The imageDirectory [" + imageDirectory + "] was created and are ready to be used.");
							}
						}
					catch(UnauthorizedAccessException exc)
						{
						ErrorLogMessage = "The current user: [" + System.Security.Principal.WindowsIdentity.GetCurrent().Name +
						"] does not have the required security permissions to access the Image directory at: " + imageDirectory +
						"\r\n " + exc.Message + " in " + exc.Source;
						Console.WriteLine("\t\t\t\t" + ErrorLogMessage);
						throw new InvalidImageFormatException(ErrorLogMessage);
						}
					catch(NotSupportedException exc)
						{
						ErrorLogMessage = "The path of Image directory [" + imageDirectory + "] contains invalid characters. Ensure that the path is valid and  contains legible path characters only. \r\n " + exc.Message + " in " + exc.Source;
						Console.WriteLine("\t\t\t\t" + ErrorLogMessage);
						throw new InvalidImageFormatException(ErrorLogMessage);
						}
					catch(DirectoryNotFoundException exc)
						{
						ErrorLogMessage = "The path of Image directory [" + imageDirectory + "] is invalid. Check that the drive is mapped and exist /r/n " + exc.Message + " in " + exc.Source;
						Console.WriteLine("\t\t\t" + ErrorLogMessage);
						throw new InvalidImageFormatException(ErrorLogMessage);
						}

					//+| Check if the Image file already exist in the local Image directory
					if(File.Exists(imageDirectory + "\\" + imageFileName))
						{
						//-|If the the image file exist just proceed...
						Console.WriteLine("\t\t\t\t The image already exist, just use it:" + imageDirectory + "\\" + imageFileName);
						}
					else //-|If the image doesn't exist, then download it...
						{
						//-|Download the relevant image from SharePoint
						WebClient objWebClient = new WebClient();
						objWebClient.Credentials = new NetworkCredential(
							userName: Properties.AppResources.DocGenerator_AccountName,
							password: Properties.AppResources.DocGenerator_Account_Password,
							domain: Properties.AppResources.DocGenerator_AccountDomain);
						try
							{
							objWebClient.DownloadFile(parImageURL, imageDirectory + "\\" + imageFileName);
							}
						catch(WebException exc)
							{
							ErrorLogMessage = "The image file could not be downloaded from SharePoint List [" + parImageURL + "]. " +
								"\n - Check that the image exist in SharePoint.\n " + exc.Message + " in " + exc.Source;
							Console.WriteLine("\t\t\t\t" + ErrorLogMessage);
							throw new InvalidImageFormatException(ErrorLogMessage);
							}
						}

					Console.WriteLine("\t\t\t\t {2} this Image:[{0}] exist in this directory:[{1}]", imageFileName, imageDirectory, File.Exists(imageDirectory + "\\" + imageFileName));
					parImageURL = imageDirectory + imageFileName;
					}
				else //-|if(parImageURL.IndexOf("/") > 0) // if it is a local file (not an URL...)
					{
					imageFileName = parImageURL.Substring(parImageURL.LastIndexOf("\\") + 1, (parImageURL.Length - parImageURL.LastIndexOf("\\") - 1));
					}

				//+| Get the image's dimensions
				var img = System.Drawing.Image.FromFile(parImageURL);
				//-| https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
				int imagePIXELheight = img.Height;
				int imagePIXELwidth = img.Width;

				Console.WriteLine("\t\t\t\t Image dimensions (H x W): {0} x {1} pixels per Inch", imagePIXELheight, imagePIXELwidth);
				Console.WriteLine("\t\t\t\t Horizontal Resolution...: {0} pixels per inch", img.HorizontalResolution);

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

				//-|If the image is wider than the Effective Width of the page
				double imageDXAwidth = 0;
				double imageDXAheight = 0;
				if ((imagePIXELwidth * 20) > parEffectivePageWidthDxa)
					{
					imageDXAwidth = Math.Round((((parEffectivePageWidthDxa / (imagePIXELwidth * 20D)) * imagePIXELwidth) * 20D) * 635D, 0);
					//imageDXAwidth = Math.Round((((parEffectivePageWidthDxa / (imagePIXELwidth * 20D)) * imagePIXELwidth) * 20D),0);
					imageDXAheight = Math.Round((((parEffectivePageWidthDxa / (imagePIXELwidth * 20D)) * imagePIXELheight) * 20D) * 635D, 0);
					//imageDXAheight = Math.Round((((parEffectivePageWidthDxa / (imagePIXELwidth * 20D)) * imagePIXELheight) * 20D),0);
					}
				else if ((imagePIXELheight * 20) > parEffectivePageHeightDxa)
					{
					imageDXAwidth = Math.Round((((parEffectivePageHeightDxa / (imagePIXELheight * 20D)) * imagePIXELwidth) * 20D) * 635D, 0);
					//imageDXAwidth = Math.Round((((parEffectivePageHeightDxa / (imagePIXELheight * 20D)) * imagePIXELwidth) * 20D),0);
					imageDXAheight = Math.Round((((parEffectivePageHeightDxa / (imagePIXELheight * 20D)) * imagePIXELheight) * 20D) * 635D, 0);
					//imageDXAheight = Math.Round((((parEffectivePageHeightDxa / (imagePIXELheight * 20D)) * imagePIXELheight) * 20D),0);
					}
				else
					{
					imageDXAwidth = Math.Round(imagePIXELwidth * 635D, 0);
					//imageDXAwidth = imagePIXELwidth;
					imageDXAheight = Math.Round(imagePIXELheight * 635D, 0);
					//imageDXAheight = imagePIXELheight;
					}
				Console.WriteLine("\t\t\t\t imageDXAwidth: {0}", imageDXAwidth);
				Console.WriteLine("\t\t\t\t imageDXAheight: {0}", imageDXAheight);

				// Define the Drawing Object instance
				DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
				
				DrwWp.Inline objInline = new DrwWp.Inline();
				objInline.DistanceFromTop = (UInt32)0U;
				objInline.DistanceFromBottom = (UInt32Value)0U;
				objInline.DistanceFromLeft = (UInt32Value)0U;
				objInline.DistanceFromRight = (UInt32Value)0U;
				objInline.AnchorId = "13286FC3";
				objInline.EditId = "114014AF";
				//!Possible issue/bug

				//-|Define Extent
				DrwWp.Extent objExtent = new DrwWp.Extent();
				objExtent.Cx = Convert.ToInt64(imageDXAwidth);
				objExtent.Cy = Convert.ToInt64(imageDXAheight);

				//-|Define Extent Effects
				DrwWp.EffectExtent objEffectExtent = new DrwWp.EffectExtent();
				objEffectExtent.LeftEdge = 0L;
				objEffectExtent.TopEdge = 0L;
				objEffectExtent.RightEdge = 2540L;
				objEffectExtent.BottomEdge = 0L;

				//-|Define the **Document Properties** by linking the image to identifier of the imaged where it was inserted in the MainDocumentPart.
				DrwWp.DocProperties objDocProperties = new DrwWp.DocProperties();
				objDocProperties.Id = Convert.ToUInt32(10000 + parPictureSeqNo);
				objDocProperties.Name = "Picture " + (10000 + parPictureSeqNo).ToString();
				
				//-|Define the **Graphic Frame Locks**
				Drw.GraphicFrameLocks objGraphicFrameLocks = new Drw.GraphicFrameLocks();
				objGraphicFrameLocks.NoChangeAspect = true;
				objGraphicFrameLocks.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

				//-|Define the **Non Visual Graphic Frame Drawing Properties**
				DrwWp.NonVisualGraphicFrameDrawingProperties objNonVisualGraphicFrameDrawingProperties = new DrwWp.NonVisualGraphicFrameDrawingProperties();
				objNonVisualGraphicFrameDrawingProperties.Append(objGraphicFrameLocks);

				//-|Define a **Graphic** object instance
				Drw.Graphic objGraphic = new Drw.Graphic();
				objGraphic.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
				Drw.GraphicData objGraphicData = new Drw.GraphicData();
				objGraphicData.Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture";

				//-|Define the Picture
				Pic.Picture objPicture = new Pic.Picture();
				objPicture.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

				//-|Define the Picture's NonVisual Properties
				Pic.NonVisualPictureProperties objNonVisualPictureProperties = new Pic.NonVisualPictureProperties();

				//-|Define the NonVisual Drawing Properties
				Pic.NonVisualDrawingProperties objNonVisualDrawingProperties = new Pic.NonVisualDrawingProperties();
				objNonVisualDrawingProperties.Id = Convert.ToUInt32(parPictureSeqNo);
				objNonVisualDrawingProperties.Name = imageFileName;

				//-|Define the Picture's NonVisual Picture Drawing Properties
				Pic.NonVisualPictureDrawingProperties objNonVisualPictureDrawingProperties = new Pic.NonVisualPictureDrawingProperties();
				objNonVisualPictureProperties.Append(objNonVisualDrawingProperties);
				objNonVisualPictureProperties.Append(objNonVisualPictureDrawingProperties);

				//-|Define a **Blib Fill** object Instance
				Pic.BlipFill objBlibFill = new Pic.BlipFill();

				//-|Define the Blib
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

				//-|Define how the image is filled
				Drw.Stretch objStretch = new Drw.Stretch();
				Drw.FillRectangle objFillRectangle = new Drw.FillRectangle();
				objStretch.Append(objFillRectangle);
				Pic.BlipFill objBlipFill = new Pic.BlipFill();
				objBlipFill.Append(objBlip);
				objBlipFill.Append(objStretch);

				//-|Define the Picture's Shape Properties
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

				//-|Define the Preset Geometry
				Drw.PresetGeometry objPresetGeometry = new Drw.PresetGeometry();
				objPresetGeometry.Preset = Drw.ShapeTypeValues.Rectangle;
				Drw.AdjustValueList objAdjustValueList = new Drw.AdjustValueList();
				objPresetGeometry.Append(objAdjustValueList);

				objShapeProperties.Append(objTransform2D);
				objShapeProperties.Append(objPresetGeometry);

				//-|Append the Definitions to the Picture Object Instance...
				objPicture.Append(objNonVisualPictureProperties);
				objPicture.Append(objBlipFill);
				objPicture.Append(objShapeProperties);

				//-|Append the the picture object to the Graphic object Instance
				objGraphicData.Append(objPicture);
				objGraphic.Append(objGraphicData);

				objInline.Append(objExtent);
				objInline.Append(objEffectExtent);
				objInline.Append(objDocProperties);
				objInline.Append(objNonVisualGraphicFrameDrawingProperties);
				objInline.Append(objGraphic);

				objDrawing.Append(objInline);

				//-|Define the Run object and append the Drawing object to it...
				DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
				objRun.Append(objDrawing);
				//-|Return the Run object which now contains the complete Image to be added to a Paragraph in the document.
				return objRun;
				}
			catch(Exception exc)
				{
				ErrorLogMessage = "The image file: [" + parImageURL + "] couldn't be located and was not inserted. \r\n " + exc.Message + " in " + exc.Source;
				Console.WriteLine(ErrorLogMessage);
				throw new InvalidImageFormatException(ErrorLogMessage);
				}
			}

		//===g
		//++ InsertHyperlinkImage
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

		public static string Insert_HyperlinkImage(
			ref MainDocumentPart parMainDocumentPart,
			ref CompleteDataSet parDataSet)
			{
			string ErrorLogMessage = "";
			string relationshipID = "";
			string imageFileName = DocGeneratorCore.Properties.AppResources.ClickLinkFileName;
			string imageDirectory = Path.GetFullPath("\\") + DocGeneratorCore.Properties.AppResources.LocalImagePath;
			string imageSharePointURL = Properties.Settings.Default.CurrentURLSharePoint + Properties.Settings.Default.CurrentURLSharePointSitePortion
					+ DocGeneratorCore.Properties.AppResources.ClickLinkImageSharePointURL;

			Console.WriteLine("\t\t\t HyperlinkImageFileName: [{0}]", DocGeneratorCore.Properties.AppResources.ClickLinkFileName);
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
				ErrorLogMessage = "The DocGenerator Account does not have the required security permissions to access the local template directory at: "
					+ imageDirectory + "\r\n " + exc.Message + " in " + exc.Source;
				Console.WriteLine("\t\t\t" + ErrorLogMessage);
				//TODO: insert code to write an error line in the document
				return null;
				}
			catch(NotSupportedException exc)
				{
				ErrorLogMessage = "The path of template directory [" + imageDirectory + "] contains invalid characters. "
					+ "Ensure that the path is valid and  contains legible path characters only. \r\n " + exc.Message + " in " + exc.Source;
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
			if(File.Exists(imageDirectory + imageFileName))
				{
				// If the the image file exist just proceed...
				Console.WriteLine("\t\t\t The image to already exist, just use it:" + imageDirectory + imageFileName);
				}
			else // If the image doesn't exist already, then download it...
				{
				// Download the relevant image from SharePoint
				WebClient objWebClient = new WebClient();

				objWebClient.Credentials = parDataSet.SDDPdatacontext.Credentials;
				//objWebClient.Credentials = new NetworkCredential(
				//	userName: Properties.AppResources.DocGenerator_AccountName,
				//	password: Properties.AppResources.DocGenerator_Account_Password,
				//	domain: Properties.AppResources.DocGenerator_AccountDomain);

				try
					{
					objWebClient.DownloadFile(address: imageSharePointURL, fileName: imageDirectory + imageFileName);
					}
				catch(WebException exc)
					{
					ErrorLogMessage = "The ClickImage file could not be downloaded from SharePoint List [" + imageSharePointURL + "]. " +
						"\n - Check that the image exist in SharePoint \n - that it is accessible \n - " +
						"and that the network connection is working. \n " + exc.Message + " in " + exc.Source;
					Console.WriteLine("\t\t\t" + ErrorLogMessage);
					return null;
					}
				}

			Console.WriteLine("\t\t\t {2} this Image:[{0}] exist in this directory:[{1}]", imageFileName, imageDirectory, File.Exists(imageDirectory + imageFileName));

			try
				{
				ImagePart objImagePart = parMainDocumentPart.AddImagePart(ImagePartType.Png);
				string hyperlinkImageURL = imageDirectory + imageFileName;

				using(FileStream objFileStream = new FileStream(path: hyperlinkImageURL, mode: FileMode.Open))
					{
					objImagePart.FeedData(objFileStream);
					}
				relationshipID = parMainDocumentPart.GetIdOfPart(part: objImagePart);

				return relationshipID;
				}
			catch(Exception exc)
				{
				ErrorLogMessage = "The image file: [" + Properties.AppResources.ClickLinkFileName + "] couldn't be located and was not inserted. \r\n "
					+ exc.Message + " in " + exc.Source;
				Console.WriteLine(ErrorLogMessage);
				return null;
				}
			}


		//===g
		//++ ConstructClickLinkHyperlink
		public static DocumentFormat.OpenXml.Wordprocessing.Drawing Construct_ClickLinkHyperlink(
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
				HyperlinkRelationship objHyperlinkRelationship = parMainDocumentPart.AddHyperlinkRelationship(
					hyperlinkUri: objUri,
					isExternal: true);
				hyperlinkID = objHyperlinkRelationship.Id;
				}

			// Define a Drawing Object instance
			DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
			// Define the Anchor object
			DrwWp.Anchor objAnchor = new DrwWp.Anchor();
			objAnchor.DistanceFromTop = (UInt32Value)0U;
			objAnchor.DistanceFromBottom = (UInt32Value)0U;
			objAnchor.DistanceFromLeft = (UInt32Value)114300U;
			objAnchor.DistanceFromRight = (UInt32Value)114300U;
			objAnchor.RelativeHeight = (UInt32Value)251659264U;
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


		//===g
		//++ Construct_BookmarkHyperlink
		public static DocumentFormat.OpenXml.Wordprocessing.Paragraph Construct_BookmarkHyperlink(
			int parBodyTextLevel,
			string parBookmarkValue)
			{
			// Create the object instances for the ParagraphProperties
			DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties objParagraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties();
			DocumentFormat.OpenXml.Wordprocessing.ParagraphStyleId objParagraPhStyleID = new DocumentFormat.OpenXml.Wordprocessing.ParagraphStyleId();
			//!Change to Templates required the removal of the Body Text Level
			//- All text will now be added as Body Text
			//- //objParagraPhStyleID.Val = "DDBodyText" + parBodyTextLevel;
			objParagraPhStyleID.Val = Properties.AppResources.Document_StyleName_BodyText;
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


		//===g
		/// <summary>
		/// 
		/// </summary>
		/// <param name="parText2insert">The text that must be inserted into the Run</param>
		/// <param name="parURL">The complete URL that must be inserted into the Run.</param>
		/// <returns>a completely formatted and formulated Hyperlink that can be added to a paragraph.</returns>
			//++ Construct_Hyperlink
		public static DocumentFormat.OpenXml.Wordprocessing.Hyperlink Construct_Hyperlink(
			ref MainDocumentPart parMainDocumentPart,
			string parText2insert,
			string parURL)
			{
			//-|Create the object instances for the Hyperlink.
			DocumentFormat.OpenXml.Wordprocessing.Hyperlink objHyperlink = new DocumentFormat.OpenXml.Wordprocessing.Hyperlink();
			
			try
				{
				//-|Create the Hyperlink index
				Uri objUri = new Uri(parURL);
				string hyperlinkID = "";
				//-| Check if the hyperlink already exist in the document
				HyperlinkRelationship hyperRelationship = parMainDocumentPart.HyperlinkRelationships.Where(h => h.Uri == objUri).FirstOrDefault();
				//-|If **HyperlinkRelationship** does not exist
				if (hyperRelationship == null)
					{
					HyperlinkRelationship objHyperlinkRelationship = parMainDocumentPart.AddHyperlinkRelationship(
						hyperlinkUri: objUri,
						isExternal: true);
					hyperlinkID = objHyperlinkRelationship.Id;
					}
				else //-|The Hyperlink already exist, we just use the Relationship Id
					{
					hyperlinkID = hyperRelationship.Id;
					}

				
				objHyperlink.History = true;
				//-|Insert the HyperlinkID
				objHyperlink.Id = hyperlinkID;

				//-|Create object instances for the **RunProperty**
				DocumentFormat.OpenXml.Wordprocessing.RunProperties objRunProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
				DocumentFormat.OpenXml.Wordprocessing.RunStyle objRunStyle = new DocumentFormat.OpenXml.Wordprocessing.RunStyle();
				objRunStyle.Val = "Hyperlink";
				Spacing objSpacing = new Spacing();
				objSpacing.Val = 14;
				objRunProperties.Append(objRunStyle);
				objRunProperties.Append(objSpacing);

				//-| Create the object instances for the Text in the Hyperlink's Run.
				DocumentFormat.OpenXml.Wordprocessing.Text objText = new DocumentFormat.OpenXml.Wordprocessing.Text();
				objText.Text = parText2insert;
				//-| Create the **Run** object instance
				DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
				objRun.Append(objRunProperties);
				objRun.Append(objText);
				objHyperlink.Append(objRun);
				}
			catch (Exception )
				{
				throw new InvalidContentFormatException("A hyperlink could not be inserted in this position. Please check in the source "
					+ "whether the hyperlink is complete, correctly formed and valid. |" + parURL + "|");
				}
			//-|Return the **Hyperlink** object which now contains the complete Hyperlink and the relevant text.
			return objHyperlink;
			}

		//===g
		//++ ConstructTable
		/// <summary>
		/// 
		/// </summary>
		/// <param name="parFirstColumn"></param>
		/// <param name="parLastColumn"></param>
		/// <param name="parFirstRow"></param>
		/// <param name="parLastRow"></param>
		/// <param name="parNoVerticalBand"></param>
		/// <param name="parNoHorizontalBand"></param>
		/// <returns></returns>
		public static DocumentFormat.OpenXml.Wordprocessing.Table Construct_Table(
			int parTableWidthInDXA,
			bool parFirstColumn = false,
			bool parLastColumn = false,
			bool parFirstRow = false,
			bool parLastRow = false,
			bool parNoVerticalBand = true,
			bool parNoHorizontalBand = false)
			{

			//- Creates a Table instance
			DocumentFormat.OpenXml.Wordprocessing.Table objTable = new DocumentFormat.OpenXml.Wordprocessing.Table();
			//- Create and set the Table Properties instance
			TableProperties objTableProperties = new TableProperties();

			//- Create and add the **Table Style**
			DocumentFormat.OpenXml.Wordprocessing.TableStyle objTableStyle = new DocumentFormat.OpenXml.Wordprocessing.TableStyle();
			objTableStyle.Val = Properties.AppResources.Document_StyleName_Table;
			objTableProperties.Append(objTableStyle);

			//- Define and add the **Table Width**
			TableWidth objTableWidth = new TableWidth();
			if(Properties.AppResources.Document_Table_Width == "")
				{
				objTableWidth.Width = "0";
				objTableWidth.Type = TableWidthUnitValues.Auto;
				}
			else
				{
				objTableWidth.Width = parTableWidthInDXA.ToString();
				objTableWidth.Type = TableWidthUnitValues.Dxa;
				}
			objTableProperties.Append(objTableWidth);

			//-|Define the **Table Layout**
			TableLayout objTableLayout = new TableLayout();
			objTableLayout.Type = TableLayoutValues.Fixed;
			objTableProperties.Append(objTableLayout);

			//-|Define and add the Table Look
			TableLook objTableLook = new TableLook();
			objTableLook.Val = "04A0";
			objTableLook.FirstRow = parFirstRow;
			objTableLook.FirstColumn = parFirstColumn;
			objTableLook.FirstRow = parFirstRow;
			objTableLook.LastColumn = parLastColumn;
			objTableLook.LastRow = parLastRow;
			objTableLook.NoVerticalBand = parNoVerticalBand;
			objTableLook.NoHorizontalBand = parNoHorizontalBand;
			objTableProperties.Append(objTableLook);
			// Append the TableProperties instance to the Table instance
			objTable.Append(objTableProperties);

			return objTable;

			}


		//===g
		//++ ConstructTableGrid

		/// <summary>
		/// Constructs a TableGrid which can then be appended to a Table object.
		/// </summary>
		/// <param name="parColumnWidthList">
		/// Pass a List of integers which contains the width of each table column. (Pixel based).
		/// </param>
		/// <param name="parTableWidthPixels">
		/// Pass the width of the table in pixel</param>
		/// <returns></returns>
		public static DocumentFormat.OpenXml.Wordprocessing.TableGrid ConstructTableGrid(
			List<int> parColumnWidthList,
			int parTableWidthPixels = 0)
			{
			//-Get the **DxaPerPixel ratio from the App Resource file
			//int dxaPerPixel = 0;
			//if(!int.TryParse(Properties.AppResources.Document_DxaPerPixel_Ratio, out dxaPerPixel))
			//	dxaPerPixel = 15;

			//- Create the **TableGrid** object instance
			TableGrid objTableGrid = new TableGrid();
			//- Process all the columns as defined in the parColumnWidthList and create a Column Grid entry per column.
			//-The value of columnItem is a percentage value
			decimal columnWidth = 0;
			foreach(int columnItem in parColumnWidthList)
				{
				GridColumn objGridColumn = new GridColumn();
				columnWidth = parTableWidthPixels * (columnItem / 100m);
				objGridColumn.Width = Convert.ToInt32(columnWidth).ToString();
				objTableGrid.Append(objGridColumn);
				};
			return objTableGrid;
			}

		//===g
		//++ ConstructTableRow
		/// <summary>
		/// 
		/// </summary>
		/// <param name="parIsFirstRow"></param>
		/// <param name="parIsLastRow"></param>
		/// <param name="parIsFirstColumn"></param>
		/// <param name="parIsLastColumn"></param>
		/// <param name="parIsOddHorizontalBand"></param>
		/// <param name="parIsEvenHorizontalBand"></param>
		/// <param name="parHasConditionalStyle"></param>
		/// <returns></returns>
		public static TableRow ConstructTableRow(
			bool parIsFirstRow = false,
			bool parIsLastRow = false,
			bool parIsFirstColumn = false,
			bool parIsLastColumn = false,
			bool parIsOddHorizontalBand = false,
			bool parIsEvenHorizontalBand = false,
			bool parHasConditionalStyle = true)
			{
			//- Create a **TableRow** object instance
			TableRow objTableRow = new TableRow();
			objTableRow.RsidTableRowAddition = "00377A72";
			objTableRow.RsidTableRowProperties = "00377A72";

			//- Create the **TableRowProperties** object instance
			TableRowProperties objTableRowProperties = new TableRowProperties();

			//if required, create and add the Conditional Format Style
			if(parHasConditionalStyle)
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

			return objTableRow;
			}


		//===g
		//++ ConstructTableCell
		/// <summary>
		/// This procedure use the parameters to construct a Cell object and then return the construced Cell as an object to the celler.
		/// </summary>
		/// <param name="parCellWidth">width of the cell in Pixels</param>
		/// <param name="parHasCondtionalFormatting">OPTIONAL, default value = FALSE, determinse whater a Conditional formatting instance will be inserted for the table cell</param>
		/// <param name="parIsFirstRow">OPTIONAL; default = FALSE</param>
		/// <param name="parIsLastRow">OPTIONAL; default = FALSE</param>
		/// <param name="parIsFirstColumn">OPTIONAL; default = FALSE</param>
		/// <param name="parIsLastColumn">OPTIONAL; default = FALSE</param>
		/// <param name="parColumnMerge">OPTIONAL; default = 1, Indicate if and how many columns to the right of the cell, needs to be merged.</param>
		/// <param name="parRowMerge">OPTIONAL; default = 1, Indicate if and how many Rows BELOW the cell, needs to be merged.</param>
		/// <param name="parFirstRowFirstColumn">OPTIONAL; default = FALSE</param>
		/// <param name="parLastRowFirstColumn">OPTIONAL; default = FALSE</param>
		/// <param name="parFirstRowLastColumn">OPTIONAL; default = FALSE</param>
		/// <param name="parLastRowLastColumn">OPTIONAL; default = FALSE</param>
		/// <param name="parEvenHorizontalBand">OPTIONAL; default = FALSE</param>
		/// <param name="parOddHorizontalBand">OPTIONAL; default = FALSE</param>
		/// <returns>returns a suitably consructed TableCell object</returns>
		public static TableCell ConstructTableCell(
			int parCellWidth,
			bool parHasCondtionalFormatting = false,
			bool parIsFirstRow = false,
			bool parIsLastRow = false,
			bool parIsFirstColumn = false,
			bool parIsLastColumn = false,
			int parColumnMerge = 0,
			enumTableRowMergeType parRowMerge = enumTableRowMergeType.None,
			string parVerticalAlignment = "Centre",
			string parHorizontalAlignment = "Left",
			bool parFirstRowFirstColumn = false,
			bool parLastRowFirstColumn = false,
			bool parFirstRowLastColumn = false,
			bool parLastRowLastColumn = false,
			bool parEvenHorizontalBand = false,
			bool parOddHorizontalBand = false)
			{

			//-Create new TableCell object instance that will be returned to the calling instruction.
			TableCell objTableCell = new TableCell();
			//- Create a new TableCellProperty object instance
			TableCellProperties objTableCellProperties = new TableCellProperties();

			ConditionalFormatStyle objConditionalFormatStyle = new ConditionalFormatStyle();
			objConditionalFormatStyle.Val = "001000000000";
			objConditionalFormatStyle.FirstRow = parIsFirstRow;
			objConditionalFormatStyle.LastRow = parIsLastRow;
			objConditionalFormatStyle.FirstColumn = parIsFirstColumn;
			objConditionalFormatStyle.LastColumn = parIsLastColumn;
			if (parIsFirstRow && parIsFirstColumn)
				objConditionalFormatStyle.FirstRowFirstColumn = true;
			else
				objConditionalFormatStyle.FirstRowFirstColumn = false;
			if (parIsFirstRow && parIsLastColumn)
				objConditionalFormatStyle.FirstRowLastColumn = true;
			else
				objConditionalFormatStyle.FirstRowLastColumn = false;
			if (parIsLastRow && parIsFirstColumn)
				objConditionalFormatStyle.LastRowFirstColumn = true;
			else
				objConditionalFormatStyle.LastRowFirstColumn = false;
			if (parIsLastRow && parIsLastColumn)
				objConditionalFormatStyle.LastRowLastColumn = true;
			else
				objConditionalFormatStyle.LastRowLastColumn = false;
			objConditionalFormatStyle.OddHorizontalBand = false;
			objConditionalFormatStyle.EvenHorizontalBand = false;
			objConditionalFormatStyle.OddVerticalBand = false;
			objConditionalFormatStyle.EvenVerticalBand = false;

			objTableCellProperties.Append(objConditionalFormatStyle);

			//-Create the **TableCellWidth** object instance
			TableCellWidth objTableCellWidth = new TableCellWidth();
			//-The parameter value is in DXA
			parCellWidth = Convert.ToInt32(parCellWidth); 
			objTableCellWidth.Width = parCellWidth.ToString();
			objTableCellWidth.Type = TableWidthUnitValues.Dxa;
			objTableCellProperties.Append(objTableCellWidth);

			//-Insert **GridSpan** if required
			if (parColumnMerge > 1)
				{
				GridSpan objGridSpan = new GridSpan();
				objGridSpan.Val = parColumnMerge;
				objTableCellProperties.Append(objGridSpan);
				}

			//-Check if the cell is a merged row...
			if (parRowMerge != enumTableRowMergeType.None)
				{ //-It is merged with a cell on a nother row...
				VerticalMerge objVerticalMerge = new VerticalMerge();
				//-Check if it is the beginning of the vertical merge, and set the value to "restart" if it is..
				if (parRowMerge == enumTableRowMergeType.Restart)
					objVerticalMerge.Val = MergedCellValues.Restart;
				else
					objVerticalMerge.Val = MergedCellValues.Continue;
				//-Append the VerticalMerge object instance
				objTableCellProperties.Append(objVerticalMerge);
				}

			//- Construct the Cell Alignment
			TableCellVerticalAlignment objTableCellVerticalAlignment = new TableCellVerticalAlignment();
			if(parIsFirstRow)
				objTableCellVerticalAlignment.Val = TableVerticalAlignmentValues.Center;
			else
				objTableCellVerticalAlignment.Val = TableVerticalAlignmentValues.Top;
			objTableCellProperties.Append(objTableCellVerticalAlignment);
			
			// Append the TableCellProperties object to the TableCell object.
			objTableCell.Append(objTableCellProperties);
			return objTableCell;
			} // end of ConstructTableCell Method


		

		} //End of oxmlDocument Class

	//***g
	//***g

	//++ oxmlWorkbook Class
	/// <summary>
	/// This object represents mostly Workbook/Worksheet related procedures
	/// </summary>

	class oxmlWorkbook : oxmlDocumentWorkbook
		{

		//++ InsertSharedStringItem
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

		//++ InsertCellInWorksheet
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

		//++ InsertHyperlink
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

		//++InsertComment
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
			objFontSize.Val = Convert.ToDouble(Properties.AppResources.Workbook_Comments_FontSize);
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


		//++PopulateCell
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
			int intColumnPosition = 0;
			string strCellReference = parColumnLetter + parRowNumber;
			SheetData objSheetData = parWorksheetPart.Worksheet.GetFirstChild<SheetData>();
			
			//Populate the Cell that must be inserted or updated
			Cell objCell = new Cell();
			// Insert the CellReference
			objCell.CellReference = strCellReference;
			// If the cell has to have a value, add it else leave it blank.
			if(parCellcontents != null)
				{
				objCell.DataType = new EnumValue<CellValues>(parCellDatatype);
				objCell.CellValue = new CellValue(parCellcontents);
				}
			// Always update the StyleID
			objCell.StyleIndex = parStyleId;

			// Populate the Cell with a hyperlink if needed Hyperlink if required
			if(parHyperlinkURL != null)
				{
				oxmlWorkbook.InsertHyperlink(
					parWorksheetPart: parWorksheetPart,
					parCellReference: strCellReference,
					parHyperLinkID: "Hyp" + parHyperlinkCounter,
					parHyperlinkURL: parHyperlinkURL);
				}
			//Console..WriteLine("\t\t\t\t\t + writing Cell: {0} - {1} \t - StyleID: {2}", strCellReference, parCellcontents, parStyleId);
			// Now determine the position where the objCell must be inserted.
			Row objRow = new Row() { RowIndex = parRowNumber };
			Row objReferenceRow = new Row();
			if(objSheetData.Elements<Row>().Where(r => r.RowIndex.Value == parRowNumber).Count() > 0)
				{
				//Console..WriteLine("\t\t\t\t\t + Row: {0} already exist... will be used as Row Number: {1} ...", parRowNumber, objReferenceRow.RowIndex);
				objRow = objSheetData.Elements<Row>().Where(r => r.RowIndex == parRowNumber).First();
				}
			else // The Row doesn't exist...
				{
				if(objSheetData.Elements<Row>().Last<Row>().Count() > 0) // Check if any rows exist...
					{
					objReferenceRow = objSheetData.Elements<Row>().Last<Row>();
					// Check if the last existing Row's RowIndex is LESS/SMALLER than parRowNumber
					if(objReferenceRow.RowIndex > parRowNumber)
						{
						//Console..WriteLine("\t\t\t\t\t + Row: {0} doesn't exist, but there are rows greater than Row {0} ... Determine where to insert it...",parRowNumber);
						objReferenceRow = null;
						foreach(Row itemRow in objSheetData.Elements<Row>().Where(r => r.RowIndex > parRowNumber))
							{
							if(itemRow.RowIndex > parRowNumber) // got the Reference row BEFORE which to insert the ROW
								{
								//Console..WriteLine("\t\t\t\t\t + Row: {0} doesn't exist... insert BEFORE Row: {1}...", parRowNumber, itemRow.RowIndex);
								objReferenceRow = itemRow;
								objSheetData.InsertBefore<Row>(newChild: objRow, refChild: objReferenceRow);
								break;
								}
							}
						if(objReferenceRow == null) // unlikely, but if no Row is found that is greater then parRowNumber
							{
							//Console..WriteLine("\t\t\t\t\t + Unlikely - but possible, Append new Row....");
							objSheetData.Append(objRow);
							}
						}
					else // The Row Index is Smaller THAN the Row, find the correct place to insert it.
						{
						objRow = new Row() { RowIndex = parRowNumber };
						objSheetData.InsertAfter<Row>(newChild: objRow, refChild: objReferenceRow);
						//Console..WriteLine("\t\t\t\t\t + Row: {0} doesn't exist... INSERTED after Row: {0}...", parRowNumber, objReferenceRow.RowIndex);
						}						
					}
				else
					{
					//Console..WriteLine("\t\t\t\t\t\t + No Rows exist, just appen a new Row...");
					objSheetData.Append(objRow);
					}
				}
			// Check if the cell specified in parCellReference parameter exist in the row, 
			// If the cell, exist remove it.
			if(objRow.Elements<Cell>().Where(c => c.CellReference.Value == strCellReference).Count() > 0)
				{
				// The cell exist, overwrite the existing cell with the objCell...
				Cell objExistingCell = objRow.Elements<Cell>().Where(c => c.CellReference.Value == strCellReference).First();
				objExistingCell.Remove();
				}

			// Cells MUST be in sequential order according to CellReference
			// Determine where to insert the new cell because the cell must be in the exact correct sequence else it will be a corrupt sheet when opened in MS Excel.
			string strCellColumnLetter ="";
			Regex objRegex = new Regex("[A-Za-z]+");
			Cell objReferenceCell = null;
			foreach(Cell itemCell in objRow.Elements<Cell>())
				{
				Match objMatch = objRegex.Match(itemCell.CellReference.Value);
				strCellColumnLetter = objMatch.ToString();
				if(parColumnLetter.Length > strCellColumnLetter.Length)
					{
					//Console..WriteLine("\t\t\t\t\t\t - Length of {0}={1} > {2}={3} :. skip...", parColumnLetter, parColumnLetter.Length, strCellColumnLetter, strCellColumnLetter.Length);
					continue;
					}

				if(parColumnLetter.Length < strCellColumnLetter.Length)
					{
					//Console..WriteLine("\t\t\t\t\t\t - Length of {0}={1} < {2}={3} :. Insert the cell before {2}..", parColumnLetter, parColumnLetter.Length, strCellColumnLetter, strCellColumnLetter.Length);
					objReferenceCell = itemCell;
					break;
					}

				intColumnPosition = string.Compare(strA: strCellColumnLetter, strB: parColumnLetter, ignoreCase: true);
				if(intColumnPosition > 0) // The objCell reference is going AFTER 
					{
					//Console..WriteLine("\t\t\t\t\t\t - {0} goes BEFORE {1} :. Insert HERE...", parColumnLetter, strCellColumnLetter);
					objReferenceCell = itemCell;
					break;
					}
				//Console..WriteLine("\t\t\t\t\t\t - {0} goes after {1} :. skip...", parColumnLetter, strCellColumnLetter);
				}
			// If the objReferenceCell == null, the cell is inserted at the end position in the objRow.
			objRow.InsertBefore(newChild: objCell, refChild: objReferenceCell);

			parWorksheetPart.Worksheet.Save();
			} // end PopulateCell procedure

		//++MergeCell
		public static void MergeCell(
			WorksheetPart parWorksheetPart,
			string parTopLeftCell,
			string parBottomRightCell)
			{

			MergeCells objMergeCells;
			// Check if a Merge Cell collection exist for the worksheet
			if(parWorksheetPart.Worksheet.Elements<MergeCells>().Count() > 0)
				objMergeCells = parWorksheetPart.Worksheet.Elements<MergeCells>().First();
			else
				{
				objMergeCells = new MergeCells();
				// Insert in the specific location in the worksheet
				if(parWorksheetPart.Worksheet.Elements<CustomSheetView>().Count() > 0)
					parWorksheetPart.Worksheet.InsertAfter(newChild: objMergeCells, refChild: parWorksheetPart.Worksheet.Elements<CustomSheetView>().First());
				else
					parWorksheetPart.Worksheet.InsertAfter(newChild: objMergeCells, refChild: parWorksheetPart.Worksheet.Elements<SheetData>().First());
				}

			// Create a MergeCell and append it to the MergeCells collection if one doesn't already exist

			// If the MergeCell, exist remove it.
			if(objMergeCells.Elements<MergeCell>().Where(c => c.Reference.Value == parTopLeftCell + ":" + parBottomRightCell).Count() > 0)
				{
				// The cell exist, overwrite the existing cell with the objCell...
				MergeCell objExistingMergeCell = objMergeCells.Elements<MergeCell>().Where(c => c.Reference.Value == parTopLeftCell + ":" + parBottomRightCell).First();
				objExistingMergeCell.Remove();
				}

			MergeCell objMergeCell = new MergeCell();
			objMergeCell.Reference = new StringValue(parTopLeftCell + ":" + parBottomRightCell);
			objMergeCells.Append(objMergeCell);

			// Save the Worksheet to preseve the merge.
			parWorksheetPart.Worksheet.Save();

			} // end of MergeCells

		} //End of oxmlWorkbook class
	//++RowColumnNumber
	/// <summary>
	/// This object is used in Workbook generating functions that require content i.e. comments to be inserted in Row then Coloumn sequence.
	/// </summary>
	class RowColumnNumber
		{
		public int RowNumber{get; set;}
		public int ColumnNumber{get; set;}
		}

	} // End of Namespace
