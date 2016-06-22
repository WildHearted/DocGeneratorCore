using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using Microsoft.SharePoint.Client;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocGeneratorCore.SDDPwebReference;
using DocGeneratorCore.SDDPServiceReference;

namespace DocGeneratorCore

	//++ enumDocumentTypes
	{/// <summary>
	 ///	Mapped to the following columns in the [Document Collection Library]  of SharePoint:
	 ///	- values less then 10 is mappaed to [Generate Service Framework Documents]
	 ///	- values between 20 and 49 is mapped to [Generate Internal Documents]
	 /// - values greater than 50 is mapped to [Generate External Documents]
	 /// - values 
	 /// </summary>
	public enum enumDocumentTypes
		{
		Service_Framework_Document_DRM_sections=1,	// class defined
		Service_Framework_Document_DRM_inline=2,	// class defined
		ISD_Document_DRM_Sections=20,				// class defined
		ISD_Document_DRM_Inline=21,				// class defined
		RACI_Workbook_per_Role=25,				// class defined
		RACI_Matrix_Workbook_per_Deliverable=26,	// class defined
		Content_Status_Workbook=30,				// class defined
		Activity_Effort_Workbook=35,				// no Class - but keep for later use
		Internal_Services_Model_Workbook=39,		// class defined
		Internal_Technology_Coverage_Dashboard=40,	// class defined
		CSD_Document_DRM_Sections=50,				// class defined
		CSD_Document_DRM_Inline=51,				// class defined
		CSD_based_on_Client_Requirements_Mapping=52,	// class defined
		Client_Requirement_Mapping_Workbook=60,		// class defined
		Contract_SoW_Service_Description=70,		// class defined
		Pricing_Addendum_Document=71,				// class defined
		External_Technology_Coverage_Dashboard=80	// class defined
		}

	//++ enumDocumentStatusses
	public enum enumDocumentStatusses
		{
		New=0,		//- The document generation is initiated
		Creating=1,	//- Busy Creating the document
		Building=2,	//- Building/generating the document
		FatalError=3,	//- An **unexpected** and/or fatal error occurred during the generation
		Error= 5,		//- An Error occurred that **prematurely** ended the generation process (not necessarity a fatal error)
		Completed=6,	//- The document generation completed normally/as expected
		Uploading=7,	//- The document uploading began
		Uploaded=8,	//- The document was successfully uploaded
		Done=9		//- Generation is completed, post generation activities to proceed
		}


	//++ Document_Workbook class
	public class Document_Workbook
		{
		// Object Fields
		public string text2Write = "";
		// Object Properties
		public int ID{get; set;}
		public enumDocumentTypes DocumentType { get; set; }
		public int DocumentCollectionID{get; set;}
		public string DocumentCollectionTitle{get; set;}
		public string IntroductionRichText{get; set;}
		public string ExecutiveSummaryRichText{get; set;}
		public String DocumentAcceptanceRichText{get; set;}
		public enumDocumentStatusses DocumentStatus{get; set;}
		public bool HyperlinkView{get; set;}
		public bool HyperlinkEdit{get; set;}
		public string Template{get; set;}
		/// <summary>
		/// This property is a List of Hierarchy objects which represent the nodes (content) that need to be included in the generated document.
		/// </summary>
		public List<Hierarchy> SelectedNodes{get; set;}
		/// <summary>
		/// This property is a list of strings that will contain all the error messages why this specific 
		/// Document instance cannot be generated.
		/// </summary>
		public List<string> ErrorMessages{get; set;}
		public enumPresentationMode PresentationMode{get; set;}
		public string LocalDocumentURI{get; set;}
		public string FileName{get; set;}
		public string URLonSharePoint{get; set;}
		public bool UnhandledError{get; set;}

		//====================
		//+ Methods:
		//====================
		//++ LogError method
		/// <summary>
		/// Use this method whenever an error occurs while preparing a Document object before it is generated,
		/// to add each fo the errors to the list of errors. 
		/// </summary>
		/// <param name="parErrorString"></param>
		public void LogError(string parErrorString)
			{
			if(this.ErrorMessages == null)
				this.ErrorMessages = new List<string>();

			this.ErrorMessages.Add(parErrorString);
			}

		//++ UploadDoc method
		public bool UploadDoc(
			int? parRequestingUserID)
			{
			try
				{
				Console.WriteLine("Uploading document to Generated Document Library");
				
				//- Construct the SharePoint Client context and authentication...
				ClientContext objSPcontext = new ClientContext(webFullUrl: Properties.AppResources.SharePointSiteURL + "/");
				objSPcontext.Credentials = new NetworkCredential(
					userName: Properties.AppResources.DocGenerator_AccountName,
					password: Properties.AppResources.DocGenerator_Account_Password,
					domain: Properties.AppResources.DocGenerator_AccountDomain);
				Web objWeb = objSPcontext.Web;

				FileCreationInformation objNewFile = new FileCreationInformation();
				objNewFile.Content = System.IO.File.ReadAllBytes(this.LocalDocumentURI);
				objNewFile.Url = this.FileName;
				objNewFile.Overwrite = true;

				List objUploadDocumentLibrary = objWeb.Lists.GetByTitle(Properties.AppResources.List_Generated_Documents_Library_SimpleName);
				Microsoft.SharePoint.Client.File objFileToUpload = objUploadDocumentLibrary.RootFolder.Files.Add(parameters: objNewFile);

				objSPcontext.Load(objFileToUpload);
				objSPcontext.ExecuteQuery();

				//- Document Uploaded
				Console.WriteLine("\t + Document upload completed...");

				//- update the relevant columns/fields of the uploaded file
				Console.WriteLine("\t + Begin to update properties...");

				//- Obtain the Generated Documents List (actually a Document Library) and all its fileds/columns.
				List objGeneratedDocumentsList = objWeb.Lists.GetByTitle("Generated Documents");
				FieldCollection objGeneratedDocumentsFields = objGeneratedDocumentsList.Fields;
				CamlQuery objCAMLquery = new CamlQuery();
				objCAMLquery.ViewXml = @"<View>  
										<Query> 
											<OrderBy><FieldRef Name='Created' Ascending='FALSE' /></OrderBy> 
										</Query> 
										<ViewFields><FieldRef Name='ID' />
											<FieldRef Name='Title' />
											<FieldRef Name='Document_Collection' /><
											FieldRef Name='Editor' />
											<FieldRef Name='Created' />
										</ViewFields> 
										<RowLimit>1</RowLimit> 
									</View>";

				ListItemCollection objListEntries = objGeneratedDocumentsList.GetItems(objCAMLquery);
				objSPcontext.Load(objListEntries, entry => entry.Include
										(listEntry => listEntry["ID"],
										 listEntry => listEntry["Document_Collection"],
										 listEntry => listEntry["Title"],
										 listEntry => listEntry["Editor"],
										 listEntry => listEntry["Created"]));

				objSPcontext.ExecuteQuery();

				Microsoft.SharePoint.Client.ListItem objListItem = objListEntries[0];

				Console.WriteLine("{0} - {1}", objListItem["ID"], objListItem["Title"]);
				//- update the Title field/column
				objListItem["Title"] = this.FileName.Replace(oldValue: "_", newValue: " ");
				objListItem.Update();
				//- update the association of the uploaded document with the Document Collection Library entry
				//- with which is associated in the Document_Collection column/field.
				FieldLookupValue objFieldLookupValueDC = objListItem["Document_Collection"] as FieldLookupValue;
				if(objFieldLookupValueDC == null)
					{objFieldLookupValueDC = new FieldLookupValue();}

				//- set the association...
				objFieldLookupValueDC.LookupId = this.DocumentCollectionID;
				objListItem["Document_Collection"] = objFieldLookupValueDC;
				//- update all the columns that were changed
				objListItem.Update();

				//- update the Editor (Modified By) column association to the person who requested the generation of the document
				FieldLookupValue objFieldLookupValueEditor = objListItem["Editor"] as FieldLookupValue;
				if(objFieldLookupValueEditor == null)
					{
					objFieldLookupValueEditor = new FieldLookupValue();
					}
				// set the association...
				objFieldLookupValueEditor.LookupId = Convert.ToInt16(parRequestingUserID);
				objListItem["Editor"] = objFieldLookupValueEditor;
				//- update all the columns that were changed
				objListItem.Update();

				objSPcontext.ExecuteQuery();

				this.URLonSharePoint = Properties.AppResources.SharePointURL
					+ "/"+ Properties.AppResources.List_DocumentLibrary_GeneratedDocuments
					+ "/" + this.FileName;
				Console.WriteLine("\t + Successfully Uploaded: {0}", this.URLonSharePoint);

				objSPcontext.Dispose();

				}
			catch(InvalidQueryExpressionException exc)
				{
				Console.WriteLine("\n*** ERROR: Invalid Query Expression Exception ***\n{0} - {1}\nInnerException: {2}\nStackTrace: {3}.", 
					exc.HResult, exc.Message, exc.InnerException, exc.StackTrace);
				return false;
				}

			catch(Exception exc)
				{
				Console.WriteLine("\n*** Exception ERROR ***\n{0} - {1}\nInnerException: {2}\nStackTrace: {3}.", exc.HResult, exc.Message, exc.InnerException, exc.StackTrace);
				return false;
				}
						
			Console.WriteLine("Upload Successful...");
			return true;
			}		
		}

	/// <summary>
	/// This is the base class for all documents. 
	/// The LOWEST level sub-class must alwasy be used to configure/setup generatable documents.
	/// </summary>
	class aDocument : Document_Workbook
		{
		private bool _introductories_Section = false;
		public bool Introductory_Section
			{
			get{return this._introductories_Section;}
			set{this._introductories_Section = value;}
			}
		private bool _introduction = false;
		public bool Introduction
			{
			get{return this._introduction;}
			set{this._introduction = value;}
			}
		private bool _executive_Summary = false;
		public bool Executive_Summary
			{
			get{return this._executive_Summary;}
			set{this._executive_Summary = value;}
			}
		private bool _Acronyms_Glossary_of_Terms_Section = false;
		public bool Acronyms_Glossary_of_Terms_Section
			{
			get{return this._Acronyms_Glossary_of_Terms_Section;}
			set{this._Acronyms_Glossary_of_Terms_Section = value;}
			}
		private bool _acronyms = false;
		public bool Acronyms
			{
			get{return this._acronyms;}
			set{this._acronyms = value;}
			}
		private Dictionary<int, string> _dictionaryGlossaryAndAcronyms = new Dictionary<int, string>();
		public Dictionary<int, string> DictionaryGlossaryAndAcronyms
			{
			get{return this._dictionaryGlossaryAndAcronyms;}
			set{this._dictionaryGlossaryAndAcronyms = value;}
			}
		private bool _glossary_of_Terms = false;
		public bool Glossary_of_Terms
			{
			get{return this._glossary_of_Terms;}
			set{this._glossary_of_Terms = value;}
			}
		private UInt32 _pageHeight = 0;
		public UInt32 PageHight
			{
			get{return this._pageHeight;}
			set{this._pageHeight = value;}
			}
		private UInt32 _pageWidth = 0;
		public UInt32 PageWith
			{
			get{return this._pageWidth;}
			set{this._pageWidth = value;}
			}
		private bool _colorCodingLayer1 = false;
		public bool ColorCodingLayer1
			{
			get{return this._colorCodingLayer1;}
			set{this._colorCodingLayer1 = value;}
			}
		private bool _colorCodingLayer2 = false;
		public bool ColorCodingLayer2
			{
			get{return this._colorCodingLayer2;}
			set{this._colorCodingLayer2 = value;}
			}
		private bool _colorCodingLayer3 = false;
		public bool ColorCodingLayer3
			{
			get{return this._colorCodingLayer3;}
			set{this._colorCodingLayer3 = value;}
			}

		}
	//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	/// <summary>
	/// all Workbooks are based on this class.
	/// </summary>
	class aWorkbook : Document_Workbook
		{
		/// <summary>
		/// Return the alphabetic letter for a worksheet column after providing a numeric column number as parameter.
		/// </summary>
		/// <param name="parColumnNo"></param>
		/// <returns></returns>
		public static string GetColumnLetter(int parColumnNo)
			{
			var intFirstLetter = ((parColumnNo) / 676) + 64;
			var intSecondLetter = ((parColumnNo % 676) / 26) + 64;
			var intThirdLetter = (parColumnNo % 26) + 65;

			var firstLetter = (intFirstLetter > 64)
			    ? (char)intFirstLetter : ' ';
			var secondLetter = (intSecondLetter > 64)
			    ? (char)intSecondLetter : ' ';
			var thirdLetter = (char)intThirdLetter;

			return string.Concat(firstLetter, secondLetter,
			    thirdLetter).Trim();
			}
		//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
		/// <summary>
		/// Provide a column letter and  gives you corresponding column number (integer)
		/// </summary>
		/// <param name="parColumnLetter"></param>
		/// <returns>an integer as the column row number</returns>
		public static int GetColumnNumber(string parColumnLetter)
			{
			Regex alphaValue = new Regex("^[A-Z]+$");
			if(!alphaValue.IsMatch(parColumnLetter))
				throw new ArgumentException();

			char[] columnLetters = parColumnLetter.ToCharArray();
			Array.Reverse(columnLetters);

			int convertedColumnNumber = 0;
			for(int i = 0; i < columnLetters.Length; i++)
				{
				char letter = columnLetters[i];
				// ASCII 'A' = 65
				int current = i == 0 ? letter - 65 : letter - 64;
				convertedColumnNumber += current * (int)Math.Pow(26, i);
				}

			return convertedColumnNumber;
			}


		///%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
		/// <summary>
		/// Adds comments to a rowksheet. Two parameters are required: 
		///    parWorksheetPart = which must be an Worksheet object.
		///    parDictionaryOfComments object with Key= string of column letter value | row number e.g. A|1 or AB|66 
		/// </summary>
		/// <param name="parWorksheetPart">Worksheet Part to which the comments must be added </param>
		/// <param name="parDictionaryOfComments">Dictionary of cell references as the key (ie. A1) and the comment text as the value</param>
		public static void InsertWorksheetComments(
			WorksheetPart parWorksheetPart, 
			Dictionary<string, string> parDictionaryOfComments)
			{
			if(parDictionaryOfComments.Any())
				{
				string strVmlXmlForAllComments = string.Empty;
				string strColumnLetter = string.Empty;
				string strRowNumber = string.Empty;
				int intShapeIndex = 1;
				// Create all the VML Shapes XML for all the comments in the Dictionary
				foreach(var commentEntry in parDictionaryOfComments)
					{
					intShapeIndex += 1;
					strColumnLetter = commentEntry.Key.Substring(startIndex: 0, length: commentEntry.Key.IndexOf("|"));
					strRowNumber = commentEntry.Key.Substring(startIndex: commentEntry.Key.IndexOf("|") + 1, length: commentEntry.Key.Length - commentEntry.Key.IndexOf("|") - 1);
                         strVmlXmlForAllComments += GetCommentVMLShapeXML(
						parColumnLetter: strColumnLetter,
						parRowNumber: strRowNumber,
						parShapeIndex: intShapeIndex
						);
					}

				//Console.WriteLine("VML for Comments: \n[{0}]", strVmlXmlForAllComments);

				// check if a VmlDrawing part already exist, if it does, delete it, and replace it with the new VmlDrawingpart
				VmlDrawingPart objVmlDrawingPart;
				string strVmlDrawingPartId = string.Empty;
				IEnumerable<VmlDrawingPart> ieVmlDrawingParts;
				//Console.WriteLine("VMLdrawingParts: {0}", parWorksheetPart.VmlDrawingParts.Count());
				if(parWorksheetPart.VmlDrawingParts.Count() > 0)
					{
					try
						{

						foreach(var item in parWorksheetPart.VmlDrawingParts)
							{
							ieVmlDrawingParts = parWorksheetPart.VmlDrawingParts;
							objVmlDrawingPart = ieVmlDrawingParts.FirstOrDefault<VmlDrawingPart>();
							strVmlDrawingPartId = parWorksheetPart.GetIdOfPart(part: objVmlDrawingPart);
							parWorksheetPart.DeletePart(id: strVmlDrawingPartId);
							}
						}
					catch(InvalidOperationException)
						{
						// just ignore the exception
						}
					}
				else
					{
					strVmlDrawingPartId = "rId2";
					}
				// The VMLDrawingPart must contain all the definitions for how to draw every comment shape for the worksheet
				objVmlDrawingPart = parWorksheetPart.AddNewPart<VmlDrawingPart>(id: strVmlDrawingPartId);
				using(XmlTextWriter writer = new XmlTextWriter(objVmlDrawingPart.GetStream(FileMode.Create), Encoding.UTF8))
					{
					writer.WriteRaw(
						"<xml " + 
							"xmlns:v=\"urn:schemas-microsoft-com:vml\"\r\n " + 
							"xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n " + 
							"xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\r\n " + 

							"<o:shapelayout v:ext=\"edit\">\r\n  " + 
								"<o:idmap v:ext=\"edit\" data=\"1\"/>\r\n " + 
							"</o:shapelayout>" +

							"<v:shapetype id=\"_x0000_t202\" " + 
								"coordsize=\"21600,21600\" " + 
								"o:spt=\"202\"\r\n  " + 
								"path=\"m,l,21600r21600,l21600,xe\">\r\n  " + 
								"<v:stroke joinstyle=\"miter\"/>\r\n  " + 
								"<v:path gradientshapeok=\"t\" " + 
								"o:connecttype=\"rect\"/>\r\n " + 
							"</v:shapetype>" + 

							strVmlXmlForAllComments + 

						"</xml>");
					}
				//Console.WriteLine("VMLdrawingPart: [{0}]", objVmlDrawingPart.OpenXmlPackage);
				
				// check if a WorksheetCommentsPart already exist, if it does, use it, else create a new one.
				WorksheetCommentsPart objWorksheetCommentsPart;
				if(parWorksheetPart.WorksheetCommentsPart != null)
					{
					objWorksheetCommentsPart = parWorksheetPart.WorksheetCommentsPart;
					}
				else
					{
					objWorksheetCommentsPart = parWorksheetPart.AddNewPart<WorksheetCommentsPart>();
					}

				// The Comments collection contains each of the comments contained in the parDictionaryOfComments
				DocumentFormat.OpenXml.Spreadsheet.Comments objComments;
				bool boolAppendComments = true;
				// if there are a Comments collection, Remove it.
				if(objWorksheetCommentsPart.Comments != null)
					{
					objWorksheetCommentsPart.Comments.RemoveAllChildren();
					objComments = objWorksheetCommentsPart.Comments;
					boolAppendComments = false;
					}
				else
					{
					objComments = new DocumentFormat.OpenXml.Spreadsheet.Comments();
					boolAppendComments = true;
					}

				// Create Authors collection and the Author
				Authors objAuthors;
				Author objAuthor = new Author();
				objAuthor.Text = Properties.AppResources.Workbook_Comment_Author_Name;
				if(objComments.Authors == null)
					{
					objAuthors = new Authors();
					objAuthors.Append(objAuthor);
					objComments.Append(objAuthors);
					}
				else
					{
					objAuthors = objComments.Authors;
					objAuthors.Append(objAuthor);

					}
				// Get the Author ID
				int intAuthorID = 0;
				foreach(Author authorEntry in objAuthors)
					{
					if(authorEntry.Text == Properties.AppResources.Workbook_Comment_Author_Name)
						break;
					intAuthorID += 1;
					}

				// Create the CommentList which is a member of the Comments collection
				CommentList objCommentList;
				bool boolAppendCommentList;
				if(objWorksheetCommentsPart.Comments != null &&
				    objWorksheetCommentsPart.Comments.Descendants<CommentList>().SingleOrDefault() != null)
					{
					objCommentList = parWorksheetPart.WorksheetCommentsPart.Comments.Descendants<CommentList>().Single();
					boolAppendCommentList = false;
					}
				else
					{
					objCommentList = new CommentList();
					boolAppendCommentList = true;
					}
					
				//UInt32Value uintShapeId = 0U;

				// Create each of the comments contained in parDictionaryOfComments
				foreach(var commentEntry in parDictionaryOfComments)
					{
					strColumnLetter = commentEntry.Key.Substring(startIndex: 0, length: commentEntry.Key.IndexOf("|"));
					strRowNumber = commentEntry.Key.Substring(startIndex: commentEntry.Key.IndexOf("|") + 1, length: commentEntry.Key.Length - commentEntry.Key.IndexOf("|") - 1);
					// Create a new Comment...
					DocumentFormat.OpenXml.Spreadsheet.Comment objComment = new DocumentFormat.OpenXml.Spreadsheet.Comment();
					objComment.Reference = strColumnLetter + strRowNumber;
					objComment.AuthorId = (UInt32Value)Convert.ToUInt32(intAuthorID);
					// objComment.ShapeId = uintShapeId;

					// Create the text structure containint the text for the comment...
					CommentText objCommentText = new CommentText();
					DocumentFormat.OpenXml.Spreadsheet.Run objRun = new DocumentFormat.OpenXml.Spreadsheet.Run();
					DocumentFormat.OpenXml.Spreadsheet.RunProperties objRunProperties = new DocumentFormat.OpenXml.Spreadsheet.RunProperties();
					DocumentFormat.OpenXml.Spreadsheet.FontSize objFontSize = new DocumentFormat.OpenXml.Spreadsheet.FontSize();
					objFontSize.Val = Convert.ToDouble(Properties.AppResources.Workbook_Comments_FontSize); // 8D;
					DocumentFormat.OpenXml.Spreadsheet.Color objColor = new DocumentFormat.OpenXml.Spreadsheet.Color();
					objColor.Indexed = (UInt32Value)81U;
					RunFont objRunFont = new RunFont();
					objRunFont.Val = Properties.AppResources.Workbook_Comments_RunFont;
					RunPropertyCharSet objRunPropertyCharSet = new RunPropertyCharSet();
					objRunPropertyCharSet.Val = 1;
					DocumentFormat.OpenXml.Spreadsheet.FontFamily objFontFamily = new DocumentFormat.OpenXml.Spreadsheet.FontFamily();
                         objFontFamily.Val = 2;

					objRunProperties.Append(objFontSize);
					objRunProperties.Append(objColor);
					objRunProperties.Append(objRunFont);
					objRunProperties.Append(objRunPropertyCharSet);
					//objRunProperties.Append(objFontFamily);

					DocumentFormat.OpenXml.Spreadsheet.Text objText = new DocumentFormat.OpenXml.Spreadsheet.Text();
					objText.Text = commentEntry.Value;

					objRun.Append(objRunProperties);
					objRun.Append(objText);

					objCommentText.Append(objRun);
					objComment.Append(objCommentText);
					objCommentList.Append(objComment);

					// increment the ShapeID
					// uintShapeId += 1;

					} //foreach(var commentItem from parDictionaryOfComments

				// Once all Comments are appended to the CommentsList collection
				// Check if the CommentsList has to be appended to the Comments and append it.
				if(boolAppendCommentList)
					{
					objComments.Append(objCommentList);
					}

				// Check if the Comments have to be appended to the WorksheetCommentsPart.
				if(boolAppendComments)
					{
					objWorksheetCommentsPart.Comments = objComments;
					}

				} // if(parDictionatyOfComments.Count > 0)
			} // InsertWorksheetComments

		///%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
		/// <summary>
		/// Creates the VML Shape XML for a comment. It determines the positioning of the
		/// comment in the excel document based on the column name and row index.
		/// </summary>
		/// <param name="parColumnLetter">Column name containing the comment</param>
		/// <param name="parRowNumber">Row index containing the comment</param>
		/// <returns>VML Shape XML for a comment</returns>
		private static string GetCommentVMLShapeXML(
			string parColumnLetter, 
			string parRowNumber,
			int parShapeIndex)
			{
			string commentVmlXml = string.Empty;

			// Parse the row index into an int so we can subtract one
			int commentRowIndex;
			int commentZindex = 0;
			;
			if(int.TryParse(parRowNumber, out commentRowIndex))
				{
				commentZindex += 1;
				commentRowIndex -= 1;

				//"<v:shape id=\"_x0000_s1" + commentRowIndex * 2 + GetColumnNumber(parColumnLetter) * 3 + 

				commentVmlXml = 
				"<v:shape id=\"_x0000_s102" + parShapeIndex + "\" " + 
					"type=\"#_x0000_t202\" " + 
					"style=\'position:absolute;\r\n  " +
						"margin-left:509.25pt;" + 
						"margin-top:110.25pt;" + 
						"width:120pt;" + 
						"height:60pt;\r\n  " + 
						"z-index:1;" +
						"visibility:hidden\' " +
						"fillcolor=\"yellow [13]\" " +
						"o:insetmode=\"auto\">\r\n " +
						"<v:fill opacity=\"43909f\" color2=\"#ffffe1\"/>\r\n  " +
						"<v:shadow color=\"black\" obscured=\"f\"/>\r\n  " + 
						"<v:path o:connecttype=\"none\"/>\r\n   " +
						"<v:textbox style=\'mso-direction-alt:auto\' inset=\"2.5mm,2.5mm,2.5mm,2.5mm\">\r\n  " +
							"<div style=\'text-align:left\'></div>\r\n  " +
						"</v:textbox>\r\n  " +
					"<x:ClientData ObjectType=\"Note\">\r\n  " +
						"<x:MoveWithCells/>\r\n   " +
						//"<x:SizeWithCells/>\r\n   " +
						"<x:Anchor>\r\n   " + GetAnchorCoordinatesForVMLCommentShape(parColumnLetter, parRowNumber) + "</x:Anchor>\r\n   " +
						"<x:AutoFill>False</x:AutoFill>\r\n   " +
                              "<x:Row>" + commentRowIndex + "</x:Row>\r\n   " +
						"<x:Column>" + GetColumnNumber(parColumnLetter) + "</x:Column>\r\n   " +
					"</x:ClientData>\r\n   " +
				"</v:shape>";
				}
			return commentVmlXml;
			}

		///%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
		/// <summary>
		/// Gets the coordinates for where on the excel spreadsheet to display the VML comment shape
		/// </summary>
		/// <param name="parColumnLetter">Column Letter of where the comment is located (ie. B)</param>
		/// <param name="parRowNumber">Row Number of where the comment is located (ie. 2)</param>
		/// <returns><see cref="<x:Anchor>"/> coordinates in the form of a comma separated list</returns>
		private static string GetAnchorCoordinatesForVMLCommentShape(
			string parColumnLetter, 
			string parRowNumber)
			{
			string strCoordinates = string.Empty;
			int intStartingRow = 0;
			int intStartingColumn = GetColumnNumber(parColumnLetter);

			// From (upper right coordinate of a rectangle)
			// [0] Left column
			// [1] Left column offset
			// [2] Left row
			// [3] Left row offset

			// To (bottom right coordinate of a rectangle)
			// [4] Right column
			// [5] Right column offset
			// [6] Right row
			// [7] Right row offset
			List<int> coordList = new List<int>(8) { 0, 0, 0, 0, 0, 0, 0, 0 };

			if(int.TryParse(parRowNumber, out intStartingRow))
				{
				// Make the row be a zero based index
				intStartingRow -= 1;
				// If starting column is A, display shape in column B
				coordList[0] = intStartingColumn + 1;
				// [1] Left column offset
				coordList[1] = 10;
				// [2] Left row
				coordList[2] = intStartingRow;
				// [3] Left row offset - 
				coordList[3] = 10;
				// To (bottom right coordinate of a rectangle)
				// [4] Right column If starting column is A, display shape till column E
				coordList[4] = intStartingColumn + 4;
				// [5] Right column offset
				coordList[5] = 5;
				// [6] Right row
				coordList[6] = intStartingRow + 4; // If starting row is 0, display 3 rows down to row 3

				// The row offsets change if the shape is defined in the first row
				if(intStartingRow == 0)
					{
					// [3] Left row offset
					coordList[3] = 2;
					// [7] Right row offset
					coordList[7] = 16;
					}
				else
					{
					// [3] Left row offset
					coordList[3] = 10;
					// [7] Right row offset
					coordList[7] = 4;
					}

				strCoordinates = string.Join(", ", coordList.ConvertAll<string>(x => x.ToString()).ToArray());
				}

			return strCoordinates;
			} //end GetAchorCoordinatesForVMLcommentShape

		} // end Workbook class

	/// <summary>
	/// This class inherits from the Document class and contain all the common properties and methods that
	/// the Predefined product documents have.
	/// </summary>
	class PredefinedProduct_Document : aDocument
		{
		private bool _service_Portfolio_Section = false;
		public bool Service_Portfolio_Section
			{
			get{return this._service_Portfolio_Section;}
			set{this._service_Portfolio_Section = value;}
			}
		private bool _service_Portfolio_Description = false;
		public bool Service_Portfolio_Description
			{
			get{return this._service_Portfolio_Description;}
			set{this._service_Portfolio_Description = value;}
			}
		private bool _service_Family_Heading = false;
		public bool Service_Family_Heading
			{
			get{return this._service_Family_Heading;}
			set{this._service_Family_Heading = value;}
			}
		private bool _service_Family_Description = false;
		public bool Service_Family_Description
			{
			get{return this._service_Family_Description;}
			set{this._service_Family_Description = value;}
			}
		private bool _service_Product_Heading = false;
		public bool Service_Product_Heading
			{
			get{return this._service_Product_Heading;}
			set{this._service_Product_Heading = value;}
			}
		private bool _service_Product_Description = false;
		public bool Service_Product_Description
			{
			get{return this._service_Product_Description;}
			set{this._service_Product_Description = value;}
			}
		private bool _drm_Heading = false;
		public bool DRM_Heading
			{
			get{return this._drm_Heading;}
			set{this._drm_Heading = value;}
			}
		private bool _Deliverables_Reports_Meetings = false;
		public bool Deliverables_Reports_Meetings
			{
			get{return this._Deliverables_Reports_Meetings;}
			set{this._Deliverables_Reports_Meetings = value;}
			}
		private bool _service_Levels = false;
		public bool Service_Levels
			{
			get{return this._service_Levels;}
			set{this._service_Levels = value;}
			}
		private bool _service_Level_Heading = false;
		public bool Service_Level_Heading
			{
			get{return this._service_Level_Heading;}
			set{this._service_Level_Heading = value;}
			}
		private bool _service_Level_Commitments_Table = false;
		public bool Service_Level_Commitments_Table
			{
			get{return this._service_Level_Commitments_Table;}
			set{this._service_Level_Commitments_Table = value;}
			}
		} // end of PredefinedProduct_Document class
	
	/// <summary>
	/// This class inherits from the PredefinedProduct_Document class and contain all the common properties and methods that
	/// the External (Client Facing) documents have.
	/// </summary>
	class External_Document : PredefinedProduct_Document
		{
		private bool _service_Feature_Heading = false;
		public bool Service_Feature_Heading
			{
			get{return this._service_Feature_Heading;}
			set{this._service_Feature_Heading = value;}
			}
		private bool _service_Feature_Description = false;
		public bool Service_Feature_Description
			{
			get{return this._service_Feature_Description;}
			set{this._service_Feature_Description = value;}
			}
		} // End of the External_Document class

	/// <summary>
	/// This class inherits from the PredefinedProduct_Document class and contain all the common properties and methods that the Internal documents have.
	/// </summary>
	class Internal_Document : PredefinedProduct_Document
		{
		private bool _service_Product_Key_Client_Benefits = false;
		public bool Service_Product_Key_Client_Benefits
			{
			get{return this._service_Product_Key_Client_Benefits;}
			set{this._service_Product_Key_Client_Benefits = value;}
			}
		private bool _service_Product_Key_DD_Benefits = false;
		public bool Service_Product_KeyDD_Benefits
			{
			get{return this._service_Product_Key_DD_Benefits;}
			set{this._service_Product_Key_DD_Benefits = value;}
			}
		private bool _service_Element_Heading = false;
		public bool Service_Element_Heading
			{
			get{return this._service_Element_Heading;}
			set{this._service_Element_Heading = value;}
			}
		private bool _service_Element_Description = false;
		public bool Service_Element_Description
			{
			get{return this._service_Element_Description;}
			set{this._service_Element_Description = value;}
			}
		private bool _service_Element_Objectives = false;
		public bool Service_Element_Objectives
			{
			get{return this._service_Element_Objectives;}
			set{this._service_Element_Objectives = value;}
			}
		private bool _service_Element_Key_Client_Benefits = false;
		public bool Service_Element_Key_Client_Benefits
			{
			get{return this._service_Element_Key_Client_Benefits;}
			set{this._service_Element_Key_Client_Benefits = value;}
			}
		private bool _service_Element_Key_Client_Advantages = false;
		public bool Service_Element_Key_Client_Advantages
			{
			get{return this._service_Element_Key_Client_Advantages;}
			set{this._service_Element_Key_Client_Advantages = value;}
			}
		private bool _service_Element_Key_DD_Benefits = false;
		public bool Service_Element_Key_DD_Benefits
			{
			get{return this._service_Element_Key_DD_Benefits;}
			set{this._service_Element_Key_DD_Benefits = value;}
			}
		private bool _service_Element_Critical_Success_Factors = false;
		public bool Service_Element_Critical_Success_Factors
			{
			get{return this._service_Element_Critical_Success_Factors;}
			set{this._service_Element_Critical_Success_Factors = value;}
			}
		private bool _service_Element_Key_Performance_Indicators = false;
		public bool Service_Element_Key_Performance_Indicators
			{
			get{return this._service_Element_Key_Performance_Indicators;}
			set{this._service_Element_Key_Performance_Indicators = value;}
			}
		private bool _service_Element_High_Level_Process = false;
		public bool Service_Element_High_Level_Process
			{
			get{return this._service_Element_High_Level_Process;}
			set{this._service_Element_High_Level_Process = value;}
			}
		private bool _activities = false;
		public bool Activities
			{
			get{return this._activities;}
			set{this._activities = value;}
			}
		private bool _activity_Heading = false;
		public bool Activity_Heading
			{
			get{return this._activity_Heading;}
			set{this._activity_Heading = value;}
			}
		private bool _activity_Description_Table = false;
		public bool Activity_Description_Table
			{
			get{return this._activity_Description_Table;}
			set{this._activity_Description_Table = value;}
			}
		private bool _document_Acceptance_Section = false;
		public bool Document_Acceptance_Section
			{
			get{return this._document_Acceptance_Section;}
			set{this._document_Acceptance_Section = value;}
			}
		} // End of the Internal_Document class


	class Pricing_Addendum_Document : aDocument
		{
		private int _pricing_Worksbook_Id = 0;
		public int Pricing_Workbook_Id
			{
			get{return _pricing_Worksbook_Id;}
			set{_pricing_Worksbook_Id = value;}
			}
		public bool Generate(
			ref CompleteDataSet parDataSet,
			DesignAndDeliveryPortfolioDataContext parSDDPdatacontext)
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			//TODO: Code to added for Pricing_Addendum_Document's Generate method.
			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}
		}

	/// <summary>
	/// This class contains all the Client Service Description (CSD) with inline DRM (Deliverable Report Meeting).
	/// </summary>
	class Internal_DRM_Inline : Internal_Document
		{
		private bool _drm_Description = false;
		public bool DRM_Description
			{
			get{return this._drm_Description;}
			set{this._drm_Description = value;}
			}
		private bool _drm_Inputs = false;
		public bool DRM_Inputs
			{
			get{return this._drm_Inputs;}
			set{this._drm_Inputs = value;}
			}
		private bool _drm_Outputs = false;
		public bool DRM_Outputs
			{
			get{return this._drm_Outputs;}
			set{this._drm_Outputs = value;}
			}
		private bool _dds_DRM_Obligations = false;
		public bool DDS_DRM_Obligations
			{
			get{return this._dds_DRM_Obligations;}
			set{this._dds_DRM_Obligations = value;}
			}
		private bool _clients_DRM_Responsibilities = false;
		public bool Clients_DRM_Responsibilities
			{
			get{return this._clients_DRM_Responsibilities;}
			set{this._clients_DRM_Responsibilities = value;}
			}
		private bool _drm_Exclusions = false;
		public bool DRM_Exclusions
			{
			get{return this._drm_Exclusions;}
			set{this._drm_Exclusions = value;}
			}
		private bool _drm_Governance_Controls = false;
		public bool DRM_Governance_Controls
			{
			get{return this._drm_Governance_Controls;}
			set{this._drm_Governance_Controls = value;}
			}

		} // end of CSD_inline DRM class

	/// <summary>
	/// This class contains all the properties and methods for Internal DRM (Deliverable Report Meeting) Sections object
	/// </summary>
	class Internal_DRM_Sections : Internal_Document
		{
		private bool _drm_Summary = false;
		public bool DRM_Summary
			{
			get{return this._drm_Summary;}
			set{this._drm_Summary = value;}
			}
		private bool _drm_Section = false;
		public bool DRM_Section
			{
			get{return this._drm_Section;}
			set{this._drm_Section = value;}
			}
		private bool _deliverables = false;
		public bool Deliverables
			{get{return this._deliverables;}
			set{this._deliverables = value;}
			}
		private bool _deliverable_Heading = false;
		public bool Deliverable_Heading
			{
			get{return this._deliverable_Heading;}
			set{this._deliverable_Heading = value;}
			}
		private bool _deliverable_Description = false;
		public bool Deliverable_Description
			{
			get{return this._deliverable_Description;}
			set{this._deliverable_Description = value;}
			}
		private bool _deliverable_Inputs = false;
		public bool Deliverable_Inputs
			{
			get{return this._deliverable_Inputs;}
			set{this._deliverable_Inputs = value;}
			}
		private bool _deliverable_Outputs = false;
		public bool Deliverable_Outputs
			{
			get{return this._deliverable_Outputs;}
			set{this._deliverable_Outputs = value;}
			}
		private bool _dds_Deliverable_Obligations = false;
		public bool DDs_Deliverable_Obligations
			{
			get{return this._dds_Deliverable_Obligations;}
			set{this._dds_Deliverable_Obligations = value;}
			}
		private bool _clients_Deliverable_Responsibilities = false;
		public bool Clients_Deliverable_Responsibilities
			{
			get{return this._clients_Deliverable_Responsibilities;}
			set{this._clients_Deliverable_Responsibilities = value;}
			}
		private bool _deliverable_Exclusions = false;
		public bool Deliverable_Exclusions
			{
			get{return this._deliverable_Exclusions;}
			set{this._deliverable_Exclusions = value;}
			}
		private bool _deliverable_Governance_Controls = false;
		public bool Deliverable_Governance_Controls
			{
			get{return this._deliverable_Governance_Controls;}
			set{this._deliverable_Governance_Controls = value;}
			}
		private bool _reports = false;
		public bool Reports
			{
			get{return this._reports;}
			set{this._reports = value;}
			}
		private bool _report_Heading = false;
		public bool Report_Heading
			{
			get{return this._report_Heading;}
			set{this._report_Heading = value;}
			}
		private bool _report_Description = false;
		public bool Report_Description
			{
			get{return this._report_Description;}
			set{this._report_Description = value;}
			}
		private bool _report_Inputs = false;
		public bool Report_Inputs
			{
			get{return this._report_Inputs;}
			set{this._report_Inputs = value;}
			}
		private bool _report_Outputs = false;
		public bool Report_Outputs
			{
			get{return this._report_Outputs;}
			set{this._report_Outputs = value;}
			}
		private bool _dds_Report_Obligations = false;
		public bool DDs_Report_Obligations
			{
			get{return this._dds_Report_Obligations;}
			set{this._dds_Report_Obligations = value;}
			}
		private bool _clients_Report_Responsibilities = false;
		public bool Clients_Report_Responsibilities
			{
			get{return this._clients_Report_Responsibilities;}
			set{this._clients_Report_Responsibilities = value;}
			}
		private bool _report_Exclusions = false;
		public bool Report_Exclusions
			{
			get{return this._report_Exclusions;}
			set{this._report_Exclusions = value;}
			}
		private bool _report_Governance_Controls = false;
		public bool Report_Governance_Controls
			{
			get{return this._report_Governance_Controls;}
			set{this._report_Governance_Controls = value;}
			}
		private bool _meetings = false;
		public bool Meetings
			{
			get{return this._meetings;}
			set{this._meetings = value;}
			}
		private bool _meeting_Heading = false;
		public bool Meeting_Heading
			{
			get{return this._meeting_Heading;}
			set{this._meeting_Heading = value;}
			}
		private bool _meeting_Description = false;
		public bool Meeting_Description
			{
			get{return this._meeting_Description;}
			set{this._meeting_Description = value;}
			}
		private bool _meeting_Inputs = false;
		public bool Meeting_Inputs
			{
			get{return this._meeting_Inputs;}
			set{this._meeting_Inputs = value;}
			}
		private bool _meeting_Outputs = false;
		public bool Meeting_Outputs
			{
			get{return this._meeting_Outputs;}
			set{this._meeting_Outputs = value;}
			}
		private bool _dds_meeting_Obligations = false;
		public bool DDs_Meeting_Obligations
			{
			get{return this._dds_meeting_Obligations;}
			set{this._dds_meeting_Obligations = value;}
			}
		private bool _clients_Meeting_Responsibilities = false;
		public bool Clients_Meeting_Responsibilities
			{
			get{return this._clients_Meeting_Responsibilities;}
			set{this._clients_Meeting_Responsibilities = value;}
			}
		private bool _meeting_Exclusions = false;
		public bool Meeting_Exclusions
			{
			get{return this._meeting_Exclusions;}
			set{this._meeting_Exclusions = value;}
			}
		private bool _meeting_Governance_Controls = false;
		public bool Meeting_Governance_Controls
			{
			get{return this._meeting_Governance_Controls;}
			set{this._meeting_Governance_Controls = value;}
			}
		private bool _service_Level_Section = false;
		public bool Service_Level_Section
			{
			get{return this._service_Level_Section;}
			set{this._service_Level_Section = value;}
			}
		private bool _service_Level_Heading_in_Section = false;
		public bool Service_Level_Heading_in_Section
			{
			get{return this._service_Level_Heading_in_Section;}
			set{this._service_Level_Heading_in_Section = value;}
			}
		private bool _service_Level_Table_in_Section = false;
		public bool Service_Level_Table_in_Section
			{
			get{return this._service_Level_Table_in_Section;}
			set{this._service_Level_Table_in_Section = value;}
			}


		} // end of Internal_DRM_Sections class

	/// <summary>
	/// This class contains all the properties and methods for DRM (Deliverable Report Meeting) Sections
	/// </summary>
	class External_DRM_Sections : External_Document
		{
		private bool _drm_Summary = false;
		public bool DRM_Summary
			{
			get
				{
				return _drm_Summary;
				}
			set
				{
				_drm_Summary = value;
				}
			}
		private bool _drm_Section = false;
		public bool DRM_Section
			{
			get
				{
				return _drm_Section;
				}
			set
				{
				_drm_Section = value;
				}
			}
		private bool _deliverables = false;
		public bool Deliverables
			{
			get
				{
				return _deliverables;
				}
			set
				{
				_deliverables = value;
				}
			}
		private bool _deliverable_Heading = false;
		public bool Deliverable_Heading
			{

			get
				{
				return _deliverable_Heading;
				}
			set
				{
				_deliverable_Heading = value;
				}
			}
		private bool _deliverable_Description = false;
		public bool Deliverable_Description
			{
			get
				{
				return _deliverable_Description;
				}
			set
				{
				_deliverable_Description = value;
				}
			}
		private bool _deliverable_Inputs = false;
		public bool Deliverable_Inputs
			{
			get
				{
				return _deliverable_Inputs;
				}
			set
				{
				_deliverable_Inputs = value;
				}
			}
		private bool _deliverable_Outputs = false;
		public bool Deliverable_Outputs
			{
			get
				{
				return _deliverable_Outputs;
				}
			set
				{
				_deliverable_Outputs = value;
				}
			}
		private bool _dds_Deliverable_Obligations = false;
		public bool DDs_Deliverable_Obligations
			{
			get
				{
				return _dds_Deliverable_Obligations;
				}
			set
				{
				_dds_Deliverable_Obligations = value;
				}
			}
		private bool _clients_Deliverable_Responsibilities = false;
		public bool Clients_Deliverable_Responsibilities
			{
			get
				{
				return _clients_Deliverable_Responsibilities;
				}
			set
				{
				_clients_Deliverable_Responsibilities = value;
				}
			}
		private bool _deliverable_Exclusions = false;
		public bool Deliverable_Exclusions
			{
			get
				{
				return _deliverable_Exclusions;
				}
			set
				{
				_deliverable_Exclusions = value;
				}
			}
		private bool _deliverable_Governance_Controls = false;
		public bool Deliverable_Governance_Controls
			{
			get
				{
				return _deliverable_Governance_Controls;
				}
			set
				{
				_deliverable_Governance_Controls = value;
				}
			}
		private bool _reports = false;
		public bool Reports
			{
			get
				{
				return _reports;
				}
			set
				{
				_reports = value;
				}
			}
		private bool _report_Heading = false;
		public bool Report_Heading
			{

			get
				{
				return _report_Heading;
				}
			set
				{
				_report_Heading = value;
				}
			}
		private bool _report_Description = false;
		public bool Report_Description
			{
			get
				{
				return _report_Description;
				}
			set
				{
				_report_Description = value;
				}
			}
		private bool _report_Inputs = false;
		public bool Report_Inputs
			{
			get
				{
				return _report_Inputs;
				}
			set
				{
				_report_Inputs = value;
				}
			}
		private bool _report_Outputs = false;
		public bool Report_Outputs
			{
			get
				{
				return _report_Outputs;
				}
			set
				{
				_report_Outputs = value;
				}
			}
		private bool _dds_Report_Obligations = false;
		public bool DDs_Report_Obligations
			{
			get
				{
				return _dds_Report_Obligations;
				}
			set
				{
				_dds_Report_Obligations = value;
				}
			}
		private bool _clients_Report_Responsibilities = false;
		public bool Clients_Report_Responsibilities
			{
			get
				{
				return _clients_Report_Responsibilities;
				}
			set
				{
				_clients_Report_Responsibilities = value;
				}
			}
		private bool _report_Exclusions = false;
		public bool Report_Exclusions
			{
			get
				{
				return _report_Exclusions;
				}
			set
				{
				_report_Exclusions = value;
				}
			}
		private bool _report_Governance_Controls = false;
		public bool Report_Governance_Controls
			{
			get
				{
				return _report_Governance_Controls;
				}
			set
				{
				_report_Governance_Controls = value;
				}
			}
		private bool _meetings = false;
		public bool Meetings
			{
			get
				{
				return _meetings;
				}
			set
				{
				_meetings = value;
				}
			}
		private bool _meeting_Heading = false;
		public bool Meeting_Heading
			{

			get
				{
				return _meeting_Heading;
				}
			set
				{
				_meeting_Heading = value;
				}
			}
		private bool _meeting_Description = false;
		public bool Meeting_Description
			{
			get
				{
				return _meeting_Description;
				}
			set
				{
				_meeting_Description = value;
				}
			}
		private bool _meeting_Inputs = false;
		public bool Meeting_Inputs
			{
			get
				{
				return _meeting_Inputs;
				}
			set
				{
				_meeting_Inputs = value;
				}
			}
		private bool _meeting_Outputs = false;
		public bool Meeting_Outputs
			{
			get
				{
				return _meeting_Outputs;
				}
			set
				{
				_meeting_Outputs = value;
				}
			}
		private bool _dds_meeting_Obligations = false;
		public bool DDs_Meeting_Obligations
			{
			get
				{
				return _dds_meeting_Obligations;
				}
			set
				{
				_dds_meeting_Obligations = value;
				}
			}
		private bool _clients_Meeting_Responsibilities = false;
		public bool Clients_Meeting_Responsibilities
			{
			get
				{
				return _clients_Meeting_Responsibilities;
				}
			set
				{
				_clients_Meeting_Responsibilities = value;
				}
			}
		private bool _meeting_Exclusions = false;
		public bool Meeting_Exclusions
			{
			get
				{
				return _meeting_Exclusions;
				}
			set
				{
				_meeting_Exclusions = value;
				}
			}
		private bool _meeting_Governance_Controls = false;
		public bool Meeting_Governance_Controls
			{
			get
				{
				return _meeting_Governance_Controls;
				}
			set
				{
				_meeting_Governance_Controls = value;
				}
			}
		private bool _service_Level_Section = false;
		public bool Service_Level_Section
			{
			get
				{
				return _service_Level_Section;
				}
			set
				{
				_service_Level_Section = value;
				}
			}
		} // end of External_DRM_Sections class




	/// <summary>
	/// The CommonProcedures class contains procedurs/methods which are utilised by various Document methods.
	/// </summary>
	class CommonProcedures
		{

		/// <summary>
		/// This function constructs a Table for activities and return the constructed Table object to the caller.
		/// </summary>
		/// <param name="parWidthColumn1">column width in DXA value</param>
		/// <param name="parWidthColumn2">column width in DXA value</param>
		/// <param name="parActivityDesciption">String containing the Description of the Activity</param>
		/// <param name="parActivityInput">String containing the Input of the Activity</param>
		/// <param name="parActivityOutput">String containing the Output of the Activity</param>
		/// <param name="parActivityAssumptions">String containing the Assumptions of the Activity</param>
		/// <param name="parActivityOptionality">String containing the Optionality value of the Activity</param>
		/// <returns> An fully formatted and populated Table object is returned to the caller which can then be inserted in the Body of the MS Word document.
		/// </returns>
		public static DocumentFormat.OpenXml.Wordprocessing.Table BuildActivityTable(
				UInt32 parWidthColumn1,
				UInt32 parWidthColumn2,
				string parActivityDesciption = "",
				string parActivityInput = "",
				string parActivityOutput = "",
				string parActivityAssumptions = "",
				string parActivityOptionality = "")
			{
			// Initialize the Activity table object
			DocumentFormat.OpenXml.Wordprocessing.Table objActivityTable = new DocumentFormat.OpenXml.Wordprocessing.Table();
			objActivityTable = oxmlDocument.ConstructTable(
				parPageWidth:0,
				parFirstRow: false,
				parNoVerticalBand: true,
				parNoHorizontalBand: true);

			// Create the TableRow, TableCell used later on.
			
			// Construct the TableGrid
			TableGrid objTableGrid = new TableGrid();
			List<UInt32> lstTableColumns = new List<UInt32>();
			lstTableColumns.Add(parWidthColumn1);
			lstTableColumns.Add(parWidthColumn2);
			objTableGrid = oxmlDocument.ConstructTableGrid(lstTableColumns);
			// Append the TableGrid object instance to the Table object instance
			objActivityTable.Append(objTableGrid);
			
			// Construct the first row of the table: Activity Description
			TableRow objTableRow = new TableRow();
			objTableRow = oxmlDocument.ConstructTableRow(parHasCondinalStyle: false);

			// Construct the first cell of the row
			TableCell objTableCell1 = new TableCell();
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
			// Add the Activity Description Title in the first Cell of the row
			Paragraph objParagraph1 = new Paragraph();
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			DocumentFormat.OpenXml.Wordprocessing.Run objRun1 = new DocumentFormat.OpenXml.Wordprocessing.Run();
			objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_ActivityTable_RowTitle_Description);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow.Append(objTableCell1);
			// Add the Activity Description value to the second Cell
			TableCell objTableCell2 = new TableCell();
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2, parHasCondtionalFormatting: false);
			Paragraph objParagraph2 = new Paragraph();
			objParagraph2 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			DocumentFormat.OpenXml.Wordprocessing.Run objRun2 = new DocumentFormat.OpenXml.Wordprocessing.Run();
			objRun2 = oxmlDocument.Construct_RunText(parText2Write: parActivityDesciption);
			objParagraph2.Append(objRun2);
			objTableCell2.Append(objParagraph2);
			objTableRow.Append(objTableCell2);
			objActivityTable.Append(objTableRow);

			// Create the Activity Input row for the table
			objTableRow = oxmlDocument.ConstructTableRow(parHasCondinalStyle: false);
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
			// Add the Activity Input Title in the first Column
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_ActivityTable_RowTitle_Inputs);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow.Append(objTableCell1);
			// Add the Activity Input value in the second Column
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2);
			objParagraph2 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun2 = oxmlDocument.Construct_RunText(parText2Write: parActivityInput);
			objParagraph2.Append(objRun2);
			objTableCell2.Append(objParagraph2);
			objTableRow.Append(objTableCell2);
			objActivityTable.Append(objTableRow);

			// Create the Activity Outputs row for the table
			objTableRow = oxmlDocument.ConstructTableRow(parHasCondinalStyle: false);
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
			// Add the Activity Outputs Title in the first Column
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_ActivityTable_RowTitle_Outputs);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow.Append(objTableCell1);
			// Add the Activity Output value in the second Column
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2);
			objParagraph2 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun2 = oxmlDocument.Construct_RunText(parText2Write: parActivityOutput);
			objParagraph2.Append(objRun2);
			objTableCell2.Append(objParagraph2);
			objTableRow.Append(objTableCell2);
			objActivityTable.Append(objTableRow);

			// Create the Activity Assumptions row for the table
			objTableRow = oxmlDocument.ConstructTableRow(parHasCondinalStyle: false);
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
			// Add the Activity Assumptions Title in the first Column
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun1 = oxmlDocument.Construct_RunText(parText2Write:Properties.AppResources.Document_ActivityTable_RowTitle_Assumptions);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow.Append(objTableCell1);
			// Add the Activity Assumptions value in the second Column
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2);
			objParagraph2 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun2 = oxmlDocument.Construct_RunText(parText2Write: parActivityAssumptions);
			objParagraph2.Append(objRun2);
			objTableCell2.Append(objParagraph2);
			objTableRow.Append(objTableCell2);
			objActivityTable.Append(objTableRow);

			// Create the Activity Optionality row for the table
			objTableRow = oxmlDocument.ConstructTableRow(parHasCondinalStyle: false);
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
			// Add the Activity Optionality Title in the first Column
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_ActivityTable_RowTitle_Optionality);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow.Append(objTableCell1);
			// Add the Activity Optionality value in the second Column
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2);
			objParagraph2 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun2 = oxmlDocument.Construct_RunText(parText2Write: parActivityOptionality);
			objParagraph2.Append(objRun2);
			objTableCell2.Append(objParagraph2);
			objTableRow.Append(objTableCell2);
			objActivityTable.Append(objTableRow);

			//Return the constructed Table object
			return objActivityTable;
			}// End of method.

		
		public static DocumentFormat.OpenXml.Wordprocessing.Table BuildSLAtable(
				int parServiceLevelID,
				UInt32 parWidthColumn1,
				UInt32 parWidthColumn2,
				string parMeasurement,
				string parMeasureMentInterval,
				string parReportingInterval,
				string parServiceHours,
				string parCalculationMethod,
				string parCalculationFormula,
				List<ServiceLevelTarget> parThresholds,
				List<ServiceLevelTarget> parTargets,
				string parBasicServiceLevelConditions,
				string parAdditionalServiceLevelConditions,
				ref List<string> parErrorMessages)
			{

			// Initialize the ServiceLevel table object
			RTdecoder objRTdecoder = new RTdecoder();

			DocumentFormat.OpenXml.Wordprocessing.Table objServiceLevelTable = new DocumentFormat.OpenXml.Wordprocessing.Table();
			objServiceLevelTable = oxmlDocument.ConstructTable(
				parPageWidth: 0,
				parNoVerticalBand: true,
				parNoHorizontalBand: true);
			
			// Construct the TableGrid
			TableGrid objTableGrid = new TableGrid();
			List<UInt32> lstTableColumns = new List<UInt32>();
			lstTableColumns.Add(parWidthColumn1);
			lstTableColumns.Add(parWidthColumn2);
			objTableGrid = oxmlDocument.ConstructTableGrid(lstTableColumns);
			// Append the TableGrid object instance to the Table object instance
			objServiceLevelTable.Append(objTableGrid);

			// Construct the first row of the table: Measurement
			TableRow objTableRow = new TableRow();
			objTableRow = oxmlDocument.ConstructTableRow(parHasCondinalStyle: false);

			// Construct the Measurement Title
			TableCell objTableCell1 = new TableCell();
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
			// Add the Measurement Title in the first Cell of the row
			Paragraph objParagraph1 = new Paragraph();
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			DocumentFormat.OpenXml.Wordprocessing.Run objRun1 = new DocumentFormat.OpenXml.Wordprocessing.Run();
			objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_SLtable_RowMeasurement_Title);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow.Append(objTableCell1);
			// Add the Measurment Description value to the second Cell
			TableCell objTableCell2 = new TableCell();
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2, parHasCondtionalFormatting: false);
			Paragraph objParagraph2 = new Paragraph();
			DocumentFormat.OpenXml.Wordprocessing.Run objRun2 = new DocumentFormat.OpenXml.Wordprocessing.Run();
			List<Paragraph> listParagraphs = new List<Paragraph>();
			if(parMeasurement == null)
				{
				objParagraph2 = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1, parIsTableParagraph: true);
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_SLtable_ValueNotSpecified_Text);
				objParagraph2.Append(objRun2);
				objTableCell2.Append(objParagraph2);
				}
			else
				{
				// Decode the RichText content using the RTdecoder object and DecodeRichText method
				try
					{
					listParagraphs = objRTdecoder.DecodeRichText(parRT2decode: parMeasurement, parIsTableText: true);
					foreach(Paragraph paragraphItem in listParagraphs)
						{
						objTableCell2.Append(paragraphItem);
						}
					}
				catch(InvalidRichTextFormatException exc)
					{
					Console.WriteLine("\n\nException occurred: {0}", exc.Message);
					// A Table content error occurred, record it in the error log.
					parErrorMessages.Add("Service Level ID: " + parServiceLevelID + " Measurements attribute " +
						" contains an error in one of its Enhance Rich Text columns. Please review the content (especially tables).");
					objParagraph2 = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 0, parIsTableParagraph: true);
					objRun2 = oxmlDocument.Construct_RunText(
						parText2Write: "A content error occurred at this position and valid content could " +
						"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it. [" + exc.Message + "]",
						parIsError: true);
					objParagraph2.Append(objRun2);
					objTableCell2.Append(objParagraph2);
					}
				}
			objTableRow.Append(objTableCell2);
			objServiceLevelTable.Append(objTableRow);

			// Create the Measurment Interval row for the table
			objTableRow = oxmlDocument.ConstructTableRow(parHasCondinalStyle: false);
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
			// Add the Measurement Interval Title to the first Column
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_SLtable_RowMeasurementInterval_Title);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow.Append(objTableCell1);
			// Add the Measurement Interval value into the second Column
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2);
			objParagraph2 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			if(parMeasureMentInterval == null)
				{
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_SLtable_ValueNotSpecified_Text);
				}
			else
				{
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: parMeasureMentInterval);
				}
			objParagraph2.Append(objRun2);
			objTableCell2.Append(objParagraph2);
			objTableRow.Append(objTableCell2);
			objServiceLevelTable.Append(objTableRow);

			// Create the Reporting Interval row for the table
			objTableRow = oxmlDocument.ConstructTableRow(parHasCondinalStyle: false);
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
			// Add the Reporting Interval Title into the first Column
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_SLtable_RowReportingInterval_Title);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow.Append(objTableCell1);
			// Add the Reporting Interval value into the second Column
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2);
			objParagraph2 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			if(parMeasureMentInterval == null)
				{
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_SLtable_ValueNotSpecified_Text);
				}
			else
				{
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: parReportingInterval);
				}
			objParagraph2.Append(objRun2);
			objTableCell2.Append(objParagraph2);
			objTableRow.Append(objTableCell2);
			objServiceLevelTable.Append(objTableRow);

			// Create the Applicable Service Hours row for the table
			objTableRow = oxmlDocument.ConstructTableRow(parHasCondinalStyle: false);
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
			// Add the Service Hours Title in the first Column
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_SLtable_RowServiceHours_Title);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow.Append(objTableCell1);
			// Add the Service Hours value into the second Column
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2);
			objParagraph2 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			if(parServiceHours == null)
				{
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_SLtable_ValueNotSpecified_Text);
				}
			else
				{
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: parServiceHours);
				}
			objParagraph2.Append(objRun2);
			objTableCell2.Append(objParagraph2);
			objTableRow.Append(objTableCell2);
			objServiceLevelTable.Append(objTableRow);

			// Create the Calculation Method row for the table
			objTableRow = oxmlDocument.ConstructTableRow(parHasCondinalStyle: false);
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
			// Add the Calculation Method Title into the first Column
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_SLtable_RowCalculationMethod_Title);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow.Append(objTableCell1);
			// Add the Calculation Method value into the second Column
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2);
			objParagraph2 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun2 = oxmlDocument.Construct_RunText(parText2Write: parCalculationMethod);
			// Decode the RichText content using the RTdecoder object and DecodeRichText method
			listParagraphs.Clear();
			if(parCalculationMethod == null)
				{
				objParagraph2 = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1, parIsTableParagraph: true);
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_SLtable_ValueNotSpecified_Text);
				objParagraph2.Append(objRun2);
				objTableCell2.Append(objParagraph2);
				}
			else
				{
				try
					{
					listParagraphs = objRTdecoder.DecodeRichText(parRT2decode: parCalculationMethod, parIsTableText: true);
					foreach(Paragraph paragraphItem in listParagraphs)
						{
						objTableCell2.Append(paragraphItem);
						}
					}
				catch(InvalidRichTextFormatException exc)
					{
					Console.WriteLine("\n\nException occurred: {0}", exc.Message);
					// A Table content error occurred, record it in the error log.
					parErrorMessages.Add("Service Level ID: " + parServiceLevelID + " Calculation Method attribute " +
						" contains an error in one of its Enhance Rich Text columns. Please review the content (especially tables).");
					objParagraph2 = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 0, parIsTableParagraph: true);
					objRun2 = oxmlDocument.Construct_RunText(
						parText2Write: "A content error occurred at this position and valid content could " +
						"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it. [" + exc.Message + "]",
						parIsError: true);
					objParagraph2.Append(objRun2);
					objTableCell2.Append(objParagraph2);
					}
				}
			objTableRow.Append(objTableCell2);
			objServiceLevelTable.Append(objTableRow);

			// Create the Calculation Formula row for the table
			objTableRow = oxmlDocument.ConstructTableRow(parHasCondinalStyle: false);
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
			// Add the Calculation Formula Title in the first Column
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_SLtable_RowCalculationFormula_Title);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow.Append(objTableCell1);
			// Add the Calculation Formula value into the second Column
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2);
			objParagraph2 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun2 = oxmlDocument.Construct_RunText(parText2Write: parCalculationFormula);
			objParagraph2.Append(objRun2);
			objTableCell2.Append(objParagraph2);
			objTableRow.Append(objTableCell2);
			objServiceLevelTable.Append(objTableRow);

			// Create the Service Level Threshold row for the table
			objTableRow = oxmlDocument.ConstructTableRow(parHasCondinalStyle: false);
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
			// Add the Service Level Threshhold Title in the first Column
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_SLtable_RowThresholds_Title);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow.Append(objTableCell1);
			// Add the Service Level Threshold value into the second Column
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2);
			// the Service Level Threshold is in a list of String, process each entry and add it as a prargraph to the Table cell
			if(parThresholds.Count > 0)
				{
				foreach(ServiceLevelTarget thresholdEntry in parThresholds)
					{
					objParagraph2 = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1, parIsTableParagraph: true);
					objRun2 = oxmlDocument.Construct_RunText(parText2Write: thresholdEntry.Title);
					objParagraph2.Append(objRun2);
					objTableCell2.Append(objParagraph2);
					}
				}
			else
				{
				objParagraph2 = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1, parIsTableParagraph: true);
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_SLtable_ValueNotSpecified_Text);
				objParagraph2.Append(objRun2);
				objTableCell2.Append(objParagraph2);
				}
			objTableRow.Append(objTableCell2);
			objServiceLevelTable.Append(objTableRow);

			// Create the Service Level Targets row for the table
			objTableRow = oxmlDocument.ConstructTableRow(parHasCondinalStyle: false);
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
			// Add the Service Level Targets Title in the first Column
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_SLtable_RowTargets_Title);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow.Append(objTableCell1);
			// Add the Service Level Targets value in the second Column
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2);
			objParagraph2 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			if(parTargets.Count > 0)
				{
				foreach(ServiceLevelTarget targetEntry in parTargets)
					{
					objParagraph2 = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1, parIsTableParagraph: true);
					objRun2 = oxmlDocument.Construct_RunText(parText2Write: targetEntry.Title);
					objParagraph2.Append(objRun2);
					objTableCell2.Append(objParagraph2);
					}
				}
			else
				{
				objParagraph2 = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1, parIsTableParagraph: true);
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_SLtable_ValueNotSpecified_Text);
				objParagraph2.Append(objRun2);
				objTableCell2.Append(objParagraph2);
				}
			objTableRow.Append(objTableCell2);
			objServiceLevelTable.Append(objTableRow);

			// Create the Service Level Conditions row for the table
			objTableRow = oxmlDocument.ConstructTableRow(parHasCondinalStyle: false);
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
			// Add the Service Level Conditions Title in the first Column
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_SLtable_RowConditions_Title);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow.Append(objTableCell1);
			// Add the Service Level Conditions content in the second Column
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2);
			// Decode the RichText content using the RTdecoder object and DecodeRichText method
			listParagraphs.Clear();
			if(parBasicServiceLevelConditions == null && parAdditionalServiceLevelConditions == null)
				{
				objParagraph2 = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1, parIsTableParagraph: true);
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_SLtable_ValueNotSpecified_Text);
				objParagraph2.Append(objRun2);
				objTableCell2.Append(objParagraph2);
				}
			else
				{
				if(parBasicServiceLevelConditions != null)
					{
					try
						{
						listParagraphs = objRTdecoder.DecodeRichText(parRT2decode: parBasicServiceLevelConditions, parIsTableText: true);
						foreach(Paragraph paragraphItem in listParagraphs)
							{
							objTableCell2.Append(paragraphItem);
							}
						}
					catch(InvalidRichTextFormatException exc)
						{
						Console.WriteLine("\n\nException occurred: {0}", exc.Message);
						// A Table content error occurred, record it in the error log.
						parErrorMessages.Add("Service Level ID: " + parServiceLevelID + " Calculation Method attribute " +
							" contains an error in one of its Enhance Rich Text columns. Please review the content (especially tables).");
						objParagraph2 = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 0, parIsTableParagraph: true);
						objRun2 = oxmlDocument.Construct_RunText(
							parText2Write: "A content error occurred at this position and valid content could " +
							"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it. [" + exc.Message + "]",
							parIsError: true);
						objParagraph2.Append(objRun2);
						objTableCell2.Append(objParagraph2);
						}
					}
				
				// Insert the additional Service Level Conditions if ther are any.
				// Decode the RichText content using the RTdecoder object and DecodeRichText method
				listParagraphs.Clear();
				if(parAdditionalServiceLevelConditions != null)
					{
					// Add the Additional Service Level conditions into the second Column
					objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2);
					objParagraph2 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
					objRun2 = oxmlDocument.Construct_RunText(parText2Write: parAdditionalServiceLevelConditions);
					objParagraph2.Append(objRun2);
					objTableCell2.Append(objParagraph2);
					}
				}
			objTableRow.Append(objTableCell2);
			objServiceLevelTable.Append(objTableRow);

			//Return the constructed Table object
			return objServiceLevelTable;
			}// End of method.

		///############################################################################################
		/// <summary>
		/// This procedure use the input parameters to construct a Table of Glossary terms and Acronyms.
		/// </summary>
		/// <param name="parDictionaryGlossaryAcronym">A glossary containing GlossaryAcronym Id as Key MUST be passed as an Input Parameter.</param>
		/// <param name="parWidthColumn1">Specify the width of the first column in Dxa</param>
		/// <param name="parWidthColumn2">Specify the width of the second column in Dxa</param>
		/// <param name="parWidthColumn3">Specify the width of the third column in Dxa</param>
		/// <param name="parErrorMessages">Pass a reference to the ErrorMessages to ensure any errors that may occur is added to the ErrorMessaged.</param>
		/// <returns>
		/// The procedure returns a formated TABLE object consisting of 3 Columns Term, Acronym Meaning and it contains multiple Rows- one for each  term.</returns>
		public static DocumentFormat.OpenXml.Wordprocessing.Table BuildGlossaryAcronymsTable(
			DesignAndDeliveryPortfolioDataContext parSDDPdatacontext,
			Dictionary<int, string> parDictionaryGlossaryAcronym,
			UInt32 parWidthColumn1,
			UInt32 parWidthColumn2,
			UInt32 parWidthColumn3,
			ref List<string> parErrorMessages)
			{

			// Initialize the ServiceLevel table object
			DocumentFormat.OpenXml.Wordprocessing.Table objGlossaryAcronymsTable = new DocumentFormat.OpenXml.Wordprocessing.Table();
			objGlossaryAcronymsTable = oxmlDocument.ConstructTable(
				parPageWidth: 0,
				parFirstRow: true,
				parNoVerticalBand: true,
				parNoHorizontalBand: false);

			// Construct the TableGrid
			TableGrid objTableGrid = new TableGrid();
			List<UInt32> lstTableColumns = new List<UInt32>();
			lstTableColumns.Add(parWidthColumn1);
			lstTableColumns.Add(parWidthColumn2);
			lstTableColumns.Add(parWidthColumn3);
			objTableGrid = oxmlDocument.ConstructTableGrid(lstTableColumns);
			// Append the TableGrid object instance to the Table object instance
			objGlossaryAcronymsTable.Append(objTableGrid);

			// Construct the Heading row of the table
			TableRow objTableRow = new TableRow();
			objTableRow = oxmlDocument.ConstructTableRow();
			// Construct the first Column Heading
			TableCell objTableCell1 = new TableCell();
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1, parIsFirstRow: true);
			// Add Column1 Title for the row
			Paragraph objParagraph1 = new Paragraph();
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			DocumentFormat.OpenXml.Wordprocessing.Run objRun1 = new DocumentFormat.OpenXml.Wordprocessing.Run();
			objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_TableColumn_GlossaryAcronyms_Column1_Heading);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow.Append(objTableCell1);
			// Construct Column2 Title for the row
			TableCell objTableCell2 = new TableCell();
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2, parIsFirstRow: true);
			Paragraph objParagraph2 = new Paragraph();
			objParagraph2 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			DocumentFormat.OpenXml.Wordprocessing.Run objRun2 = new DocumentFormat.OpenXml.Wordprocessing.Run();
			objRun2 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_TableColumn_GlossaryAcronyms_Column2_Heading);
			objParagraph2.Append(objRun2);
			objTableCell2.Append(objParagraph2);
			objTableRow.Append(objTableCell2);
			// Add Column3 Title for the row
			TableCell objTableCell3 = new TableCell();
			objTableCell3 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn3, parIsFirstRow: true);
			Paragraph objParagraph3 = new Paragraph();
			objParagraph3 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			DocumentFormat.OpenXml.Wordprocessing.Run objRun3 = new DocumentFormat.OpenXml.Wordprocessing.Run();
			objRun3 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_TableColumn_GlossaryAcronyms_Column3_Heading);
			objParagraph3.Append(objRun3);
			objTableCell3.Append(objParagraph3);
			objTableRow.Append(objTableCell3);
			// append the Row object to the Table object
			objGlossaryAcronymsTable.Append(objTableRow);

			// Process the Terms and Acronyms passed in the parDictionaryGlossaryAcronyms
			List<GlossaryAcronym> listGlosaryAcronym = new List<GlossaryAcronym>();
			foreach(var item in parDictionaryGlossaryAcronym)
				{
				Console.WriteLine("\t ID: {0} - {1} was read...", item.Key, item.Value);
				var rsGlossaryAcronyms =
					from term in parSDDPdatacontext.GlossaryAndAcronyms
					where term.Id == item.Key
					select new
						{
						term.Id,
						term.Title,
						term.Acronym,
						term.Definition
						};
				var recGlossaryAcronym = rsGlossaryAcronyms.FirstOrDefault();
				if(recGlossaryAcronym == null)
					{
					Console.WriteLine("\t\t ### ENTRY NOT FOUND ###");
					continue; // process the next entry
					}
				Console.WriteLine("\t\t + {0} - {1} \n\t\t - {2}", recGlossaryAcronym.Acronym, recGlossaryAcronym.Title, recGlossaryAcronym.Definition);
				// populate the Glossary and Acronym object...
				GlossaryAcronym objGlossaryAcronym = new GlossaryAcronym();
				objGlossaryAcronym.ID = recGlossaryAcronym.Id;
				objGlossaryAcronym.Term = recGlossaryAcronym.Title;
				objGlossaryAcronym.Acronym = recGlossaryAcronym.Acronym;
				objGlossaryAcronym.Meaning = recGlossaryAcronym.Definition;
				// add the Glossary and Acronym object to the List of Glossary and Acronym objects.
				listGlosaryAcronym.Add(objGlossaryAcronym);

				} //foreach Loop

			Console.WriteLine("Total Glossary and Acronyms processed: {0}", listGlosaryAcronym.Count);

			// Sort the list Alphabetically by Term
			listGlosaryAcronym.Sort(delegate (GlossaryAcronym x, GlossaryAcronym y)
			{
				if(x.Term == null && y.Term == null)
					return 0;
				else if(x.Term == null)
					return -1;
				else if(y.Term == null)
					return 1;
				else
					return x.Term.CompareTo(y.Term);
			});

			// Process the sorted List of Glossary and Acronym Objects.
			foreach(GlossaryAcronym item in listGlosaryAcronym)
				{
				objTableRow = oxmlDocument.ConstructTableRow(parHasCondinalStyle: true);
				// Construct the first Column cell with the Term
				objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
				objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
				objRun1 = oxmlDocument.Construct_RunText(parText2Write: item.Term);
				objParagraph1.Append(objRun1);
				objTableCell1.Append(objParagraph1);
				objTableRow.Append(objTableCell1);
				// Construct Column2 cell with the Acronym
				objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2);
				objParagraph2 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: item.Acronym);
				objParagraph2.Append(objRun2);
				objTableCell2.Append(objParagraph2);
				objTableRow.Append(objTableCell2);
				// Construct Column3 cell with the Definition/Meaning
				objTableCell3 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn3);
				objParagraph3 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
				objRun3 = oxmlDocument.Construct_RunText(parText2Write: item.Meaning);
				objParagraph3.Append(objRun3);
				objTableCell3.Append(objParagraph3);
				objTableRow.Append(objTableCell3);
				// append the Row object to the Table object
				objGlossaryAcronymsTable.Append(objTableRow);

				} //foreach(GlossaryAcronym item in objListGlosaryAcronym)
				  // return the constructed table object
			return objGlossaryAcronymsTable;
			} // end of method

		//############################################################################################
		/// <summary>
		/// This procedure use the input parameters to construct a Table of Mapping Risks.
		/// </summary>
		/// <param name="parMappingRisk">An object containing MappingRisk MUST be passed as an Input Parameter.</param>
		/// <param name="parWidthColumn1">Specify the width of the first column in Dxa</param>
		/// <param name="parWidthColumn2">Specify the width of the second column in Dxa</param>
		/// <param name="parErrorMessages">Pass a reference to the ErrorMessages to ensure any errors that may occur is added to the ErrorMessaged.</param>
		/// <returns>
		/// The procedure returns a formated TABLE object consisting of 2 Columns Title and value - it contains multiple Rows- one for each risk.</returns>
		public static DocumentFormat.OpenXml.Wordprocessing.Table BuildRiskTable(
			MappingRisk parMappingRisk,
			UInt32 parWidthColumn1,
			UInt32 parWidthColumn2)
			{

			// Initialize the Mapping table object
			DocumentFormat.OpenXml.Wordprocessing.Table objMappingRiskTable = new DocumentFormat.OpenXml.Wordprocessing.Table();
			objMappingRiskTable = oxmlDocument.ConstructTable(
				parPageWidth: 0,
				parNoVerticalBand: true,
				parNoHorizontalBand: true);

			// Construct the TableGrid
			TableGrid objTableGrid = new TableGrid();

			List<UInt32> lstTableColumns = new List<UInt32>();
			lstTableColumns.Add(parWidthColumn1);
			lstTableColumns.Add(parWidthColumn2);
			objTableGrid = oxmlDocument.ConstructTableGrid(lstTableColumns);
			// Append the TableGrid object instance to the Table object instance
			objMappingRiskTable.Append(objTableGrid);

			// Process the Risk passed in the parMapping
			
			TableCell objTableCell1 = new TableCell();
			TableCell objTableCell2 = new TableCell();
			Paragraph objParagraph1 = new Paragraph();
			Paragraph objParagraph2 = new Paragraph();
			DocumentFormat.OpenXml.Wordprocessing.Run objRun1 = new DocumentFormat.OpenXml.Wordprocessing.Run();
			DocumentFormat.OpenXml.Wordprocessing.Run objRun2 = new DocumentFormat.OpenXml.Wordprocessing.Run();

			TableRow objTableRow1 = new TableRow();
			objTableRow1 = oxmlDocument.ConstructTableRow(parHasCondinalStyle: false);
			// Construct the first Column cell for the Risk Statement Row.
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_RequirementsMapping_RiskTable_RiskStatement);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow1.Append(objTableCell1);
			// Construct Column2 cell with the Risk Statement Value
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2);
			objParagraph2 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			if(parMappingRisk.Statement == null)
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: " ");
			else
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: parMappingRisk.Statement);
			objParagraph2.Append(objRun2);
			objTableCell2.Append(objParagraph2);
			objTableRow1.Append(objTableCell2);
			// append the Row object to the Table object
			objMappingRiskTable.Append(objTableRow1);

			// Construct the first Column cell for the Risk Mitigation Row.
			TableRow objTableRow2 = new TableRow();
			objTableRow2 = oxmlDocument.ConstructTableRow(parHasCondinalStyle: false);
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_RequirementsMapping_RiskTable_RiskMitigation);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow2.Append(objTableCell1);
			// Construct Column2 cell with the Risk Mitigation Value
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2);
			objParagraph2 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			if(parMappingRisk.Mitigation == null)
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: " ");
			else
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: parMappingRisk.Mitigation);
			objParagraph2.Append(objRun2);
			objTableCell2.Append(objParagraph2);
			objTableRow2.Append(objTableCell2);
			// append the Row object to the Table object
			objMappingRiskTable.Append(objTableRow2);

			// Construct the first Column cell for the Risk Exposure Row.
			TableRow objTableRow3 = new TableRow();
			objTableRow3 = oxmlDocument.ConstructTableRow(parHasCondinalStyle: false);
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_RequirementsMapping_RiskTable_RiskExposure);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow3.Append(objTableCell1);
			// Construct Column2 cell with the Risk Exposure Value
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2);
			objParagraph2 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			if(parMappingRisk.Exposure == null)
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: " ");
			else
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: parMappingRisk.Exposure);
			objParagraph2.Append(objRun2);
			objTableCell2.Append(objParagraph2);
			objTableRow3.Append(objTableCell2);
			// append the Row object to the Table object
			objMappingRiskTable.Append(objTableRow3);

			// Construct the first Column cell for the Risk Exposure Value Row.
			TableRow objTableRow4 = new TableRow();
			objTableRow4 = oxmlDocument.ConstructTableRow(parHasCondinalStyle: false);
			objTableCell1 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn1);
			objParagraph1 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			objRun1 = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_RequirementsMapping_RiskTable_RiskExposureValue);
			objParagraph1.Append(objRun1);
			objTableCell1.Append(objParagraph1);
			objTableRow4.Append(objTableCell1);
			// Construct Column2 cell with the Risk Exposure Value
			objTableCell2 = oxmlDocument.ConstructTableCell(parCellWidth: parWidthColumn2);
			objParagraph2 = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
			if(parMappingRisk.ExposureValue == null)
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: " ");
			else
				objRun2 = oxmlDocument.Construct_RunText(parText2Write: parMappingRisk.ExposureValue.ToString());
			objParagraph2.Append(objRun2);
			objTableCell2.Append(objParagraph2);
			objTableRow4.Append(objTableCell2);
			// append the Row object to the Table object
			objMappingRiskTable.Append(objTableRow4);
			// return the constructed table object
			return objMappingRiskTable;
			} // end of method

		} // end of CommonProcedures Class

	} // End of NameSpace