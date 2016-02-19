using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using mshtml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DrwWp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DrwWp2010 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using Drw = DocumentFormat.OpenXml.Drawing;
using Drw2010 = DocumentFormat.OpenXml.Office2010.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading.Tasks;

/// <summary>
///	Mapped to the [Content Layer Colour Coding Option] column in SharePoint List
/// </summary>
enum enumCaptionType
	{
	Image = 1,
	Table = 2
	}

namespace DocGenerator
	{
	class HTMLdecoder
		{
		// ------------------
		// Object Properties
		// ------------------
		/// <summary>
		/// Set the WordProcessing Body immediately after declaring an instance of the HTMLdecoder object
		/// The oXMLencoder requires the WPBody object by reference to add the decoded HTML to the oXML document.
		/// </summary>
		private Body _wpbody = new Body();
		public Body WPbody
			{
			get { return this._wpbody; }
			set { this._wpbody = value; }
			}

		/// <summary>
		/// The Document Hierarchical Level provides the stating Hierarchical level at which new content will be added to the document.
		/// </summary>	
		private int _documentHierarchyLevel = 0;
		public int DocumentHierachyLevel
			{
			get{return this._documentHierarchyLevel;}
			set{this._documentHierarchyLevel = value;}
			}
		/// <summary>
		/// The Additional Hierarchical Level property contains the number of additional levels that need to be added to the 
		/// Document Hierarchical Level when processing the HTML contained in a Enhanced Rich Text column/field.
		/// </summary>
		private int _additionalHierarchicalLevel = 0;
		private int AdditionalHierarchicalLevel
			{
			get { return this._additionalHierarchicalLevel; }
			set { this._additionalHierarchicalLevel = value; }
			}
		/// <summary>
		/// The PageWidth property contains the page width of the OXML page into which the decoded HTML content 
		/// will be inserted. It is mostly used for image and table positioning on the page in the OXML document.
		/// </summary>

		private int _tableCaptionCounter = 0;
		public int TableCaptionCounter
			{
			get{return this._tableCaptionCounter;}
			set{this._tableCaptionCounter = value;}
			}

		private int _imageCaptionCounter = 0;
		public int ImageCaptionCounter
			{
			get{return this._imageCaptionCounter;}
			set{this._imageCaptionCounter = value;}
			}

		private UInt32 _pageWidth = 0;
		private UInt32 PageWidth
			{
			get { return this._pageWidth; }
			set { this._pageWidth = value; }
			}

		private UInt32 _pageHeight = 0;
		private UInt32 PageHeight
			{
			get{return this._pageHeight;}
			set{this._pageHeight = value;}
			}

		/// <summary>
		/// When working with a table, this property contains the width of the table
		/// </summary>
		private UInt32 _tableWidth = 0;
		public UInt32 TableWidth
			{
			get{return this._tableWidth;}
			set{this._tableWidth = value;}
			}

		/// <summary>
		/// The InTableMode property is set to TRUE as soon as a table is in process, and
		/// it is set to FALSE as soon as the processing of a table ends/ is completed.
		/// </summary>
		public bool _inTableMode = false;
		private bool InTableMode
			{
			get{return this._inTableMode;}
			set{this._inTableMode = value;}
			}

		/// <summary>
		/// The TableColumnWidths is a List (array) containing and entry/occurrance representing the width of every column in the table.
		/// </summary>
		private List<UInt32> _tableColumnWidths = new List<UInt32>();
		public List<UInt32> TableColumnWidths
			{
			get{return this._tableColumnWidths;}
			set{this._tableColumnWidths = value;}
			}

		/// <summary>
		/// The TableColumnUnit describe the units used for the TableColumn widths.
		/// </summary>
		private string _tableColumnUnit = "";
		public string TableColumnUnit
			{
			get{return this._tableColumnUnit;}
			set{this._tableColumnUnit = value;}
			}

		/// <summary>
		/// The WPdocTable property is an WordProcessing.Table type object and it will contain a completely constructed OXML table
		/// while it is constructed until it is completely build, after which it will be appended to a WPbody object.
		/// </summary>
		private DocumentFormat.OpenXml.Wordprocessing.Table _wpdocTable;
		public DocumentFormat.OpenXml.Wordprocessing.Table WPdocTable
			{
			get{return this._wpdocTable;}
			set{this._wpdocTable = value;}
			}

		private bool _tableGridDone = false;
		public bool TableGridDone
			{
			get{return this._tableGridDone;}
			set{this._tableGridDone = value;}
			}

		/// <summary>
		/// This propoerty indicates the type of row that are build.
		/// </summary>
		private string _currentTableRowType = "";
		public string CurrentTableRowType
			{
			get {return this._currentTableRowType;}
			set{this._currentTableRowType = value;}
			}

		/// <summary>
		/// Indicates whether the Table has a Header Row or Table Header.
		/// </summary>
		private bool _TableHasFisrtRow = false;
		public bool TableHasFirstRow
			{
			get{return this._TableHasFisrtRow;}
			set{this._TableHasFisrtRow = value;}
			}

		/// <summary>
		/// Indicates whether the Table has a Last Row or Table Footer.
		/// </summary>
		private bool _TableHasLastRow = false;
		public bool TableHasLastRow
			{
			get{return this._TableHasLastRow;}
			set{this._TableHasLastRow = value;}
			}

		/// <summary>
		/// Indicates whether the Table has a First Column.
		/// </summary>
		private bool _TableHasFisrtColumn = false;
		public bool TableHasFirstColumn
			{
			get{return this._TableHasFisrtColumn;}
			set{this._TableHasFisrtColumn = value;}
			}

		/// <summary>
		/// Indicates whether the Table has a Last Column.
		/// </summary>
		private bool _TableHasLastColumn = false;
		public bool TableHasLastColumn
			{
			get{return this._TableHasLastColumn;}
			set{this._TableHasLastColumn = value;}
			}

		/// <summary>
		/// This property will contain the text of the Caption Text to be added after an image or table
		/// </summary>
		private string _captionText;
		public string CaptionText
			{
			get{return this._captionText;}
			set{this._captionText = value;}
			}

		/// <summary>
		///	This property indicates the type of caption that need to be inserted after an image or table.
		/// </summary>
		private enumCaptionType _captionType;
		public enumCaptionType CaptionType
			{
			get{return this._captionType;}
			set{this._captionType = value;}
			}

		private string _hyperlinkImageRelationshipID = "";
		public string HyperlinkImageRelationshipID
			{
			get{return this._hyperlinkImageRelationshipID;}
			set{this._hyperlinkImageRelationshipID = value;}
			}
		/// <summary>
		/// contains the actual hyperlink URL that will be inserted if required.
		/// </summary>
		private string _hyperlinkURL = "";
		public string HyperlinkURL
			{
			get{return this._hyperlinkURL;}
			set{this._hyperlinkURL = value;}
			}

		/// <summary>
		/// Indicator property that are set once a Hyperlink was inserted for an HTML run
		/// </summary>
		private bool _hypelinkInserted = false;
		public bool HyperlinkInserted
			{
			get{return this._hypelinkInserted;}
			set{this._hypelinkInserted = value;}
			}

// ---------------------
//--- Object Methods ---
// ---------------------

	//------------------
	//--- DecodeHTML ---
	//------------------
	/// <summary>
	/// Use this method once a new HTMLdecoder object is initialised and the 
	/// EndodedHTML property was set to the value of the HTML that has to be decoded.
	/// </summary>
	/// <param name="parDocumentLevel">
	/// Provide the document's hierarchical level at which the HTML has to be inserted.
	/// </param>
	/// <param name="parPageWidth">
	/// </param>
	/// <param name="parHTML2Decode">
	/// </param>
	/// <returns>
	/// returns a boolean value of TRUE if insert was successfull and FALSE if there was any form of failure during the insertion.
	/// </returns>
		public bool DecodeHTML(
			ref MainDocumentPart parMainDocumentPart,
			int parDocumentLevel, 
			UInt32 parPageWidthTwips, 
			UInt32 parPageHeightTwips,
			string parHTML2Decode,
			ref int parTableCaptionCounter,
			ref int parImageCaptionCounter,
			string parHyperlinkURL = "",
			string parHyperlinkImageRelationshipID = "")
			
			{
			Console.WriteLine("HTML to decode:\n{0}", parHTML2Decode);
			this.DocumentHierachyLevel = parDocumentLevel;
			this.AdditionalHierarchicalLevel = 0;
			this.PageWidth = parPageWidthTwips;
			this.PageHeight = parPageHeightTwips;
			this.TableCaptionCounter = parTableCaptionCounter;
			this.ImageCaptionCounter = parImageCaptionCounter;
			this.HyperlinkImageRelationshipID = parHyperlinkImageRelationshipID;
			this.HyperlinkURL = parHyperlinkURL;
			
			// http://stackoverflow.com/questions/11250692/how-can-i-parse-this-html-to-get-the-content-i-want
			IHTMLDocument2 objHTMLDocument2 = (IHTMLDocument2) new HTMLDocument();
			objHTMLDocument2.write(parHTML2Decode);

			//objHTMLDocument.body.innerHTML = this.EncodedHTML;
			//Console.WriteLine("{0}", objHTMLDocument2.body.innerHTML);
			Paragraph objParagraph = new Paragraph();
			objParagraph = oxmlDocument.Construct_Paragraph(1, false);
			ProcessHTMLelements(ref parMainDocumentPart, objHTMLDocument2.body.children, ref objParagraph, false);
			// Update the counters before returning
			parTableCaptionCounter = this.TableCaptionCounter;
			parImageCaptionCounter = this.ImageCaptionCounter;
			return true;
			}

	/// <summary>
	/// 
	/// </summary>
	/// <param name="parHTMLElements"></param>
	/// <param name="parExistingParagraph"></param>
	/// <param name="parAppendToExistingParagraph"></param>
		private void ProcessHTMLelements(
			ref MainDocumentPart parMainDocumentPart,
			IHTMLElementCollection parHTMLElements, 
			ref Paragraph parExistingParagraph, 
			bool parAppendToExistingParagraph)
			{

			Paragraph objNewParagraph = new Paragraph();

			if(parAppendToExistingParagraph)
				objNewParagraph = parExistingParagraph;
			
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();

			if(parHTMLElements.length > 0)
				{
				foreach(IHTMLElement objHTMLelement in parHTMLElements)
					{
					//Console.WriteLine("HTMLlevel: {0} - html.tag=<{1}>\n\r\t|{2}|", this.AdditionalHierarchicalLevel, objHTMLelement.tagName,objHTMLelement.innerHTML);
					switch(objHTMLelement.tagName)
						{
						//-----------------------
						case "DIV":
						//-----------------------
							if(objHTMLelement.children.length > 0)
								ProcessHTMLelements(
									ref parMainDocumentPart,
									objHTMLelement.children, ref objNewParagraph, false);
							else
								{
								if (objHTMLelement.innerText != null)
									objRun = oxmlDocument.Construct_RunText	(parText2Write: objHTMLelement.innerText);
								}
							break;
						//---------------------------
						case "P": // Paragraph Tag
						//---------------------------
							objNewParagraph = oxmlDocument.Construct_Paragraph(this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel);
							if(objHTMLelement.children.length > 0) // check if there are more html tags in the HTMLelement
								{
								//Console.WriteLine("\t{0} child nodes to process", objHTMLelement.children.length);
								// use the DissectHTMLstring method to process the paragraph.
								List<TextSegment> listTextSegments = new List<TextSegment>();
								listTextSegments = TextSegment.DissectHTMLstring (objHTMLelement.innerHTML);
								// Process the list to insert the content into the document.
								foreach(TextSegment objTextSegment in listTextSegments)
									{
									if(objTextSegment.Image) // Check if it is an image
										{
										IHTMLDocument2 objHTMLDocument2 = (IHTMLDocument2) new HTMLDocument();
										objHTMLDocument2.write(objTextSegment.Text);
										objNewParagraph = oxmlDocument.Construct_Paragraph(1, false);
										ProcessHTMLelements(ref parMainDocumentPart, objHTMLDocument2.body.children, ref objNewParagraph, false);
										}
									else // not an image
										{
										objRun = oxmlDocument.Construct_RunText
											(parText2Write: objTextSegment.Text,
											parBold: objTextSegment.Bold,
											parItalic: objTextSegment.Italic,
											parUnderline: objTextSegment.Undeline,
											parSubscript: objTextSegment.Subscript,
											parSuperscript: objTextSegment.Superscript);
										// Check if a hyperlink must be inserted
										if(this.HyperlinkImageRelationshipID != "")
											{
											if(this.HyperlinkInserted == false)
												{
												DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
													parMainDocumentPart: ref parMainDocumentPart,
													parImageRelationshipId: this.HyperlinkImageRelationshipID,
													parClickLinkURL: this.HyperlinkURL);
												objRun.Append(objDrawing);
												this.HyperlinkInserted = true;
                                                            }
											}
										objNewParagraph.Append(objRun);
										}
									}
								}
							else  // there are no cascading tags, just write the text if there are any
								{
								if(objHTMLelement.innerText.Length > 0)
									{
									objRun = oxmlDocument.Construct_RunText(parText2Write: objHTMLelement.innerText);
									// Check if a hyperlink must be inserted
									if(this.HyperlinkImageRelationshipID != "")
										{
										if(this.HyperlinkInserted == false)
											{
											DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
												parMainDocumentPart: ref parMainDocumentPart,
												parImageRelationshipId: this.HyperlinkImageRelationshipID,
												parClickLinkURL: this.HyperlinkURL);
											objRun.Append(objDrawing);
											this.HyperlinkInserted = true;
											}
										}
									objNewParagraph.Append(objRun);
									}
								}
							if(parAppendToExistingParagraph)
								//ignore because only a new Paragraph needs to be appended to the body
								Console.WriteLine("Skip the appending of the existing paragraph to the Body");
							else
								{
								this.WPbody.Append(objNewParagraph);
								}
							break;
						//------------------------------------
						case "TABLE":
							Console.WriteLine("Tag: TABLE\n{0}", objHTMLelement.outerHTML);
							if(this.InTableMode)
								{
								//TODO: Check for cascading tables and generate an ERROR when an occurrance of a cascading table is discovered.
								}
							else
								this.InTableMode = true;
							// Set the TableGridDone property to false, in order to get the grid defined.

							this.TableGridDone = false;

							UInt32 iTableWidth = 0;
							string TableWithUnit = "";
							if(objHTMLelement.outerHTML.IndexOf("WIDTH", 1) >= 0)
								{
								TableWithUnit = objHTMLelement.style.width;
								if(TableWithUnit.IndexOf("%", 1) > 0)
									{
									Console.WriteLine("\t The % is in position {0}", TableWithUnit.IndexOf("%", 0));
									Console.WriteLine("\t Numeric Value: {0}", TableWithUnit.Substring(0, (TableWithUnit.Length - TableWithUnit.IndexOf("%", 0)) + 1));
									if(!UInt32.TryParse(TableWithUnit.Substring(0, (TableWithUnit.Length - TableWithUnit.IndexOf("%", 1)) + 1), out iTableWidth))
										iTableWidth = 100;
									this.TableWidth = (this.PageWidth * iTableWidth / 100);
									TableWithUnit = "px";
									}
								else if(TableWithUnit.IndexOf("px", 1) > 0)
									{
									Console.WriteLine("\t The px is in position {0}", TableWithUnit.IndexOf("px", 0));
									Console.WriteLine("\t Numeric Value: {0}", TableWithUnit.Substring(0, (TableWithUnit.Length - TableWithUnit.IndexOf("px", 0)) + 1));
									if(!UInt32.TryParse(TableWithUnit.Substring(0, (TableWithUnit.Length - TableWithUnit.IndexOf("px", 1)) + 1), out iTableWidth))
										iTableWidth = this.PageWidth;
									TableWithUnit = "px";
									}
								}
							else // if the table's width is not defined.
								{
								iTableWidth = this.PageWidth;
								TableWithUnit = "px";
								}

							// Calculate the width of the table on the page.
							Console.WriteLine("\t Pagewidth: {0}", this.PageWidth);
							Console.WriteLine("\t Table Width: {0}%", iTableWidth);
							this.WPdocTable = oxmlDocument.ConstructTable(parPageWidth: this.TableWidth, 
								parFirstRow: false, 
								parFirstColumn: false, 
								parLastColumn: false, 
								parLastRow: false, 
								parNoVerticalBand: true, 
								parNoHorizontalBand: false);
							
							if(objHTMLelement.children.length > 0)
								ProcessHTMLelements(
									ref parMainDocumentPart,
									objHTMLelement.children, 
									ref objNewParagraph, 
									false);
							// Append the table to the WordProcessing.Body
							WPbody.Append(this.WPdocTable);
							//Get the Table Summary tag value and store it in the CaptionText value
							Console.WriteLine("\t Table Syummary: {0}", objHTMLelement.getAttribute("summary", 0));
							if(objHTMLelement.getAttribute("summary", 0) != "")
								{
								this.TableCaptionCounter += 1;
								objNewParagraph = oxmlDocument.Construct_Caption(
									parCaptionType: "Table",
									parCaptionSequence: this.TableCaptionCounter,
									parCaptionText: objHTMLelement.getAttribute("summary", 0));
								this.WPbody.Append(objNewParagraph);
								}
							this.WPdocTable = null;
							this.InTableMode = false;
							break;
						//------------------------------------
						case "TBODY": // Table Body
							Console.WriteLine("Tag: TABLE Body \n{0}", objHTMLelement.outerHTML);
							if(objHTMLelement.children.length > 0)
								ProcessHTMLelements(
									ref parMainDocumentPart,
									objHTMLelement.children, 
									ref objNewParagraph,
									false);
							break;
						//------------------------------------
						case "TR":     // Table Row
							Console.WriteLine("Tag: TR [Table Row]: {0}\n{1}", objHTMLelement.className, objHTMLelement.outerHTML);
							//if the table grid has NOT been defined yet, Define the Table Grid, before continue with processing
							if(!this.TableGridDone)
								{
								DetermineTableGrid(objHTMLelement.children);
								DocumentFormat.OpenXml.Wordprocessing.TableGrid objTableGrid = new TableGrid();
								objTableGrid = oxmlDocument.ConstructTableGrid(this.TableColumnWidths, this.TableColumnUnit, this.PageWidth);
								this.WPdocTable.Append(objTableGrid);
								this.TableGridDone = true;
								}

							//Check the type of Table row
							if(objHTMLelement.className.Contains("TableHeaderRow"))
								{
								this.CurrentTableRowType = "Header";
								this.TableHasFirstRow = true;
								TableProperties objTableProperties = this.WPdocTable.GetFirstChild<TableProperties>();
								TableLook objTableLook = objTableProperties.GetFirstChild<TableLook>();
								objTableLook.FirstRow = true;
								// Append a Table Header row to the table if the TableRow is the Header Row 
								DocumentFormat.OpenXml.Wordprocessing.TableRow objTableRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
								objTableRow = oxmlDocument.ConstructTableRow(
									parIsFirstRow: true,
									parIsLastRow: false,
									parIsFirstColumn: false,
									parIsLastColumn: false,
									parIsEvenHorizontalBand: false,
									parIsOddHorizontalBand: false);
								this.WPdocTable.Append(objTableRow);
								}
							else if(objHTMLelement.className.Contains("TableFooterRow"))
								{
								this.CurrentTableRowType = "Footer";
								this.TableHasLastRow = true;
								TableProperties objTableProperties = this.WPdocTable.GetFirstChild<TableProperties>();
								TableLook objTableLook = objTableProperties.GetFirstChild<TableLook>();
								objTableLook.LastRow = true;
								// Append a Table Header row to the table
								DocumentFormat.OpenXml.Wordprocessing.TableRow objTableRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
								objTableRow = oxmlDocument.ConstructTableRow(
									parIsFirstRow: false,
									parIsLastRow: true,
									parIsFirstColumn: false,
									parIsLastColumn: false,
									parIsEvenHorizontalBand: false,
									parIsOddHorizontalBand: false);
								this.WPdocTable.Append(objTableRow);
								}
							else if(objHTMLelement.className.Contains("OddRow"))
								{
								this.CurrentTableRowType = "NormalOdd";
								// Append a Table Header row to the table
								DocumentFormat.OpenXml.Wordprocessing.TableRow objTableRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
								objTableRow = oxmlDocument.ConstructTableRow(
									parIsFirstRow: false,
									parIsLastRow: false,
									parIsFirstColumn: false,
									parIsLastColumn: false,
									parIsEvenHorizontalBand: false,
									parIsOddHorizontalBand: true);
								this.WPdocTable.Append(objTableRow);
								}
							else if(objHTMLelement.className.Contains("EvenRow"))
								{
								this.CurrentTableRowType = "NormalEven";
								// Append a Table Header row to the table
								DocumentFormat.OpenXml.Wordprocessing.TableRow objTableRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
								objTableRow = oxmlDocument.ConstructTableRow(
									parIsFirstRow: false,
									parIsLastRow: false,
									parIsFirstColumn: false,
									parIsLastColumn: false,
									parIsEvenHorizontalBand: true,
									parIsOddHorizontalBand: false);
								this.WPdocTable.Append(objTableRow);
								}
							else
								{
								this.CurrentTableRowType = "";
								// Append a Table Header row to the table
								DocumentFormat.OpenXml.Wordprocessing.TableRow objTableRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
								objTableRow = oxmlDocument.ConstructTableRow(
									parIsFirstRow: false,
									parIsLastRow: false,
									parIsFirstColumn: false,
									parIsLastColumn: false,
									parIsEvenHorizontalBand: false,
									parIsOddHorizontalBand: false);
								this.WPdocTable.Append(objTableRow);
								}

							// Process the children (TH and TD) of the Table Row
							if(objHTMLelement.children.length > 0)
								{
								ProcessHTMLelements(
									ref parMainDocumentPart,
									objHTMLelement.children,
									ref objNewParagraph,
									false);
								}
							break;
						//------------------------------------
						case "TH":     // Table Header
						case "TD":     // Table Cell
							//Console.WriteLine("Tag: TH [Table Header]: {0}\n{1}",objHTMLelement.className, objHTMLelement.outerHTML);
							//Console.WriteLine("\tStyle=width: {0}", objHTMLelement.style.width);
							//Console.WriteLine("\tinnerText: {0}", objHTMLelement.innerText);
							TableCell objTableCell = new TableCell();
							// Determine the width of the Cell
							UInt32 iCellWidthValue = 0;
							string cellWithUnit = "";
							if(objHTMLelement.outerHTML.IndexOf("WIDTH", 1) >= 0)
								{
								cellWithUnit = objHTMLelement.style.width;
								if(cellWithUnit.IndexOf("%", 1) > 0)
									{
									Console.WriteLine("\t The % is in position {0}", cellWithUnit.IndexOf("%", 0));
									Console.WriteLine("\t Numeric Value: {0}", cellWithUnit.Substring(0, (cellWithUnit.Length - cellWithUnit.IndexOf("%", 0)) + 1));
									if(!UInt32.TryParse(cellWithUnit.Substring(0, (cellWithUnit.Length - cellWithUnit.IndexOf("%", 1)) + 1), out iCellWidthValue))
										iCellWidthValue = 25;
									iCellWidthValue = (this.TableWidth * iCellWidthValue) / 100;
									cellWithUnit = "px";
									}
								else if(cellWithUnit.IndexOf("px", 1) > 0)
									{
									Console.WriteLine("\t The px is in position {0}", cellWithUnit.IndexOf("px", 0));
									Console.WriteLine("\t Numeric Value: {0}", cellWithUnit.Substring(0, (cellWithUnit.Length - cellWithUnit.IndexOf("px", 0)) + 1));
									if(!UInt32.TryParse(cellWithUnit.Substring(0, (cellWithUnit.Length - cellWithUnit.IndexOf("px", 1)) + 1), out iCellWidthValue))
										iCellWidthValue = 600;
									cellWithUnit = "px";
									}
								}
							//Console.WriteLine("\t The Cell Width = {0}{1}", iCellWidthValue, cellWithUnit);
							//Console.WriteLine("\t Parent Element Classname: {0}", objHTMLelement.parentElement.className);
							//Console.WriteLine("\t Current Element Classname: {0}", objHTMLelement.className);
							if(objHTMLelement.parentElement.className.Contains("TableHeaderRow"))
								{
								if(objHTMLelement.className.Contains("TableHeaderFirstCol"))
									{
									TableProperties objTableProperties = this.WPdocTable.GetFirstChild<TableProperties>();
									TableLook objTableLook = objTableProperties.GetFirstChild<TableLook>();
									objTableLook.FirstRow = true;
									objTableLook.FirstColumn = true;
									// add the table cell to the LAST TableRow
									objTableCell = oxmlDocument.ConstructTableCell(iCellWidthValue, parIsFirstColumn: true);
									}
								else if(objHTMLelement.className.Contains("TableHeaderLastCol"))
									{
									TableProperties objTableProperties = this.WPdocTable.GetFirstChild<TableProperties>();
									TableLook objTableLook = objTableProperties.GetFirstChild<TableLook>();
									objTableLook.FirstRow = true;
									objTableLook.LastColumn = true;
									objTableCell = oxmlDocument.ConstructTableCell(iCellWidthValue, parFirstRowLastColumn: true);
									}
								else
									{
									TableProperties objTableProperties = this.WPdocTable.GetFirstChild<TableProperties>();
									TableLook objTableLook = objTableProperties.GetFirstChild<TableLook>();
									objTableLook.FirstRow = true;
									objTableCell = oxmlDocument.ConstructTableCell(iCellWidthValue, parIsFirstRow: true);
									}
								}
							else if(objHTMLelement.parentElement.className.Contains("TableFooterRow"))
								{
								if(objHTMLelement.className.Contains("TableFooterFirstCol"))
									{
									TableProperties objTableProperties = this.WPdocTable.GetFirstChild<TableProperties>();
									TableLook objTableLook = objTableProperties.GetFirstChild<TableLook>();
									objTableLook.LastRow = true;
									objTableLook.FirstColumn = true;
									// add the table cell to the LAST TableRow
									objTableCell = oxmlDocument.ConstructTableCell(iCellWidthValue, parIsFirstColumn: true, parLastRowFirstColumn: true);
									}
								else if(objHTMLelement.className.Contains("TableFooterLastCol"))
									{
									TableProperties objTableProperties = this.WPdocTable.GetFirstChild<TableProperties>();
									TableLook objTableLook = objTableProperties.GetFirstChild<TableLook>();
									objTableLook.LastRow = true;
									objTableLook.LastColumn = true;
									objTableCell = oxmlDocument.ConstructTableCell(iCellWidthValue, parFirstRowLastColumn: true, parLastRowLastColumn: true);
									}
								else
									{
									TableProperties objTableProperties = this.WPdocTable.GetFirstChild<TableProperties>();
									TableLook objTableLook = objTableProperties.GetFirstChild<TableLook>();
									objTableLook.LastRow = true;
									objTableCell = oxmlDocument.ConstructTableCell(iCellWidthValue, parIsLastRow: true);
									}
								}
							else   // not a table Header or Footer column
								{
								if(objHTMLelement.className.Contains("TableFirstCol"))
									{
									objTableCell = oxmlDocument.ConstructTableCell(iCellWidthValue, parIsFirstColumn: true);
									}
								else if(objHTMLelement.className.Contains("TableLastCol"))
									{
									objTableCell = oxmlDocument.ConstructTableCell(iCellWidthValue, parIsLastColumn: true);
									}
								else
									{
									objTableCell = oxmlDocument.ConstructTableCell(iCellWidthValue);
									}
								}

							// Check if the TableHeader has Children...
							objNewParagraph = oxmlDocument.Construct_Paragraph(0, true);
							
							if(objHTMLelement.children.length > 0) // check if there are more html tags in the HTMLelement
								{
								Console.WriteLine("\t{0} child nodes to process", objHTMLelement.children.length);
								// use the DissectHTMLstring method to process the paragraph.
								List<TextSegment> listTextSegments = new List<TextSegment>();
								listTextSegments = TextSegment.DissectHTMLstring(objHTMLelement.innerHTML);
								foreach(TextSegment objTextSegment in listTextSegments)
									{
									objRun = oxmlDocument.Construct_RunText
											(parText2Write: objTextSegment.Text,
											parBold: objTextSegment.Bold,
											parItalic: objTextSegment.Italic,
											parUnderline: objTextSegment.Undeline,
											parSubscript: objTextSegment.Subscript,
											parSuperscript: objTextSegment.Superscript);
									// Check if a hyperlink must be inserted
									if(this.HyperlinkImageRelationshipID != "")
										{
										if(this.HyperlinkInserted == false)
											{
											DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
												parMainDocumentPart: ref parMainDocumentPart,
												parImageRelationshipId: this.HyperlinkImageRelationshipID,
												parClickLinkURL: this.HyperlinkURL);
											objRun.Append(objDrawing);
											this.HyperlinkInserted = true;
											}
										}
									objNewParagraph.Append(objRun);
									}
								objTableCell.Append(objNewParagraph);
								}
							else  // there are no cascading tags, just write the text if there are any
								{
								if(objHTMLelement.innerText.Length > 0)
									{
									objRun = oxmlDocument.Construct_RunText(parText2Write: objHTMLelement.innerText);
									// Check if a hyperlink must be inserted
									if(this.HyperlinkImageRelationshipID != "")
										{
										if(this.HyperlinkInserted == false)
											{
											DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
												parMainDocumentPart: ref parMainDocumentPart,
												parImageRelationshipId: this.HyperlinkImageRelationshipID,
												parClickLinkURL: this.HyperlinkURL);
											objRun.Append(objDrawing);
											this.HyperlinkInserted = true;
											}
										}
									objNewParagraph.Append(objRun);
									}
								objTableCell.Append(objNewParagraph);
								}
							Console.WriteLine("\tLastChild in Table: {0}", this.WPdocTable.LastChild);
							this.WPdocTable.LastChild.Append(objTableCell);
							break;

						//------------------------------------
						case "UL":     // Unorganised List (Bullets to follow) Tag
							Console.WriteLine("Tag: UNORGANISED LIST\n\r{0}", objHTMLelement.outerHTML);
							if(objHTMLelement.children.length > 0)
								{
								ProcessHTMLelements(
									ref parMainDocumentPart,
									objHTMLelement.children,
									ref objNewParagraph,
									false);
								}
							else
								{
								objRun = oxmlDocument.Construct_RunText(parText2Write: objHTMLelement.innerText);
								// Check if a hyperlink must be inserted
								if(this.HyperlinkImageRelationshipID != "")
									{
									if(this.HyperlinkInserted == false)
										{
										DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref parMainDocumentPart,
											parImageRelationshipId: this.HyperlinkImageRelationshipID,
											parClickLinkURL: this.HyperlinkURL);
										objRun.Append(objDrawing);
										this.HyperlinkInserted = true;
										}
									}
								}
							break;
						//------------------------------------
						case "OL":     // Orginised List (numbered list) Tag
							Console.WriteLine("Tag: ORGANISED LIST\n\r{0}", objHTMLelement.outerHTML);
							if(objHTMLelement.children.length > 0)
								{
								ProcessHTMLelements(
									ref parMainDocumentPart,
									objHTMLelement.children,
									ref objNewParagraph,
									false);
								}
							else
								{
								objRun = oxmlDocument.Construct_RunText(parText2Write: objHTMLelement.innerText);
								// Check if a hyperlink must be inserted
								if(this.HyperlinkImageRelationshipID != "")
									{
									if(this.HyperlinkInserted == false)
										{
										DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref parMainDocumentPart,
											parImageRelationshipId: this.HyperlinkImageRelationshipID,
											parClickLinkURL: this.HyperlinkURL);
										objRun.Append(objDrawing);
										this.HyperlinkInserted = true;
										}
									}
								}
							break;
						//------------------------------------
						case "LI":     // List Item (an entry from a organised or unorginaised list
							Console.WriteLine("Tag: LIST ITEM\n\r{0}", objHTMLelement.outerHTML);
							// Construct the paragraph with the bullet or number...
							if (objHTMLelement.parentElement.tagName == "OL") // number list
								objNewParagraph = oxmlDocument.Construct_BulletNumberParagraph(parIsBullet: false,parBulletLevel: this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel);
							else // "UL" == Unorganised/Bullet list item
								objNewParagraph = oxmlDocument.Construct_BulletNumberParagraph(parIsBullet: true, parBulletLevel: this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel);

							if(objHTMLelement.children.length > 0) // check if there are more html tags in the HTMLelement
								{
								Console.WriteLine("\t{0} child nodes to process", objHTMLelement.children.length);
								// use the DissectHTMLstring method to process the paragraph.
								List<TextSegment> listTextSegments = new List<TextSegment>();
								listTextSegments = TextSegment.DissectHTMLstring(objHTMLelement.innerHTML);
								foreach(TextSegment objTextSegment in listTextSegments)
									{
									objRun = oxmlDocument.Construct_RunText
										(parText2Write: objTextSegment.Text,
										parBold: objTextSegment.Bold,
										parItalic: objTextSegment.Italic,
										parUnderline: objTextSegment.Undeline,
										parSubscript: objTextSegment.Subscript,
										parSuperscript: objTextSegment.Superscript);
									// Check if a hyperlink must be inserted
									if(this.HyperlinkImageRelationshipID != "")
										{
										if(this.HyperlinkInserted == false)
											{
											DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
												parMainDocumentPart: ref parMainDocumentPart,
												parImageRelationshipId: this.HyperlinkImageRelationshipID,
												parClickLinkURL: this.HyperlinkURL);
											objRun.Append(objDrawing);
											this.HyperlinkInserted = true;
											}
										}
									objNewParagraph.Append(objRun);
									}
								}
							else  // there are no cascading tags, just write the text if there are any
								{
								if(objHTMLelement.innerText.Length > 0)
									{
									objRun = oxmlDocument.Construct_RunText(parText2Write: objHTMLelement.innerText);
									// Check if a hyperlink must be inserted
									if(this.HyperlinkImageRelationshipID != "")
										{
										if(this.HyperlinkInserted == false)
											{
											DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
												parMainDocumentPart: ref parMainDocumentPart,
												parImageRelationshipId: this.HyperlinkImageRelationshipID,
												parClickLinkURL: this.HyperlinkURL);
											objRun.Append(objDrawing);
											this.HyperlinkInserted = true;
											}
										}
									objNewParagraph.Append(objRun);
									}
								}
							if(parAppendToExistingParagraph)
								//ignore because only a new Paragraph needs to be appended to the body
								Console.WriteLine("Skip the appending of the existing paragraph to the Body");
							else
								{
								this.WPbody.Append(objNewParagraph);
								}

							break;
						//------------------------------------
						case "IMG":    // Image Tag
							Console.WriteLine("Tag:IMAGE \n\r{0}", objHTMLelement.outerHTML);
							// Increment the image counter
							ImageCaptionCounter += 1;
							// Check if the image has a Caption that needs to be inserted.
							string imageCaption = "...";
							if(objHTMLelement.getAttribute("alt", 0) != "")
								imageCaption = objHTMLelement.getAttribute("alt", 0);
							
                                   objNewParagraph = oxmlDocument.Construct_Caption(
								parCaptionType: "Image",
								parCaptionSequence: this.ImageCaptionCounter,
								parCaptionText: ": " + imageCaption);

							Console.WriteLine("{0}", objHTMLelement.getAttribute("src", 0));
							Console.WriteLine("{0}", objHTMLelement.getAttribute("src", 4));
							string fileURL = objHTMLelement.getAttribute("src",1);
							if(fileURL.StartsWith("about"))
								fileURL = fileURL.Substring(6,fileURL.Length - 6);

							Console.WriteLine("\t Image URL: {0}", fileURL);
							objRun = oxmlDocument.InsertImage(
								parMainDocumentPart: ref parMainDocumentPart,
								parParagraphLevel: this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel,
								parPictureSeqNo: this.ImageCaptionCounter,
								parImageURL: Properties.AppResources.SharePointURL + fileURL,
								parEffectivePageTWIPSheight: this.PageHeight,
								parEffectivePageTWIPSwidth: this.PageWidth);
							if(objRun != null)
								objNewParagraph.Append(objRun);
							else
								objRun = oxmlDocument.Construct_RunText("ERROR: Unable to insert the image - an error occurred");

							//this.WPbody.AppendChild<Paragraph>(objNewParagraph);
							this.WPbody.Append(objNewParagraph);
							break;
						case "STRONG": // Bold Tag
							Console.WriteLine("TAG: BOLD\n\r{0}", objHTMLelement.outerHTML);

							break;
						//------------------------------------
						case "SPAN":   // Underline is embedded in the Span tag
							Console.WriteLine("Tag: Span\n\r{0}", objHTMLelement.outerHTML);
							break;
						//------------------------------------
						case "EM":     // Italic Tag
							Console.WriteLine("Tag: ITALIC\n\r{0}", objHTMLelement.outerHTML);
//							this.ItalicsOn = true;
//							if(objHTMLelement.children.length > 0)
//								{
//								// use the DissectHTMLstring method to process the paragraph.
//								List<TextSegment> listTextSegments = new List<TextSegment>();
//								listTextSegments = TextSegment.DissectHTMLstring(objHTMLelement.innerHTML);
//								foreach(TextSegment objTextSegment in listTextSegments)
//									{
//									objRun = oxmlDocument.Construct_RunText
//											(parText2Write: objTextSegment.Text,
//											parBold: objTextSegment.Bold,
//											parItalic: objTextSegment.Italic,
//											parUnderline: objTextSegment.Undeline,
//											parSubscript: objTextSegment.Subscript,
//											parSuperscript: objTextSegment.Superscript);
//									objNewParagraph.Append(objRun);
//									}
//}
//							else  // there are no cascading tags, just append the text to an existing paragrapg object
//								{
//								if(objHTMLelement.innerText.Length > 0)
//									{
//									objRun = oxmlDocument.Construct_RunText
//										(parText2Write: objHTMLelement.innerText,
//										parBold: this.BoldOn,
//										parItalic: this.ItalicsOn,
//										parUnderline: this.UnderlineOn);
//									objNewParagraph.Append(objRun);
//									}
//								}
//							this.ItalicsOn = false;
							break;
						//------------------------------------
						case "SUB":    // Subscript Tag
							Console.WriteLine("Tag: SUPERSCRIPT\n\r{0}", objHTMLelement.outerHTML);
							break;
						//------------------------------------
						case "SUP":    // Super Script Tag
							Console.WriteLine("Tag: SUPERSCRIPT\n\r{0}", objHTMLelement.outerHTML);
							break;
						//------------------------------------
						case "H1":     // Heading 1
						case "H1A":    // Alternate Heading 1
							Console.WriteLine("Tag: H1\n\r{0}", objHTMLelement.outerHTML);
							this.AdditionalHierarchicalLevel = 1;
							objNewParagraph = oxmlDocument.Insert_Heading(
								parHeadingLevel: this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel);

							objRun = oxmlDocument.Construct_RunText(parText2Write: objHTMLelement.innerText);
							// Check if a hyperlink must be inserted
							if(this.HyperlinkImageRelationshipID != "")
								{
								if(this.HyperlinkInserted == false)
									{
									DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
										parMainDocumentPart: ref parMainDocumentPart,
										parImageRelationshipId: this.HyperlinkImageRelationshipID,
										parClickLinkURL: this.HyperlinkURL);
									objRun.Append(objDrawing);
									this.HyperlinkInserted = true;
									}
								}
							objNewParagraph.Append(objRun);
							this.WPbody.Append(objNewParagraph);
							break;
						//------------------------------------
						case "H2":     // Heading 2
						case "H2A":    // Alternate Heading 2
							Console.WriteLine("Tag: H2\n\r{0}", objHTMLelement.outerHTML);
							this.AdditionalHierarchicalLevel = 2;
							objNewParagraph = oxmlDocument.Insert_Heading(
								parHeadingLevel: this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel);

							objRun = oxmlDocument.Construct_RunText(parText2Write: objHTMLelement.innerText);
							// Check if a hyperlink must be inserted
							if(this.HyperlinkImageRelationshipID != "")
								{
								if(this.HyperlinkInserted == false)
									{
									DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
										parMainDocumentPart: ref parMainDocumentPart,
										parImageRelationshipId: this.HyperlinkImageRelationshipID,
										parClickLinkURL: this.HyperlinkURL);
									objRun.Append(objDrawing);
									this.HyperlinkInserted = true;
									}
								}
							objNewParagraph.Append(objRun);
							this.WPbody.Append(objNewParagraph);
							break;
						//------------------------------------
						case "H3":     // Heading 3
						case "H3A":    // Alternate Heading 3
							Console.WriteLine("Tag: H3\n\r{0}", objHTMLelement.outerHTML);
							this.AdditionalHierarchicalLevel = 3;
							objNewParagraph = oxmlDocument.Insert_Heading(
								parHeadingLevel: this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel);

							objRun = oxmlDocument.Construct_RunText(parText2Write: objHTMLelement.innerText);
							// Check if a hyperlink must be inserted
							if(this.HyperlinkImageRelationshipID != "")
								{
								if(this.HyperlinkInserted == false)
									{
									DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
										parMainDocumentPart: ref parMainDocumentPart,
										parImageRelationshipId: this.HyperlinkImageRelationshipID,
										parClickLinkURL: this.HyperlinkURL);
									objRun.Append(objDrawing);
									this.HyperlinkInserted = true;
									}
								}
							objNewParagraph.Append(objRun);
							this.WPbody.Append(objNewParagraph);
							break;
						//------------------------------------
						case "H4":     // Heading 4
						case "H4A":    // Alternate Heading 4
							Console.WriteLine("Tag: H4\n\r{0}", objHTMLelement.outerHTML);
							this.AdditionalHierarchicalLevel = 4;
							objNewParagraph = oxmlDocument.Insert_Heading(
								parHeadingLevel: this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel);

							objRun = oxmlDocument.Construct_RunText(parText2Write: objHTMLelement.innerText);
							// Check if a hyperlink must be inserted
							if(this.HyperlinkImageRelationshipID != "")
								{
								if(this.HyperlinkInserted == false)
									{
									DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
										parMainDocumentPart: ref parMainDocumentPart,
										parImageRelationshipId: this.HyperlinkImageRelationshipID,
										parClickLinkURL: this.HyperlinkURL);
									objRun.Append(objDrawing);
									this.HyperlinkInserted = true;
									}
								}
							objNewParagraph.Append(objRun);
							this.WPbody.Append(objNewParagraph);
							break;
						default:
							Console.WriteLine("**** ignoring tag: {0}", objHTMLelement.tagName);
							break;

						} // switch(objHTMLelement.tagName)


					} // foreach(IHTMLElement objHTMLelement in parHTMLElements)


				} // if (parHTMLElements.length > 0)


			} // end of Method

		public void DetermineTableGrid(IHTMLElementCollection parHTMLelements)
			{
			// First clear the TableColumn widths.
			if(this.TableColumnWidths.Count > 0)
				this.TableColumnWidths.Clear();
			string sWidth = "";
			UInt32 iWidth = 0;
			this.TableHasFirstRow = false;
			this.TableHasLastRow = false;
			this.TableHasFirstColumn = false;
			this.TableHasLastColumn = false;
						
			/// gaan hier aan....
               foreach(IHTMLElement tableColumnItem in parHTMLelements)
				{
				Console.WriteLine("\t\t\t {0} - {1}", tableColumnItem.tagName, tableColumnItem.outerHTML);

				// determine the width of each column

				sWidth = tableColumnItem.style.width;
				if(sWidth.IndexOf("%", 0) > 0)
					{
					this.TableColumnUnit = "%";
					Console.WriteLine("\t\t\t The % is in position {0}", sWidth.IndexOf("%", 0));
					Console.WriteLine("\t\t\t Numeric Value: {0}", sWidth.Substring(0, (sWidth.Length - sWidth.IndexOf("%", 0)) + 1));
					if(!UInt32.TryParse(sWidth.Substring(0, (sWidth.Length - sWidth.IndexOf("%", 1)) + 1), out iWidth))
						iWidth = 100;
					}
				else if(sWidth.IndexOf("px", 0) > 0)
					{
					this.TableColumnUnit = "px";
					Console.WriteLine("\t\t\t The px is in position {0}", sWidth.IndexOf("px", 0));
					Console.WriteLine("\t\t\t Numeric Value: {0}", sWidth.Substring(0, (sWidth.Length - sWidth.IndexOf("px", 0)) + 1));
					if(!UInt32.TryParse(sWidth.Substring(0, (sWidth.Length - sWidth.IndexOf("px", 1)) + 1), out iWidth))
						iWidth = 100;
					}
				else
					{
					iWidth = 100;
					}

				this.TableColumnWidths.Add(iWidth);
				}

			} // end of Method, DetermineTableGrid

		}    // end of Class

	/// <summary>
	/// TextSegment Class
	/// </summary>
	class TextSegment
		{
		private string _text;
		public string Text
			{
			get{return this._text;}
			set{this._text = value;}
			}

		private bool _bold;
		public bool Bold
			{
			get {return this._bold;}
			set{this._bold = value;}
			}

		private bool _italic;
		public bool Italic
			{
			get{return this._italic;}
			set{this._italic = value;}
			}

		private bool _undeline;
		public bool Undeline
			{
			get{return this._undeline;}
			set{this._undeline = value;}
			}

		private bool _subscript;
		public bool Subscript
			{
			get{return this._subscript;}
			set{this._subscript = value;}
			}

		private bool _superscript;
		public bool Superscript
			{
			get{return this._superscript;}
			set{this._superscript = value;}
			}
		private bool _image;
		public bool Image
			{
			get{return this._image;}
			set{this._image = value;}
			}

		public static List<TextSegment> DissectHTMLstring(string parTextString)
			{
			int i = 0;
			int iPointer = 0;
			int iOpenTagStart = 0;
			int iOpenTagEnds = 0;
			string sOpenTag = "";
			int iCloseTagStart = 0;
			int iCloseTagEnds = 0;
			string sCloseTag = "";
			bool bBold = false;
			bool bItalic = false;
			bool bUnderline = false;
			bool bSuperScript = false;
			bool bSubscript = false;
			int iNextTagStart = 0;
			int iNextTagEnds = 0;
			string sNextTag = "";
			
			List<TextSegment> listTextSegments = new List<TextSegment>();
			//-----------------------------------------------------------
			// replace and/or remove special strings before processing the Text Segment... 
			parTextString = parTextString.Replace(oldValue: "&quot;", newValue: Convert.ToString(value: (char) 22));
			parTextString = parTextString.Replace(oldValue: "&nbsp;", newValue: "");
			parTextString = parTextString.Replace(oldValue: "&#160;", newValue: "");
			parTextString = parTextString.Replace(oldValue: "  ", newValue: " ");
			Console.WriteLine("\t\t\tString to examine:\r\t\t\t|{0}|", parTextString);

			do
				{
				iNextTagStart = parTextString.IndexOf("<", iPointer);
				if(iNextTagStart < 0) // Check if there are any tags left to process
					break;
				iNextTagEnds = parTextString.IndexOf(">", iPointer);
				sNextTag = parTextString.Substring(iNextTagStart, (iNextTagEnds - iNextTagStart) + 1);
				if(sNextTag.IndexOf("<IMG") >= 0)
					{
					// Extract the Image tah and place it in the text string
					TextSegment objTextSegment = new TextSegment();
					objTextSegment.Bold = false;
					objTextSegment.Italic = false;
					objTextSegment.Undeline = false;
					objTextSegment.Subscript = false;
					objTextSegment.Superscript = false;
					objTextSegment.Image = true;
					objTextSegment.Text = sNextTag;
					listTextSegments.Add(objTextSegment);
					Console.WriteLine("\t\t\t-- IMG: {0}", objTextSegment.Text);
					iPointer = iPointer + sNextTag.Length;
					}
				else
					{
					if(sNextTag.IndexOf("/") < 0) // it is an Open tag
						{
						// Check if there are any text BEFORE the tag
						if(iNextTagStart > iPointer)
							{
							//extract the text before the first tag and place it in the List of TextSegments
							TextSegment objTextSegment = new TextSegment();
							objTextSegment.Text = parTextString.Substring(iPointer, (iNextTagStart - iPointer));
							objTextSegment.Bold = bBold;
							objTextSegment.Italic = bItalic;
							objTextSegment.Undeline = bUnderline;
							objTextSegment.Subscript = bSubscript;
							objTextSegment.Superscript = bSuperScript;
							objTextSegment.Image = false;
							listTextSegments.Add(objTextSegment);
							Console.WriteLine("\t\t\t** {0}", objTextSegment.Text);
							iPointer = iNextTagStart;
							}
						// Determine the START
						iOpenTagStart = iNextTagStart;
						iOpenTagEnds = iNextTagEnds;
						sOpenTag = sNextTag;
						Console.WriteLine("\t\t\t\t- OpenTag: {0} = {1} - {2}", sOpenTag, iOpenTagStart, iOpenTagEnds);
						// Define the corresponding closing tag
						if(sOpenTag.IndexOf("STRONG") > 0)
							{
							sCloseTag = "</STRONG>";
							bBold = true;
							}
						else if(sOpenTag.IndexOf("EM>") > 0)
							{
							sCloseTag = "</EM>";
							bItalic = true;
							}
						else if(sOpenTag.IndexOf("underline") > 0)
							{
							sCloseTag = "</SPAN>";
							bUnderline = true;
							}
						else if(sOpenTag.IndexOf("SUB") > 0)
							{
							sCloseTag = "</SUB>";
							bSubscript = true;
							}
						else if(sOpenTag.IndexOf("SUP") > 0)
							{
							sCloseTag = "</SUP>";
							bSuperScript = true;
							}
						else if(sOpenTag.IndexOf("SPAN") >= 0)
							sCloseTag = "</SPAN>";
						else
							sCloseTag = "";

						iCloseTagStart = parTextString.IndexOf(value: sCloseTag, startIndex: iOpenTagStart + sOpenTag.Length);
						if(iCloseTagStart < 0)
							// the close tag was not found?
							Console.WriteLine("ERROR: {0} - not found!", sCloseTag);
						else
							{
							iCloseTagEnds = iCloseTagStart + sCloseTag.Length - 1;
							Console.WriteLine("\t\t\t\t- CloseTag: {0} = {1} - {2}", sCloseTag, iCloseTagStart, iCloseTagEnds);
							//iPointer = iOpenTagEnds + 1;
							}
						iPointer = iOpenTagEnds + 1;
						}
					else  // it is a CLOSE tag
						{
						// Check if there are any text BEFORE the tag
						if(iNextTagStart > iPointer)
							{
							//extract the text before the first tag and place it in the List of TextSegments
							TextSegment objTextSegment = new TextSegment();
							objTextSegment.Text = parTextString.Substring(iPointer, (iNextTagStart - iPointer));
							objTextSegment.Bold = bBold;
							objTextSegment.Italic = bItalic;
							objTextSegment.Undeline = bUnderline;
							objTextSegment.Subscript = bSubscript;
							objTextSegment.Superscript = bSuperScript;
							objTextSegment.Image = false;
							listTextSegments.Add(objTextSegment);
							Console.WriteLine("\t\t\t** {0}", objTextSegment.Text);
							}
						// Obtain the Close Tag
						iCloseTagStart = iNextTagStart;
						iCloseTagEnds = iNextTagEnds;
						sCloseTag = sNextTag;
						Console.WriteLine("\t\t\t\t- CloseTag: {0} = {1} - {2}", sCloseTag, iCloseTagStart, iCloseTagEnds);
						// Depending on the closing tag set the text format off
						if(sCloseTag.IndexOf("/STRONG") > 0)
							bBold = false;
						if(sCloseTag.IndexOf("/EM") > 0)
							bItalic = false;
						if(sCloseTag.IndexOf("/SPAN") > 0)
							bUnderline = false;
						if(sCloseTag.IndexOf("/SUB") > 0)
							bSubscript = false;
						if(sCloseTag.IndexOf("/SUP") > 0)
							bSuperScript = false;
						iPointer = iNextTagEnds + 1;
						} // if it is a Close Tag
					}
				} while(iPointer < parTextString.Length);

			//checked if there are trailing characters that need to be processed.
			if(iPointer < parTextString.Length)
				{
				//extract the text pointer until the end of the string place it in the List of TextSegments
				TextSegment objTextSegment = new TextSegment();
				objTextSegment.Text = parTextString.Substring(iPointer, (parTextString.Length - iPointer));
				objTextSegment.Bold = bBold;
				objTextSegment.Italic = bItalic;
				objTextSegment.Undeline = bUnderline;
				objTextSegment.Subscript = bSubscript;
				objTextSegment.Superscript = bSuperScript;
				listTextSegments.Add(objTextSegment);
				iPointer = parTextString.Length;
				Console.WriteLine("\t\t\t** {0}", objTextSegment.Text);
				}

			i = 0;
			foreach(TextSegment objTextSegmentItem in listTextSegments)
				{
				i += 1;
				Console.WriteLine("\t\t+ {0}: {1} (Bold:{2} Italic:{3} Underline:{4} Subscript:{5} Superscript:{6} Image:{7})",
					i, objTextSegmentItem.Text, objTextSegmentItem.Bold, objTextSegmentItem.Italic, objTextSegmentItem.Undeline, objTextSegmentItem.Subscript,
					objTextSegmentItem.Subscript, objTextSegmentItem.Image);
				}

			return listTextSegments;

			} // end method

		} // end class
	}
