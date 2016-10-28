using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace DocGeneratorCore
	{

	public enum enumCaptionType
		{
		Image,
		Table
		}

	public enum enumAlignmentHorizontal
		{
		Left,
		Centre,
		Right
		}

	public enum enumAlignmentVertical
		{
		Top,
		Centre,
		Bottom
		}

	public enum enumWidthHeightType
		{
		Pixels,
		Percent
		}

	public class WorkCell
		{
		public enumAlignmentHorizontal AlignHorizontal { get; set; } = enumAlignmentHorizontal.Left;

		public enumAlignmentVertical AlignmVertical { get; set; } = enumAlignmentVertical.Centre;

		public bool Bold { get; set; } = false;

		public bool EvenColumn { get; set; } = false;

		public bool FirstColumn { get; set; } = false;

		public string HTMLcontent { get; set; } = String.Empty;

		public bool Italic { get; set; } = false;

		public bool LastColumn { get; set; } = false;

		public bool LineThrough { get; set; } = false;

		public int Number { get; set; } = 0;

		public bool OddColumn { get; set; } = false;

		public int MergeColumns { get; set; } = 0;

		public enumTableRowMergeType MergeRow { get; set; } = enumTableRowMergeType.None;
		
		public bool Subscript { get; set; } = false;

		public bool Superscript { get; set; } = false;

		public bool Underline { get; set; } = false;

		public int WidthPercentage { get; set; } = 0;

		public int WidthInDxa { get; set; } = 0;

		}

	//++
	public class WorkRow
		{
		public List<WorkCell> Cells { get; set; } = new List<WorkCell>();
		public bool EvenRow { get; set; } = false;
		public bool FirstRoW { get; set; } = false;
		public bool LastRow { get; set; } = false;
		public int Number { get; set; } = 0;
		public bool OddRow { get; set; } = false;
		}

	//++ These classes are used to construct tables
	public class WorkTable
		{
		/// <summary>
		///      This flag MUST be set to TRUE as soon as a WorkTable object is instanciated.
		/// </summary>
		public bool Active { get; set; } = false;

		/// <summary>
		///      May contain a string which can be used as the Table's Caption.
		/// </summary>
		public String Caption { get; set; } = String.Empty;

		/// <summary>
		///      Set to TRUE (the default is FALSE) if the table has a FIRST Column defined.
		/// </summary>
		public bool FirstColumn { get; set; } = false;

		/// <summary>
		///      Set to True (the default is FALSE) if the table has a FIRST or HEADER Row defined.
		/// </summary>
		public bool FirstRow { get; set; } = false;

		/// <summary>
		///      This list of integers defined the table's column grids from the left to the right
		///      column... The Tuple consist of two intances:
		///      1. Column Width as a n integer value
		///      2. Span indicator which is **True** if the width was derived from a merged/spanned
		///         cell and False if it was from an unmerged cell
		/// </summary>
		public List<int> GridColumnWidthPercentages { get; set; }

		/// <summary>
		///      Set to TRUE (the default is FALSE) as soon as the Table Grid was determined and defined.
		/// </summary>
		public bool GridDone { get; set; } = false;

		/// <summary>
		///      Set to TRUE (the default is FALSE) if the table has a LAST or TOTAL column defined.
		/// </summary>
		public bool LastColumn { get; set; } = false;

		/// <summary>
		///      Set to TRUE (the default is FALSE) if the table has a Last or FOOTER Row defined.
		/// </summary>
		public bool LastRow { get; set; } = false;

		/// <summary>
		///      Needs to be set to indicate the WidthType of the the original HTML table. Will be
		///      used during the calculation of the Table's width scaling...
		/// </summary>
		public enumWidthHeightType OriginalTableWidthType { get; set; } = enumWidthHeightType.Percent;

		/// <summary>
		///      The original value of the table width as it was defined in the HTML source...
		/// </summary>
		public int OriginalTableWidthValue { get; set; } = 0;

		/// <summary>
		///      This list is a collection of WorkRow objects which represents each row of the table
		///      from top to bottom of the table...
		/// </summary>
		public List<WorkRow> Rows { get; set; } = new List<WorkRow>();

		public int WidthInDXA { get; set; } = 0;

		/// <summary>
		///      The percentage of the Document page that the table will occupy...
		/// </summary>
		public int WidthPrecentage { get; set; } = 0;


		}

	/// <summary>
	///      HTML Decoder class is used to instansiate a HTMLdecoder object. The object is used to
	///      decode HTML structure and translate it into Open XML document or Workbook content. Set
	///      the properties begore
	/// </summary>
	public class HTMLdecoder
		{
		/// <summary>
		///      The Additional Hierarchical Level property contains the number of additional levels
		///      that need to be added to the Document Hierarchical Level when processing the HTML
		///      contained in a Enhanced Rich Text column/field.
		/// </summary>
		public int AdditionalHierarchicalLevel { get; set; }

		//++ Object Properties
		public bool Bold { get; set; } = false;

		/// <summary>
		///      This property will contain the text of the Caption Text to be added after an image
		///      or table
		/// </summary>
		public string CaptionText { get; set; } = string.Empty;

		/// <summary>
		///      This property indicates the type of caption that need to be inserted after an image
		///      or table. It used by the DECODEhtml and ProcessHTMLelements methods to handle
		///      Captions when encoding of OpenXML documents.
		/// </summary>
		public enumCaptionType CaptionType { get; set; } = enumCaptionType.Table;

		public string ClientName { get; set; } = String.Empty;

		/// <summary>
		///      This property is used to control Content Layer colour coding.
		/// </summary>
		public string ContentLayer { get; set; } = "None";

		/// <summary>
		///      The Document Hierarchical Level provides the stating Hierarchical level at which new
		///      content will be added to the document.
		/// </summary>
		public int DocumentHierachyLevel { get; set; }

		public bool Italics { get; set; } = false;

		/// The PageHeight property contains the page Height of the OXML page into which the decoded
		/// HTML content will be inserted. It is mostly used for image and table positioning on the
		/// page in the OXML document. </summary>
		public uint PageHeightDxa { get; set; }

		/// <summary>
		///      The PageWidth property contains the page width of the OXML page into which the
		///      decoded HTML content will be inserted. It is mostly used for image and table
		///      positioning on the page in the OXML document.
		/// </summary>
		public uint PageWidthDxa { get; set; }

		public bool StrikeTrough { get; set; } = false;

		public bool Subscript { get; set; } = false;

		public bool SuperScript { get; set; } = false;

		public WorkTable TableToBuild { get; set; }

		public bool Underline { get; set; } = false;

		public string WorkString { get; set; } = String.Empty;

		/// <summary>
		///      Set the WordProcessing Body immediately after declaring an instance of the
		///      HTMLdecoder object The oXMLencoder requires the WPBody object by reference to add
		///      the decoded HTML to the oXML document.
		/// </summary>
		public Body WPbody { get; set; }

		/// <summary>
		///      Contains the Level of the bullet list
		/// </summary>
		public int BulletLevel { get; set; } = 0;

		/// <summary>
		///      The HyperlinkRelationshipID is used by the DECODEhtml method to handle Hyperlinks in
		///      the encoding of OpenXML documents.
		/// </summary>
		private string HyperlinkImageRelationshipID { get; set; } = string.Empty;

		/// <summary>
		///      Indicator property that are set once a Hyperlink was inserted for an HTML run
		/// </summary>
		public bool HyperlinkInserted { get; set; } = false;

		/// <summary>
		///      The HyperlinkURL contains the ACTUAL hyperlink URL that will be inserted if
		///      required. It used by the DECODEhtml and ProcessHTMLelements methods to to handle
		///      Hyperlinks in the encoding of OpenXML documents.
		/// </summary>
		public string HyperlinkURL { get; set; } = string.Empty;

		/// <summary>
		///      Contains the level of the numbered list
		/// </summary>
		public bool RestartNumbering { get; set; } = false;

		public string SharePointSiteURL { get; set; } = string.Empty;

	
		//++ Object Methods
		public Table BuildTable(
			ref MainDocumentPart parMainDocumentPart,
			WorkTable parWorkTable,
			string parContentLayer,
			ref int parNumberingCounter)
			{
			int cellWidth = 0;
			decimal cellWidthPixels = 0m;
			int columnCounter = 0;
			int rowCounter = 0;
			TableRow objTableRow = new TableRow();
			TableCell objTableCell = new TableCell();
			Paragraph objParagraph = new Paragraph();
			List<Paragraph> paragraphs = new List<Paragraph>();
			Run objRun = new Run();
			
			//+Create the OpenXML Table object instance...
			DocumentFormat.OpenXml.Wordprocessing.Table objTable = new DocumentFormat.OpenXml.Wordprocessing.Table();

			Console.Write("\t\t Initialise the Table Object... ");
			//Construct the table in the
			objTable = oxmlDocument.Construct_Table(
				parTableWidthInDXA: parWorkTable.WidthInDXA,
				parFirstRow: parWorkTable.FirstRow,
				parFirstColumn: parWorkTable.FirstColumn,
				parLastColumn: parWorkTable.LastColumn,
				parLastRow: parWorkTable.LastRow,
				parNoVerticalBand: true,
				parNoHorizontalBand: true);
			Console.Write("\t Done!");

			//+Construct the Table Grid
			Console.Write("\n\t\t Create the Table Grid Object... ");
			TableGrid objTableGrid = new TableGrid();
			objTableGrid = oxmlDocument.ConstructTableGrid(
				parColumnWidthList: parWorkTable.GridColumnWidthPercentages, 
				parTableWidthPixels: parWorkTable.WidthInDXA);
			objTable.Append(objTableGrid);
			Console.Write("\t Done!");

			//+ Process all the table Rows...
			foreach (WorkRow rowItem in parWorkTable.Rows)
				{
				rowCounter += 1;
				Console.Write("\n\t\t\t => Row {0}", rowCounter);
				//-Get the Table's properties to set the table's LookProperties...
				TableProperties objTableProperties = objTable.GetFirstChild<TableProperties>();
				//-Get the **TableLook** from in the TableProperties...
				TableLook objTableLook = objTableProperties.GetFirstChild<TableLook>();

				if (rowItem.FirstRoW)
					{
					//-Update the **FirsRow** value...
					objTableLook.FirstRow = true;
					}

				if (rowItem.LastRow)
					{
					//-Update the **LastRow** value...
					objTableLook.LastRow = true;
					}

				//- Create a **TableRow** object instance... row to the table
				objTableRow = new TableRow();
				objTableRow = oxmlDocument.ConstructTableRow(
					parIsFirstRow: rowItem.FirstRoW,
					parIsLastRow: rowItem.LastRow);

				columnCounter = 0;

				//+Process all the Cells for each Row
				foreach (WorkCell cellItem in rowItem.Cells.OrderBy(c => c.Number))
					{
					Console.Write("\n\t\t\t\t + Column {0}", columnCounter + 1);

					//-Update the **TableLook** with First- and LastColumns id requried...
					if (cellItem.FirstColumn)
						objTableLook.FirstColumn = true;

					if (cellItem.LastColumn)
						objTableLook.LastColumn = true;

					//+Determine the width of the cell
					Console.Write("\t MergeColumns:{0}", cellItem.MergeColumns);
					Console.Write("\t MergeRow:{0}", cellItem.MergeRow);
					//-Check if the cell spans multiple columns
					if (cellItem.MergeColumns > 1)
						{ //-the cell spans multiple columns...
						//-Add the width of the spanned columns
						cellWidth = 0;
						for (int i = columnCounter; i <= columnCounter + cellItem.MergeColumns - 1; i++)
							{
							cellWidth += parWorkTable.GridColumnWidthPercentages[i];
							}
						}
					else
						{ //- it is a single column cell
						cellWidth = parWorkTable.GridColumnWidthPercentages[columnCounter];
						}
					Console.Write("\t Width:{0}%", cellWidth);
					//-The values in the **parWorkTable.GridColumnWidths** are percentages (%)
					//-Therefore it must be converted to pixels
					cellWidthPixels = (parWorkTable.WidthInDXA * cellWidth) / 100m;
					Console.Write(" - {0}px", cellWidthPixels);

					//- Add the **TableCell** to the row...
					objTableCell = new TableCell();
					objTableCell = oxmlDocument.ConstructTableCell(
						parCellWidth: Convert.ToInt16(cellWidthPixels),
						parHasCondtionalFormatting: false,
						parIsFirstRow: rowItem.FirstRoW,
						parIsLastRow: rowItem.LastRow,
						parIsFirstColumn: cellItem.FirstColumn,
						parIsLastColumn: cellItem.LastColumn,
						parRowMerge: cellItem.MergeRow,
						parColumnMerge: cellItem.MergeColumns,
						parHorizontalAlignment: cellItem.AlignHorizontal.ToString(),
						parVerticalAlignment: cellItem.AlignmVertical.ToString());

					Console.Write("\t Cell Created...");
					//- Check if there are any content to insert into the cell
					if (string.IsNullOrWhiteSpace(cellItem.HTMLcontent))
						{
						//- MS Word requires a paragraph to exist for every table cell. 
						//- Therefore we just add a blank paragraph to ensure that the MS Word document is not invalidated.
						objParagraph = new Paragraph();

						objParagraph = oxmlDocument.Construct_Paragraph(
							parBodyTextLevel: 0,
							parIsTableParagraph: true,
							parIsTableHeader: rowItem.FirstRoW,
							parFirstRow: rowItem.FirstRoW,
							parLastRow: rowItem.LastRow,
							parFirstColumn: cellItem.FirstColumn,
							parLastColumn: cellItem.LastColumn);

						Run workRun = oxmlDocument.Construct_RunText(
							parText2Write: " ",
							parContentLayer: parContentLayer);

						objParagraph.Append(workRun);
						objTableCell.Append(objParagraph);

						Console.Write("\t Empty Paragraph inserted...");
						}
					else
						{ //-there is content in the cell, process it
						paragraphs = DissectHTMLstring(
							parMainDocumentPart: ref parMainDocumentPart,
							parHTMLString: cellItem.HTMLcontent, 
							parNumberingCounter: ref parNumberingCounter, 
							parIsTable: true,
							parIsTableHeader: rowItem.FirstRoW,
							parAdditionalHierarchyLevel: this.AdditionalHierarchicalLevel,
							parClientName: this.ClientName,
							parContentLayer: this.ContentLayer,
							parHierarchyLevel: this.DocumentHierachyLevel,
							parRestartNumbering: this.RestartNumbering
							);

						//-|Table Cells must contain a *Paragraph* therefore a blank paragraph need to be added even if the HTML cell is empty
						if (paragraphs.Count == 0)
							{
							objParagraph = oxmlDocument.Construct_Paragraph(
								parBodyTextLevel: 1,
								parFirstRow: rowItem.FirstRoW,
								parFirstColumn: cellItem.FirstColumn,
								parIsTableHeader: rowItem.FirstRoW,
								parLastColumn: cellItem.LastColumn,
								parIsTableParagraph: true);
							objTableCell.Append(objParagraph);
							}
						else
							{ //-|Insert all the paragraphs...
							foreach (Paragraph paragraphItem in paragraphs)
								{
								objTableCell.Append(paragraphItem);
								}
							}
						Console.Write("\t Paragraphs inserted.");
						}
					objTableRow.Append(objTableCell);
					Console.Write("\t Appended Cell...");
					columnCounter += 1;
					}
				//-once all the **TableColumns are added, append the the TableRow to the Table object.
				objTable.Append(objTableRow);
				}
			Console.WriteLine("\n\t{0} Table Construction completed {0}", new string('=', 30));
			return objTable;
			}

		//===g
		/// <summary>
		///      Use this method once a new HTMLdecoder object is initialised and provide at least
		///      all the compulasory(non-optional) properties to the method.
		/// </summary>
		/// <param name="parMainDocumentPart">
		///      [Compulsory] provide the REFERENCE to an already declared MainDocumentPart object of
		///      the document into which the content HTML content needs to be inserted.
		/// </param>
		/// <param name="parDocumentLevel">
		///      [Compulsory] An interger value of 0 to 9 indicating the Heading level or BodyText
		///      Level at which insertion must be inserted.
		/// </param>
		/// <param name="parPageWidthTwips">
		///      [Compulsory] provide the page width into which the content must be inserted. This
		///      needs to be the actual width less margins, gutters and indentations, which and
		///      wherever applicable.
		/// </param>
		/// <param name="parPageHeightTwips">
		///      [Compulsory] provide the page height into which the content must be inserted. This
		///      needs to be the actual height less margins and header/footer offsets, which and
		///      wherever applicable.
		/// </param>
		/// <param name="parHTML2Decode">
		///      [Compulsory] assign the actual HTML string that need to be decoded and inserted to
		///      this parameter.
		/// </param>
		/// <param name="parTableCaptionCounter">
		///      [Compulsory] provide the REFERENCE to an integer containing the current
		///      TableCaptionCounter that need to be used if the HTML contains a table.
		/// </param>
		/// <param name="parImageCaptionCounter">
		///      [Compulsory] provide the REFERENCE to an integer containing the current
		///      ImageCaptionCounter that need to be used if the HTML contains a table.
		/// </param>
		/// <param name="parPictureNo">
		///      [Compulsory] provide the REFERENCE to an integer containing the current Picture
		///      Number/Counter that need to be used if the HTML contains a imagetable.
		/// </param>
		/// <param name="parHyperlinkID">
		///      [Compulsory] hyperlinks needs to be unique in the OpenXML document, therefore a
		///      unique value must ALWAYS be inserted into the OpenXML document else it becomes invalid.
		/// </param>
		/// <param name="parHyperlinkURL">
		///      [Compulsory] but can be null if no hyperlink needs to be inserted referencing the
		///      complete HTML section.
		/// </param>
		/// <param name="parHyperlinkImageRelationshipID">
		///      [Compulsory] but can be null, if the parHyperlinkURL is null.this is the
		///      Relationship id, provided used in the OPENXml document where the image to which the
		///      hyperlink in parHyperlink and parHyperlinkID needs to be linked.
		/// </param>
		/// <param name="parContentLayer">
		///      [optional] defaults to "None" else it must be one of the following string values:
		///      "Layer1", "Layer2" or "Layer3"
		/// </param>
		/// <returns>
		///      returns a boolean value of TRUE if insert was successfull and FALSE if there was any
		///      form of failure during the insertion.
		/// </returns>
		public void DecodeHTML(
			ref MainDocumentPart parMainDocumentPart,
			int parDocumentLevel,
			string parHTML2Decode,
			ref int parTableCaptionCounter,
			ref int parImageCaptionCounter,
			ref int parPictureNo,
			ref int parHyperlinkID,
			ref int parNumberingCounter,
			string parSharePointSiteURL,
			string parClientName,
			uint parPageHeightDxa,
			uint parPageWidthDxa,
			string parHyperlinkURL = "",
			string parHyperlinkImageRelationshipID = "",
			string parContentLayer = "None")
			
			{
			//+Update the Properties of the object...
			this.SharePointSiteURL = parSharePointSiteURL;
			this.DocumentHierachyLevel = parDocumentLevel;
			this.AdditionalHierarchicalLevel = 0;
			this.HyperlinkImageRelationshipID = parHyperlinkImageRelationshipID;
			this.HyperlinkURL = parHyperlinkURL;
			this.ContentLayer = parContentLayer;
			this.PageHeightDxa = Convert.ToUInt16(parPageHeightDxa);
			this.PageWidthDxa = Convert.ToUInt16(parPageWidthDxa);
			this.HyperlinkInserted = false;
			//+Variables:
			//- The sequence of the values in the tuple variable is:
			//-**1** *Bullet Levels*, 
			//-**2** *Number Levels*
			Tuple<int, int> bulletNumberLevels = new Tuple<int, int>(0, 0);
			int bulletLevel = 0;
			int numberLevel = 0;
			int headingLevel = 0;
			int cellWidthPixels = 0;
			int cellWidthPercentage = 0;
			int columnCounter = 1;			//-| variable which keeps track of the number of columns in a table row
			int rowCounter = 1;             //-| variable which keeps track of the number of rows in a table
			int rowSpan = 1;				//-| variable used to termporary keep row span values when processing tables
			int spanValue = 1;
			bool isDIVempty = true;
			int imageWidth = 0;
			string hyperlinkURL = string.Empty;
			enumWidthHeightType imageWidthType = enumWidthHeightType.Percent;
			int imageHeight = 0;
			enumWidthHeightType imageHeightType = enumWidthHeightType.Percent;
			
			Paragraph objNewParagraph = null;
			Run objRun = new Run();
			//-Working Table variables...
			WorkRow workRow = new WorkRow();
			WorkCell workCell = new WorkCell();
			//-|Initialise the TableToBuild...
			this.TableToBuild = null;

			try
				{
				//-Declare the htmlData object instance
				HtmlAgilityPack.HtmlDocument htmlData = new HtmlAgilityPack.HtmlDocument();
				//-Load the HTML content into the htmlData object.
				//-Load a string into the HTMLDocument

				htmlData.LoadHtml(parHTML2Decode);
				//Console.WriteLine(new string('_', 120));
				//Console.WriteLine("\nHTML: {0}", htmlData.DocumentNode.OuterHtml);
				//Console.WriteLine(new string('_', 120));
				//- Set the ROOT of the data loaded in the htmlData
				var htmlRoot = htmlData.DocumentNode;
				Console.WriteLine("\nRoot Node Tag..............: {0}", htmlRoot.Name);
				Console.WriteLine("______________________________HTML decoding iterations begin ______________________________________");
				foreach (HtmlNode node in htmlData.DocumentNode.DescendantsAndSelf())
					{
					//Console.WriteLine(">-- {0} --<", node.Name);

					Application.DoEvents();

					//! Check whether a table needs to be written to the document
					if (this.TableToBuild != null)
						{
						if (this.TableToBuild.Active
						&& !node.XPath.Contains("table"))
							{
							Console.WriteLine("\n\t{0} Table Mapped {0}", new string('-', 60));
							Table objTable = new Table();
							Console.WriteLine("\n\t{0} Constructing TABLE for insertion into Document {0}", new string('-', 25));
							objTable = BuildTable(
								parMainDocumentPart: ref parMainDocumentPart,
								parWorkTable: this.TableToBuild, 
								parNumberingCounter: ref parNumberingCounter, 
								parContentLayer: parContentLayer);
							Console.WriteLine("\n\t {0} Insert constructed TABLE into the document {0}", new string('-', 25));
							this.WPbody.Append(objTable);
							//-If the table has a caption, insert it...
							if (!string.IsNullOrWhiteSpace(this.TableToBuild.Caption))
								{
								parTableCaptionCounter += 1;
								Console.WriteLine("\t Table Caption... |{0} {1}: {2}|", "Table", parTableCaptionCounter, this.TableToBuild.Caption);
								objNewParagraph = oxmlDocument.Construct_Caption(
									parCaptionType: "Table",
									parCaptionText: "Table" + parTableCaptionCounter + ": " + this.TableToBuild.Caption);
								this.WPbody.Append(objNewParagraph);
								}
							Console.WriteLine("\n\t {0} Inserted TABLE and Caption into the document {0}", new string('=', 25));
							this.TableToBuild = null;
							objNewParagraph = null;
							objRun = null;
							}
						}

					//===g
					switch (node.Name)
						{
						//---g
						//+<DIV>
						case "div":
						//-Check for other tags in the **div** tag
						isDIVempty = true;
						if (node.HasChildNodes)
							{
							//-|<div> is embedded in a **table** and will be processed by another method at a later stage.
							//-Therefore we skip over the ,div> when we process tables.
							if (node.XPath.IndexOf(value: "table", startIndex: 0, comparisonType: StringComparison.OrdinalIgnoreCase) > 0)
								break;

							//-Check if there are any other valid HTML tags succeeding the **<DIV>** tag
							foreach (HtmlNode decendentNode in node.Descendants())
								{
								//-Check for any valid content that need to be processed...
								if (Regex.IsMatch(decendentNode.Name, "p")			//-|paragraph
								|| Regex.IsMatch(decendentNode.Name, "ol")			//-|Organised List
								|| Regex.IsMatch(decendentNode.Name, "ul")			//-|Unorganised List
								|| Regex.IsMatch(decendentNode.Name, "li")			//-|ListItem
								|| Regex.IsMatch(decendentNode.Name, "img")			//-|Image
								|| Regex.IsMatch(decendentNode.Name, "table")		//-|Table
								|| Regex.IsMatch(decendentNode.Name, "h1-4")		//-|Heading 1-4
								)
									{
									//- Just CONSUME the div tag and process the other tags in the subsequent node cycles...
									isDIVempty = false;
									break;
									}
								}
							}

						//-|Reaching this point, means that the <div> tag is *NOT* part of any valid HTML tags and is an **isolated** <div> tag
						if (isDIVempty)
							{
							//-It actually means the text cannot be properly formatted because it doesn't contain valid HTML tags which indicates how to format the text
							//-|although a **Content Error** should be produced, we temporary deactivate a *blatant* content error and try to auto-interpret the content.
							
							//-First check if there is not a populated paragraph that need to be inserted into the document before the content error is raised
							if (objNewParagraph != null && !string.IsNullOrEmpty(objNewParagraph.InnerText))
							{ //-Write the *paragraph* to the document...
							this.WPbody.Append(objNewParagraph);
							}
							//-|We just create a new document paragraph
							objNewParagraph = oxmlDocument.Construct_Paragraph();
							//-|The text will processed and added to the paragraph in the next tag processing loop
				
							/*
							//-HTML tags - **THEREFORE** a format issue is raised...
							throw new InvalidContentFormatException("The content was probably pasted from an external source without "
								+ "formatting it after pasting. This error usually occur when a HTML <DIV> tag contains text but no paragraph a <p> </p> HTML tag. "
								+ "Please inspect the content and resolve the issue, by formatting the relevant content using the relevant "
								+ "Enhanced Rich Text styles. |" + node.InnerText + "|");
							*/
							}
						break;

						//---g
						//+<#text> **TEXT**
						case "#text":
						//- Text in tables are embedded in a *workCell* and is processed later by another method.
						//-Therefore we skip over paragraphs when we process **tables** or when a **HyperLink** was processed.
						if (node.XPath.IndexOf(value: "table", startIndex: 0, comparisonType: StringComparison.OrdinalIgnoreCase) > 0
						|| node.XPath.IndexOf(value: @"/a", startIndex: 0 ,comparisonType: StringComparison.Ordinal) > 0)
							break;
						//-First, clean the string by removing unwanted and unnecessary characters
						this.WorkString = CleanText(node.InnerText, this.ClientName);
						//-Check if the Workstring is blank **OR** if the text is part of a **table**, then **DON'T** process it
						if (string.IsNullOrWhiteSpace(this.WorkString))
							break;

						if (node.XPath.Contains("/strong"))
							{
							Console.Write("|BOLD|");
							this.Bold = true;
							}
						else
							this.Bold = false;

						if (node.XPath.Contains("/em"))
							{
							Console.Write("|ITALICS|");
							this.Italics = true;
							}
						else
							this.Italics = false;

						if (node.XPath.Contains("/span"))
							if (node.ParentNode.HasAttributes)
								{
								foreach (HtmlAttribute attributeItem in node.ParentNode.Attributes)
									{
									if (attributeItem.Value.Contains("underline"))
										{
										this.Underline = true;
										Console.Write("|UNDERLINE|");
										}
									else
										this.Underline = false;
									}
								}
							else
								this.Underline = false;
						else
							this.Underline = false;

						if (node.XPath.Contains("/sub"))
							{
							Console.Write("|SUBSCRIPT|");
							this.Subscript = true;
							}
						else
							this.Subscript = false;

						if (node.XPath.Contains("/sup"))
							{
							Console.Write("|SUPERSCRIPT|");
							this.SuperScript = true;
							}
						else
							this.SuperScript = false;

						//-Insert the **text** if the string is *not* empty
						if (!String.IsNullOrWhiteSpace(this.WorkString))
							{
							Console.Write("[{0}]", this.WorkString);
							objRun = oxmlDocument.Construct_RunText(
								parText2Write: this.WorkString,
								parContentLayer: this.ContentLayer,
								parBold: this.Bold,
								parItalic: this.Italics,
								parUnderline: this.Underline,
								parSubscript: this.Subscript,
								parSuperscript: this.SuperScript);

							// Check if a hyperlink must be inserted
							if (string.IsNullOrWhiteSpace(this.HyperlinkImageRelationshipID)
							&& string.IsNullOrWhiteSpace(this.HyperlinkURL))
								{ //-|ignore the hyperlink 
								}
							else
								{
								if (this.HyperlinkInserted == false)
									{
									parHyperlinkID += 1;
									DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing =
										oxmlDocument.Construct_ClickLinkHyperlink(
										parMainDocumentPart: ref parMainDocumentPart,
										parImageRelationshipId: this.HyperlinkImageRelationshipID,
										parClickLinkURL: this.HyperlinkURL,
										parHyperlinkID: parHyperlinkID);
									objRun.Append(objDrawing);
									this.HyperlinkInserted = true;
									}
								}
							objNewParagraph.Append(objRun);
							}
						break;
						//---g
						//+<a> - **Hyperlink** tag
						case "a":
						hyperlinkURL = string.Empty;
						//-Process the **Hyperlink** attributes
						foreach (HtmlAttribute hyperlinkAttr in node.Attributes)
							{ //-|use the **href** attribute to obtain and url
							if(hyperlinkAttr.Name == "href")
								{
								hyperlinkURL = hyperlinkAttr.Value;
								Console.Write("<Hyperlink> {0} - Text{1}", hyperlinkURL, node.InnerText);
								break;
								}
							}

						if (hyperlinkURL.StartsWith("about"))
							hyperlinkURL = hyperlinkURL.Substring(6, hyperlinkURL.Length - 6);

						if (hyperlinkURL.Substring(0, 7) == "/sites/")
							hyperlinkURL = Properties.AppResources.SharePointSiteURL + hyperlinkURL;
						
						//-|Check if an initialised **Paragraph** exist
						if (objNewParagraph == null)
							{
							objNewParagraph = oxmlDocument.Construct_Paragraph();
							}

						try
							{
							objNewParagraph.Append(oxmlDocument.Construct_Hyperlink(
								parMainDocumentPart: ref parMainDocumentPart,
								parText2insert: node.InnerText,
								parURL: hyperlinkURL));
							}
						catch(InvalidContentFormatException)
							{
							throw new InvalidContentFormatException("The content was probably pasted from an external source without "
								+ "formatting it after pasting. This error usually occur when a HTML <DIV> tag contains text but no paragraph a <p> </p> HTML tag. "
								+ "Please inspect the content and resolve the issue, by formatting the relevant content using the relevant "
								+ "Enhanced Rich Text styles. |" + node.InnerText + "|");
							}
						
						break;
						//---g
						//+<p> - **Paragraph**
						case "p":
						//- **Normal** Paragraphs and **Table** paragraphs are handled a little different because,
						//- table paragrpahs are embedded in a *workCell* and is processed later by another method.
						//-Therefore we skip over paragraphs when we process tables.
						if (node.XPath.IndexOf(value: "table", startIndex: 0, comparisonType: StringComparison.OrdinalIgnoreCase) > 0)
							break;

						//-|Check if the **objNewParagraph** is *NOT* null **AND** that it actually contains text and commit it to the document before 
						//-|before initiating a new 
						if (objNewParagraph != null && !String.IsNullOrEmpty(objNewParagraph.InnerText))
							{ //-Write the *paragraph* to the document...
							this.WPbody.Append(objNewParagraph);
							objNewParagraph = null;
							}

						//-Check if the paragraph is part of a bullet- number- list
						if (node.XPath.Contains("/ol") || node.XPath.Contains("/ul") && node.XPath.Contains("/li"))
							{ //-If it is :. get the number of bullet - and number - levels in the xPath
							bulletNumberLevels = GetBulletNumberLevels(node.XPath);
							bulletLevel = bulletNumberLevels.Item1;
							numberLevel = bulletNumberLevels.Item2;
							//- now exit the loop, bto process the **"#text"** or other child tags...
							break;
							}
						else
							{
							bulletLevel = 0;
							numberLevel = 0;
							}

						//-Check if a **NEW** paragraph must be initialised or whether the existing paragraph needs to be used to add run text.
						if (objNewParagraph == null)
							objNewParagraph = oxmlDocument.Construct_Paragraph();

						if (node.HasChildNodes)
							{
							Console.Write("\n\t <{0}> ", node.Name);
							}
						else
							{//?Does this means it is an empty paragraph?
							this.WorkString = node.InnerText;
							if (this.WorkString != string.Empty)
								Console.Write(" <{0}>", node.Name);
							}
						break;

						//---g
						//+<h1-4> **Heading 1-4**
						case "h1":
						case "h2":
						case "h3":
						case "h4":
						//-Get the *Heading level* and set the **headingLevel** value
						if (!int.TryParse(node.Name.Substring(1, (node.Name.Length - 1)), out headingLevel))
							{ headingLevel = 0; }
						Console.Write("\n {0} + <{1}>", new String('\t', headingLevel * 2), node.Name);
						//- Set the **this.AdditionalHierarchicalLevel** to the headingLevel value
						this.AdditionalHierarchicalLevel = headingLevel;

						//-Check if there is a populated paragraph
						if (objNewParagraph != null
						&& objNewParagraph.InnerText != null)
							{ //-Write the *paragraph* to the document...
							this.WPbody.Append(objNewParagraph);
							}
						//-Create the paragrpah for the **Heading**
						objNewParagraph = oxmlDocument.Construct_Heading(
							parHeadingLevel: this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel);

						//-if there are no child nodes, check if the innterText is also blank
						if (!node.HasChildNodes)
							{
							if (node.InnerText == String.Empty)
								{//-Destroy the paragraph because it will be a blank heading in the document...
								objNewParagraph = null;
								}
							}
						break;

						//---g
						//+ <UL> **Unorganised List**
						case "ul":
						//-Check if there is a populated paragraph
						if (objNewParagraph != null
						&& objNewParagraph.InnerText != null)
							{ //-Write the *paragraph* to the document...
							this.WPbody.Append(objNewParagraph);
							objNewParagraph = null;
							}
						else
							objNewParagraph = null;

						if (node.HasChildNodes)
							{
							//Console.Write("\n {0} <{1}>", Tabs(headingLevel) + Tabs(bulletLevel), node.Name);
							}
						break;

						//---g
						//+<ol> **Organised List**
						case "ol":
						//-Check if there is a populated paragraph that **CONTAINS** text...
						if (objNewParagraph != null
						&& !String.IsNullOrWhiteSpace(objNewParagraph.InnerText))
							{ //-Write the *paragraph* to the document...
							this.WPbody.Append(objNewParagraph);
							objNewParagraph = null;
							}
						else
							objNewParagraph = null;

						if (node.HasChildNodes)
							{
							//Console.Write("\n {0} <{1}>", Tabs(headingLevel) + Tabs(bulletLevel), node.Name);
							//-Determine the number of bullet- and number- levels from the xPath
							bulletNumberLevels = GetBulletNumberLevels(node.XPath);
							numberLevel = bulletNumberLevels.Item2;
							if(numberLevel == 1)
								{
								this.RestartNumbering = true;
								parNumberingCounter += 1;
								}
							}

						break;

						//---g
						//+<li>  **List Item**
						case "li":

						if (node.XPath.IndexOf(value: "table", startIndex: 0, comparisonType: StringComparison.OrdinalIgnoreCase) > 0)
							break;

						//-Check if there is a populated paragraph that contains text and write it to the Document before initiating a new paragraph...
						if (objNewParagraph != null
						&& !String.IsNullOrEmpty(objNewParagraph.InnerText))
							{ //-Write the *paragraph* to the document...
							this.WPbody.Append(objNewParagraph);
							objNewParagraph = null;
							}

						//-Determine the number of bullet- and number- levels from the xPath
						bulletNumberLevels = GetBulletNumberLevels(node.XPath);
						bulletLevel = bulletNumberLevels.Item1;
						numberLevel = bulletNumberLevels.Item2;
						
						//-Construct the paragraph with the bullet or number :. depends on the value of the bulletLevel...
						if (bulletLevel > 0)
							{//- if it is a **Bullet** list entry, create a new **Pargraph** *Bullet* object...
							objNewParagraph = oxmlDocument.Construct_BulletParagraph(
								parIsTableBullet: false,
								parBulletLevel: bulletLevel);
							}
						//- check if it is **Organised/Number list** item
						else if (numberLevel > 0)
							{ //-|if it is a **Number** list entry, create a new **Pargraph** *Number* object instance...
							if (this.RestartNumbering)
								parNumberingCounter += 1;
							objNewParagraph = oxmlDocument.Construct_NumberParagraph(
								parMainDocumentPart: ref parMainDocumentPart,
								parIsTableNumber: false,
								parRestartNumbering: this.RestartNumbering,
								parNumberingId: parNumberingCounter,
								parNumberLevel: numberLevel);
							this.RestartNumbering = false;
							}

						if (node.HasChildNodes)
							{
							if (bulletLevel > 0)
								{
								Console.Write("\n {0} - <{1}>", new String('\t', headingLevel + bulletLevel), node.Name);
								}
							else if (numberLevel > 0)
								{
								Console.Write("\n {0} {1}. <{2}>", new String('\t', (headingLevel + numberLevel)), numberLevel, node.Name);
								}
							}
						break;

						//---g
						//++Image
						case "img":
							{

							if (node.XPath.IndexOf(value: "table", startIndex: 0, comparisonType: StringComparison.OrdinalIgnoreCase) > 0)
								break;

							Console.WriteLine("\n\t\t\t {0} Process Image {0}", new string('-', 25));
							//-Check if the **objNewParagraph** is *NOT * null * *AND * *that it actually contains text
							if (objNewParagraph != null && !String.IsNullOrEmpty(objNewParagraph.InnerText))
									{ //-Write the *paragraph* to the document...
									this.WPbody.Append(objNewParagraph);
									objNewParagraph = null;
									}

							this.CaptionText = string.Empty;

							string imageFileURL = string.Empty;
							//-Process the table attributes to determine how to format the image...
							foreach (HtmlAttribute imageAttr in node.Attributes)
								{
								switch (imageAttr.Name)
									{
									//-use the **alt** attribute to obtain and set the **Image Caption**
									case "alt":
									this.CaptionType = enumCaptionType.Image;
									if (imageAttr.Value == null)
										this.CaptionText = string.Empty;
									else
										{
										this.CaptionText = imageAttr.Value;
										parImageCaptionCounter += 1;
										}

									Console.WriteLine("\t\t\t\t Caption: {0}", this.CaptionText);
									break;
									case "src":
									imageFileURL = imageAttr.Value;
									if (imageFileURL.StartsWith("about"))
										imageFileURL = imageFileURL.Substring(6, imageFileURL.Length - 6);
									Console.WriteLine("\t\t\t\t <img> URL: {0}", imageFileURL);
									break;
									}
								}
							
							try
								{
								//-|create a new paragraph into which the image will be inserted
								objNewParagraph = oxmlDocument.Construct_Paragraph();
								//-|load the image and insert it into a run object instance
								objRun = oxmlDocument.Insert_Image(
									parMainDocumentPart: ref parMainDocumentPart,
									parParagraphLevel: this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel,
									parPictureSeqNo: parImageCaptionCounter,
									parImageURL: this.SharePointSiteURL + imageFileURL,
									parEffectivePageHeightDxa: this.PageHeightDxa,
									parEffectivePageWidthDxa: this.PageWidthDxa,
									parImageHeight: imageHeight,
									parImageHeightType: imageHeightType,
									parImageWidth: imageWidth,
									parImageWidthType: imageWidthType);
								//-|Append the image in the run, to the pargraph
								objNewParagraph.Append(objRun);
								//-|Insert the pargraph into the document
								this.WPbody.Append(objNewParagraph);

								//-|Insert the image caption if required.
								if (!string.IsNullOrWhiteSpace(this.CaptionText))
									{
									//-|Create the Caption Paragraph...
									objNewParagraph = oxmlDocument.Construct_Caption(
										parCaptionType: "Image",
										parCaptionText: Properties.AppResources.Document_Caption_Image_Text + parImageCaptionCounter + ": " + this.CaptionText);
									//-|Insert the paragrpah into the document...
									this.WPbody.Append(objNewParagraph);
									}
								}
							catch (InvalidImageFormatException exc)
								{
								throw new InvalidImageFormatException(exc.Message);
								}
							objNewParagraph = null;
							Console.WriteLine("\t\t\t {0} Image Processed {0}", new string('-', 25));
							}

						break;

						//---g
						//++Table
						case "table":

						//-Check if the **objNewParagraph** is *NOT* null **AND** that it actually contains text
						if (objNewParagraph != null && !String.IsNullOrEmpty(objNewParagraph.InnerText))
							{ //-Write the *paragraph* to the document...
							this.WPbody.Append(objNewParagraph);
							objNewParagraph = null;
							}

						//-Define the Table Instance
						if (this.TableToBuild == null)
							{
							this.TableToBuild = new WorkTable();
							this.TableToBuild.Active = true;      //-Set Table Mode to Active
							this.TableToBuild.GridDone = false;   //-The table's **GRID** is not yet determined...
							}
						else
							{
							if (this.TableToBuild.Active)
								{
								Console.WriteLine("\n *********ERROR********** - Cascading table, No attributes defined for the table");
								throw new InvalidContentFormatException("The TABLE that is suppose to appear here, contains a cascading table "
									+ " (a table within a table) The DocGenerator is not designed to produce cascading tables. "
									+ "Please inspect the content and ensure there are not table embedded into another table.");
								}
							}
						Console.WriteLine("\n\t{0} Begin to build a table {0}", new string('=', 30));
						Console.Write("\n\n\t <Table> ");
						Tuple<int, enumWidthHeightType> tableWidthValue = new Tuple<int, enumWidthHeightType>(0, enumWidthHeightType.Percent);
						//-Check if the table has **attributes** define to obtain the table's **width**...
						if (!node.HasAttributes)
							{//- The table **doesn't have** any attributes
							Console.WriteLine("\n ERROR - No attributes defined for the table");
							throw new InvalidContentFormatException("The TABLE's width is missing, therefore the table cannot be inserted into the "
								+ "document. Please inspect the content and resolve the issue, by formatting the relevant content with the "
								+ "Enhanced Rich Text styles.");
							}
						else
							{
							//-Process the table attributes to determine how to format the table...
							foreach (HtmlAttribute tableAttr in node.Attributes)
								{
								switch (tableAttr.Name)
									{
									//-use the **summary** to obtain and set the **Table Caption**
									case "summary": //- get the table caption
									if (tableAttr.Value == null)
										this.TableToBuild.Caption = string.Empty;
									else
										this.TableToBuild.Caption = tableAttr.Value;
									break;

									//-get the table **width** from the style attribute
									case "style":
									//- Check that the style contains the table width as part of the style
									if (tableAttr.Value.Contains("width"))
										{
										tableWidthValue = GetHTMLwidthAttribute(tableAttr.Value);
										this.TableToBuild.OriginalTableWidthValue = tableWidthValue.Item1;
										this.TableToBuild.OriginalTableWidthType = tableWidthValue.Item2;
										}
									break;

									//-The table Width is specified as an attribute
									case "width":
									tableWidthValue = GetHTMLwidthAttribute(tableAttr.Value);
									this.TableToBuild.OriginalTableWidthValue = tableWidthValue.Item1;
									this.TableToBuild.OriginalTableWidthType = tableWidthValue.Item2;
									break; //break out of the SWITCH
									}
								}
							}

						if (this.TableToBuild.OriginalTableWidthValue <= 0)
							{
							throw new InvalidTableFormatException("The TABLE's width is NOT specified as a percentage value, therefore the table cannot be inserted into the "
								+ "document. Please inspect the content and resolve the issue, by formatting the table with the "
								+ "Enhanced Rich Text styles and ensure that the table width and al cell widths are specified as a percentage.");
							}

						if (this.TableToBuild.OriginalTableWidthType == enumWidthHeightType.Percent)
							{
							this.TableToBuild.WidthPrecentage = this.TableToBuild.OriginalTableWidthValue;
							this.TableToBuild.WidthInDXA = Convert.ToInt16((this.PageWidthDxa * this.TableToBuild.OriginalTableWidthValue) / 100);
							}
						else //-OriginalTableWidthType is **Pixels**
							{
							throw new InvalidTableFormatException("The TABLE's width is specified in pixels while it should be specified as a percentage. "
								+ "Therefore the table cannot be scalled and inserted into the document. "
								+ "Please inspect the at this position and resolve the issue, by specifying table WIDTH as a %.");
							}

						//-Check if the table's width is defined, if not raise an exception and exit
						if (this.TableToBuild.WidthInDXA == 0 || this.TableToBuild.WidthPrecentage == 0)
							{
							Console.WriteLine("\n ERROR - Could Not determine the table's width.");
							throw new InvalidContentFormatException("The TABLE's width could not be determined, therefore the table cannot be "
								+ "inserted into the document. Please inspect the content and resolve the issue, by formatting the relevant "
								+ "content with the Enhanced Rich Text styles.");
							}

						Console.Write("\t ...Original Table Width: {0} {1}\t translated to: {2}%\t {3}dxa ", this.TableToBuild.OriginalTableWidthValue, this.TableToBuild.OriginalTableWidthType,
						this.TableToBuild.WidthPrecentage, this.TableToBuild.WidthInDXA);

						//-Determine the Table Grid...
						Console.WriteLine("\n{0} Begin to Define Table Grid {0}", new string('-', 50));
						DetermineTableGrid(parHTMLnodes: node.DescendantsAndSelf());
						rowCounter = 0;
						Console.WriteLine("{0} Begin to Map the TABLE {0}", new string('-', 50));
						break;

						//---g
						//+Table Body = **<tb>**
						case "tbody":
						//-Just ignore it.
						break;

						//---g
						//+Table Row = **<tr>**
						case "tr":
						//-Before processing the *Table Row* check if the row is already defined in the table.
						//-The row may have been defined when a cell spanned multiple rows
						rowCounter += 1;
						//workRow = new WorkRow();
						workRow = this.TableToBuild.Rows.SingleOrDefault(r => r.Number == rowCounter);
						if (workRow == null)
							{
							workRow = new WorkRow();
							workRow.Number = rowCounter;
							this.TableToBuild.Rows.Add(workRow);
							workRow = this.TableToBuild.Rows.SingleOrDefault(r => r.Number == rowCounter);
							}

						Console.Write("\n\t\t <Row> No:{0}", workRow.Number);

						//+Set the Row's properties...
						if (node.HasAttributes)
							{ //- The *row* has attributes defined
							  //-Determine if the row is a **HEADER** row...
							foreach (HtmlAgilityPack.HtmlAttribute tableAttr in node.Attributes)
								{
								//-Check the **class** attribute to dertermine whether the row is specified as a *TableHeader*
								if (tableAttr.Name.Contains("class"))
									{
									if (tableAttr.Value.Contains("HeaderRow"))
										workRow.FirstRoW = true;

									//-Check if the *class* attribute identifies the Row as an **OddRow**
									else if (tableAttr.Value.Contains("OddRow"))
										workRow.OddRow = true;

									//-Check if the *class* attribute identifies the Row as an **EvenRow**
									else if (tableAttr.Value.Contains("EvenRow"))
										workRow.EvenRow = true;

									//-Check if the *class* attribute identifies the Row as a **EvenRow**
									else if (tableAttr.Value.Contains("FooterRow"))
										workRow.LastRow = true;
									}
								}
							}
						Console.Write("\tHeader:{0}", workRow.FirstRoW);
						Console.Write("\t Footer:{0}", workRow.LastRow);
						Console.Write("\t Odd:{0}", workRow.OddRow);
						Console.Write("\t Even:{0}", workRow.EvenRow);
						//-The **columnNo** variable is used to keep track of the number of cells being processed in a table row.
						rowSpan = 1;
						columnCounter = 0;
						break;

						//---g
						//+<th> <td> - **Header Cell** and **Normal Cell**
						case "th":
						case "td":
						//-| There are 2 type of tags covered in this section of code.
						//-|     1. Table **header/footer** cells which are identified by **<th>** tags
						//-|     2. ***Normal** calls which are identified by **<td>** tags
						//-| This section of code applies to both types of tags, therefore it is covered in the same code block
Repeat_Cell:
						//-Increment the **columnNo * *counter which keeps track of the number of columns in a row...
						columnCounter += 1;
						rowSpan = 1;
						Tuple<int, enumWidthHeightType> cellWidth = new Tuple<int, enumWidthHeightType>(0, enumWidthHeightType.Percent);

						Console.Write("\n\t\t\t Cell:<{0}> \tNo:{1}", node.Name, columnCounter);
						//-Check if the cell already exist in the *workRow*
						//workCell = new WorkCell();
						workCell = workRow.Cells.SingleOrDefault(c => c.Number == columnCounter);
						if (workCell == null)
							{
							//-Initialiase the instance of the workCell object...
							workCell = new WorkCell();
							workCell.Number = columnCounter;
							workCell.MergeRow = enumTableRowMergeType.None;
							workRow.Cells.Add(workCell);
							}
						else if (workCell.MergeRow == enumTableRowMergeType.Continue)
							{
							Console.Write("\tFirst:{0}", workCell.FirstColumn);
							Console.Write("\tOdd:{0}", workCell.OddColumn);
							Console.Write("\tEven:{0}", workCell.EvenColumn);
							Console.Write("\tLast:{0}", workCell.LastColumn);
							Console.Write("\tColSpan:{0}", workCell.MergeColumns);
							Console.Write("\tWidth:{0}%", workCell.WidthPercentage);
							Console.Write("\tWidth:{0}px", workCell.WidthInDxa);
							Console.Write("\tRowSpan:{0}", workCell.MergeRow);
							goto Repeat_Cell;
							}

						workCell = workRow.Cells.SingleOrDefault(c => c.Number == columnCounter);
						
						//+Set the **Cell's** properties...
						foreach (HtmlAttribute nodeAttribute in node.Attributes)
							{
							Application.DoEvents();

							switch (nodeAttribute.Name)
								{
								//+ *class* attribute value to determine the type of cell...
								case "class":
								//-Check if the *class* attribute identifies the Cell as a **First Column Cell**
								if (nodeAttribute.Value.Contains("FirstCol"))
									workCell.FirstColumn = true;

								//-Check if the *class* attribute identifies the Cell as a **Last Column Cell**
								else if (nodeAttribute.Value.Contains("LastCol"))
									workCell.LastColumn = true;

								//-Check if the *class* attribute identifies the Cell as an **Odd Column Cell**
								else if (nodeAttribute.Value.Contains("OddCol"))
									workCell.OddColumn = true;

								//-Check if the *class* attribute identifies the Cell as an **Even Column Cell**
								else if (nodeAttribute.Value.Contains("EvenCol"))
									workCell.EvenColumn = true;
								break;

								//+Determine the **Row Span**
								case "rowspan":
								if (int.TryParse(nodeAttribute.Value, out spanValue))
									rowSpan = spanValue;
								else
									rowSpan = 1;
								break;

								//+Determine the **Column Spans**
								case "colspan":
								//-Parse the **colspan** value, if the parse fail set it to 1
								if (int.TryParse(nodeAttribute.Value, out spanValue))
									workCell.MergeColumns = spanValue;
								else
									workCell.MergeColumns = 1;
								break;

								//+Determine the **Width** in the Attribite
								case "width":
								if (!string.IsNullOrWhiteSpace(nodeAttribute.Value))
									{ //-The *cell* has a defined **width** value - it is **NOT** *Null* or *space*...
									cellWidth = GetHTMLwidthAttribute(nodeAttribute.Value);
									if (cellWidth.Item2 == enumWidthHeightType.Pixels)
										cellWidthPixels = cellWidth.Item1;
									else if (cellWidth.Item2 == enumWidthHeightType.Percent)
										cellWidthPercentage = cellWidth.Item1;
									}
								break;

								//+Determine the cell **Width** in style
								case "style":
								//-First check if the Cell's **width** is specified in a *style*...
								if (!string.IsNullOrWhiteSpace(nodeAttribute.Value))
									{
									//-check if the value is the width...
									if (nodeAttribute.Value.Contains("width"))
										{
										cellWidth = GetHTMLwidthAttribute(nodeAttribute.Value);
										if (cellWidth.Item2 == enumWidthHeightType.Pixels)
											cellWidthPixels = cellWidth.Item1;
										else if (cellWidth.Item2 == enumWidthHeightType.Percent)
											cellWidthPercentage = cellWidth.Item1;
										}
									}
								break;
								}
							}

						//-Check that the column has a **width** defined, if it doesn't have a width defined,
						//-Get a width from the Table's Grid Definition, else use the defined width.
						if (cellWidthPixels == 0)
							{
							//-Derive the column width from the Grid
							workCell.WidthPercentage = this.TableToBuild.GridColumnWidthPercentages[columnCounter - 1];
							//-Check if the cell spans multiple columns and add those column's values to the initial columns width value.
							if (workCell.MergeColumns > 1)
								{
								for (int i = 2; i < workCell.MergeColumns + 1; i++)
									{
									workCell.WidthPercentage = this.TableToBuild.GridColumnWidthPercentages[i - 1];
									}
								}
							}
						else
							workCell.WidthPercentage = cellWidthPixels;

						//-Set the columnSpan of the cell to 1 if it is 0.
						if (workCell.MergeColumns == 0)
							workCell.MergeColumns = 1;

						//-Set the column number for the sell
						workCell.Number = columnCounter;
						//-Calculate the Column Width pixels...
						workCell.WidthInDxa = (this.TableToBuild.WidthInDXA * workCell.WidthPercentage) / 100;
						//-Embed the text in the HTMLcontent property
						workCell.HTMLcontent = node.InnerHtml;
						Console.Write("\tFirst:{0}", workCell.FirstColumn);
						Console.Write("\tOdd:{0}", workCell.OddColumn);
						Console.Write("\tEven:{0}", workCell.EvenColumn);
						Console.Write("\tLast:{0}", workCell.LastColumn);
						Console.Write("\tColSpan:{0}", workCell.MergeColumns);
						Console.Write("\tWidth:{0}%", workCell.WidthPercentage);
						Console.Write("\tWidth:{0}dxa", workCell.WidthInDxa);

						//+Populate the rows and cells for each spanned row
						//-Check if the row span greater than 1, meaning the call contains a **VerticalMerge**
						if (rowSpan > 1)
							{
							workCell.MergeRow = enumTableRowMergeType.Restart;
							Console.Write("\tRowSpan:{0}", workCell.MergeRow);
							//-The creation of the rows and the cells need to occur the *number* of times that the row span/merge value
							for (int rowIncrement = 2; rowIncrement <= rowSpan; rowIncrement++) //-|start at 2 because the first occurrence is already added.
								{
								//-Check if the row already exist
								WorkRow newWorkRow = new WorkRow();
								if (this.TableToBuild.Rows.Where(r => r.Number == (rowCounter + rowIncrement - 1)).FirstOrDefault() == null)
									{ //-The row doesn't exist yet...
									newWorkRow.Number = rowCounter + (rowIncrement - 1);
									this.TableToBuild.Rows.Add(newWorkRow);
									}
								else
									{
									//-Now retrieve the **ROW** to check if the column exist...
									newWorkRow = TableToBuild.Rows.SingleOrDefault(r => r.Number == (rowCounter + rowIncrement - 1));
									}

								//+Check if the specific cell already exist in the row.
								WorkCell newWorkCell = new WorkCell();
								if (newWorkRow.Cells.Where(c => c.Number == columnCounter).FirstOrDefault() == null)
									{
									//-cell doesn't exist, number it
									newWorkCell.Number = columnCounter;
									newWorkCell.MergeRow = enumTableRowMergeType.Continue;
									newWorkCell.MergeColumns = workCell.MergeColumns;
									newWorkCell.WidthPercentage = workCell.WidthPercentage;
									newWorkCell.WidthInDxa = workCell.WidthInDxa;
									newWorkRow.Cells.Add(newWorkCell);
									}
								else
									{ //-Cell exist, update it...
									newWorkCell = newWorkRow.Cells.FirstOrDefault(c => c.Number == columnCounter);
									//-set Table Row Merge Type to **continue**
									newWorkCell.MergeRow = enumTableRowMergeType.Continue;
									newWorkCell.HTMLcontent = string.Empty;
									}
								}
							}
						else
							{
							Console.Write("\tRowSpan:{0}", workCell.MergeRow);
							}
						break;

						//---g
						default:
						//Console.WriteLine("\n ~~~ skip {0} ~~~", node.Name);
						continue;
						}
					}
				if (objNewParagraph != null)
					{
					if (!string.IsNullOrWhiteSpace(objNewParagraph.InnerText))
						{
						this.WPbody.Append(objNewParagraph);
						objNewParagraph = null;
						}
					}
				
				}
			catch (InvalidContentFormatException exc)
				{
				Console.WriteLine("\n\nInvalid Content Format Exception: {0} - {1}", exc.Message, exc.Data);
				// Update the counters before returning
				throw new InvalidContentFormatException(exc.Message);
				}
			catch (InvalidTableFormatException exc)
				{
				Console.WriteLine("\n\nException: {0} - {1}", exc.Message, exc.Data);
				// Update the counters before returning
				throw new InvalidContentFormatException(exc.Message);
				}
			catch (InvalidImageFormatException exc)
				{
				Console.WriteLine("\n\nException: {0} - {1}", exc.Message, exc.Data);
				// Update the counters before returning
				throw new InvalidContentFormatException(exc.Message);
				}
			catch (Exception exc)
				{
				// Update the counters before returning
				Console.WriteLine("\n**** Exception **** \n\t{0} - {1}\n\t{2}", exc.HResult, exc.Message, exc.StackTrace);
				throw new InvalidContentFormatException("An unexpected error occurred at this point, in the document generation. \nError detail: " + exc.Message);
				}
			finally
				{
				Console.WriteLine("\n{0} HTML decoding iterations ENDed {0}", new string('_', 25));
				}
			return;
			}

		public static string CleanText(string parText, string parClientName)
			{
			//!The sequence in which these statements appear is important
			string cleanText = string.Empty;
			if (!string.IsNullOrWhiteSpace(parText))
				{
				cleanText = parText;
				//-|Strip out any HTML tags..
				cleanText = Regex.Replace(input: cleanText, pattern: "<[^>].+?", replacement: string.Empty);		//-|**Any HTML** tags
				cleanText = Regex.Replace(input: cleanText, pattern: "&#160;", replacement: " ");                   //-|**Fixed Space** character
				cleanText = Regex.Replace(input: cleanText, pattern: "\u200b", replacement: string.Empty);          //-|**Zero Width space** DEC8203
				cleanText = Regex.Replace(input: cleanText, pattern: @"\p{Cc}", replacement: string.Empty);         //-|**Any Control** characters 
				cleanText = Regex.Replace(input: cleanText, pattern: "&#58;", replacement: ":");                    //-|**Colon**
				cleanText = Regex.Replace(input: cleanText, pattern: "&#39;", replacement: "'");                    //-|**apostrophy s**
				cleanText = Regex.Replace(cleanText, @"[\u0016]", string.Empty);                                    //-|remove **illegal** XML charaters
				cleanText = Regex.Replace(cleanText, @"[^\u0020-\u007E]+", string.Empty);                           //-|remove any other **illegal** XML charaters
				cleanText = Regex.Replace(input: cleanText, pattern: "&lt;", replacement: "less than symbol ");     //-|**Less than** symbol (<)
				cleanText = Regex.Replace(input: cleanText, pattern: "&gt;", replacement: "greater than symbol ");  //-|**Greater than** symbol
				cleanText = Regex.Replace(input: cleanText, pattern: "\n", replacement: "");                        //-|**New Line**
				cleanText = Regex.Replace(input: cleanText, pattern: "\r", replacement: "");                        //-|**Carraige Return**
				cleanText = Regex.Replace(input: cleanText, pattern: "\t", replacement: "");                        //-|**TAB**
				cleanText = Regex.Replace(input: cleanText, pattern: "&quot;", replacement: "'");                   //-|**Quote**
				cleanText = Regex.Replace(input: cleanText, pattern: "&amp;", replacement: "&");                    //-|**Ampersand**
				cleanText = Regex.Replace(input: cleanText,
							pattern: "#ClientName#", replacement: parClientName, options: RegexOptions.IgnoreCase); //-|**#ClientName#**
				//-|Remove repeating spaces
				cleanText = cleanText.Replace("     ", " ");        //-|cleanup any *5* consecutive spaces
				cleanText = cleanText.Replace("   ", " ");          //-|cleanup any *triple* consecutive spaces
				cleanText = cleanText.Replace("  ", " ");           //-|cleanup any *double* spaces
				if (cleanText == " " || cleanText == "  " || cleanText == "   ")
					cleanText = "";                                 //-|cleanup the string if it contains only a space.
				}
			return cleanText;
			}

		public static string CleanHTML(string parHTMLText, string parClientName)
			{
			string cleanHTML = string.Empty;
			if (string.IsNullOrWhiteSpace(parHTMLText) == false)
				{
				cleanHTML = parHTMLText;
				cleanHTML = Regex.Replace(input: cleanHTML, pattern: "&#160;", replacement: " ");                   //-|**Fixed Space** character
				cleanHTML = Regex.Replace(input: cleanHTML, pattern: "\u200b", replacement: string.Empty);          //-|**Zero Width space** DEC8203
				cleanHTML = Regex.Replace(input: cleanHTML, pattern: @"\p{Cc}", replacement: string.Empty);         //-|**Any Control** characters 
				cleanHTML = Regex.Replace(input: cleanHTML, pattern: "&#58;", replacement: ":");                    //-|**Colon**
				cleanHTML = Regex.Replace(input: cleanHTML, pattern: "&#39;", replacement: "'");                    //-|**apostrophy s**
				cleanHTML = Regex.Replace(cleanHTML, @"[\u0016]", string.Empty);									//-|remove **illegal** XML charaters
				cleanHTML = Regex.Replace(cleanHTML, @"[^\u0020-\u007E]+", string.Empty);							//-|remove any other **illegal** XML charaters
				cleanHTML = Regex.Replace(input: cleanHTML, pattern: "&lt;", replacement: "less than symbol ");     //-|**Less than** symbol (<)
				cleanHTML = Regex.Replace(input: cleanHTML, pattern: "&gt;", replacement: "greater than symbol ");  //-|**Greater than** symbol
				cleanHTML = Regex.Replace(input: cleanHTML, pattern: "\n", replacement: "");                        //-|**New Line**
				cleanHTML = Regex.Replace(input: cleanHTML, pattern: "\r", replacement: "");                        //-|**Carraige Return**
				cleanHTML = Regex.Replace(input: cleanHTML, pattern: "\t", replacement: "");                        //-|**TAB**
				cleanHTML = Regex.Replace(input: cleanHTML, pattern: "&quot;", replacement: "'");                   //-|**Quote**
				cleanHTML = Regex.Replace(input: cleanHTML, pattern: "&amp;", replacement: "&");                    //-|**Ampersand**
				cleanHTML = Regex.Replace(input: cleanHTML, 
										pattern: "#ClientName#", 
										replacement: parClientName,
										options: RegexOptions.IgnoreCase);											//-|**#ClientName#**
				}
			return cleanHTML;
			}

		public static Tuple<int, int> GetBulletNumberLevels(string parXpath)
			{
			int bulletLevels = 0;
			int numberLevels = 0;
			int positionInString = 0;

			//+Check the number of **Headings** </h> tags
			//-Check if the xPath contains any bullets
			positionInString = 0;

			//+Check the number of **Unorganised List** </ul> tags
			//-Check if the xPath contains any bullets
			positionInString = 0;
			if (parXpath.Contains("/ul"))
				{ //- if it contains bullets, count the number of bullets
				for (int i = 0; i < parXpath.Length - 1;)
					{
					//-get the ocurrences of bullets
					positionInString = parXpath.IndexOf("/ul", i);
					if (positionInString >= 0)
						{
						bulletLevels += 1;
						i = positionInString + 3;
						}
					else if (positionInString < 0)
						break;
					}
				}
			else //-If it doesn't contain any bullets, set **bulletLevels** to 0
				{
				bulletLevels = 0;
				}

			//+Check the number of **Organised List** </ol> tags
			//-Check if the xPath contains any **Organised Lists** (Numbered tags).
			positionInString = 0;
			if (parXpath.Contains("/ol"))
				{ //- if it contains any, count the number of occurrences
				for (int i = 0; i < parXpath.Length - 1;)
					{
					//-get the ocurrences of tags
					positionInString = parXpath.IndexOf("/ol", i);
					if (positionInString >= 0)
						{
						numberLevels += 1;
						i = positionInString + 3;
						}
					else if (positionInString < 0)
						break;
					}
				}
			else //-If it doesn't contain any numbers, set **numberLevels** to 0
				{
				numberLevels = 0;
				}

			//-Return the counted occurrences
			return new Tuple<int, int>(bulletLevels, numberLevels);
			}

		//===R

		/// <summary>
		/// string parameter typically coming from a HTML attribute.
		/// </summary>
		/// <param name="parAttribute"></param>
		/// <returns>Tupple(int,enumWidthHeightType) Item1 is the Width value; Item2 is the type of value (Percentage or Pixel) </returns>
		public static Tuple<int, enumWidthHeightType> GetHTMLwidthAttribute (string parAttribute)
			{
			int returnWidth = 0;
			enumWidthHeightType returnWidthType = enumWidthHeightType.Percent;
			decimal workWidth = 0;
			//-|Check if the parameter contains a specified "width" or whether it is only a value... 
			if (parAttribute.Contains("width"))
				{ //-|It contains the word "**width**" therefore process it in this code block...
				parAttribute = parAttribute.Replace("&#58;", "");
				parAttribute = parAttribute.Replace(":", "");
				parAttribute = parAttribute.Replace(" ", "");

				int startPosition = parAttribute.IndexOf(value: "width", startIndex: 0) + 5;
				int endPosition = parAttribute.IndexOf(";", startIndex: parAttribute.IndexOf("width", 0) + 5);

				if (parAttribute.Substring(startIndex: startPosition, length: endPosition - startPosition).Contains("px") == true)
					{
					parAttribute = parAttribute.Substring(startIndex: startPosition, length: endPosition - startPosition - 2);
					decimal.TryParse(parAttribute, out workWidth);
					returnWidthType = enumWidthHeightType.Pixels;
					}
				else if (parAttribute.Substring(startIndex: startPosition, length: endPosition - startPosition).Contains("%") == true)
					{
					parAttribute = parAttribute.Substring(startIndex: startPosition, length: endPosition - startPosition - 1);
					decimal.TryParse(parAttribute, out workWidth);
					returnWidthType = enumWidthHeightType.Percent;
					}
				else
					{
					returnWidth = 0;
					}
				}
			else
				{ //-|It doesn't contain the words "**width**", therefore process in this code block...
				if(parAttribute.IndexOf("px",0) > 0)
					{
					//-|remove the **px** and parse the value
					parAttribute = parAttribute.Replace("px", "");
					decimal.TryParse(parAttribute, out workWidth);
					returnWidthType = enumWidthHeightType.Pixels;
					}
				else if(parAttribute.IndexOf("%",0) > 0)
					{
					//-|remove the **px** and parse the value
					parAttribute = parAttribute.Replace("%", "");
					decimal.TryParse(parAttribute, out workWidth);
					returnWidthType = enumWidthHeightType.Percent;
					}
				else
					{
					workWidth = 0;
					}
				}
			returnWidth = Convert.ToInt16(workWidth);
			Tuple<int, enumWidthHeightType> returnValue = new Tuple<int, enumWidthHeightType>(returnWidth, returnWidthType);
			return returnValue;
			}

		//***R

		public void DetermineTableGrid(IEnumerable<HtmlNode> parHTMLnodes)
			{
			try
				{
				//- First clear the GridColumnWidths...
				//-| The **Tuple** consist of two intances:
				//-- |    1. Column **Width** as a n integer value
				//--|    2. **Span** indicator which is **True** if the width was derived from a *merged/spanned* cell and **False** if it was from an unmerged cell
				this.TableToBuild.GridColumnWidthPercentages = new List<int>();

				//-Initialise Variables and Properties
				int cellCounter = 0;
				int rowCounter = 0;
				//-Working Variables...
				string spanValue;
				int colSpan = 1;
				int rowSpan = 1;
				string cellWidthValue = "";
				int cellWidthPercentage = 0;
				int cellWidthPixels = 0;
				//-| The **Tuple** consist of two intances:
				//-- |    1. Column **Width** as a n integer value
				//--|    2. **Span** indicator which is **True** if the width was derived from a *merged/spanned* cell and **False** if it was from an unmerged cell
				//--|    3. **RowSpan** number Working variable which only has a value during processing, used to indicate the number of spanned rows.
				Tuple<int, bool, int> gridColumn;

				//-This list contains the **width** of each column once the method completed.
				List<Tuple<int, bool, int>> gridColumns = new List<Tuple<int, bool, int>>();

				//?This is what happens in this section of code...
				//- | 1. All the rows of the table is processed to determine the GRID of the table that need to be created in the MS Document.
				//- | 2. Each Column is each Row is inspected to determine how many columns is in the table and the **WIDTH** of each column.
				//-| 3. The complicating factors are:
				//- |     a) ... some columns may **NOT** have their width specified in the HTML.
				//-|     b) ...columns may **SPAN** multiple columns which means their width may not apply to each of the spanned columns..

				//- Process the collection of columns that were send as parameter.
				foreach (HtmlNode node in parHTMLnodes.Where(n => n.Name != "#text" && n.Name != "p"))
					{
					switch (node.Name)
						{
						//+ Table Row
						case "tr":
							{
							cellCounter = 0;
							rowCounter += 1;
							Console.Write("\n\t + <{0}>", node.Name);
							break;
							}

						//+Table Header or Table Cell
						case "th":
						case "td":
							{
Repeat_for_RowSpans:
							Console.Write("\n\t\t - <{0}>", node.Name);
							spanValue = String.Empty;
							colSpan = 1;
							rowSpan = 1;
							cellWidthValue = String.Empty;
							cellWidthPercentage = 0;
							cellWidthPixels = 0;
							cellCounter += 1;

							//-Check if RowSpan value in the **columnGrid** is greater or equal to the than 1.
							if (gridColumns.Count >= cellCounter)
								{
								//-Check if the RowSpan value is > 1, which means this column is preceded by a column previously defined by a spanned row...
								if (gridColumns[cellCounter - 1].Item3 > 1)
									{ //-Yes it is preceded by a column from a merged row.
									gridColumn = new Tuple<int, bool, int>(gridColumns[cellCounter - 1].Item1, gridColumns[cellCounter - 1].Item2,
										gridColumns[cellCounter - 1].Item3 - 1);
									gridColumns[cellCounter - 1] = gridColumn;
									Console.Write("\t Merge:| RowSpan:{0}\tColSpan:{1}\tWidth:{2}% | -->  |Keep column {3} as {4}%|",
										gridColumns[(cellCounter - 1)].Item3, 1, gridColumns[cellCounter - 1].Item1, cellCounter, gridColumns[cellCounter - 1].Item1);
									//-Repeat the code above to increase the columnCounter and process the next column completely,
									//-also to perform these conditions again in the to determine if the next column is merged row as well.
									goto Repeat_for_RowSpans;
									}
								}

							//+Check if there is a **colspan** defined for the cell...
							if (node.ChildAttributes("colspan").FirstOrDefault() == null)
								{
								spanValue = "1";
								}
							else
								{
								spanValue = node.ChildAttributes("colspan").FirstOrDefault().Value.ToString();
								}

							//-Check if the *cell* has a **colspan** value defined
							if (!string.IsNullOrWhiteSpace(spanValue))
								{
								colSpan = Convert.ToInt16(spanValue);
								}

							//+Check if there is a **rowspan** defined for the cell...
							if (node.ChildAttributes("rowspan").FirstOrDefault() == null)
								{
								spanValue = "1";
								}
							else
								{
								spanValue = node.ChildAttributes("rowspan").FirstOrDefault().Value.ToString();
								}

							//-Translate the alphanumeric value to an integer value
							if (!string.IsNullOrWhiteSpace(spanValue))
								{
								rowSpan = Convert.ToInt16(spanValue);
								}

							//+Check if a cell **width** is defined...
							//- Cell **WIDTH** can be defined as an *attribute* or in a *style* value...
							//-First Check for the **width** value defined in the *attribute*...
							Tuple<int, enumWidthHeightType> cellWidth = new Tuple<int, enumWidthHeightType>(0, enumWidthHeightType.Percent);
							if (node.ChildAttributes("width").FirstOrDefault() != null)
								{ //-The **width** attribute was found...
								cellWidthValue = node.ChildAttributes("width").FirstOrDefault().Value.ToString();
								if (!string.IsNullOrWhiteSpace(cellWidthValue))
									{ //-The *cell* has a defined **width** value - it is **NOT** *Null* or *space*...
									cellWidth = GetHTMLwidthAttribute(cellWidthValue);
									if (cellWidth.Item2 == enumWidthHeightType.Pixels)
										cellWidthPixels = cellWidth.Item1;
									else if(cellWidth.Item2 == enumWidthHeightType.Percent)
										cellWidthPercentage = cellWidth.Item1;
									}
								}
							else   //- |The **width** atrribute was not found, now check if there is a **width** defined in the **style**
								{ 
								if (node.ChildAttributes("style").Where(v => v.Value.Contains("width")).FirstOrDefault() != null)
									{
									cellWidthValue = node.ChildAttributes("style").Where(v => v.Value.Contains("width")).FirstOrDefault().Value;
									if (!String.IsNullOrWhiteSpace(cellWidthValue))
										{  //-The *cell* has a defined **width** value - it is **NOT** *Null* or *space*...
										cellWidth = GetHTMLwidthAttribute(cellWidthValue);
										if (cellWidth.Item2 == enumWidthHeightType.Pixels)
											cellWidthPixels = cellWidth.Item1;
										else if (cellWidth.Item2 == enumWidthHeightType.Percent)
											cellWidthPercentage = cellWidth.Item1;
										}
									}
								}

							//+By now the cell's **WIDTH**, **Column Span** and **Row Span** should be known.
							//-Note: each cells doesn't always have a **width**, **column span** or **row span** specified :. validation and extrapolation is required.
							//-If the **column span** is 0, set it to 1 which is the default value
							if (colSpan < 1)
								colSpan = 1;

							//-Calculate the **cell width percentage** (we only work with % from here on, to ensure scalability to document width).
							if (cellWidthPixels > 0)
								{
								//-Determine the percentage value of the cell if it was specified in pixels...
								throw new InvalidTableFormatException("There are cells in the table with a WIDTH specified in pixels instead of a percentage. "
									+ "Please inspect the content and resolve the issue, by ensuring that ALL columns in the table have % values.)");
								}

							//-If the **row span** is 0, set it to 1 which is the default value
							if (rowSpan < 1)
								rowSpan = 1;

							Console.Write("\t Read: | RowSpan:{0}\tColSpan:{1}\tWidth:{2}% |", rowSpan, colSpan, cellWidthPercentage);
							Console.Write(" --> ");

							//+Add the cell to the **gridColumns** if applicable
							//-Check if the column already exist in the columnGrid
							if (cellCounter > gridColumns.Count)
								{ //-- Column doesn't exist yet, therefore **Add** it to the columnGrid.
								cellWidthPercentage /= colSpan;

								//-The cell can span multiple (merged column) therefore multiple columns may have to be added,
								//--Therefore we loop every column span :. starting at 0 up to the number of spanned cells.
								for (int i = 0; i < colSpan; i++)
									{
									Console.Write(" |Stored column {0} as {1}%", cellCounter + i, cellWidthPercentage);
									if (rowSpan > 1)
										rowSpan -= 1;
									gridColumn = new Tuple<int, bool, int>(cellWidthPercentage, colSpan > 1, rowSpan);
									gridColumns.Add(gridColumn);
									}
								Console.Write("|");
								}

							//-An entry already exist in the gridColumns, check what needs to change if any....
							else
								{ //-The column already exist - therefore we can try to obtain a more accurate value of its width...
								  //- Locate the column in the tableGrid; and compare values
								  //-Check if the column spans multiple cells
								if (colSpan > 1)
									{ //- the cell is a merge cell and spans multiple column, which means it may still be an inaccurate value,
									  //-but possibly a more accurate than the previous estimate...
									cellWidthPercentage /= colSpan;

									//-Because the cell spans, we need to check for the values for the other spanned/merged columns as well...
									for (int i = 0; i < colSpan; i++)
										{
										//- Check if current cell's width calculation is **more** accurate that the previously stored WIDTH
										//--(use **cellCounter - 1** because the tableGrid is a List which has a 0 based index reference)
										if (gridColumns[(cellCounter - 1) + i].Item2 == false)
											{ //- This means that the previous value was derrived from a cell that was **NOT a Merged** cell,
											  //-which is considered to have a more accurate width value than the calculated value from a merged/spanned cell.

											//-Check if a the current cell has a **rowspan** (if it is part of a merged row)...
											if (rowSpan > 1)
												{
												//-Check if the RowSpan value is > 1, which means this column is preceded by a column previously defined by a spanned row...
												//if (gridColumns[cellCounter - 1].Item3 > 1)
													{ //-Yes it is preceded by a column from a merged row.
													gridColumn = new Tuple<int, bool, int>(
														gridColumns[cellCounter - 1 + i].Item1,
														gridColumns[cellCounter - 1 + i].Item2,
														rowSpan);
													gridColumns[cellCounter - 1 + i] = gridColumn;
													Console.Write(" |Keep column {0} as {1}% update Rowspan:{2}",
														cellCounter + i, gridColumns[cellCounter - 1 + i].Item1, rowSpan);
													}
												}
											else
												{
												Console.Write(" |Keep column {0} as {1}%", (cellCounter) + i, gridColumns[(cellCounter - 1) + i].Item1);
												}
											}
										else
											{ //- This means that the previous value was **also** derrived from a merged/spanned cell
											  //- which is considered **Not** to be an absolute accurate width value. Therefore we check
											  //-if this value is *smaller* than the previous width and overwrite the width if it is the case.
											if (gridColumns[(cellCounter - 1) + i].Item1 > cellWidthPercentage)
												{
												Console.Write(" |Overwrite column {0} as {1}% update rowspan:{2}", (cellCounter) + i, cellWidthPercentage, rowSpan);
												if (rowSpan > 1)
													rowSpan -= 1;
												gridColumn = new Tuple<int, bool, int>(cellWidthPercentage, true, rowSpan);
												gridColumns[(cellCounter - 1) + i] = gridColumn;
												}
											else
												{ //-This means that the saved value for the column was less than the current width of the cell,
												  //which is considered to have a more acceptable/accurate width value than the previous
												  //- calculated value from another //- merged/spanned cell.
												Console.Write(" |Keep column {0} as {1}%", (cellCounter) + i, gridColumns[(cellCounter - 1) + i].Item1);
												}
											}
										}
									cellCounter += (colSpan - 1);
									Console.Write("|");
									}
								else
									{ //-The cell doesn't have any merge/span multiple columns
									  //- therefore this **width** value is considered to be more //- accurate if the previous values was based on the
									  //- calculations of a merged/span cell
									if (gridColumns[(cellCounter - 1)].Item2 == true) //-- previous saved value was based on a merged//spanned cell, therfore
										{ //- overwrite it with a more accurate value of a unmeged/not spanned cell
										Console.Write(" |Overwrite column {0} as {1}%", (cellCounter), cellWidthPercentage);
										if (rowSpan > 1)
											rowSpan -= 1;
										gridColumn = new Tuple<int, bool, int>(cellWidthPercentage, false, rowSpan);
										gridColumns[(cellCounter - 1)] = gridColumn;
										}
									else
										{ //- This means that the saved value for the column stored from a single cell (not a merged/span cell) of the cell,
										  //- AND that this cell is ALSO a unmerged cell value...
										  //? which cell is the most accurate?
										  //? Not sure - pick one... (the smallest or the larges value **or** former or latter?) latter?
										if (cellWidthPercentage < 1)
											{ //-We don't want to end up with empty cells, rather use the spanned cell value than 0
											Console.Write(" |Keep column {0} as {1}%", cellCounter, gridColumns[(cellCounter - 1)].Item1);
											}
										else if (cellWidthPercentage == gridColumns[cellCounter - 1].Item1)
											{
											//-Check if the rowspan requires an update to the gridColumn...
											if (rowSpan > gridColumns[cellCounter - 1].Item3)
												{
												gridColumn = new Tuple<int, bool, int>(gridColumns[cellCounter - 1].Item1, gridColumns[cellCounter - 1].Item2, rowSpan);
												gridColumns[cellCounter - 1] = gridColumn;
												Console.Write(" |Keep column {0} as {1}% rowspan:{2}", cellCounter, gridColumns[(cellCounter - 1)].Item1, rowSpan);
												}
											else
												{
												Console.Write(" |Keep column {0} as {1}%", cellCounter, gridColumns[(cellCounter - 1)].Item1);
												}
											}
										else
											{
											Console.Write(" |Overwrite column {0} as {1}%", cellCounter, cellWidthPercentage);
											if (rowSpan > 1)
												rowSpan -= 1;
											gridColumn = new Tuple<int, bool, int>(cellWidthPercentage, false, rowSpan);
											gridColumns[(cellCounter - 1)] = gridColumn;
											}
										}
									Console.Write("|");
									}
								}
							break;
							}
						}
					}

				Console.WriteLine("\n\t{0}", new String('_', 120));
				Console.Write("\t Columns: |");

				//+ Check if there are any columns in the tableGrid that have 0 cellWidthPercentage values.
				int totalTableColumnWidths = 0;
				foreach (Tuple<int, bool, int> gridColumnItem in gridColumns)
					{
					Console.Write("\t {0}% ", gridColumnItem.Item1);
					if (gridColumnItem.Item1 == 0)
						{
						throw new InvalidTableFormatException("The table contains columns/cells that don't have a column width specified. Please revise the content and correct the content error. Ensure that all columns and cells have their width specified as a percentage of the table's width and that the total width for each row totals 100% of the table's width.");
						}
					else
						totalTableColumnWidths += gridColumnItem.Item1;
					}
				Console.Write(" | \t = Total: {0}%|\n", totalTableColumnWidths);
				Console.WriteLine("\t{0}", new String('_', 120));

				//- Also check that the total column of the column width doesn't exceed 100%
				if (totalTableColumnWidths > 101)
					{
					throw new InvalidTableFormatException("The total width values of one or more of the table's columns EXCEEDS 101% of the width of the table. Therefore the table cannot be accurately scalled and formatted to fit the document's page width. Please revise the content and ensure that all columns in the table widths are specified as a percentage of the table's width AND that the total width of all cells in each row don't  eceed 101% of the width specified for the table.");
					}

				//- Also check that the total width of all columns in a row is not less than 99% of the table's width.
				if (totalTableColumnWidths < 99)
					{
					throw new InvalidTableFormatException("The total values of one or more of the table's columns is LESS than 99% of the width of the table. Therefore the table cannot be accurately scalled and formatted to fit the document's page width. Please revise the content and ensure that the width of each column in a row is NOT less than 99% of the width specified for the table.");
					}

				//-At this point all is in order with the table grid, therefore insert it into the **this.TableToBeBuild**.
				foreach (var item in gridColumns)
					{
					this.TableToBuild.GridColumnWidthPercentages.Add(item.Item1);
					}
				this.TableToBuild.GridDone = true;
				}
			catch (InvalidTableFormatException exc)
				{
				Console.WriteLine("Exception: {0} - {1}", exc.Message, exc.Data);
				this.TableToBuild.Active = false;
				throw new InvalidTableFormatException(exc.Message);
				}
			catch (Exception exc)
				{
				Console.WriteLine("\n\nException ERROR: {0} - {1} - {2} - {3}", exc.HResult, exc.Source, exc.Message, exc.Data);
				}

			Console.WriteLine("{0} Table Grid Defined {0}", new String('-', 50));
			}


		//===R

		public static List<Paragraph> DissectHTMLstring(
			ref MainDocumentPart parMainDocumentPart,
			ref int parNumberingCounter,
			string parHTMLString,
			string parClientName,
			string parContentLayer,
			int parHierarchyLevel,
			int parAdditionalHierarchyLevel,
			bool parRestartNumbering,
			bool parIsTable = false,
			bool parIsTableHeader = false)
			{

			List<Paragraph> paragraphs = new List<Paragraph>();		//-List of all the paragrpahs that will be returned to the calling method.
			bool isDIVempty = false;                                //-Indicator which facilitate the validation of empty DIV tags
			Paragraph workParagraph = null;              //-Work variable used to construct a Paragraph
			Run workRun = new Run();                                //-Work variable which is used to construct a Run.
			string workString = string.Empty;                       //-A string variable used to process text
			bool bold = false;
			bool italic = false;
			bool underline = false;
			bool superScript = false;
			bool subscript = false;
			bool strikethrough = false;
			string workText = string.Empty;
			//-The sequence of the values in the tuple variable is:
			//-**1** *Bullet Levels*, 
			//-**2** *Number Levels*
			Tuple<int, int> bulletNumberLevels = new Tuple<int, int>(0, 0);
			int bulletLevel = 0;
			int numberLevel = 0;
			int headingLevel = 0;

			//Console.WriteLine("{0} Disecting HTML {0}", new string('_', 25));
			//-Declare the htmlData object instance
			HtmlAgilityPack.HtmlDocument htmlData = new HtmlAgilityPack.HtmlDocument();
			//-Load the HTML content into the htmlData object.
			//-Load a string into the HTMLDocument
			htmlData.LoadHtml(parHTMLString);
			//Console.WriteLine("HTML: {0}", htmlData.DocumentNode.OuterHtml);
			//- Set the ROOT of the data loaded in the htmlData
			var htmlRoot = htmlData.DocumentNode;

			try
				{
				foreach (HtmlNode node in htmlData.DocumentNode.DescendantsAndSelf())
					{

					//-Check is a new paragraph is initialised and it contains content

					switch (node.Name)
						{
						//---g
						//+<DIV>
						case "div":
						//-Check for other tags in the **div** tag
						isDIVempty = true;
						if (node.HasChildNodes)
							{
							//-Check if there are any other valid HTML closing tags succeeding the **<DIV>**
							foreach (HtmlNode decendentNode in node.Descendants())
								{
								//-Check for any valid content that need to be processed...
								if (decendentNode.Name == "#text"   //-//text
								|| decendentNode.Name == "p"        //-//paragraph
								|| decendentNode.Name == "ol"       //-//Organised List
								|| decendentNode.Name == "ul"       //-//Unorganised List
								|| decendentNode.Name == "li"       //-//ListItem
								|| node.XPath.Contains("/img")      //-//Image
								|| node.XPath.Contains("/table")    //-//Table
								|| decendentNode.Name.Contains("h")) //-//Heading
									{
									//- Just CONSUME the div tag and process the tags in the subsequent node cycles...
									isDIVempty = false;
									break;
									}
								}
							}

						if (isDIVempty)
							{
							//-When the code reach this point, it means that there is text in an isolated **DIV** tag
							//-which means the text cannot be properly formatted because it doesn't contain valid...
							//-No HTML tags - **THEREFORE** a format issue is raised...
							if (workParagraph != null
							&& !String.IsNullOrEmpty(workParagraph.InnerText))
								{ //-Add the *paragraph* to the list of Paragraphs...
								paragraphs.Add(workParagraph);
								workParagraph = null;
								}
							//-Construct a paragraph in which to place the content error's run content.
							throw new InvalidContentFormatException("The content was probably pasted from an external source without "
								+ "formatting it after pasting. "
								+ "Please inspect the content and resolve the issue, by formatting the relevant content using the relevant "
								+ "Enhanced Rich Text styles. |" + node.InnerText + "|");
							}
						break;
						//---g
						//+<#text> **TEXT**
						case "#text":
						//-First, clean the string by removing unwanted and unnecessary characters
						workString = CleanText(parText: node.InnerText, parClientName: parClientName);
						//-Check if the **workString** is blank, then **DON'T** process it
						if (String.IsNullOrWhiteSpace(workString))
							break;

						if (node.XPath.Contains("/strong"))
							{
							//Console.Write("|BOLD|");
							bold = true;
							}
						else
							bold = false;

						if (node.XPath.Contains("/em"))
							{
							//Console.Write("|ITALICS|");
							italic = true;
							}
						else
							italic = false;

						if (node.ParentNode.Attributes.Where(a => a.Name == "style" && a.Value.Contains("line-through")).FirstOrDefault() != null)
							{
							strikethrough = true;
							//Console.Write("|STRIKE-THROUGH|");
							}
						else
							strikethrough = false;

						if (node.ParentNode.Attributes.Where(a => a.Name == "style" && a.Value.Contains("underline")).FirstOrDefault() != null)
							{
							underline = true;
							//Console.Write("|UNDERLINE|");
							}
						else
							underline = false;

						if (node.XPath.Contains("/sub"))
							{
							//Console.Write("|SUBSCRIPT|");
							subscript = true;
							}
						else
							subscript = false;

						if (node.XPath.Contains("/sup"))
							{
							//Console.Write("|SUPERSCRIPT|");
							superScript = true;
							}
						else
							superScript = false;

						//-Insert the **text** if the string is *not* empty
						if (!string.IsNullOrWhiteSpace(workString))
							{
							//-Check if there is an initialised paragraph
							if (workParagraph == null)
								{
								workParagraph = oxmlDocument.Construct_Paragraph(
									parBodyTextLevel: 0,
									parIsTableParagraph: parIsTable,
									parIsTableHeader: parIsTableHeader);
								}

							//Console.Write("[{0}]", workString);
							workRun = oxmlDocument.Construct_RunText(
								parText2Write: workString,
								parContentLayer: parContentLayer,
								parBold: bold,
								parItalic: italic,
								parUnderline: underline,
								parSubscript: subscript,
								parSuperscript: superScript,
								parStrikeTrough: strikethrough);
							}
						workParagraph.Append(workRun);
						break;

						//+<p> - **Paragraph**
						case "p":
						//! **Normal** Paragraphs and **Table** paragraphs are a little different...
						if (workParagraph != null
						&& !string.IsNullOrEmpty(workParagraph.InnerText))
							{ //-Write the *paragraph* to the **paragraphs list**
							paragraphs.Add(workParagraph);	
							}
						workParagraph = null;

						//-Check if the paragraph contains any usable content.
						if (!node.HasChildNodes)
							{ //-Check if it contains any usable text
							workString = CleanText(node.InnerText, parClientName);
							if (string.IsNullOrWhiteSpace(workString))
								break;
							}

						//-Check if the paragraph is part of a bullet- or number- list
						if ((node.XPath.Contains("/ol")
						|| node.XPath.Contains("/ul"))
						&& node.XPath.Contains("/li"))
							{ //-If is :. get the number of bullet- or number level in the xPath
							bulletNumberLevels = GetBulletNumberLevels(node.XPath);
							bulletLevel = bulletNumberLevels.Item1;
							numberLevel = bulletNumberLevels.Item2;
							//- now exit the loop, to process the **"#text"** or other child tags...
							break;
							}
						else
							{
							bulletLevel = 0;
							numberLevel = 0;
							}

						if (bulletLevel > 0) //-**Bulleted** paragraph
							workParagraph = oxmlDocument.Construct_BulletParagraph(parBulletLevel: bulletLevel, parIsTableBullet: parIsTable);
						else if(numberLevel > 0) //-**Numbered** paragraph
							workParagraph = oxmlDocument.Construct_BulletParagraph(parBulletLevel: bulletLevel, parIsTableBullet: parIsTable);
						else //-**Normal** Paragraph.
							workParagraph = oxmlDocument.Construct_Paragraph(parIsTableParagraph: parIsTable, parIsTableHeader: parIsTableHeader);
						break;

						//---g
						//+<h1-4> **Heading 1 - 4**
						case "h1":
						case "h2":
						case "h3":
						case "h4":
						//-**Headings** are not allowed if processing **table** content
						if (parIsTable)
							{
							//-Construct a paragraph in which to place the content error's run content.
							throw new InvalidTableFormatException(message:
							"Heading  Styles cannot be included in tables, because it will disrupt the numbering system in the document. "
							+ "Rather use bullets or numbers if required."
							+ "Please inspect the content and resolve the issue, by removing Style Headings from the table.");
							}

						//Console.Write("\n {0} + <{1}>", new String('\t', headingLevel * 2), node.Name);
						//- Set the **this.AdditionalHierarchicalLevel** to the headingLevel value
						parAdditionalHierarchyLevel = headingLevel;

						//-Check if there is a populated paragraph that need to be committed before initialising the new one...
						if (workParagraph != null
						&& workParagraph.InnerText != null)
							{ //-Add the *paragraph* to the **documents**
							paragraphs.Add(workParagraph);
							workParagraph = null;
							}

						//-Check if there are child nodes, and check if the innterText is also blank
						if (!node.HasChildNodes)
							{
							if (node.InnerText == string.Empty)
								{//-skip the paragraph because it will be a blank heading in the document...
								break;
								}
							}

						//-Ininitialise a new paragrpah for the **Heading**
						workParagraph = oxmlDocument.Construct_Heading(
							parHeadingLevel: parHierarchyLevel + parAdditionalHierarchyLevel);
						break;

						//---g

						//+ <UL> **Unorganised List**
						case "ul":
						//-Check if there is a populated paragraph that need to be processed
						if (workParagraph != null
						&& workParagraph.InnerText != null)
							{ //-Write the *paragraph* to the document...
							paragraphs.Add(workParagraph);
							workParagraph = null;
							}
						else
							workParagraph = null;

						break;

						//---g
						//+<ol> **Organised List**
						case "ol":
						//-Check if there is a populated paragraph that **CONTAINS** text...
						if (workParagraph != null
						&& !string.IsNullOrWhiteSpace(workParagraph.InnerText))
							{ //-Write the *paragraph* to the document...
							paragraphs.Add(workParagraph);
							workParagraph = null;
							}
						else
							workParagraph = null;

						//-Determine the number of bullet - and number - levels usng the occurrences in xPath
						bulletNumberLevels = GetBulletNumberLevels(node.XPath);
						bulletLevel = bulletNumberLevels.Item1;
						numberLevel = bulletNumberLevels.Item2;

						//-If the number level is equal to 1, then a new number list must begin at 1.
						if (numberLevel == 1)
							{
							parRestartNumbering = true;
							parNumberingCounter += 1;
							}
						break;

						//---g
						//+<li>  **List Item**
						case "li":
						//-Determine the number of bullet- and number- levels usng the occurrences in xPath
						bulletNumberLevels = GetBulletNumberLevels(node.XPath);
						bulletLevel = bulletNumberLevels.Item1;
						numberLevel = bulletNumberLevels.Item2;

						//-Check if there is a populated paragraph that contains text and write it to the Document before initiating a new paragraph...
						if (workParagraph != null
						&& !string.IsNullOrEmpty(workParagraph.InnerText))
							{ //-Write the *paragraph* to the document...
							paragraphs.Add(workParagraph);
							workParagraph = null;
							}

						//-Construct the paragraph with the bullet level depending bulletLevel value
						if (bulletLevel > 0)
							{//- if it is a **Bullet** list entry, create a new **Pargraph** *Bullet* object...
							workParagraph = oxmlDocument.Construct_BulletParagraph(
								parIsTableBullet: parIsTable,
								parBulletLevel: bulletLevel);
							}

						//- check if it is **Organised/Number list** item
						else if (numberLevel > 0)
							{//-if it is a **Number** list entry, create a new **Pargraph** *Number* object instance...
							if (parRestartNumbering)
								parNumberingCounter += 1;
							workParagraph = oxmlDocument.Construct_NumberParagraph(
								parMainDocumentPart: ref parMainDocumentPart,
								parIsTableNumber: parIsTable,
								parNumberLevel: numberLevel,
								parRestartNumbering: parRestartNumbering,
								parNumberingId: parNumberingCounter);
							parRestartNumbering = false;
							}
						break;

						//---g
						//++Image
						case "img":
						break;
						}
					}
				//-Commit the last paragraph if it has not been written yet.
				if (workParagraph != null
				&& !string.IsNullOrEmpty(workParagraph.InnerText))
					paragraphs.Add(workParagraph);

				}
			catch (InvalidContentFormatException exc)
				{
				Console.WriteLine("\n\nInvalid Content Format Exception: {0} - {1}", exc.Message);
				// Update the counters before returning
				throw new InvalidContentFormatException(exc.Message);
				}
			catch (InvalidTableFormatException exc)
				{
				Console.WriteLine("\n\n Invalid Table Format Exception: {0} - {1}", exc.Message);
				// Update the counters before returning
				throw new InvalidContentFormatException(exc.Message);
				}
			catch (InvalidImageFormatException exc)
				{
				Console.WriteLine("\n\nInvalid Image Exception: {0} - {1}", exc.Message);
				// Update the counters before returning
				throw new InvalidContentFormatException(exc.Message);
				}
			catch (Exception exc)
				{
				// Update the counters before returning
				Console.WriteLine("\n**** Exception **** \n\t{0} - {1}\n\t{2}", exc.HResult, exc.Message, exc.StackTrace);
				throw new InvalidContentFormatException("An unexpected error occurred at this point, in the document generation. \nError detail: " + exc.Message);
				}

			//Console.WriteLine("\n{0} Done dissecting HTML {0}", new string('_', 25));
			return paragraphs;
			}
		} 
	}