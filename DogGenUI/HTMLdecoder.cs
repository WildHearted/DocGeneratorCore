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

namespace DogGenUI
	{
	class HTMLdecoder
		{
		// ------------------
		// Object Properties
		// ------------------
		private string _encodedHTML;
		public string EncodedHTML
			{
			get {return this._encodedHTML;}
			set {this._encodedHTML = value;}
			}
		private Body _wpbody;
		public Body WPbody
			{
			get {return this._wpbody;}
			set { this._wpbody = value;}
			}

		private int _htmlLevel;
		private int HTMLlevel
			{
			get {return this._htmlLevel;}
			set {this._htmlLevel = value;}
			}
		private int _documentHierarchyLevel;

		public int DocumentHierachyLevel
			{
			get { return this._documentHierarchyLevel;}
			set { this._documentHierarchyLevel = value;}
			}

		private bool _paragraphOn;
		private bool ParagraphOn
			{
			get{return this._paragraphOn;}
			set{this._paragraphOn = value;}
			}

		private bool _boldOn;
		private bool BoldOn
			{
			get{return this._boldOn;}
			set{this._boldOn = value;}
			}

		private bool _UnderlineOn;
		private bool UnderlineOn
			{
			get{return this._UnderlineOn;}
			set{this._UnderlineOn = value;}
			}

		private bool _italicsOn;
		private bool ItalicsOn
			{
			get{return this._italicsOn;}
			set{this._italicsOn = value;}
			}

		// ----------------
		// Object Methods
		// ---------------

		/// <summary>
		/// Use this method once a new HTMLdecoder object is initialised and the 
		/// EndodedHTML property was set to the value of the HTML that has to be decoded.
		/// </summary>
		/// <param name="parDocumentLevel">
		/// Provide the document's hierarchical level at which the HTML has to be inserted.
		/// </param>
		/// <returns>
		/// returns a boolean value of TRUE if insert was successfull and FALSE if there was any for of failure during the insertion.
		/// </returns>
		public bool DecodeHTML(int parDocumentLevel, string parHTML2Decode)
			{
			Console.WriteLine("HTML to decode: \n\r{0}", parHTML2Decode);
			this.DocumentHierachyLevel = parDocumentLevel;
			this.HTMLlevel = 0;
			HTMLDocument objHTMLDocument = new HTMLDocument();
			IHTMLDocument2 objHTMLDocument2 = (IHTMLDocument2) objHTMLDocument;
			objHTMLDocument2.write(parHTML2Decode);

			//objHTMLDocument.body.innerHTML = this.EncodedHTML;
			Console.WriteLine("{0}", objHTMLDocument.body.innerHTML);
			ProcessHTMLelements(objHTMLDocument.body.children);
			return true;
			}

		private void ProcessHTMLelements(IHTMLElementCollection parHTMLElements)
			{
			Paragraph objParagraph = new Paragraph();
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
			this.HTMLlevel += 1;
			if(parHTMLElements.length > 0)
				{
				foreach(IHTMLElement objHTMLelement in parHTMLElements)
					{
					Console.WriteLine("HTMLlevel: {0} - html.tag=<{1}>", this.HTMLlevel, objHTMLelement.tagName);
					Console.WriteLine("outerHTML: {0}", objHTMLelement.innerHTML);
					
					switch(objHTMLelement.tagName)
						{
						case "DIV":
							if(objHTMLelement.children.length > 0)
								ProcessHTMLelements(objHTMLelement.children);
							else
								if(objParagraph == null)
								{
								oxmlDocument.Construct_Paragraph(this.DocumentHierachyLevel);
								this.WPbody.Append(objParagraph);
								}
							else
								{
								oxmlDocument.Construct_RunText
									(parText2Write: objHTMLelement.innerText, 
									parBold: this.BoldOn, 
									parItalic: this.ItalicsOn, 
									parUnderline: this.UnderlineOn);

								}
							break;
						//------------------------------------
						case "P": // Paragraph Tag
							if(ParagraphOn)
								{
								oxmlDocument.Construct_Paragraph(this.DocumentHierachyLevel);
								this.WPbody.Append(objParagraph);
								}

							oxmlDocument.Construct_Paragraph(this.DocumentHierachyLevel);
							this.WPbody.Append(objParagraph);
							this.ParagraphOn = true;
							Console.WriteLine("Children.length: {0}", objHTMLelement.children.length);
							break;
						//------------------------------------
						case "TABLE":

							break;
						//------------------------------------
						case "TBody": // Table Body

							break;
						//------------------------------------
						case "TR":     // Table Row

							break;
						//------------------------------------
						case "TH":     // Table Header

							break;
						//------------------------------------
						case "TD":     // Table Cell

							break;
						//------------------------------------
						case "UL":     // Unorganised List (Bullets to follow) Tag

							break;
						//------------------------------------
						case "OL":     // Orginised List (numbered list) Tag

							break;
						//------------------------------------
						case "LI":     // List Item (an entry from a organised or unorginaised list

							break;
						//------------------------------------
						case "IMG":    // Image Tag

							break;
						case "Strong": // Bold Tag

							break;
						//------------------------------------
						case "Span":   // Underline is embedded in the Span tag

							break;
						//------------------------------------
						case "EM":     // Italic Tag

							break;
						//------------------------------------
						case "SUB":    // Subscript Tag

							break;
						//------------------------------------
						case "SUP":    // Super Script Tag

							break;
						//------------------------------------
						case "H1":	// Heading 1
						case "H1A":	// Alternate Heading 1

							break;
						//------------------------------------
						case "H2":     // Heading 2
						case "H2A":    // Alternate Heading 2

							break;
						//------------------------------------
						case "H3":     // Heading 3
						case "H3A":    // Alternate Heading 3

							break;
						//------------------------------------
						case "H4":	// Heading 4
						case "H4A":	// Alternate Heading 4

							break;
						default:

							break;
						
						} // switch(objHTMLelement.tagName)


					} // foreach(IHTMLElement objHTMLelement in parHTMLElements)


				} // if (parHTMLElements.length > 0)


			}

		public static void  WriteTextToDoc(string parText2Write)
			{
			}
		}
	}
