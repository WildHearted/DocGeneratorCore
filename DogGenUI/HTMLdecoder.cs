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

		private Body _wpbody;
		public Body WPbody
			{
			get { return this._wpbody; }
			set { this._wpbody = value; }
			}
		/// <summary>
		/// The Additional Hierarchical Level property contains the number of additional levels that need to be added to the Document Hierarchical Level when processing the HTML contained in a Enhanced Rich Text column/field.
		/// </summary>
		private int _additionalHierarchicalLevel;
		private int AdditionalHierarchicalLevel
			{
			get { return this._additionalHierarchicalLevel; }
			set { this._additionalHierarchicalLevel = value; }
			}

		/// <summary>
		/// The Document Hierarchical Level provides the stating Hierarchical level at which new content will be added to the document.
		/// </summary>
		private int _documentHierarchyLevel;
		public int DocumentHierachyLevel
			{
			get { return this._documentHierarchyLevel; }
			set { this._documentHierarchyLevel = value; }
			}

		private string _textBeforeFirstTag;
		private string TextBeforeFirstTag
			{
			get { return this._textBeforeFirstTag; }
			set { this._textBeforeFirstTag = value; }
			}

		//private string _textAfterFirstTag;
		//private string TextAfterFirstTag
		//	{
		//	get { return this._textAfterFirstTag; }
		//	set { this._textAfterFirstTag = value; }
		//	}

		//private string _textBeforeLastTag;
		//private string TextBeforeLastTag
		//	{
		//	get { return this._textBeforeLastTag; }
		//	set { this._textBeforeLastTag = value; }
		//	}

		private string _textAfterLastTag;
		private string TextAfterLastTag
			{
			get { return this._textAfterLastTag; }
			set { this._textAfterLastTag = value; }
			}

		private int _spanTags;
		private int SpanTags
			{
			get{return this._spanTags;}
			set{this._spanTags = value;}
			}

		private int _brTags;
		private int BRtags
			{
			get{return this._brTags;}
			set{this._brTags = value;}
			}

		private int _otherTags;
		private int OtherTags
			{
			get{return this._otherTags;}
			set{this._otherTags = value;}
			}

		private bool _paragraphOn;
		private bool ParagraphOn
			{
			get { return this._paragraphOn; }
			set { this._paragraphOn = value; }
			}

		private bool _boldOn;
		private bool BoldOn
			{
			get { return this._boldOn; }
			set { this._boldOn = value; }
			}

		private bool _UnderlineOn;
		private bool UnderlineOn
			{
			get { return this._UnderlineOn; }
			set { this._UnderlineOn = value; }
			}

		private bool _italicsOn;
		private bool ItalicsOn
			{
			get { return this._italicsOn; }
			set { this._italicsOn = value; }
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
			this.AdditionalHierarchicalLevel = 0;

			// http://stackoverflow.com/questions/11250692/how-can-i-parse-this-html-to-get-the-content-i-want
			IHTMLDocument2 objHTMLDocument2 = (IHTMLDocument2) new HTMLDocument();
			objHTMLDocument2.write(parHTML2Decode);

			//objHTMLDocument.body.innerHTML = this.EncodedHTML;
			//Console.WriteLine("{0}", objHTMLDocument2.body.innerHTML);
			ProcessHTMLelements(objHTMLDocument2.all);
			return true;
			}

		private void ProcessHTMLelements(IHTMLElementCollection parHTMLElements)
			{
			Paragraph objParagraph = new Paragraph();
			DocumentFormat.OpenXml.Wordprocessing.Run objRun = new DocumentFormat.OpenXml.Wordprocessing.Run();

			if(parHTMLElements.length > 0)
				{
				foreach(IHTMLElement objHTMLelement in parHTMLElements)
					{
					Console.WriteLine("HTMLlevel: {0} - html.tag=<{1}>", this.AdditionalHierarchicalLevel, objHTMLelement.tagName);
					//Console.WriteLine("\tinnerHTML: {0}", objHTMLelement.innerHTML);

					switch(objHTMLelement.tagName)
						{
						//-----------------------
						case "DIV":
						//-----------------------
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
						//---------------------------
						case "P": // Paragraph Tag
						//---------------------------
							//this.ParagraphOn = true;
							if(objParagraph != null)
								{
								this.WPbody.Append(objParagraph);
								objParagraph = null;
								}
							objParagraph = oxmlDocument.Construct_Paragraph(this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel);
							if(objHTMLelement.children.length > 0) // check if there are more html tags in the HTMLelement
								{
								this.CheckForPrePostHTMLtags(objHTMLelement.innerHTML);
								if(this.OtherTags > 0)
									{
									if(this.TextBeforeFirstTag != "")
										{
										objRun = oxmlDocument.Construct_RunText
											(parText2Write: this.TextBeforeFirstTag, parBold: this.BoldOn, parItalic: this.ItalicsOn,
											parUnderline: this.UnderlineOn);
										objParagraph.Append(objRun);
										}
									//Process the cascading tags.
									ProcessHTMLelements(parHTMLElements: objHTMLelement.children);
									//-----------------------------
									if(this.TextAfterLastTag != "")
										{
										oxmlDocument.Construct_RunText
											(parText2Write: this.TextAfterLastTag, parBold: this.BoldOn, parItalic: this.ItalicsOn,
											parUnderline: this.UnderlineOn);
										objParagraph.Append(objRun);
										}
									} //if(this.OtherTags > 0)
								else   // 
									{
									if(objHTMLelement.innerText != "")
										{
										oxmlDocument.Construct_RunText
												(parText2Write: objHTMLelement.innerText, parBold: this.BoldOn, parItalic: this.ItalicsOn,
												parUnderline: this.UnderlineOn);
										objParagraph.Append(objRun);
										}
									}
								}
							else  // there are no cascading tags, just write the text if there are any
								{
								if(objHTMLelement.innerText != "")
									{
									oxmlDocument.Construct_RunText
											(parText2Write: objHTMLelement.innerText, parBold: this.BoldOn, parItalic: this.ItalicsOn,
											parUnderline: this.UnderlineOn);
									objParagraph.Append(objRun);
									}
								}
							if(objParagraph != null)
								{
								this.WPbody.Append(objParagraph);
								objParagraph = null;
								}
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
							this.BoldOn = true;
							if(objHTMLelement.children.length > 0)
								{
								this.CheckForPrePostHTMLtags(objHTMLelement.innerHTML);
								if(this.OtherTags > 0)
									{
									if(this.TextBeforeFirstTag != "")
										{
										objRun = oxmlDocument.Construct_RunText
											(parText2Write: this.TextBeforeFirstTag, parBold: this.BoldOn, parItalic: this.ItalicsOn,
											parUnderline: this.UnderlineOn);
										objParagraph.Append(objRun);
										}
									//Process the cascading tags.
									ProcessHTMLelements(parHTMLElements: objHTMLelement.children);
									//-----------------------------
									if(this.TextAfterLastTag != "")
										{
										oxmlDocument.Construct_RunText
											(parText2Write: this.TextAfterLastTag, parBold: this.BoldOn, parItalic: this.ItalicsOn,
											parUnderline: this.UnderlineOn);
										objParagraph.Append(objRun);
										}
									} //if(this.OtherTags > 0)
								else   // 
									{
									if(objHTMLelement.innerText != "")
										{
										oxmlDocument.Construct_RunText
												(parText2Write: objHTMLelement.innerText, parBold: this.BoldOn, parItalic: this.ItalicsOn,
												parUnderline: this.UnderlineOn);
										objParagraph.Append(objRun);
										}
									}
								}
							else  // there are no cascading tags, just write the text if there are any
								{
								if(objHTMLelement.innerText != "")
									{
									oxmlDocument.Construct_RunText
											(parText2Write: objHTMLelement.innerText, parBold: this.BoldOn, parItalic: this.ItalicsOn,
											parUnderline: this.UnderlineOn);
									objParagraph.Append(objRun);
									}
								}
							this.BoldOn = false;
							break;
						//------------------------------------
						case "Span":   // Underline is embedded in the Span tag
							if (objHTMLelement.tagName == "span style=" + "" + "text-styleTextDecoration;underline;" + "" +">" )
								{
								Console.WriteLine("Underline: {0}\n{1}", objHTMLelement.tagName, objHTMLelement.innerHTML);
								}
							break;
						//------------------------------------
						case "EM":     // Italic Tag
							this.ItalicsOn = true;
							if(objHTMLelement.children.length > 0)
								{
								this.CheckForPrePostHTMLtags(objHTMLelement.innerHTML);
								if(this.OtherTags > 0)
									{
									if(this.TextBeforeFirstTag != "")
										{
										objRun = oxmlDocument.Construct_RunText
											(parText2Write: this.TextBeforeFirstTag, parBold: this.BoldOn, parItalic: this.ItalicsOn,
											parUnderline: this.UnderlineOn);
										objParagraph.Append(objRun);
										}
									//Process the cascading tags.
									ProcessHTMLelements(parHTMLElements: objHTMLelement.children);
									//-----------------------------
									if(this.TextAfterLastTag != "")
										{
										oxmlDocument.Construct_RunText
											(parText2Write: this.TextAfterLastTag, parBold: this.BoldOn, parItalic: this.ItalicsOn,
											parUnderline: this.UnderlineOn);
										objParagraph.Append(objRun);
										}
									} //if(this.OtherTags > 0)
								else   // 
									{
									if(objHTMLelement.innerText != "")
										{
										oxmlDocument.Construct_RunText
												(parText2Write: objHTMLelement.innerText, parBold: this.BoldOn, parItalic: this.ItalicsOn,
												parUnderline: this.UnderlineOn);
										objParagraph.Append(objRun);
										}
									}
								}
							else  // there are no cascading tags, just write the text if there are any
								{
								if(objHTMLelement.innerText != "")
									{
									oxmlDocument.Construct_RunText
											(parText2Write: objHTMLelement.innerText, parBold: this.BoldOn, parItalic: this.ItalicsOn,
											parUnderline: this.UnderlineOn);
									objParagraph.Append(objRun);
									}
								}
							this.ItalicsOn = false;
							break;
						//------------------------------------
						case "SUB":    // Subscript Tag

							break;
						//------------------------------------
						case "SUP":    // Super Script Tag

							break;
						//------------------------------------
						case "H1":     // Heading 1
						case "H1A":    // Alternate Heading 1

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
						case "H4":     // Heading 4
						case "H4A":    // Alternate Heading 4

							break;
						default:
							Console.WriteLine("\t - ignoring tag: {0}", objHTMLelement.tagName);
							break;

						} // switch(objHTMLelement.tagName)


					} // foreach(IHTMLElement objHTMLelement in parHTMLElements)


				} // if (parHTMLElements.length > 0)


			}

		public static void WriteTextToDoc(string parText2Write)
			{
			}

		public void CheckForPrePostHTMLtags(string parTextString)
			{
			int iTagCheck = 1;
			int iTagStart = 0;
			string sCheckString = "";
			this.TextBeforeFirstTag = "";
			this.TextAfterLastTag = "";
			this.SpanTags = 0;
			this.BRtags = 0;
			this.OtherTags = 0;

			// Check if there are any text BEFORE the first tag
			iTagStart = parTextString.IndexOf("<");
			//check if any tags were found in parText2Check
			if(iTagStart > iTagCheck)
				{
				//this means that there is text before the first HTML tag
				sCheckString = parTextString.Substring(startIndex: iTagCheck, length: (iTagStart - iTagCheck));
				sCheckString = sCheckString.Replace(oldValue: " ", newValue: "");
				sCheckString = sCheckString.Replace(oldValue: "&nbsp;", newValue: "");
				sCheckString = sCheckString.Replace(oldValue: "&#160;", newValue: "");
				sCheckString = sCheckString.Replace(oldChar: (char)63, newChar: Convert.ToChar(value: ""));
				// Check if the sCheckString is greater than 0 length after the replacements of invalid characters
				if(sCheckString.Length > 0)
					{
					this.TextBeforeFirstTag = parTextString.Substring(startIndex: iTagStart, length: (iTagStart - iTagCheck));
					this.TextBeforeFirstTag = this.TextBeforeFirstTag.Replace(oldValue: "  ", newValue: " ");
					this.TextBeforeFirstTag = this.TextBeforeFirstTag.Replace(oldValue: "&nbsp;", newValue: "");
					this.TextBeforeFirstTag = this.TextBeforeFirstTag.Replace(oldValue: "&#160", newValue: "");
					}
				}

			// check how many other tags are in parText2Check
			IHTMLDocument2 objHTMLworkDoc = (IHTMLDocument2) new HTMLDocument();
			objHTMLworkDoc.write(parTextString);
			IHTMLElementCollection objHTMLworkElements = (IHTMLElementCollection) objHTMLworkDoc.body.children;
			if(objHTMLworkElements.length > 0)
				{
				foreach(IHTMLElement objHTMLworkElement in objHTMLworkElements)
					{
					switch(objHTMLworkElement.tagName)
						{
						case "SPAN":
							this.SpanTags += 1;
							break;
						case "BR":
							this.BRtags += 1;
							break;
						default:
							this.OtherTags += 1;
							break;
						}
					}
				}

			// Check if there is any text AFTER the last tag
			iTagStart = parTextString.LastIndexOf(">");
			//check if any ">" tags were found in parText2Check
			if(iTagStart < parTextString.Length)
				{
				//this means that there is text AFTER the last HTML tag
				sCheckString = parTextString.Substring(startIndex: iTagStart + 1, length: (parTextString.Length - iTagStart));
				sCheckString = sCheckString.Replace(oldValue: " ", newValue: "");
				sCheckString = sCheckString.Replace(oldValue: "&nbsp;", newValue: "");
				sCheckString = sCheckString.Replace(oldValue: "&#160;", newValue: "");
				sCheckString = sCheckString.Replace(oldChar: (char) 63, newChar: Convert.ToChar(value: ""));
				// Check if the sCheckString is greater than 0 length after the replacements of invalid characters
				if(sCheckString.Length > 0)
					{
					this.TextBeforeFirstTag = parTextString.Substring(startIndex: iTagStart + 1, length: (parTextString.Length - iTagStart));
					this.TextBeforeFirstTag = this.TextBeforeFirstTag.Replace(oldValue: "  ", newValue: " ");
					this.TextBeforeFirstTag = this.TextBeforeFirstTag.Replace(oldValue: "&nbsp;", newValue: "");
					this.TextBeforeFirstTag = this.TextBeforeFirstTag.Replace(oldValue: "&#160", newValue: "");
					}
				}

			} // end of Class
		}
	}
