using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using mshtml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using System.Threading.Tasks;

namespace DocGenerator
	{
	/// <summary>
	/// The RTdecoder or Rich Text Decoder is used to decode Rich Text (not Enhanced Rich Text - use the HTMLdecoder to decode Enhanced RichText).
	/// RTdecoder will not process images and tables.
	/// </summary>
	class RTdecoder
		{
		// ------------------
		// Object Properties
		// ------------------
		private List<Paragraph> _paragraphList;
		public List<Paragraph> PargraphList
			{
			get{return this._paragraphList;}
			set{this._paragraphList = value;}
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
			get{return this._additionalHierarchicalLevel;}
			set{this._additionalHierarchicalLevel = value;}
			}

		private bool _isTableText = false;
		public bool IsTableText
			{
			get{return this._isTableText;}
			set{this._isTableText = value;}
			}
		private string _contentLayer = "None";
		public string ContentLayer
			{
			get{return this._contentLayer;}
			set{this._contentLayer = value;}
			}

		/// <summary>
		/// 
		/// </summary>
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
		/// The unique ID of the hyperlink if it need to be inserted. Works in concjunction with the HyperlinkURL and HyoperlinkImageRelationshipID
		/// </summary>
		private int _hyperlinkID = 0;
		public int HyperlinkID
			{
			get{return this._hyperlinkID;}
			set{this._hyperlinkID = value;}
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

		// Procedures/Methods
		public List<Paragraph> DecodeRichText(
			string parRT2decode,
               string parContentLayer = "None")
			{
			this.ContentLayer = parContentLayer;

			// move the content to be decoded into a IHTMLDocument in order to process the HTML hierarchy.
			IHTMLDocument2 objHTMLDocument2 = (IHTMLDocument2)new HTMLDocument();
			objHTMLDocument2.write(parRT2decode);

			// Process the HTML contained in the RT and validate whether it was successfull.
			if(ProcessElements(objHTMLDocument2.body.children))
                    {

				}
			else // if the processing failed
				{

				}

			//Console.WriteLine("{0}", objHTMLDocument2.body.innerHTML);
			Paragraph objParagraph = new Paragraph();



			return
               }

		public bool ProcessElements(IHTMLElementCollection parHTMLelements)
			{
			Paragraph objParagraph = new Paragraph();
			Run objRun = new Run();
			try
				{
				Console.WriteLine("parHTMLElements.length = {0}", parHTMLelements.length);
				if(parHTMLelements.length < 1) // there are no cascading HTML content
					{
					objRun = oxmlDocument.Construct_RunText(parText2Write: " ");
					objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: this.DocumentHierachyLevel, parIsTableParagraph: this.IsTableText);
					objParagraph.Append(objRun);
					this.PargraphList.Add(objParagraph);
					}
				else // There are cascading HTML content to process
					{

					}

				foreach(IHTMLElement objHTMLelement in parHTMLelements)
					{
					Console.WriteLine("HTMLlevel: {0} - html.tag=<{1}>\n\t|{2}|", this.AdditionalHierarchicalLevel, 
						objHTMLelement.tagName, 
						objHTMLelement.innerHTML);
					switch(objHTMLelement.tagName)
						{
						//-----------
						case "DIV":
						//-----------
							if(objHTMLelement.children.length > 0)
								ProcessElements(objHTMLelement.children);
							else
								{
								if(objHTMLelement.innerText != null)
									{
									objParagraph = oxmlDocument.Construct_Paragraph(
										parBodyTextLevel: this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel,
										parIsTableParagraph: this.IsTableText);
									objRun = oxmlDocument.Construct_RunText(parText2Write: objHTMLelement.innerText);
									objParagraph.Append(objRun);
									this.PargraphList.Add(objParagraph);
									}
								}
							break;
						// ---------------------------
						case "P": // Paragraph Tag
						//---------------------------
							if(objHTMLelement.innerText != null)
								{
								objParagraph = oxmlDocument.Construct_Paragraph(this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel, this.IsTableText);
								if(objHTMLelement.children.length > 0) // check if there are more html tags in the HTMLelement
									{
									// use the DissectHTMLstring method to process the paragraph.
									List<TextSegment> listTextSegments = new List<TextSegment>();
									listTextSegments = TextSegment.DissectHTMLstring(objHTMLelement.innerHTML);
									// Process the list to insert the content into Paragraph List
									foreach(TextSegment objTextSegment in listTextSegments)
										{
										if(objTextSegment.Image) // Check if it is an image
											{
											if(this.IsTableText)
												throw new InvalidRichTextFormatException("Attempted to insert a image into a table.");
											else
												throw new InvalidRichTextFormatException("Rich Text is not suppose to contain an Image");
											}
										else // not an image
											{
											objRun = oxmlDocument.Construct_RunText
												(parText2Write: objTextSegment.Text,
												parContentLayer: this.ContentLayer,
												parBold: objTextSegment.Bold,
												parItalic: objTextSegment.Italic,
												parUnderline: objTextSegment.Undeline,
												parSubscript: objTextSegment.Subscript,
												parSuperscript: objTextSegment.Superscript);
											objParagraph.Append(objRun);
											}
										} // foreach loop end
									this.PargraphList.Add(objParagraph);
									}
								else  // there are no cascading tags, just write the text if there are any
									{
									if(objHTMLelement.innerText.Length > 0)
										{
										if(!objHTMLelement.outerHTML.Contains("<P></P>"))
											{
											objRun = oxmlDocument.Construct_RunText(parText2Write:
												objHTMLelement.innerText,
												parContentLayer: this.ContentLayer);
											objParagraph.Append(objRun);
											this.PargraphList.Add(objParagraph);
											}
										}
									} // there are no cascading tags
								} // if(objHTMLelement.innerText != null)
							break;
						//-----------------
						case "TABLE":
						//-----------------
							if(this.IsTableText)
								{
								throw new InvalidRichTextFormatException("Attempted to insert a table into a table (no cascading tables allowed).");
								}
							else
								{
								throw new InvalidRichTextFormatException("Rich Text is not suppose to contain a Table.");
								}
						//----------------------------
						case "TBODY": // Table Body
						case "TR":     // Table Row
						case "TH":     // Table Header
						case "TD":     // Table Cell
						//----------------------------
							Console.WriteLine("Ingnore all Table related tags.");
							break;
						//------------------------------------
						case "OL": // Orginised List (numbered list) Tag
						//-----------------------------------
							//Console.WriteLine("Tag: ORGANISED LIST\n{0}", objHTMLelement.outerHTML);
							if(objHTMLelement.children.length > 0)
								{
								ProcessElements(objHTMLelement.children);
								}
							break;
						//----------------------
						case "LI":    // List Item (an entry from a organised or unorginaised list
						//----------------------
							if(objHTMLelement.parentElement.tagName == "OL") // number list
								{
								objParagraph = oxmlDocument.Construct_BulletNumberParagraph(
									parIsBullet: this.IsTableText, 
									parBulletLevel: this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel);
								}
							else // "UL" == Unorganised/Bullet list item
								{
								objParagraph = oxmlDocument.Construct_BulletNumberParagraph(
									parIsBullet: this.IsTableText, 
									parBulletLevel: this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel);
								}
							if(objHTMLelement.children.length > 0) // check if there are more html tags in the HTMLelement
								{
								// use the DissectHTMLstring method to process the paragraph.
								List<TextSegment> listTextSegments = new List<TextSegment>();
								listTextSegments = TextSegment.DissectHTMLstring(objHTMLelement.innerHTML);
								foreach(TextSegment objTextSegment in listTextSegments)
									{
									objRun = oxmlDocument.Construct_RunText
										(parText2Write: objTextSegment.Text,
										parContentLayer: this.ContentLayer,
										parBold: objTextSegment.Bold,
										parItalic: objTextSegment.Italic,
										parUnderline: objTextSegment.Undeline,
										parSubscript: objTextSegment.Subscript,
										parSuperscript: objTextSegment.Superscript);
									objParagraph.Append(objRun);
									}
								this.PargraphList.Add(objParagraph);
								}
							else  // there are no cascading tags, just write the text if there are any
								{
								if(objHTMLelement.innerText.Length > 0)
									{
									objRun = oxmlDocument.Construct_RunText(parText2Write: objHTMLelement.innerText);
									objParagraph.Append(objRun);
									this.PargraphList.Add(objParagraph);
									}
								}

							break;
						// -------------------------
						case "IMG":    // Image Tag
						//---------------------------
							if(this.IsTableText)
								{
								throw new InvalidRichTextFormatException("Attempted to insert an Image into a table.");
								}
							else
								{
								throw new InvalidRichTextFormatException("Rich Text is not suppose to contain an Image.");
								}

						//----------------------------------
						case "STRONG":	// Bold Text
						//-------------------------------
							if(objHTMLelement.innerText != null)
								{
								objParagraph = oxmlDocument.Construct_Paragraph(this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel, this.IsTableText);
								if(objHTMLelement.children.length > 0) // check if there are more html tags in the HTMLelement
									{
									// use the DissectHTMLstring method to process the paragraph.
									List<TextSegment> listTextSegments = new List<TextSegment>();
									listTextSegments = TextSegment.DissectHTMLstring(objHTMLelement.innerHTML);
									// Process the list to insert the content into Paragraph List
									foreach(TextSegment objTextSegment in listTextSegments)
										{
										if(objTextSegment.Image) // Check if it is an image
											{
											if(this.IsTableText)
												throw new InvalidRichTextFormatException("Attempted to insert a image into a table.");
											else
												throw new InvalidRichTextFormatException("Rich Text is not suppose to contain an Image");
											}
										else // not an image
											{
											objRun = oxmlDocument.Construct_RunText
												(parText2Write: objTextSegment.Text,
												parContentLayer: this.ContentLayer,
												parBold: true,
												parItalic: objTextSegment.Italic,
												parUnderline: objTextSegment.Undeline,
												parSubscript: objTextSegment.Subscript,
												parSuperscript: objTextSegment.Superscript);
											objParagraph.Append(objRun);
											}
										} // foreach loop end
									this.PargraphList.Add(objParagraph);
									}
								else  // there are no cascading tags, just write the text if there are any
									{
									if(objHTMLelement.innerText.Length > 0)
										{
										if(!objHTMLelement.outerHTML.Contains("<P></P>"))
											{
											objRun = oxmlDocument.Construct_RunText(
												parText2Write:objHTMLelement.innerText,
												parContentLayer: this.ContentLayer,
												parBold: true);
											objParagraph.Append(objRun);
											this.PargraphList.Add(objParagraph);
											}
										}
									} // there are no cascading tags
								} // if(objHTMLelement.innerText != null)
							break;
						// --------------------------
						case "EM":  // Italic Tag
						//---------------------------
							if(objHTMLelement.innerText != null)
								{
								objParagraph = oxmlDocument.Construct_Paragraph(this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel, this.IsTableText);
								if(objHTMLelement.children.length > 0) // check if there are more html tags in the HTMLelement
									{
									// use the DissectHTMLstring method to process the paragraph.
									List<TextSegment> listTextSegments = new List<TextSegment>();
									listTextSegments = TextSegment.DissectHTMLstring(objHTMLelement.innerHTML);
									// Process the list to insert the content into Paragraph List
									foreach(TextSegment objTextSegment in listTextSegments)
										{
										if(objTextSegment.Image) // Check if it is an image
											{
											if(this.IsTableText)
												throw new InvalidRichTextFormatException("Attempted to insert a image into a table.");
											else
												throw new InvalidRichTextFormatException("Rich Text is not suppose to contain an Image");
											}
										else // not an image
											{
											objRun = oxmlDocument.Construct_RunText
												(parText2Write: objTextSegment.Text,
												parContentLayer: this.ContentLayer,
												parBold: objTextSegment.Bold,
												parItalic: true,
												parUnderline: objTextSegment.Undeline,
												parSubscript: objTextSegment.Subscript,
												parSuperscript: objTextSegment.Superscript);
											objParagraph.Append(objRun);
											}
										} // foreach loop end
									this.PargraphList.Add(objParagraph);
									}
								else  // there are no cascading tags, just write the text if there are any
									{
									if(objHTMLelement.innerText.Length > 0)
										{
										if(!objHTMLelement.outerHTML.Contains("<P></P>"))
											{
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: objHTMLelement.innerText,
												parContentLayer: this.ContentLayer,
												parItalic: true);
											objParagraph.Append(objRun);
											this.PargraphList.Add(objParagraph);
											}
										}
									} // there are no cascading tags
								} // if(objHTMLelement.innerText != null)
							break;
						//------------------------
						case "SUB":  // Subscript
						//------------------------
							if(objHTMLelement.innerText != null)
								{
								objParagraph = oxmlDocument.Construct_Paragraph(this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel, this.IsTableText);
								if(objHTMLelement.children.length > 0) // check if there are more html tags in the HTMLelement
									{
									// use the DissectHTMLstring method to process the paragraph.
									List<TextSegment> listTextSegments = new List<TextSegment>();
									listTextSegments = TextSegment.DissectHTMLstring(objHTMLelement.innerHTML);
									// Process the list to insert the content into Paragraph List
									foreach(TextSegment objTextSegment in listTextSegments)
										{
										if(objTextSegment.Image) // Check if it is an image
											{
											if(this.IsTableText)
												throw new InvalidRichTextFormatException("Attempted to insert a image into a table.");
											else
												throw new InvalidRichTextFormatException("Rich Text is not suppose to contain an Image");
											}
										else // not an image
											{
											objRun = oxmlDocument.Construct_RunText
												(parText2Write: objTextSegment.Text,
												parContentLayer: this.ContentLayer,
												parBold: objTextSegment.Bold,
												parItalic: objTextSegment.Italic,
												parUnderline: objTextSegment.Undeline,
												parSubscript: true,
												parSuperscript: objTextSegment.Superscript);
											objParagraph.Append(objRun);
											}
										} // foreach loop end
									this.PargraphList.Add(objParagraph);
									}
								else  // there are no cascading tags, just write the text if there are any
									{
									if(objHTMLelement.innerText.Length > 0)
										{
										if(!objHTMLelement.outerHTML.Contains("<P></P>"))
											{
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: objHTMLelement.innerText,
												parContentLayer: this.ContentLayer,
												parSubscript: true);
											objParagraph.Append(objRun);
											this.PargraphList.Add(objParagraph);
											}
										}
									} // there are no cascading tags
								} // if(objHTMLelement.innerText != null)
							break;
						//------------------------
						case "SUP":  // Superscript
						//------------------------
							if(objHTMLelement.innerText != null)
								{
								objParagraph = oxmlDocument.Construct_Paragraph(this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel, this.IsTableText);
								if(objHTMLelement.children.length > 0) // check if there are more html tags in the HTMLelement
									{
									// use the DissectHTMLstring method to process the paragraph.
									List<TextSegment> listTextSegments = new List<TextSegment>();
									listTextSegments = TextSegment.DissectHTMLstring(objHTMLelement.innerHTML);
									// Process the list to insert the content into Paragraph List
									foreach(TextSegment objTextSegment in listTextSegments)
										{
										if(objTextSegment.Image) // Check if it is an image
											{
											if(this.IsTableText)
												throw new InvalidRichTextFormatException("Attempted to insert a image into a table.");
											else
												throw new InvalidRichTextFormatException("Rich Text is not suppose to contain an Image");
											}
										else // not an image
											{
											objRun = oxmlDocument.Construct_RunText
												(parText2Write: objTextSegment.Text,
												parContentLayer: this.ContentLayer,
												parBold: objTextSegment.Bold,
												parItalic: objTextSegment.Italic,
												parUnderline: objTextSegment.Undeline,
												parSubscript: objTextSegment.Subscript,
												parSuperscript: true);
											objParagraph.Append(objRun);
											}
										} // foreach loop end
									this.PargraphList.Add(objParagraph);
									}
								else  // there are no cascading tags, just write the text if there are any
									{
									if(objHTMLelement.innerText.Length > 0)
										{
										if(!objHTMLelement.outerHTML.Contains("<P></P>"))
											{
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: objHTMLelement.innerText,
												parContentLayer: this.ContentLayer,
												parSuperscript: true);
											objParagraph.Append(objRun);
											this.PargraphList.Add(objParagraph);
											}
										}
									} // there are no cascading tags
								} // if(objHTMLelement.innerText != null)
							break;
						//-----------------------------------------------------
						case "SPAN":   // Underline is embedded in the Span tag
						//-----------------------------------------------------
							if(objHTMLelement.innerText != null)
								{
								Console.WriteLine("innerText.Length: {0} - [{1}]", objHTMLelement.innerText.Length, objHTMLelement.innerText);
								if(objHTMLelement.id.Contains("rangepaste"))
									{
									Console.WriteLine("Tag: SPAN - rangepaste ignored [{0}]", objHTMLelement.innerText);
									}
								else if(objHTMLelement.style.color != null)
									{
									Console.WriteLine("Tag: SPAN Style COLOR ignored [{0}]", objHTMLelement.innerText);
									}
								else if(objHTMLelement.id.Contains("underline"))
									{
									objParagraph = oxmlDocument.Construct_Paragraph(
										parBodyTextLevel: this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel,
										parIsTableParagraph: this.IsTableText);
									if(objHTMLelement.children.length > 0) // check if there are more html tags in the HTMLelement
										{
										// use the DissectHTMLstring method to process the paragraph.
										List<TextSegment> listTextSegments = new List<TextSegment>();
										listTextSegments = TextSegment.DissectHTMLstring(objHTMLelement.innerHTML);
										// Process the list to insert the content into Paragraph List
										foreach(TextSegment objTextSegment in listTextSegments)
											{
											if(objTextSegment.Image) // Check if it is an image
												{
												if(this.IsTableText)
													throw new InvalidRichTextFormatException("Attempted to insert a image into a table.");
												else
													throw new InvalidRichTextFormatException("Rich Text is not suppose to contain an Image");
												}
											else // not an image
												{
												objRun = oxmlDocument.Construct_RunText
													(parText2Write: objTextSegment.Text,
													parContentLayer: this.ContentLayer,
													parBold: objTextSegment.Bold,
													parItalic: objTextSegment.Italic,
													parUnderline: true,
													parSubscript: objTextSegment.Subscript,
													parSuperscript: objTextSegment.Superscript);
												objParagraph.Append(objRun);
												}
											} // foreach loop end
										this.PargraphList.Add(objParagraph);
										}
									else  // there are no cascading tags, just write the text if there are any
										{
										if(objHTMLelement.innerText.Length > 0)
											{
											if(!objHTMLelement.outerHTML.Contains("<P></P>"))
												{
												objRun = oxmlDocument.Construct_RunText(
													parText2Write: objHTMLelement.innerText,
													parContentLayer: this.ContentLayer,
													parUnderline: true);
												objParagraph.Append(objRun);
												this.PargraphList.Add(objParagraph);
												}
											}
										} // there are no cascading tags
									} //if(objHTMLelement.id.Contains("underline"))
								}
							break;
						//--------------------------
						case "H1":     // Heading 1
						case "H2":     // Heading 2
						case "H3":     // Heading 3
						case "H4":     // Heading 4

							//Console.WriteLine("Tag: H1\n{0}", objHTMLelement.outerHTML);
							if (this.IsTableText)
								{
								this.AdditionalHierarchicalLevel = 0;
								objParagraph = oxmlDocument.Construct_Heading(
									parHeadingLevel: this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: objHTMLelement.innerText, 
									parContentLayer: this.ContentLayer,
									parBold: true);
								}
							else
								{
								this.AdditionalHierarchicalLevel = Convert.ToInt16(objHTMLelement.tagName.Substring(1, 1));
								objParagraph = oxmlDocument.Construct_Heading(
									parHeadingLevel: this.DocumentHierachyLevel + this.AdditionalHierarchicalLevel);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: objHTMLelement.innerText, 
									parContentLayer: this.ContentLayer);
								}
								
							objParagraph.Append(objRun);
							this.PargraphList.Add(objParagraph);
							break;
						} // end switch
					} // end foreach loop
				}//Try
               catch(InvalidRichTextFormatException exc)
				{
				Console.WriteLine("Exception: {0}", exc.Message);
				throw new InvalidRichTextFormatException(exc.Message, exc);
				}
			catch(Exception exc)
				{
				Console.WriteLine("EXCEPTION ERROR: {0} - {1} - {2} - {3}", exc.HResult, exc.Source, exc.Message, exc.Data);
				}

			return true;
			}
	}
