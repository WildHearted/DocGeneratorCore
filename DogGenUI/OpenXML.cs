using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;

// Reference sources:
// https://msdn.microsoft.com/en-us/library/office/ff478255.aspx (Baic Open XML Documents)
// https://msdn.microsoft.com/en-us/library/dd469465%28v=office.12%29.aspx (Examples with merging and Presentations)

namespace DogGenUI
	{
	public class oxmlDocument
		{

		public static bool LoadDocumentFromTemplate(string parTemplateURL, ref WordprocessingDocument parOXMLdocument)
			{
			string filename = @"C:\Users\ben.vandenberg\Desktop\AnotherSampleWordDocument.docx";
			//filename = parTemplateURL;
			try
				{
				using(WordprocessingDocument wdTemplate = WordprocessingDocument.Open(filename, true))
					{
					Console.WriteLine("The file {0} was successfully opened as the wdTemplate.");
					parOXMLdocument = wdTemplate;
					return true;
					}
				}
			catch(OpenXmlPackageException exc)
				{
				Console.WriteLine("OpenXmlPackageException Source: {0}", exc.Source);
				parOXMLdocument = null;
				return false;
				}
			catch(ArgumentNullException exc)
				{
				Console.WriteLine("ArgumentNullException Source: {0}", exc.Source);
				parOXMLdocument = null;
				return false;
				}
			}
		}
	class oxmlWorkbook
		{
		}
	}
