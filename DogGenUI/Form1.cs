using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Services.Client;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.SharePoint;
using DogGenUI.SDDPServiceReference;
using DogGenUI;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
namespace DogGenUI
	{

	public partial class Form1 : Form
		{
		/// <summary>
		///	Declare the SharePoint connection as a DataContext
		/// </summary>
		private static string websiteURL = "https://teams.dimensiondata.com/sites/ServiceCatalogue";
		DesignAndDeliveryPortfolioDataContext datacontexSDDP = new DesignAndDeliveryPortfolioDataContext(new Uri(websiteURL + "/_vti_bin/listdata.svc"));
		
		public Form1()
			{
			InitializeComponent();
			datacontexSDDP.Credentials = CredentialCache.DefaultCredentials;
			}

		private void btnSDDP_Click(object sender, EventArgs e)
			{
			Cursor.Current = Cursors.WaitCursor;
			string returnResult = "";
			List<DocumentCollection> docCollectionsToGenerate = new List<DocumentCollection>();
			try
				{
				returnResult = DocumentCollection.GetCollectionsToGenerate(ref docCollectionsToGenerate);
				if (returnResult.Substring(0,4) == "Good")
					{
					Console.WriteLine("There are {0} Document Collections to generate.", docCollectionsToGenerate.Count());
					lblConnect.Text = "There are " + docCollectionsToGenerate.Count + " Document Collections to generate...";
					}
				else if(returnResult.Substring(0,5) == "Error")
					{
					Console.WriteLine("\n\n ERROR: There was an error accessing the Document Collections. \n{0}", returnResult);
					lblConnect.Text = "ERROR: There was an error processing the Document Collections.";
					}
				}
			catch(InvalidProgramException ex)
				{
				Console.WriteLine("Exception occurred [{0}] \n Inner Exception: {1}", ex.Message, ex.InnerException);
				}
			// Continue here if there are any Document Collections to generate...
			lblConnect.Refresh();
			try
				{
				if(docCollectionsToGenerate.Count > 0)
					{
					foreach(DocumentCollection docToGen in docCollectionsToGenerate)
						{
						Console.WriteLine("Ready to generate entry: {0} - {1}", docToGen.ID.ToString(), docToGen.Title);
						lblConnect.Text = "Generating " + docToGen.ID.ToString() + " - " + docToGen.Title + "...";
						lblConnect.Refresh();
						}

					Console.WriteLine("\n\n{0} Document Collection(s) were Generated.", docCollectionsToGenerate.Count);
					lblConnect.Text = "Document Generation completed for " + docCollectionsToGenerate.Count + " document collections.";
					lblConnect.Refresh();
					}
				else
					{
					Console.WriteLine("Sorry, nothing to generate at this stage.");
					lblConnect.Text = "Sorry, nothing to generate at this stage.";
					lblConnect.Refresh();
					}
				}
			catch(Exception ex)	// if the List is empty - nothing to generate
				{
				Console.WriteLine("Exception Error: {0} occurred and means {1}", ex.Source, ex.Message);
				lblConnect.Text = "Exception error" + ex.HResult + " - " + ex.Message;
				lblConnect.Refresh();
				}
			finally
				{
				Cursor.Current = Cursors.Default;
				}
			}

		private void btnOpenMSwordDocument(object sender, EventArgs e)
			{
			if(textBoxFileName.Text == null | textBoxFileName.Text.Length == 0)
				{
				string message = "Specify a MS Word document path";
				string caption = "File cannot be empty";
				MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
				textBoxFileName.Focus();
				return;
				}
			else
				{
				// validate if the file exist
				if(File.Exists(textBoxFileName.Text))
					{
					string message = "The document does not exit, please specify an exisiting document path and filename.";
					string caption = "Document not found";
					MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
					}
				else
					{
					string filename = DateTime.Now.ToShortDateString();
					Console.Write("filename: [{0}]", filename);
					filename = filename.Replace("/", "-") + "_" + DateTime.Now.ToShortTimeString();
					Console.Write("filename: [{0}]", filename);
					filename = filename.Replace(":", "-");
					filename = filename.Replace(" ", "_");
					filename = "newDoc_" + filename;
					Console.Write("filename: [{0}]", filename);

					WordprocessingDocument objDocument = WordprocessingDocument.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true);
					// Load the Document Template...
					if(oxmlDocument.LoadDocumentFromTemplate(textBoxFileName.Text, ref objDocument))
						{
						Console.WriteLine("New MS Word document created from template.");
						}
					else // unable to load from the template
						{
						Console.WriteLine("Unable to assign the remplate to the new document");
						}
					}
				}
			}

		private void Form1_Load(object sender, EventArgs e)
			{
			textBoxFileName.Text = "C:\\Users\ben.vandenberg\\Desktop\\AnotherSampleWordDocument.docx";
               }
		}
	}
