using System;
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


		private void btnButton2_Click(object sender, EventArgs e)
			{
			IEnumerable<Comic> comics = Comic.BuildCatalogue();
			// declare an instance called comics object as an IEnumerable (collection) based on the values returned from calling the Comic.BuildCatalogue method.
			Dictionary<int, decimal> values = Comic.GetPrices();   // declare an instance called values obect as a Dictionaty (collection) based on the values returned from calling the Comic.GetPrices() method 

			// return a collection of comics from the comics object where the values for its issue is > 500 and sort the result in decending order per issue price.
			var MostExpensive = from comic in comics where values [comic.Issue] > 500 orderby values [comic.Issue] descending select comic;

			foreach(Comic comic in MostExpensive)
				Console.WriteLine("{0} is worth {1:c}", comic.Name, values [comic.Issue]);
			}

		private void btnSDDP_Click(object sender, EventArgs e)
			{
			Cursor.Current = Cursors.WaitCursor;
			List<DocumentCollection> docCollectionsToGenerate = new List<DocumentCollection>();
			try
				{
				if(DocumentCollection.GetCollectionsToGenerate(ref docCollectionsToGenerate))
					{
					lblConnect.Text = "There are " + docCollectionsToGenerate.Count + " Document Collections to generate...";
					}
				else
					{
					Console.WriteLine("At this stage there are no Document Collections to generate.");
					}
				}
			catch(InvalidProgramException ex)
				{
				Console.WriteLine("Exception occurred [{0}] \n Inner Exception: {1}", ex.Message, ex.InnerException);
				}
			// Continue here if there are any Document Collections to generate...
			lblConnect.Text = "There are " + docCollectionsToGenerate.Count.ToString() + " Document Collections to generate...\n\n";
			lblConnect.Refresh();
			try
				{
				foreach(DocumentCollection docToGen in docCollectionsToGenerate)
					{
					Console.WriteLine("Ready to generate entry: {0} - {1}", docToGen.ID.ToString(), docToGen.Title);
					lblConnect.Text = "Generating " + docToGen.ID.ToString() + " - " + docToGen.Title+ "...";
					lblConnect.Refresh();
					}

				Console.WriteLine("\n\n{0} Document Collection(s) were Generated.", docCollectionsToGenerate.Count);
				lblConnect.Text = "Document Generation completed for " + docCollectionsToGenerate.Count + " document collections.";
				lblConnect.Refresh();

				}
			catch(Exception ex)	// if the List is empty - nothing to generate
				{
				Console.WriteLine("Sorry, nothing to generate at this stage.");
				Console.WriteLine("Exception Error: {0} occurred and means {1}", ex.Source, ex.Message);
				}
			finally
				{
				if(docCollectionsToGenerate.Count > 0)
					{
					Cursor.Current = Cursors.Default;
					}
				}
			}
		}
	}
