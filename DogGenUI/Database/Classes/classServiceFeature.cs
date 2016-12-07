using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Indexing;
using VelocityDb.Session;

namespace DocGeneratorCore.Database.Classes
	{
	public class ServiceFeature : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a ServiceFeature as mapped to the SharePoint List named ServiceFeatures.
		/// </summary>
		
		#region Properties
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		public int IDsp {
			get { return this._IDsp; }
			set { Update(); this._IDsp = value; }
			}

		private string _Title;
		public string Title {
			get { return this._Title; }
			set { Update(); this._Title = value; }
			}

		private double? _SortOrder;
		public double? SortOrder {
			get {return this._SortOrder;}
			set {Update(); this._SortOrder = value;}
			}

		private int _ServiceProductIDsp;
		public int ServiceProductIDsp {
			get {return this._ServiceProductIDsp;}
			set {Update();this._ServiceProductIDsp = value;}
			}

		private string _CSDheading;
		public string CSDheading {
			get { return this._CSDheading; }
			set {Update();this._CSDheading = value;}
			}

		private string _CSDdescription;
		public string CSDdescription {
			get {return this._CSDdescription;}
			set {Update();this._CSDdescription = value;}
			}

		private string _SOWheading;
		public string SOWheading {
			get {return this._SOWheading;}
			set {Update();this._SOWheading = value;}
			}

		private string _SOWdescription;
		public string SOWdescription {
			get {return this._SOWdescription;}
			set {Update();this._SOWdescription = value;}
			}

		private string _ContentLayer;
		public string ContentLayer {
			get {return this._ContentLayer;}
			set {Update();this._ContentLayer = value; }
			}

		private int? _ContentPredecessorFeatureIDsp;
		public int? ContentPredecessorFeatureIDsp {
			get {return this._ContentPredecessorFeatureIDsp;}
			set {Update();this._ContentPredecessorFeatureIDsp = value;}
			}

		private string _ContentStatus;
		public string ContentStatus {
			get {return this._ContentStatus;}
			set {Update();this._ContentStatus = value;}
			}
		#endregion

			#region Methods
			//++Store
			/// <summary>
			/// Store/Save a new Object in the database, use the same Store method for New and Updates.
			/// </summary>
		public static bool Store(
			int parIDsp,
			string parTitle,
			double? parSortOrder,
			int parServiceProductIDsp,
			string parCSDheading,
			string parCSDdescription,
			string parSOWheading,
			string parSOWdescription,
			string parContentLayer,
			int? parContentPredecessorFeatureIDsp,
			string parContentStatus
			)
			{
			ServiceFeature newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<ServiceFeature>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();
					if (newEntry == null)
						newEntry = new ServiceFeature();
					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.SortOrder = parSortOrder;
					newEntry.ServiceProductIDsp = parServiceProductIDsp;
					newEntry.SOWheading = parSOWheading;
					newEntry.SOWdescription = parSOWdescription;
					newEntry.CSDheading = parCSDheading;
					newEntry.CSDdescription = parCSDdescription;
					newEntry.ContentLayer = parContentLayer;
					newEntry.ContentPredecessorFeatureIDsp = parContentPredecessorFeatureIDsp;
					newEntry.ContentStatus = parContentStatus;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					return true;
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database persisting Service Feature ### - {0} - {1}", exc.HResult, exc.Message);
				return false;
				}
			}

		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>DataStatus object is retrieved if it exist, else null is retured.</returns>
		public static ServiceFeature Read(int parIDsp)
			{
			ServiceFeature result = new ServiceFeature();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<ServiceFeature>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					}
				}
			catch (Exception exc)
				{
				result = null;
				Console.WriteLine("### Exception Database reading Service Feature [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
				}
			return result;
			}

		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database. 
		/// Specify a List of intergers containing the SharePoint ID values of all the Service Feature objects
		/// that need to be retrived and added to the list.
		/// </summary>
		/// <param name="parIDs">pass a List<int> of all the IDsp (SharePoint ID) that need to be retrieved and returned.
		/// If all ServiceFeatures must be retrieve, pass an empty List (with count = 0) to return all objects.</int> </param>
		/// <returns>a List<ServiceFeature> objects are retrurned.</returns>
		public static List<ServiceFeature> ReadAll(List<int> parIDs)
			{
			List<ServiceFeature> results = new List<ServiceFeature>();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Return all Products if no product is specified
					if (parIDs.Count == 0)
						{
						foreach (ServiceFeature entry in dbSession.AllObjects<ServiceFeature>())
							{
							results.Add(entry);
							}
						}
					else //-| Specific entries were specified.
						{
						foreach (int item in parIDs)
							{
							ServiceFeature entry = new ServiceFeature();
							entry = (from thisEntry in dbSession.AllObjects<ServiceFeature>()
									 where thisEntry.IDsp == item
									 select thisEntry).FirstOrDefault();
							results.Add(entry);
							}
						}
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all Service Feature ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}

		//++ReadAllForProduct
		/// <summary>
		/// Read/retrieve all the ServiceFeature objects from the specified ServiceProduct.
		/// </summary>
		/// <param name="parProductIDsp">pass an integer containing the ServiceProduct IDsp (SharePoint ID) for which Service Features must be retrieved.
		/// If a 0(zero) is passed ALL the entries will be returned.</param>
		/// <returns>a List containing all the relevant ServiceFeature objects are retrurned.</returns>
		public static List<ServiceFeature> ReadAllForProduct(int parProductIDsp)
			{
			List<ServiceFeature> results = new List<ServiceFeature>();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Return all Service features if no Service Product is specified
					if (parProductIDsp == 0)
						{
						foreach (ServiceFeature entry in dbSession.AllObjects<ServiceFeature>())
							{
							results.Add(entry);
							}
						}
					else //-| Specific entry was specified.
						{
						foreach (ServiceFeature entry in (from theEntry in dbSession.AllObjects<ServiceFeature>()
														  where theEntry.ServiceProductIDsp == parProductIDsp
														  select theEntry))
							{							
							results.Add(entry);
							}
						}
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all Service Feature ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}

		#endregion
		}
	}
