using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Indexing;
using VelocityDb.Session;

namespace DocGeneratorCore.Database.Classes
	{
	public class ServiceElement : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a ServiceElement as mapped to the SharePoint List named ServiceElements.
		/// </summary>

		#region Properties
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		public int IDsp {
			get { return this._IDsp; }
			set { Update(); this._IDsp = value; }
			}

		private double? _SortOrder;
		private string _Title;
		public string Title {
			get { return this._Title; }
			set { UpdateNonIndexField(); this._Title = value; }
			}

		private int? _ServiceProductIDsp;
		public double? SortOrder {
			get {return this._SortOrder;}
			set {UpdateNonIndexField(); this._SortOrder = value;}
			}

		private string _ISDheading;
		public int? ServiceProductIDsp {
			get {return this._ServiceProductIDsp;}
			set {UpdateNonIndexField();this._ServiceProductIDsp = value;}
			}

		private string _ISDdescription;
		public string ISDheading {
			get { return this._ISDheading; }
			set {UpdateNonIndexField();this._ISDheading = value;}
			}

		private string _Objectives;
		public string ISDdescription {
			get {return this._ISDdescription;}
			set {UpdateNonIndexField();this._ISDdescription = value;}
			}

		private string _KeyClientBenefits;
		public string Objectives {
			get {return this._Objectives;}
			set {UpdateNonIndexField();this._Objectives = value;}
			}

		private string _KeyClientAdvantages;
		public string KeyClientBenefits {
			get {return this._KeyClientBenefits;}
			set {UpdateNonIndexField();this._KeyClientBenefits = value;}
			}

		private string _KeyDDbenefits;
		public string KeyClientAdvantages {
			get {return this._KeyClientAdvantages;}
			set {UpdateNonIndexField();this._KeyClientAdvantages = value;}
			}

		private string _KeyPerformanceIndicators;
		public string KeyDDbenefits {
			get {return this._KeyDDbenefits;}
			set {UpdateNonIndexField();this._KeyDDbenefits = value;}
			}

		private string _CriticalSuccessFactors;
		public string KeyPerformanceIndicators {
			get {return this._KeyPerformanceIndicators;}
			set {UpdateNonIndexField();this._KeyPerformanceIndicators = value;}
			}

		private string _ProcessLink;
		public string CriticalSuccessFactors {
			get {return this._CriticalSuccessFactors;}
			set {UpdateNonIndexField();this._CriticalSuccessFactors = value;}
			}

		private string _ContentLayer;
		public string ProcessLink {
			get {return this._ProcessLink;}
			set {UpdateNonIndexField();this._ProcessLink = value;}
			}
		public string ContentLayer {
			get {return this._ContentLayer;}
			set {UpdateNonIndexField();this._ContentLayer = value;
				}
			}

		private int? _ContentPredecessorElementIDsp;
		public int? ContentPredecessorElementIDsp {
			get {return this._ContentPredecessorElementIDsp;}
			set {UpdateNonIndexField();this._ContentPredecessorElementIDsp = value;}
			}

		private string _ContentStatus;
		public string ContentStatus {
			get {return this._ContentStatus;}
			set {UpdateNonIndexField();this._ContentStatus = value;}
			}
		#endregion

			#region Methods
			
		//---g
		//++Store
		/// <summary>
		/// Store/Save a new Object in the database, use the same Store method for New and Updates.
		/// </summary>
		public static bool Store(
			int parIDsp,
			string parTitle,
			double? parSortOrder,
			int parServiceProductIDsp,
			string parISDheading,
			string parISDdescription,
			string parKeyDDbenefits,
			string parKeyClientBenefits,
			string parKeyClientAdvantages,
			string parKeyPerformanceIndicators,
			string parCriticalSuccessFactors,
			string parProcessLink,
			string parContentLayer,
			int? parContentPredecessorElementIDsp,
			string parContentStatus
			)
			{
			bool result = false;
			ServiceElement newEntry;
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<ServiceElement>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();
					if (newEntry == null)
						newEntry = new ServiceElement();
					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.SortOrder = parSortOrder;
					newEntry.ServiceProductIDsp = parServiceProductIDsp;
					//-|Use the **ServicePoroductIDsp** to retrieve the ServiceProduct Object instance.
					newEntry._ServiceProductIDsp = parServiceProductIDsp;
					newEntry.KeyDDbenefits = parKeyDDbenefits;
					newEntry.KeyClientBenefits = parKeyClientBenefits;
					newEntry.ISDheading = parISDheading;
					newEntry.ISDdescription = parISDdescription;
					newEntry.KeyClientAdvantages = parKeyClientAdvantages;
					newEntry.KeyClientBenefits = parKeyClientBenefits;
					newEntry.KeyDDbenefits = parKeyDDbenefits;
					newEntry.KeyPerformanceIndicators = parKeyPerformanceIndicators;
					newEntry.CriticalSuccessFactors = parCriticalSuccessFactors;
					newEntry.ProcessLink = parProcessLink;
					newEntry.ContentLayer = parContentLayer;
					newEntry.ContentPredecessorElementIDsp = parContentPredecessorElementIDsp;
					newEntry.ContentStatus = parContentStatus;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					result = true;
					}
				catch (Exception exc)
					{
					Console.WriteLine("### Exception Database persisting Service Product ### - {0} - {1}", exc.HResult, exc.Message);
					dbSession.Abort();
					}
				}
			return result;
			}

		//---g
		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>object is returned if it exist, else null is retured.</returns>
		public static ServiceElement Read(int parIDsp)
			{
			ServiceElement result = new ServiceElement();
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();
					result = (from thisEntry in dbSession.AllObjects<ServiceElement>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					dbSession.Commit();
					}
				catch (Exception exc)
					{
					result = null;
					Console.WriteLine("### Exception Database reading ServiceElement [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
					dbSession.Abort();
					}
				}
			return result;
			}

		//---g
		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database. 
		/// Specify a List of intergers containing the SharePoint ID values of all the Service Element objects
		/// that need to be retrived and added to the list.
		/// </summary>
		/// <param name="parIDs">pass a List<int> of all the IDsp (SharePoint ID) that need to be retrieved and returned.
		/// If all ServiceElements must be retrieve, pass an empty List (with count = 0) to return all objects.</int> </param>
		/// <returns>a List<ServiceElement> objects are retrurned.</returns>
		public static List<ServiceElement> ReadAll(List<int> parIDs)
			{
			List<ServiceElement> results = new List<ServiceElement>();
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();
					//-|Return all object if no product is specified
					if (parIDs.Count == 0)
						{
						foreach (ServiceElement entry in dbSession.AllObjects<ServiceElement>())
							{
							results.Add(entry);
							}
						}
					else //-| Specific entries were specified.
						{
						foreach (int item in parIDs)
							{
							ServiceElement entry = new ServiceElement();
							entry = (from thisEntry in dbSession.AllObjects<ServiceElement>()
									 where thisEntry.IDsp == item
									 select thisEntry).FirstOrDefault();
							results.Add(entry);
							}
						}
					dbSession.Commit();
					}
				catch (Exception exc)
					{
					Console.WriteLine("### Exception Database reading all ServiceElement ### - {0} - {1}", exc.HResult, exc.Message);
					dbSession.Abort();
					}
				}
			return results;
			}
		//===G
		//++ReadAllForProduct
		/// <summary>
		/// Read/retrieve all the Service Element Entries for a specific Service Product. 
		/// Specify a interger containing the SharePoint ID values of all the Service Product for which the Elements must be returned.
		/// </summary>
		/// <param name="parIDsp">pass an integer of the Service Product IDsp (SharePoint ID).</param>
		/// <returns>a List of ServiceElement objects are retrurned.</returns>
		public static List<ServiceElement> ReadAllForProduct(int parIDsp)
			{
			List<ServiceElement> results = new List<ServiceElement>();
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();

					//-|Return all Service Element objects for the specified Service Product
	
					foreach (ServiceElement entry in (from theEntry in dbSession.AllObjects<ServiceElement>()
													where theEntry.ServiceProductIDsp == parIDsp
													select theEntry))
						{
						results.Add(entry);
						}

					dbSession.Commit();
					}
				catch (Exception exc)
					{
					Console.WriteLine("### Exception Database reading all ServiceElement ### - {0} - {1}", exc.HResult, exc.Message);
					dbSession.Abort();
					}
				}
			return results;
			}

		#endregion
		}
	}
