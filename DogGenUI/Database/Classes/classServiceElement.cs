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
		#region Variables
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		private string _Title;
		private double? _SortOrder;
		[Index]
		private ServiceProduct _ServiceProduct;
		private int? _ServiceProductIDsp;
		private string _ISDheading;
		private string _ISDdescription;
		private string _Objectives;
		private string _KeyClientBenefits;
		private string _KeyClientAdvantages;
		private string _KeyDDbenefits;
		private string _KeyPerformanceIndicators;
		private string _CriticalSuccessFactors;
		private string _ProcessLink;
		private string _ContentLayer;
		[Index]
		private int? _ContentPredecessorElementIDsp;
		private string _ContentStatus;
		#endregion

		#region Properties
		public int IDsp {
			get { return this._IDsp; }
			set { Update(); this._IDsp = value; }
			}
		public string Title {
			get { return this._Title; }
			set { Update(); this._Title = value; }
			}
		public double? SortOrder {
			get {return this._SortOrder;}
			set {Update(); this._SortOrder = value;}
			}
		public ServiceProduct ServiceProduct {
			get { return this._ServiceProduct; }
			set { Update(); this._ServiceProduct = value; }
			}
		public int? ServiceProductIDsp {
			get {return this._ServiceProductIDsp;}
			set {Update();this._ServiceProductIDsp = value;}
			}
		public string ISDheading {
			get { return this._ISDheading; }
			set {Update();this._ISDheading = value;}
			}
		public string ISDdescription {
			get {return this._ISDdescription;}
			set {Update();this._ISDdescription = value;}
			}
		public string Objectives {
			get {return this._Objectives;}
			set {Update();this._Objectives = value;}
			}
		public string KeyClientBenefits {
			get {return this._KeyClientBenefits;}
			set {Update();this._KeyClientBenefits = value;}
			}
		public string KeyClientAdvantages {
			get {return this._KeyClientAdvantages;}
			set {Update();this._KeyClientAdvantages = value;}
			}
		public string KeyDDbenefits {
			get {return this._KeyDDbenefits;}
			set {Update();this._KeyDDbenefits = value;}
			}
		public string KeyPerformanceIndicators {
			get {return this._KeyPerformanceIndicators;}
			set {Update();this._KeyPerformanceIndicators = value;}
			}
		public string CriticalSuccessFactors {
			get {return this._CriticalSuccessFactors;}
			set {Update();this._CriticalSuccessFactors = value;}
			}
		public string ProcessLink {
			get {return this._ProcessLink;}
			set {Update();this._ProcessLink = value;}
			}
		public string ContentLayer {
			get {return this._ContentLayer;}
			set {Update();this._ContentLayer = value;
				}
			}
		public int? ContentPredecessorElementIDsp {
			get {return this._ContentPredecessorElementIDsp;}
			set {Update();this._ContentPredecessorElementIDsp = value;}
			}
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
			ServiceElement newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
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
					newEntry.ServiceProduct = ServiceProduct.Read(parIDsp: parServiceProductIDsp);
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
					return true;
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database persisting Service Product ### - {0} - {1}", exc.HResult, exc.Message);
				return false;
				}
			}

		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>DataStatus object is retrieved if it exist, else null is retured.</returns>
		public static ServiceElement Read(int parIDsp)
			{
			ServiceElement result = new ServiceElement();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<ServiceElement>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					}
				}
			catch (Exception exc)
				{
				result = null;
				Console.WriteLine("### Exception Database reading ServiceElement [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
				}
			return result;
			}

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
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Return all Products if no product is specified
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
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all ServiceElement ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}
		#endregion
		}
	}
