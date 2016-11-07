using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Indexing;
using VelocityDb.Session;

namespace DocGeneratorCore.Database.Classes
	{
	public class ServiceProduct : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a ServiceProduct as mapped to the SharePoint List named ServiceProduct.
		/// </summary>
		#region Variables
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		private string _Title;
		[Index]
		private ServiceFamily _ServiceFamily;
		private int? _ServiceFamilyIDsp;
		private string _ISDheading;
		private string _ISDdescription;
		private string _KeyDDbenefits;
		private string _KeyClientBenefits;
		private string _CSDheading;
		private string _CSDdescription;
		private string _SOWheading;
		private string _SOWdescription;
		private double? _PlannedElements;
		private double? _PlannedFeatures;
		private double? _PlannedDeliverables;
		private double? _PlannedMeetings;
		private double? _PlannedReports;
		private double? _PlannedServiceLevels;
		private double? _PlannedActivities;
		private double? _PlannedActivityEffortDrivers;
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
		public ServiceFamily ServiceFamily {
			get { return this._ServiceFamily; }
			set { Update(); this._ServiceFamily = value; }
			}
		public int? ServiceFamilyIDsp {
			get {return this._ServiceFamilyIDsp;}
			set {Update();this._ServiceFamilyIDsp = value;}
			}
		public string ISDheading {
			get { return this._ISDheading; }
			set {Update();this._ISDheading = value;}
			}
		public string ISDdescription {
			get {return this._ISDdescription;}
			set {Update();this._ISDdescription = value;}
			}
		public string KeyDDbenefits {
			get {return this._KeyDDbenefits;}
			set {Update();this._KeyDDbenefits = value;}
			}
		public string KeyClientBenefits {
			get {return this._KeyClientBenefits;}
			set {Update();this._KeyClientBenefits = value;}
			}
		public string CSDheading {
			get {return this._CSDheading;}
			set {Update();this._CSDheading = value;}
			}
		public string CSDdescription {
			get {return this._CSDdescription;}
			set {Update();this._CSDdescription = value;}
			}
		public string SOWheading {
			get {return this._SOWheading;}
			set {Update();this._SOWheading = value;}
			}
		public string SOWdescription {
			get {return this._SOWdescription;}
			set {Update();this._SOWdescription = value;}
			}
		public double? PlannedElements {
			get {return this._PlannedElements;}
			set {Update();this._PlannedElements = value;}
			}
		public double? PlannedFeatures {
			get {return this._PlannedFeatures;}
			set {Update();this._PlannedFeatures = value;
				}
			}
		public double? PlannedDeliverables {
			get {return this._PlannedDeliverables;}
			set {Update();this._PlannedDeliverables = value;}
			}
		public double? PlannedServiceLevels {
			get {return this._PlannedServiceLevels;}
			set {Update();this._PlannedServiceLevels = value;}
			}
		public double? PlannedMeetings {
			get {return this._PlannedMeetings;}
			set {Update();this._PlannedMeetings = value;}
			}
		public double? PlannedReports {
			get {return this._PlannedReports;}
			set {Update();this._PlannedReports = value;}
			}
		public double? PlannedActivities {
			get {return this._PlannedActivities;}
			set {Update();this._PlannedActivities = value;}
			}
		public double? PlannedActivityEffortDrivers {
			get {return this._PlannedActivityEffortDrivers;}
			set {Update();this._PlannedActivityEffortDrivers = value;}
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
			int parServiceFamilyIDsp,
			string parISDheading,
			string parISDdescription,
			string parKeyDDbenefits,
			string parKeyClientBenefits,
			string parCSDheading,
			string parCSDdescription,
			string parSOWheading,
			string parSOWdescription,
			double? parPlannedElements,
			double? parPlannedFeatures,
			double? parPlannedDeliverables,
			double? parPlannedReports,
			double? parPlannedMeetings,
			double? parPlannedActivities,
			double? parPlannedServiceLevels,
			double? parPlannedActivityEffortDrivers
			)
			{
			ServiceProduct newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<ServiceProduct>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();
					if (newEntry == null)
						newEntry = new ServiceProduct();
					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.ServiceFamilyIDsp = parServiceFamilyIDsp;
					//-|Use the **ServicePortfolioIDsp** to retrieve the ServicePortfolio Object instance.
					newEntry.ServiceFamily = ServiceFamily.Read(parIDsp: parServiceFamilyIDsp);
					newEntry.KeyDDbenefits = parKeyDDbenefits;
					newEntry.KeyClientBenefits = parKeyClientBenefits;
					newEntry.ISDheading = parISDheading;
					newEntry.ISDdescription = parISDdescription;
					newEntry.CSDheading = parCSDheading;
					newEntry.CSDdescription = parCSDdescription;
					newEntry.SOWheading = parSOWheading;
					newEntry.SOWdescription = parSOWdescription;
					newEntry.PlannedElements = parPlannedElements;
					newEntry.PlannedFeatures = parPlannedFeatures;
					newEntry.PlannedDeliverables = parPlannedDeliverables;
					newEntry.PlannedReports = parPlannedReports;
					newEntry.PlannedMeetings = parPlannedMeetings;
					newEntry.PlannedServiceLevels = parPlannedServiceLevels;
					newEntry.PlannedActivities = parPlannedActivities;
					newEntry.PlannedActivityEffortDrivers = parPlannedActivityEffortDrivers;
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
		public static ServiceProduct Read(int parIDsp)
			{
			ServiceProduct result = new ServiceProduct();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<ServiceProduct>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					}
				}
			catch (Exception exc)
				{
				result = null;
				Console.WriteLine("### Exception Database reading ServiceProduct [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
				}
			return result;
			}

		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database. Specify a List of intergers containing the SharePoint ID values of all the product objects 
		/// that need to be retrived and added to the list.
		/// </summary>
		/// <param name="parIDs">pass a List<int> of all the IDsp (SharePoint ID) that need to be retrieved and returned.
		/// If all ServiceProducts must be retrieve, pass an empty List (with count = 0) to return all objects.</int> </param>
		/// <returns>a List<ServiceProduct> objects is retrurned.</returns>
		public static List<ServiceProduct> ReadAll(List<int> parIDs)
			{
			List<ServiceProduct> results = new List<ServiceProduct>();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Return all Products if no product is specified
					if (parIDs.Count == 0)
						{
						foreach (ServiceProduct entry in dbSession.AllObjects<ServiceProduct>())
							{
							results.Add(entry);
							}
						}
					else //-| Specific entries were specified.
						{
						foreach (int item in parIDs)
							{
							ServiceProduct entry = new ServiceProduct();
							entry = (from thisEntry in dbSession.AllObjects<ServiceProduct>()
									 where thisEntry.IDsp == item
									 select thisEntry).FirstOrDefault();
							results.Add(entry);
							}
						}
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all ServiceProduct ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}
		#endregion
		}
	}
