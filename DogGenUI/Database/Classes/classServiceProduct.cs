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
			set { UpdateNonIndexField(); this._Title = value; }
			}

		private int? _ServiceFamilyIDsp;
		public int? ServiceFamilyIDsp {
			get {return this._ServiceFamilyIDsp;}
			set {UpdateNonIndexField();this._ServiceFamilyIDsp = value;}
			}

		private string _ISDheading;
		public string ISDheading {
			get { return this._ISDheading; }
			set {UpdateNonIndexField();this._ISDheading = value;}
			}

		private string _ISDdescription;
		public string ISDdescription {
			get {return this._ISDdescription;}
			set {UpdateNonIndexField();this._ISDdescription = value;}
			}

		private string _KeyDDbenefits;
		public string KeyDDbenefits {
			get {return this._KeyDDbenefits;}
			set {UpdateNonIndexField();this._KeyDDbenefits = value;}
			}

		private string _KeyClientBenefits;
		public string KeyClientBenefits {
			get {return this._KeyClientBenefits;}
			set {UpdateNonIndexField();this._KeyClientBenefits = value;}
			}

		private string _CSDheading;
		public string CSDheading {
			get {return this._CSDheading;}
			set {UpdateNonIndexField();this._CSDheading = value;}
			}

		private string _CSDdescription;
		public string CSDdescription {
			get {return this._CSDdescription;}
			set {UpdateNonIndexField();this._CSDdescription = value;}
			}

		private string _SOWheading;
		public string SOWheading {
			get {return this._SOWheading;}
			set {UpdateNonIndexField();this._SOWheading = value;}
			}

		private string _SOWdescription;
		public string SOWdescription {
			get {return this._SOWdescription;}
			set {UpdateNonIndexField();this._SOWdescription = value;}
			}

		private double? _PlannedElements;
		public double? PlannedElements {
			get {return this._PlannedElements;}
			set {UpdateNonIndexField();this._PlannedElements = value;}
			}

		private double? _PlannedFeatures;
		public double? PlannedFeatures {
			get {return this._PlannedFeatures;}
			set {UpdateNonIndexField();this._PlannedFeatures = value; }
			}

		private double? _PlannedDeliverables;
		public double? PlannedDeliverables {
			get {return this._PlannedDeliverables;}
			set {UpdateNonIndexField();this._PlannedDeliverables = value;}
			}

		private double? _PlannedMeetings;
		public double? PlannedServiceLevels {
			get {return this._PlannedServiceLevels;}
			set {Update();this._PlannedServiceLevels = value;}
			}

		private double? _PlannedReports;
		public double? PlannedMeetings {
			get {return this._PlannedMeetings;}
			set {UpdateNonIndexField();this._PlannedMeetings = value;}
			}

		private double? _PlannedServiceLevels;
		public double? PlannedReports {
			get {return this._PlannedReports;}
			set {UpdateNonIndexField();this._PlannedReports = value;}
			}

		private double? _PlannedActivities;
		public double? PlannedActivities {
			get {return this._PlannedActivities;}
			set {UpdateNonIndexField();this._PlannedActivities = value;}
			}

		private double? _PlannedActivityEffortDrivers;
		public double? PlannedActivityEffortDrivers {
			get {return this._PlannedActivityEffortDrivers;}
			set {UpdateNonIndexField();this._PlannedActivityEffortDrivers = value;}
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
			bool result = false;
			ServiceProduct newEntry;
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
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
					result = true;
					}
				catch (Exception exc)
					{
					Console.WriteLine("### Exception Database persisting ServiceProduct ### - {0} - {1}", exc.HResult, exc.Message);
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
		/// <returns>the object is retrieved if it exist, else null is retured.</returns>
		public static ServiceProduct Read(int parIDsp)
			{
			ServiceProduct result = new ServiceProduct();
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();
					result = (from thisEntry in dbSession.AllObjects<ServiceProduct>()
						  where thisEntry.IDsp == parIDsp
						  select thisEntry).FirstOrDefault();
					dbSession.Commit();
					}
				catch (Exception exc)
					{
					result = null;
					Console.WriteLine("### Exception Database reading ServiceProduct [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
					dbSession.Abort();
					}
				}
			return result;
			}

		//---g
		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database. Specify a List of intergers containing the SharePoint ID values of all the product objects 
		/// that need to be retrived and added to the list.
		/// </summary>
		/// <param name="parIDs">pass a List<int> of all the IDsp (SharePoint ID) that need to be retrieved and returned.
		/// If all ServiceProducts must be retrieve, pass an empty List (with count = 0) to return all objects.</int> </param>
		/// <returns>a List of ServiceProduct objects is retrurned.</returns>
		public static List<ServiceProduct> ReadAll(List<int> parIDs)
			{
			List<ServiceProduct> results = new List<ServiceProduct>();
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
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
					dbSession.Commit();
					}
				catch (Exception exc)
					{
					Console.WriteLine("### Exception Database reading all ServiceProduct ### - {0} - {1}", exc.HResult, exc.Message);
					dbSession.Abort();
					}
				}
			return results;
			}
		#endregion
		}
	}
