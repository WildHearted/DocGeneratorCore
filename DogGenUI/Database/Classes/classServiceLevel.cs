using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Indexing;
using VelocityDb.Session;


namespace DocGeneratorCore.Database.Classes
	{
	public class ServiceLevel : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a ServiceLevel as mapped to the SharePoint List named ServiceLevels.
		/// </summary>

		#region Properties
		[Index]
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

		private int _CategoryIDsp;
		public int CategoryIDsp {
			get { return this._CategoryIDsp;}
			set { UpdateNonIndexField(); this._CategoryIDsp = value;}
			}

		private int? _ServiceProductIDsp;
		public int? ServiceProductIDsp {
			get { return this._ServiceProductIDsp;}
			set { UpdateNonIndexField();this._ServiceProductIDsp = value;}
			}

		private string _ISDheading;
		public string ISDheading {
			get { return this._ISDheading; }
			set {UpdateNonIndexField();this._ISDheading = value;}
			}

		private string _ISDdescription;
		public string ISDdescription {
			get { return this._ISDdescription;}
			set { UpdateNonIndexField();this._ISDdescription = value;}
			}

		private string _ISDsummary;
		public string ISDsummary {
			get { return this._ISDsummary; }
			set { UpdateNonIndexField(); this._ISDsummary = value; }
			}

		private string _CSDheading;
		public string CSDheading {
			get { return this._CSDheading;}
			set { UpdateNonIndexField();this._CSDheading = value;}
			}
		private string _CSDdescription;
		public string CSDdescription {
			get { return this._CSDdescription; }
			set { UpdateNonIndexField();this._CSDdescription = value; }
			}

		private string _CSDsummary;
		public string CSDsummary {
			get { return this._CSDsummary; }
			set { UpdateNonIndexField();this._CSDsummary = value; }
			}

		private string _SOWheading;
		public string SOWheading {
			get { return this._SOWheading; }
			set { UpdateNonIndexField();this._SOWheading = value; }
			}

		private string _SOWdescription;
		public string SOWdescription {
			get { return this._SOWdescription; }
			set { UpdateNonIndexField();this._SOWdescription = value; }
			}

		private string _SOWsummary;
		public string SOWsummary {
			get { return this._SOWsummary; }
			set { UpdateNonIndexField();this._SOWsummary = value; }
			}

		private string _Measurement;
		public string Measurement {
			get { return this._Measurement; }
			set { UpdateNonIndexField();this._Measurement = value; }
			}

		private string _MeasurementInterval;
		public string MeasurementInterval {
			get { return this._MeasurementInterval; }
			set { UpdateNonIndexField(); this._MeasurementInterval = value; }
			}

		private string _ReportingInterval;
		public string ReportingInterval {
			get { return this._ReportingInterval; }
			set { UpdateNonIndexField(); this._ReportingInterval = value; }
			}

		private string _CalculationMethod;
		public string CalculationMethod {
			get { return this._CalculationMethod; }
			set { UpdateNonIndexField(); this._CalculationMethod = value; }
			}

		private string _CaculationFormula;
		public string CalculationFormula {
			get { return this._CaculationFormula; }
			set { UpdateNonIndexField(); this._CaculationFormula = value; }
			}

		private string _ServiceHours;
		public string ServiceHours {
			get { return this._ServiceHours; }
			set { UpdateNonIndexField(); this._ServiceHours = value; }
			}

		private string _BasicConditions;
		public string BasicConditions {
			get { return this._BasicConditions;}
			set { UpdateNonIndexField(); this._BasicConditions = value; }
			}

		private List<ServiceLevelTarget> _PerformanceTargets;
		public List<ServiceLevelTarget> PerformanceTargets {
			get { return this._PerformanceTargets; }
			set { UpdateNonIndexField(); this._PerformanceTargets = value; }
			}

		private List<ServiceLevelTarget> _PerformanceThresholds;
		public List<ServiceLevelTarget> PerformanceThresholds {
			get { return this._PerformanceThresholds; } 
			set { UpdateNonIndexField(); this._PerformanceThresholds = value; }
			}

		private string _ContentStatus;
		public string ContentStatus {
			get {return this._ContentStatus;}
			set {UpdateNonIndexField();this._ContentStatus = value;}
			}
		#endregion

		//===g
		#region Methods
		//---g
		//++Store
		/// <summary>
		/// Store/Save a new Object in the database, use the same Store method for New and Updates.
		/// </summary>
		public static bool Store(
			int parIDsp,
			string parTitle,
			int parCategoryIDsp,
			int parServiceProductIDsp,
			string parISDheading,
			string parISDdescription,
			string parISDsummary,
			string parCSDheading,
			string parCSDdescription,
			string parCSDsummary,
			string parSOWheading,
			string parSOWdescription,
			string parSOWsummary,
			string parMeasurement,
			string parMeasurementInterval,
			string parReportingInterval,
			string parCalculationMethod,
			string parCalculationFormula,
			string parServiceHours,
			string parBasicConditions,
			List<ServiceLevelTarget> parPerformanceTargets,
			List<ServiceLevelTarget> parPerformanceThresholds,
			string parContentStatus)
			{
			bool result = false;
			ServiceLevel newEntry;
			using (SessionNoServerShared dbSession = new SessionNoServerShared(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<ServiceLevel>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();
					if (newEntry == null)
						newEntry = new ServiceLevel();

					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.CategoryIDsp = parCategoryIDsp;
					newEntry.ServiceProductIDsp = parServiceProductIDsp;
					newEntry.ISDheading = parISDheading;
					newEntry.ISDdescription = parISDdescription;
					newEntry.ISDsummary = parISDsummary;
					newEntry.CSDheading = parCSDheading;
					newEntry.CSDdescription = parCSDdescription;
					newEntry.CSDsummary = parCSDsummary;
					newEntry.SOWheading = parSOWheading;
					newEntry.SOWdescription = parSOWdescription;
					newEntry.SOWsummary = parSOWsummary;
					newEntry.Measurement = parMeasurement;
					newEntry.MeasurementInterval = parMeasurementInterval;
					newEntry.ReportingInterval = parReportingInterval;
					newEntry.CalculationMethod = parCalculationMethod;
					newEntry.CalculationFormula = parCalculationFormula;
					newEntry.ServiceHours = parServiceHours;
					newEntry.BasicConditions = parBasicConditions;
					newEntry.PerformanceTargets = parPerformanceTargets;
					newEntry.PerformanceThresholds = parPerformanceThresholds;
					newEntry.ContentStatus = parContentStatus;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					result = true;
					}
				catch (Exception exc)
					{
					Console.WriteLine("### Exception Database persisting Service Product ### - {0} - {1}", exc.HResult, exc.Message);
					result = false;
					dbSession.Abort();
					}
				}
			return result;
			}

		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>ServiceLevel object is retrieved if it exist, else null is retured.</returns>
		public static ServiceLevel Read(int parIDsp)
			{
			ServiceLevel result = new ServiceLevel();
			try
				{
				using (SessionNoServerShared dbSession = new SessionNoServerShared(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<ServiceLevel>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					}
				}
			catch (Exception exc)
				{
				result = null;
				Console.WriteLine("### Exception Database reading ServiceLevel [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
				}
			return result;
			}

		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database. 
		/// Specify a List of intergers containing the SharePoint ID values of all the ServiceLevel objects
		/// that need to be retrived and added to the list.
		/// </summary>
		/// <param name="parIDs">pass a List<int> of all the IDsp (SharePoint ID) that need to be retrieved and returned.
		/// If all ServiceLevel must be retrieve, pass an empty List (with count = 0) to return all objects.</int> </param>
		/// <returns>a List<Deliverrable> objects are retrurned.</returns>
		public static List<ServiceLevel> ReadAll(List<int> parIDs)
			{
			List<ServiceLevel> results = new List<ServiceLevel>();
			try
				{
				using (SessionNoServerShared dbSession = new SessionNoServerShared(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Return all Products if no product is specified
					if (parIDs.Count == 0)
						{
						foreach (ServiceLevel entry in dbSession.AllObjects<ServiceLevel>())
							{
							results.Add(entry);
							}
						}
					else //-| Specific entries were specified.
						{
						foreach (int item in parIDs)
							{
							ServiceLevel entry = new ServiceLevel();
							entry = (from thisEntry in dbSession.AllObjects<ServiceLevel>()
									 where thisEntry.IDsp == item
									 select thisEntry).FirstOrDefault();
							results.Add(entry);
							}
						}
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all ServiceLevel ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}
		#endregion
		}
	}
