using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Collection;
using VelocityDb.Collection.BTree;
using VelocityDb.Indexing;
using VelocityDb.Session;
using VelocityDb.TypeInfo;
using VelocityDBExtensions;

namespace DocGeneratorCore.Database.Classes
	{
	public class ServiceLevel : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a ServiceLevel as mapped to the SharePoint List named ServiceLevels.
		/// </summary>
		#region Variables
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		private string _Title;
		private ServiceLevelCategory _Category;
		[Index]
		private double? _SortOrder;
		[Index]
		private int? _ServiceProductIDsp;
		private string _ISDheading;
		private string _ISDdescription;
		private string _ISDsummary;
		private string _CSDheading;
		private string _CSDdescription;
		private string _CSDsummary;
		private string _SOWheading;
		private string _SOWdescription;
		private string _SOWsummary;
		private string _Measurement;
		private string _MeasurementInterval;
		private string _ReportingInterval;
		private string _CalculationMethod;
		private string _CaculationFormula;
		private string _ServiceHours;
		private string _BasicConditions;
		private List<ServiceLevelTarget> _PerformanceThresholds;
		private List<ServiceLevelTarget> _PerformanceTargets;
		private string _ContentStatus;
		#endregion

		#region Properties
		public int IDsp {
			get { return this._IDsp; }
			set { Update(); this._IDsp = value; }
			}
		public string Title {
			get { return this._Title; }
			set { UpdateNonIndexField(); this._Title = value; }
			}
		public ServiceLevelCategory Category {
			get { return this._Category;}
			set { Update(); this._Category = value;}
			}
		public int? ServiceProductIDsp {
			get { return this._ServiceProductIDsp;}
			set { Update();this._ServiceProductIDsp = value;}
			}
		public string ISDheading {
			get { return this._ISDheading; }
			set {UpdateNonIndexField();this._ISDheading = value;}
			}
		public string ISDdescription {
			get { return this._ISDdescription;}
			set { UpdateNonIndexField();this._ISDdescription = value;}
			}
		public string ISDsummary {
			get { return this._ISDsummary; }
			set { UpdateNonIndexField(); this._ISDsummary = value; }
			}
		public string CSDheading {
			get { return this._CSDheading;}
			set { UpdateNonIndexField();this._CSDheading = value;}
			}
		public string CSDdescription {
			get { return this._CSDdescription; }
			set { UpdateNonIndexField();this._CSDdescription = value; }
			}
		public string CSDsummary {
			get { return this._CSDsummary; }
			set { UpdateNonIndexField();this._CSDsummary = value; }
			}
		public string SOWheading {
			get { return this._SOWheading; }
			set { UpdateNonIndexField();this._SOWheading = value; }
			}
		public string SOWdescription {
			get { return this._SOWdescription; }
			set { UpdateNonIndexField();this._SOWdescription = value; }
			}
		public string SOWsummary {
			get { return this._SOWsummary; }
			set { UpdateNonIndexField();this._SOWsummary = value; }
			}
		public string Measurement {
			get { return this._Measurement; }
			set { UpdateNonIndexField();this._Measurement = value; }
			}
		public string MeasurementInterval {
			get { return this._MeasurementInterval; }
			set { UpdateNonIndexField(); this._MeasurementInterval = value; }
			}
		public string ReportingInterval {
			get { return this._ReportingInterval; }
			set { UpdateNonIndexField(); this._ReportingInterval = value; }
			}
		public string CalculationMethod {
			get { return this._CalculationMethod; }
			set { UpdateNonIndexField(); this._CalculationMethod = value; }
			}
		public string CalculationFormula {
			get { return this._CaculationFormula; }
			set { UpdateNonIndexField(); this._CaculationFormula = value; }
			}
		public string ServiceHours {
			get { return this._ServiceHours; }
			set { UpdateNonIndexField(); this._ServiceHours = value; }
			}
		public string BasicConditions {
			get { return this._BasicConditions;}
			set { UpdateNonIndexField(); this._BasicConditions = value; }
			}
		public List<ServiceLevelTarget> PerformanceTargets {
			get { return this._PerformanceTargets; }
			set { UpdateNonIndexField(); this._PerformanceTargets = value; }
			} 
		public List<ServiceLevelTarget> PerformanceThresholds {
			get { return this._PerformanceThresholds; } 
			set { UpdateNonIndexField(); this._PerformanceThresholds = value; }
			}
		public string ContentStatus {
			get {return this._ContentStatus;}
			set {Update();this._ContentStatus = value;}
			}
		#endregion

		//===g
		#region Methods
		//++Store
		/// <summary>
		/// Store/Save a new Object in the database, use the same Store method for New and Updates.
		/// </summary>
		public static bool Store(
			int parIDsp,
			string parTitle,
			ServiceLevelCategory parCategory,
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
			string parContentStatus
			)
			{
			ServiceLevel newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<ServiceLevel>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();
					if (newEntry == null)
						newEntry = new ServiceLevel();

					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.Category = parCategory;
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
		/// <returns>ServiceLevel object is retrieved if it exist, else null is retured.</returns>
		public static ServiceLevel Read(int parIDsp)
			{
			ServiceLevel result = new ServiceLevel();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
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
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
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
