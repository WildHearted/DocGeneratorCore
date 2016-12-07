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
	public class Activity : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a Activity as mapped to the SharePoint List named Activitys.
		/// </summary>
		
		#region Properties
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		public int IDsp {
			get { return this._IDsp; }
			set { Update(); this._IDsp = value; }
			}

		private string _Category;
		public string Category {
			get { return this._Category; }
			set { UpdateNonIndexField(); this._Category = value;}
			}

		private string _Title;
		public string Title {
			get { return this._Title; }
			set { UpdateNonIndexField(); this._Title = value; }
			}

		private double? _SortOrder;
		public double? SortOrder {
			get { return this._SortOrder;}
			set { UpdateNonIndexField(); this._SortOrder = value;}
			}

		private List<int?> _ActivityDependenciesIDsp;
		public List<int?> ActivityDepencenciesIDsp {
			get { return this._ActivityDependenciesIDsp;}
			set { UpdateNonIndexField();this._ActivityDependenciesIDsp = value;}
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

		private string _Assumptions;
		public string Inputs {
			get { return this._Inputs; }
			set { UpdateNonIndexField();this._Inputs = value; }
			}

		private string _Inputs;
		public string Outputs {
			get { return this._Outputs; }
			set { UpdateNonIndexField(); this._Outputs = value; }
			}

		private string _Outputs;
		public string Assumptions {
			get { return this._Assumptions; }
			set { UpdateNonIndexField(); this._Assumptions = value; }
			}

		private string _Optionality;
		public string OLA {
			get { return this._Ola; }
			set { UpdateNonIndexField(); this._Ola = value; }
			}

		private string _Ola;
		public string OLAvariations {
			get { return this._OLAvariations; }
			set { UpdateNonIndexField(); this._OLAvariations = value; }
			}

		private string _OLAvariations;
		public string Optionality {
			get { return this._Optionality; }
			set { UpdateNonIndexField(); this._Optionality = value; }
			}

		private string _OwningEntity;
		public string OwningEntity {
			get { return this._OwningEntity; }
			set { UpdateNonIndexField(); this._OwningEntity = value; }
			}

		private List<int?> _RACIaccountables;
		public List<int?> RACIaccountables {
			get { return this._RACIaccountables; }
			set { UpdateNonIndexField(); this._RACIaccountables = value; }
			}

		private List<int?> _RACIresponsibles;
		public List<int?> RACIresponsibles {
			get { return this._RACIresponsibles; }
			set { UpdateNonIndexField(); this._RACIresponsibles = value; }
			}

		private List<int?> _RACIconsulteds;
		public List<int?> RACIconsulteds {
			get { return this._RACIconsulteds; }
			set { UpdateNonIndexField(); this._RACIconsulteds = value; }
			}

		private List<int?> _RACIinformeds;
		public List<int?> RACIinformeds {
			get { return this._RACIinformeds; }
			set { UpdateNonIndexField(); this._RACIinformeds = value; }
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
		/// Store a new or Save an existing Object in the database, use the same Store method for New and Updates.
		/// </summary>
		public static bool Store(
			int parIDsp,
			string parTitle,
			string parCategory,
			double? parSortOrder,
			List<int?> parActivityDependenciesIDsp,
			string parISDheading,
			string parISDdescription,
			string parISDsummary,
			string parInputs,
			string parOutputs,
			string parAssumptions,
			string parOptionality,
			string parOLA,
			string parOLAvariations,
			string parCSDheading,
			string parCSDdescription,
			string parCSDsummary,
			string parSOWheading,
			string parSOWdescription,
			string parSOWsummary,
			List<int?> parRACIaccountables,
			List<int?> parRACIresponsibles,
			List<int?> parRACIconsulteds,
			List<int?> parRACIinformeds,
			string parContentStatus)
			{
			bool result = false;
			Activity newEntry;
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<Activity>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();
					if (newEntry == null)
						newEntry = new Activity();

					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.Category = parCategory;
					newEntry.SortOrder = parSortOrder;
					newEntry.ActivityDepencenciesIDsp = parActivityDependenciesIDsp;
					newEntry.ISDheading = parISDheading;
					newEntry.ISDdescription = parISDdescription;
					newEntry.ISDsummary = parISDsummary;
					newEntry.Inputs = parInputs;
					newEntry.Outputs = parOutputs;
					newEntry.Assumptions = parAssumptions;
					newEntry.Optionality = parOptionality;
					newEntry.OLA = parOLA;
					newEntry.OLAvariations = parOLAvariations;
					newEntry.CSDheading = parCSDheading;
					newEntry.CSDdescription = parCSDdescription;
					newEntry.CSDsummary = parCSDsummary;
					newEntry.SOWheading = parSOWheading;
					newEntry.SOWdescription = parSOWdescription;
					newEntry.SOWsummary = parSOWsummary;
					newEntry.RACIaccountables = parRACIaccountables;
					newEntry.RACIresponsibles = parRACIresponsibles;
					newEntry.RACIconsulteds = parRACIconsulteds;
					newEntry.RACIinformeds = parRACIinformeds;
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
		/// <returns>Activity object is retrieved if it exist, else null is retured.</returns>
		public static Activity Read(int parIDsp)
			{
			Activity result = new Activity();
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();
					result = (from thisEntry in dbSession.AllObjects<Activity>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					dbSession.Commit();
					}
				catch (Exception exc)
					{
					result = null;
					Console.WriteLine("### Exception Database reading Activity [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
					dbSession.Abort();
					}
				}
			return result;
			}

		//---g
		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database. 
		/// Specify a List of intergers containing the SharePoint ID values of all the Activity objects
		/// that need to be retrived and added to the list.
		/// </summary>
		/// <param name="parIDs">pass a List<int> of all the IDsp (SharePoint ID) that need to be retrieved and returned.
		/// If all Activity must be retrieve, pass an empty List (with count = 0) to return all objects.</int> </param>
		/// <returns>a List of Activity objects are retrurned.</returns>
		public static List<Activity> ReadAll(List<int> parIDs)
			{
			List<Activity> results = new List<Activity>();
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();
					//-|Return all if none is specified
					if (parIDs.Count == 0)
						{
						foreach (Activity entry in dbSession.AllObjects<Activity>())
							{
							results.Add(entry);
							}
						}
					else //-| Specific entries were specified.
						{
						foreach (int item in parIDs)
							{
							Activity entry = new Activity();
							entry = (from thisEntry in dbSession.AllObjects<Activity>()
									 where thisEntry.IDsp == item
									 select thisEntry).FirstOrDefault();
							results.Add(entry);
							}
						}
					dbSession.Commit();
					}
				catch (Exception exc)
					{
					Console.WriteLine("### Exception Database reading all Activity ### - {0} - {1}", exc.HResult, exc.Message);
					dbSession.Abort();
					}
				}
			return results;
			}
		#endregion
		}
	}
