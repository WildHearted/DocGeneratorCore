using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Indexing;
using VelocityDb.Session;

namespace DocGeneratorCore.Database.Classes
	{
	public class Deliverable : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a Deliverable as mapped to the SharePoint List named Deliverables.
		/// </summary>
		
		#region Properties
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		public int IDsp {
			get { return this._IDsp; }
			set { Update(); this._IDsp = value; }
			}

		private string _DeliverableType;
		public string DeliverableType {
			get { return this._DeliverableType; }
			set { UpdateNonIndexField(); this._DeliverableType = value;}
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

		private int? _ServiceProductIDsp;
		public int? ServiceProductIDsp {
			get { return this._ServiceProductIDsp;}
			set { UpdateNonIndexField();this._ServiceProductIDsp = value;}
			}

		private string _ContentLayer;
		public string ContentLayer {
			get { return this._ContentLayer; }
			set { UpdateNonIndexField(); this._ContentLayer = value; }
			}

		private int? _ContentPredecessorDeliverableIDsp;
		public int? ContentPredecessorDeliverableIDsp {
			get { return this._ContentPredecessorDeliverableIDsp; }
			set { UpdateNonIndexField(); this._ContentPredecessorDeliverableIDsp = value; }
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

		private string _Inputs;
		public string Inputs {
			get { return this._Inputs; }
			set { UpdateNonIndexField();this._Inputs = value; }
			}

		private string _Outputs;
		public string Outputs {
			get { return this._Outputs; }
			set { UpdateNonIndexField(); this._Outputs = value; }
			}

		private string _DDobligations;
		public string DDobligations {
			get { return this._DDobligations; }
			set { UpdateNonIndexField(); this._DDobligations = value; }
			}

		private string _ClientResponsibilities;
		public string ClientResponsibilities {
			get { return this._ClientResponsibilities; }
			set { UpdateNonIndexField(); this._ClientResponsibilities = value; }
			}

		private string _Exclusions;
		public string Exclusions {
			get { return this._Exclusions; }
			set { UpdateNonIndexField(); this._Exclusions = value; }
			}

		private string _GovernanceControls;
		public string GovernanceControls {
			get { return this._GovernanceControls; }
			set { UpdateNonIndexField(); this._GovernanceControls = value; }
			}

		private string _TransitionDescription;
		public string TransitionDescription {
			get { return this._TransitionDescription;}
			set { UpdateNonIndexField(); this._TransitionDescription = value; }
			}

		private string _WhatHasChanged;
		public string WhatHasChanged {
			get { return this._WhatHasChanged; }
			set { UpdateNonIndexField(); this._WhatHasChanged = value; }
			}

		private List<string> _SupportingSystems;
		public List<string> SupportingSystems {
			get { return this._SupportingSystems; }
			set { UpdateNonIndexField(); this._SupportingSystems = value; }
			}

		private List<int> _GlossaryAndAcronyms;
		public List<int> GlossaryAndAcronyms {
			get { return this._GlossaryAndAcronyms; } 
			set { UpdateNonIndexField(); this._GlossaryAndAcronyms = value; }
			}

		private List<int> _RACIaccountables;
		public List<int> RACIaccountables {
			get { return this._RACIaccountables; }
			set { UpdateNonIndexField(); this._RACIaccountables = value; }
			}

		private List<int> _RACIresponsibles;
		public List<int> RACIresponsibles {
			get { return this._RACIresponsibles; }
			set { UpdateNonIndexField(); this._RACIresponsibles = value; }
			}

		private List<int> _RACIconsulteds;
		public List<int> RACIconsulteds {
			get { return this._RACIconsulteds; }
			set { UpdateNonIndexField(); this._RACIconsulteds = value; }
			}

		private List<int> _RACIinformeds;
		public List<int> RACIinformeds {
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
		/// Store/Save a new Object in the database, use the same Store method for New and Updates.
		/// </summary>
		public static bool Store(
			int parIDsp,
			string parTitle,
			string parDeliverableType,
			int parServiceProductIDsp,
			double? parSortOrder,
			string parContentLayer,
			int? parContentPredecessorDeliverableIDsp,
			string parISDheading,
			string parISDdescription,
			string parISDsummary,
			string parInputs,
			string parOutputs,
			string parDDobligations,
			string parClientResponsibilities,
			string parExceptions,
			string parGovernanceControls,
			string parCSDheading,
			string parCSDdescription,
			string parCSDsummary,
			string parSOWheading,
			string parSOWdescription,
			string parSOWsummary,
			string parTransitionDescription,
			string parWhatHasChanged,
			List<string> parSupportingSystems,
			List<int> parGlossaryAndAcronyms,
			List<int> parRACIaccountables,
			List<int> parRACIresponsibles,
			List<int> parRACIconsulteds,
			List<int> parRACIinformeds,
			string parContentStatus)
			{
			bool result = false;
			Deliverable newEntry;
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<Deliverable>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();
					if (newEntry == null)
						newEntry = new Deliverable();

					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.SortOrder = parSortOrder;
					newEntry.DeliverableType = parDeliverableType;
					newEntry.ServiceProductIDsp = parServiceProductIDsp;
					newEntry.ContentLayer = parContentLayer;
					newEntry.ContentPredecessorDeliverableIDsp = parContentPredecessorDeliverableIDsp;

					newEntry.ISDheading = parISDheading;
					newEntry.ISDdescription = parISDdescription;
					newEntry.ISDsummary = parISDsummary;
					newEntry.Inputs = parInputs;
					newEntry.Outputs = parOutputs;
					newEntry.DDobligations = parDDobligations;
					newEntry.ClientResponsibilities = parClientResponsibilities;
					newEntry.Exclusions = parExceptions;
					newEntry.GovernanceControls = parGovernanceControls;
					newEntry.CSDheading = parCSDheading;
					newEntry.CSDdescription = parCSDdescription;
					newEntry.CSDsummary = parCSDsummary;
					newEntry.SOWheading = parSOWheading;
					newEntry.SOWdescription = parSOWdescription;
					newEntry.SOWsummary = parSOWsummary;
					newEntry.TransitionDescription = parTransitionDescription;
					newEntry.WhatHasChanged = parWhatHasChanged;
					newEntry.SupportingSystems = parSupportingSystems;
					newEntry.GlossaryAndAcronyms = parGlossaryAndAcronyms;
					newEntry.RACIaccountables = parRACIaccountables;
					newEntry.RACIresponsibles = parRACIresponsibles;
					newEntry._RACIconsulteds = parRACIconsulteds;
					newEntry._RACIinformeds = parRACIinformeds;
					newEntry.ContentStatus = parContentStatus;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					result = true;
					}
				catch (Exception exc)
					{
					Console.WriteLine("### Exception Database persisting Service Product ### - {0} - {1}", exc.HResult, exc.Message);
					result = false;
					}
				}
			return result;
			}

		//---g
		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>Deliverable object is retrieved if it exist, else null is retured.</returns>
		public static Deliverable Read(int parIDsp)
			{
			Deliverable result = new Deliverable();
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();
					result = (from thisEntry in dbSession.AllObjects<Deliverable>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					dbSession.Commit();
					}
				catch (Exception exc)
					{
					result = null;
					Console.WriteLine("### Exception Database reading Deliverable [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
					dbSession.Abort();
					}
				}
			return result;
			}

		//---g
		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database. 
		/// Specify a List of intergers containing the SharePoint ID values of all the Deliverable objects
		/// that need to be retrived and added to the list.
		/// </summary>
		/// <param name="parIDs">pass a List<int> of all the IDsp (SharePoint ID) that need to be retrieved and returned.
		/// If all Deliverable must be retrieve, pass an empty List (with count = 0) to return all objects.</int> </param>
		/// <returns>a List<Deliverrable> objects are retrurned.</returns>
		public static List<Deliverable> ReadAll(List<int> parIDs)
			{
			List<Deliverable> results = new List<Deliverable>();
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();
					//-|Return all if none is specified
					if (parIDs.Count == 0)
						{
						foreach (Deliverable entry in dbSession.AllObjects<Deliverable>())
							{
							results.Add(entry);
							}
						}
					else //-| Specific entries were specified.
						{
						foreach (int item in parIDs)
							{
							Deliverable entry = new Deliverable();
							entry = (from thisEntry in dbSession.AllObjects<Deliverable>()
									 where thisEntry.IDsp == item
									 select thisEntry).FirstOrDefault();
							results.Add(entry);
							}
						}
					dbSession.Commit();
					}
				catch (Exception exc)
					{
					Console.WriteLine("### Exception Database reading all Deliverable ### - {0} - {1}", exc.HResult, exc.Message);
					dbSession.Abort();
					}
				}
			return results;
			}
		#endregion
		}
	}
