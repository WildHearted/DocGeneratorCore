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
	public class Deliverable : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a Deliverable as mapped to the SharePoint List named Deliverables.
		/// </summary>
		#region Variables
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		[Index]
		private string _Title;
		private string _DeliverableType;
		[Index]
		private double? _SortOrder;
		[Index]
		private int? _ServiceProductIDsp;
		private string _ContentLayer;
		[Index]
		private int? _ContentPredecessorDeliverableIDsp;
		private Deliverable _ContentPredecessorDeliverable;
		private string _ISDheading;
		private string _ISDdescription;
		private string _ISDsummary;
		private string _CSDheading;
		private string _CSDdescription;
		private string _CSDsummary;
		private string _SOWheading;
		private string _SOWdescription;
		private string _SOWsummary;
		private string _Inputs;
		private string _Outputs;
		private string _DDobligations;
		private string _ClientResponsibilities;
		private string _Exclusions;
		private string _GovernanceControls;
		private string _TransitionDescription;
		private string _WhatHasChanged;
		private List<string> _SupportingSystems;
		private List<int?> _RACIaccountables;
		private List<int?> _RACIresponsibles;
		private List<int?> _RACIconsulteds;
		private List<int?> _RACIinformeds;
		private List<int> _GlossaryAndAcronyms;
		private string _ContentStatus;
		#endregion

		#region Properties
		public int IDsp {
			get { return this._IDsp; }
			set { Update(); this._IDsp = value; }
			}
		public string DeliverableType {
			get { return this._DeliverableType; }
			set { UpdateNonIndexField(); this._DeliverableType = value;}
			}
		public string Title {
			get { return this._Title; }
			set { Update(); this._Title = value; }
			}
		public double? SortOrder {
			get { return this._SortOrder;}
			set { Update(); this._SortOrder = value;}
			}
		public int? ServiceProductIDsp {
			get { return this._ServiceProductIDsp;}
			set { Update();this._ServiceProductIDsp = value;}
			}
		public string ContentLayer {
			get { return this._ContentLayer; }
			set { Update(); this._ContentLayer = value; }
			}
		public int? ContentPredecessorDeliverableIDsp {
			get { return this._ContentPredecessorDeliverableIDsp; }
			set { Update(); this._ContentPredecessorDeliverableIDsp = value; }
			}
		public Deliverable ContentPredecessorDeliverable {
			get {
				Session?.LoadFields(pObj: _ContentPredecessorDeliverable);
				return this._ContentPredecessorDeliverable;
				}
			set {
				Update();
				this._ContentPredecessorDeliverable = value;
				}
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
		public string Inputs {
			get { return this._Inputs; }
			set { UpdateNonIndexField();this._Inputs = value; }
			}
		public string Outputs {
			get { return this._Outputs; }
			set { UpdateNonIndexField(); this._Outputs = value; }
			}
		public string DDobligations {
			get { return this._DDobligations; }
			set { UpdateNonIndexField(); this._DDobligations = value; }
			}
		public string ClientResponsibilities {
			get { return this._ClientResponsibilities; }
			set { UpdateNonIndexField(); this._ClientResponsibilities = value; }
			}
		public string Exclusions {
			get { return this._Exclusions; }
			set { UpdateNonIndexField(); this._Exclusions = value; }
			}
		public string GovernanceControls {
			get { return this._GovernanceControls; }
			set { UpdateNonIndexField(); this._GovernanceControls = value; }
			}
		public string TransitionDescription {
			get { return this._TransitionDescription;}
			set { UpdateNonIndexField(); this._TransitionDescription = value; }
			}
		public string WhatHasChanged {
			get { return this._WhatHasChanged; }
			set { UpdateNonIndexField(); this._WhatHasChanged = value; }
			}
		public List<string> SupportingSystems {
			get { return this._SupportingSystems; }
			set { UpdateNonIndexField(); this._SupportingSystems = value; }
			} 
		public List<int> GlossaryAndAcronyms {
			get { return this._GlossaryAndAcronyms; } 
			set { UpdateNonIndexField(); this._GlossaryAndAcronyms = value; }
			}
		public List<int?> RACIaccountables {
			get { return this._RACIaccountables; }
			set { UpdateNonIndexField(); this._RACIaccountables = value; }
			}
		public List<int?> RACIresponsibles {
			get { return this._RACIresponsibles; }
			set { UpdateNonIndexField(); this._RACIresponsibles = value; }
			}
		public List<int?> RACIconsulteds {
			get { return this._RACIconsulteds; }
			set { Update(); this._RACIconsulteds = value; }
			}
		public List<int?> RACIinformeds {
			get { return this.RACIinformeds; }
			set { Update(); this._RACIconsulteds = value; }
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
			List<int?> parRACIaccountables,
			List<int?> parRACIresponsibles,
			List<int?> parRACIconsulteds,
			List<int?> parRACIinformeds,
			string parContentStatus
			)
			{
			Deliverable newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
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
					newEntry.DeliverableType  = parDeliverableType;
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
		/// <returns>Deliverable object is retrieved if it exist, else null is retured.</returns>
		public static Deliverable Read(int parIDsp)
			{
			Deliverable result = new Deliverable();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<Deliverable>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					}
				}
			catch (Exception exc)
				{
				result = null;
				Console.WriteLine("### Exception Database reading Deliverable [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
				}
			return result;
			}

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
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Return all Products if no product is specified
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
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all Deliverable ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}
		#endregion
		}
	}
