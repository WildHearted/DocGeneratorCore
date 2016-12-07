using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Indexing;
using VelocityDb.Session;

namespace DocGeneratorCore.Database.Classes
	{
	public class MappingDeliverable : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a MappingDeliverable as mapped to the SharePoint List named MappingServicePowers.
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

		[Index]
		private int? _MappingRequirementIDsp;
		public int? MappingRequirementIDsp {
			get { return this._MappingRequirementIDsp; }
			set { Update(); this._MappingRequirementIDsp = value; }
			}

		private bool _NewDeliverable;
		public bool NewDeliverable {
			get { return this._NewDeliverable; }
			set { UpdateNonIndexField(); this._NewDeliverable = value; }
			}

		private string _NewRequirement;
		public string NewRequirement {
			get { return this._NewRequirement; }
			set { UpdateNonIndexField(); this._NewRequirement= value; }
			}

		private int? _MappedDeliverableIDsp;
		public int? MappedDeliverableID {
			get { return this._MappedDeliverableIDsp; }
			set { UpdateNonIndexField(); this._MappedDeliverableIDsp = value; }
			}

		private string _ComplianceComments;
		public string ComplianceComments {
			get { return this._ComplianceComments; }
			set { UpdateNonIndexField(); this._ComplianceComments = value; }
			}

		#endregion

		//===G
		#region Methods
		//++Store
		/// <summary>
		/// Store/Save a new Object in the database, use the same Store method for New and Updates.
		/// </summary>
		public static bool Store(
			int parIDsp,
			string parTitle,
			int? parMappingRequirementIDsp,
			bool parNewDeliverable,
			string parNewRequirements,
			int? parMappedDeliverableIDsp,
			string parComplianceComments)

			{
			MappingDeliverable newEntry;
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<MappingDeliverable>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();

					if (newEntry == null)
						newEntry = new MappingDeliverable();
					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.MappingRequirementIDsp = parMappingRequirementIDsp;
					newEntry.NewDeliverable = parNewDeliverable;
					newEntry.NewRequirement = parNewRequirements;
					newEntry._MappedDeliverableIDsp = parMappedDeliverableIDsp;
					newEntry.ComplianceComments = parComplianceComments;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					return true;
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					Console.WriteLine("### Exception Database persisting MappingDeliverable ### - {0} - {1}", exc.HResult, exc.Message);
					return false;
					}
				}
			}

		//++Read
		/// <summary>
		/// Read/retrieve a specific entry from the database.
		/// </summary>
		/// <returns>MappingDeliverable object is retrieved if it exist, else null is retured.</returns>
		public static MappingDeliverable Read(int parIDsp)
			{
			MappingDeliverable result = new MappingDeliverable();
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<MappingDeliverable>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();

					dbSession.Commit();
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					result = null;
					Console.WriteLine("### Exception Database reading MappingDeliverable [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
					}
				}
			return result;
			}

		//++ReadAllForRequirement
		/// <summary>
		/// Read/retrieve all the Mapping Deliverables entries for a specific Mapping Requirement from the database.
		/// </summary>
		/// <param name="parMappingRequirementIDsp">pass an int? of the MappingSRequirement IDsp (SharePoint ID) that need to be retrieved.</param>
		/// <returns>a List of MappingDeliverable objects is retrurned.</returns>
		public static List<MappingDeliverable> ReadAllForRequirement(int? parMappingRequirementIDsp)
			{
			List<MappingDeliverable> results = new List<MappingDeliverable>();
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();
					var mappingRequirements = from thisEntry in dbSession.AllObjects<MappingDeliverable>()
											  where thisEntry.MappingRequirementIDsp == parMappingRequirementIDsp
											  select thisEntry;

					foreach (MappingDeliverable item in mappingRequirements)
						{
						results.Add(item);
						}

					dbSession.Commit();
					return results;
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					Console.WriteLine("### Exception Database reading all MappingDeliverable ### - {0} - {1}", exc.HResult, exc.Message);
					}
				}
			return results;
			}


		//+DeleteAll
		/// <summary>
		/// Delete all the entries from the database. 
		/// </summary>
		/// <returns>a boolean value TRUE = success FALSE = failure</returns>
		public static bool DeleteAll()
			{
			bool result = false;

			using (ServerClientSession dbSession = new ServerClientSession(
				systemHost: Properties.Settings.Default.CurrentDatabaseHost,
				systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginUpdate();

					foreach (MappingDeliverable entry in dbSession.AllObjects<MappingDeliverable>())
						{
						dbSession.Unpersist(entry);
						}

					dbSession.Commit();
					result = true;
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					Console.WriteLine("### Exception Database deleting all Mapping Deliverables  ### - {0} - {1}", exc.HResult, exc.Message);
					result = false;
					}
				}
			return result;
			}
		#endregion
		}
	}
