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
		#region Variables
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		private string _Title;
		[Index]
		private int? _MappingRequirementIDsp;
		private bool _NewDeliverable;
		private string _NewRequirement;
		private int? _MappedDeliverableIDsp;
		private string _ComplianceComments;
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
		public int? MappingRequirementIDsp {
			get { return this._MappingRequirementIDsp; }
			set { Update(); this._MappingRequirementIDsp = value; }
			}
		public bool NewDeliverable {
			get { return this._NewDeliverable; }
			set { UpdateNonIndexField(); this._NewDeliverable = value; }
			}
		public string NewRequirement {
			get { return this._NewRequirement; }
			set { UpdateNonIndexField(); this._NewRequirement= value; }
			}
		public int? MappedDeliverableID {
			get { return this._MappedDeliverableIDsp; }
			set { UpdateNonIndexField(); this._MappedDeliverableIDsp = value; }
			}
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
			string parComplianceComments
			)

			{
			MappingDeliverable newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
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
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database persisting MappingDeliverable ### - {0} - {1}", exc.HResult, exc.Message);
				return false;
				}
			}

		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>MappingDeliverable object is retrieved if it exist, else null is retured.</returns>
		public static MappingDeliverable Read(int parIDsp)
			{
			MappingDeliverable result = new MappingDeliverable();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<MappingDeliverable>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					}
				}
			catch (Exception exc)
				{
				result = null;
				Console.WriteLine("### Exception Database reading MappingDeliverable [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
				}
			return result;
			}

		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database.
		/// Specify an interger containing the SharePoint ID values of a MaapingRequirement to retrieve all the related MappingDeliverable objects.
		/// </summary>
		/// <param name="parMappingRequirementIDsp">pass an int? of the MappingSRequirement IDsp (SharePoint ID) that need to be retrieved.
		/// If all MappingDeliverables must be retrieve, pass a null value as the parameter to return all objects.</param>
		/// <returns>a List<MappingDeliverable> objects are retrurned.</returns>
		public static List<MappingDeliverable> ReadAll(int? parMappingRequirementIDsp)
			{
			List<MappingDeliverable> results = new List<MappingDeliverable>();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Return all MappingDeliverables Null is specified
					if (parMappingRequirementIDsp == null)
						{
						var mappingRequirements = from thisEntry in dbSession.AllObjects<MappingDeliverable>() select thisEntry;
						if (mappingRequirements.Count() > 0)
							{
							foreach (MappingDeliverable item in mappingRequirements)
								{
								results.Add(item);
								}
							}
						}
					else
						{
						var mappingRequirements = from thisEntry in dbSession.AllObjects<MappingDeliverable>()
												  where thisEntry.MappingRequirementIDsp == parMappingRequirementIDsp
												  select thisEntry;
						if (mappingRequirements.Count() > 0)
							{
							foreach (MappingDeliverable item in mappingRequirements)
								{
								results.Add(item);
								}
							}
						}
					return results;
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all MappingDeliverable ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}
		#endregion
		}
	}
