using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Indexing;
using VelocityDb.Session;

namespace DocGeneratorCore.Database.Classes
	{
	public class MappingServiceLevel : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a MappingServiceLevel as mapped to the SharePoint List named MappingServicePowers.
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

		private int? _MappingDeliverableIDsp;
		public int? MappingDeliverableIDsp {
			get { return this._MappingDeliverableIDsp; }
			set { Update(); this._MappingDeliverableIDsp = value; }
			}

		private bool? _NewServiceLevel;
		public bool? NewServiceLevel {
			get { return this._NewServiceLevel; }
			set { UpdateNonIndexField(); this._NewServiceLevel = value; }
			}

		private string _ServiceLevelRequirement;
		public string RequirementText {
			get { return this._ServiceLevelRequirement; }
			set { UpdateNonIndexField(); this._ServiceLevelRequirement= value; }
			}

		private int? _MappedServiceLevelIDsp;
		public int? MappedServiceLevelIDsp {
			get { return this._MappedServiceLevelIDsp; }
			set { Update(); this._MappedServiceLevelIDsp = value; }
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
			int? parMappedServiceLevelIDsp
			)

			{
			MappingServiceLevel newEntry;
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<MappingServiceLevel>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();

					if (newEntry == null)
						newEntry = new MappingServiceLevel();
					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.MappingDeliverableIDsp = parMappingRequirementIDsp;
					newEntry.NewServiceLevel = parNewDeliverable;
					newEntry.RequirementText = parNewRequirements;
					newEntry.MappedServiceLevelIDsp = parMappedServiceLevelIDsp;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					return true;
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					Console.WriteLine("### Exception Database persisting MappingServiceLevel ### - {0} - {1}", exc.HResult, exc.Message);
					return false;
					}
				}
			}

		//++Read
		/// <summary>
		/// Read/retrieve a specific entries from the database
		/// </summary>
		/// <returns>MappingServiceLevel object is retrieved if it exist, else null is retured.</returns>
		public static MappingServiceLevel Read(int parIDsp)
			{
			MappingServiceLevel result = new MappingServiceLevel();
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<MappingServiceLevel>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();

					dbSession.Commit();
					}
				catch (Exception exc)
					{
					result = null;
					dbSession.Abort();
					Console.WriteLine("### Exception Database reading MappingServiceLevel [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
					}
				}
			return result;
			}

		//++ReadAllForMappingDeliverable
		/// <summary>
		/// Read/retrieve all the entries from the database.
		/// Specify an interger containing the SharePoint ID values of a MaapingRequirement to retrieve all the related MappingServiceLevel objects.
		/// </summary>
		/// <param name="parMappingDeliverableIDsp">pass an int? of the MappingSRequirement IDsp (SharePoint ID) that need to be retrieved.
		/// If all MappingServiceLevels must be retrieve, pass a null value as the parameter to return all objects.</param>
		/// <returns>a List<MappingServiceLevel> objects are retrurned.</returns>
		public static List<MappingServiceLevel> ReadAllForMappingDeliverable(int? parMappingDeliverableIDsp)
			{
			List<MappingServiceLevel> results = new List<MappingServiceLevel>();
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();
					var mappingRequirements = from thisEntry in dbSession.AllObjects<MappingServiceLevel>()
											  where thisEntry.MappingDeliverableIDsp == parMappingDeliverableIDsp
											  select thisEntry;

					foreach (MappingServiceLevel item in mappingRequirements)
						{
						results.Add(item);
						}
					dbSession.Commit();
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

					foreach (MappingServiceLevel entry in dbSession.AllObjects<MappingServiceLevel>())
						{
						dbSession.Unpersist(entry);
						}

					dbSession.Commit();
					result = true;
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					Console.WriteLine("### Exception Database deleting all Mapping Service Levels  ### - {0} - {1}", exc.HResult, exc.Message);
					result = false;
					}
				}
			return result;
			}

		#endregion
		}
	}
