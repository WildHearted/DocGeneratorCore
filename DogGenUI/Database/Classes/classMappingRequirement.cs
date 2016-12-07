using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Indexing;
using VelocityDb.Session;

namespace DocGeneratorCore.Database.Classes
	{
	public class MappingRequirement : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a MappingRequirement as mapped to the SharePoint List named MappingServicePowers.
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
		private int? _MappingServiceTowerIDsp;
		public int? MappingServiceTowerIDsp {
			get { return this._MappingServiceTowerIDsp; }
			set { Update(); this._MappingServiceTowerIDsp = value; }
			}

		private double? _SortOrder;
		public double? SortOrder {
			get { return this._SortOrder; }
			set { Update(); this._SortOrder = value; }
			}

		private string _RequirementText;
		public string RequirementText {
			get { return this._RequirementText; }
			set { UpdateNonIndexField(); this._RequirementText= value; }
			}

		private string _RequirementServiceLevel;
		public string RequirementServiceLevel {
			get { return this._RequirementServiceLevel; }
			set { UpdateNonIndexField(); this._RequirementServiceLevel = value; }
			}

		private string _SourceReference;
		public string SourceReference {
			get { return this._SourceReference; }
			set { UpdateNonIndexField(); this._SourceReference = value; }
			}

		private string _ComplianceStatus;
		public string ComplianceStatus {
			get { return this._ComplianceStatus; }
			set { UpdateNonIndexField(); this._ComplianceStatus = value; }
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
			int? parMappingServiceTowerIDsp,
			double? parSortOrder,
			string parRequirementText,
			string parRequirementServiceLevel,
			string parSourceReference,
			string parComplianceStatus,
			string parComplianceComments)

			{
			MappingRequirement newEntry;
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<MappingRequirement>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();

					if (newEntry == null)
						newEntry = new MappingRequirement();

					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.MappingServiceTowerIDsp = parMappingServiceTowerIDsp;
					newEntry.SortOrder = parSortOrder;
					newEntry.RequirementText = parRequirementText;
					newEntry.RequirementServiceLevel = parRequirementServiceLevel;
					newEntry.SourceReference = parSourceReference;
					newEntry.ComplianceStatus = parComplianceStatus;
					newEntry.ComplianceComments = parComplianceComments;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					return true;
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					Console.WriteLine("### Exception Database persisting MappingRequirement ### - {0} - {1}", exc.HResult, exc.Message);
					return false;
					}
				}
			}

		//===G
		//++Read
		/// <summary>
		/// Read/retrieve an entry from the database
		/// </summary>
		/// <returns>MappingRequirement object is retrieved if it exist, else null is retured.</returns>
		public static MappingRequirement Read(int parIDsp)
			{
			MappingRequirement result = new MappingRequirement();
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<MappingRequirement>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();

					dbSession.Commit();
					}
				catch (Exception exc)
					{
					result = null;
					dbSession.Abort();
					Console.WriteLine("### Exception Database reading MappingRequirement [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
					}
				}
			return result;
			}

		//===G
		//++ReadAllForServiceTower
		/// <summary>
		/// Read/retrieve all the Mapping Requirements associated with a specific Service Tower form the database.
		/// </summary>
		/// <param name="parMappingServiceTowerIDsp">pass an int? of the MappingServiceTower IDsp (SharePoint ID) that need to be retrieved.</param>
		/// <returns>a List of MappingRequirement objects are retrurned.</returns>
		public static List<MappingRequirement> ReadAllForServiceTower(int? parMappingServiceTowerIDsp)
			{
			List<MappingRequirement> results = new List<MappingRequirement>();
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();
					//-|Return all MappingRequirements for the Service Tower
					var mappingRequirements = from thisEntry in dbSession.AllObjects<MappingRequirement>()
											where thisEntry.MappingServiceTowerIDsp == parMappingServiceTowerIDsp
											select thisEntry;

					foreach (MappingRequirement item in mappingRequirements)
						{
						results.Add(item);
						}
					dbSession.Commit();
					return results;
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					Console.WriteLine("### Exception Database reading all MappingRequirement ### - {0} - {1}", exc.HResult, exc.Message);
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

					foreach (MappingRequirement entry in dbSession.AllObjects<MappingRequirement>())
						{
						dbSession.Unpersist(entry);
						}

					dbSession.Commit();
					result = true;
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					Console.WriteLine("### Exception Database deleting all Mapping Requirements  ### - {0} - {1}", exc.HResult, exc.Message);
					result = false;
					}
				}
			return result;
			}
		#endregion
		}
	}
