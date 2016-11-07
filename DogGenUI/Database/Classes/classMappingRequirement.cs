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
		#region Variables
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		private string _Title;
		[Index]
		private int? _MappingServiceTowerIDsp;
		private double? _SortOrder;
		private string _RequirementText;
		private string _RequirementServiceLevel;
		private string _SourceReference;
		private string _ComplianceStatus;
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
		public int? MappingServiceTowerIDsp {
			get { return this._MappingServiceTowerIDsp; }
			set { Update(); this._MappingServiceTowerIDsp = value; }
			}
		public double? SortOrder {
			get { return this._SortOrder; }
			set { Update(); this._SortOrder = value; }
			}
		public string RequirementText {
			get { return this._RequirementText; }
			set { UpdateNonIndexField(); this._RequirementText= value; }
			}
		public string RequirementServiceLevel {
			get { return this._RequirementServiceLevel; }
			set { UpdateNonIndexField(); this._RequirementServiceLevel = value; }
			}
		public string SourceReference {
			get { return this._SourceReference; }
			set { UpdateNonIndexField(); this._SourceReference = value; }
			}
		public string ComplianceStatus {
			get { return this._ComplianceStatus; }
			set { UpdateNonIndexField(); this._ComplianceStatus = value; }
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
			int? parMappingServiceTowerIDsp,
			double? parSortOrder,
			string parRequirementText,
			string parRequirementServiceLevel,
			string parSourceReference,
			string parComplianceStatus,
			string parComplianceComments
			)

			{
			MappingRequirement newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
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
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database persisting MappingRequirement ### - {0} - {1}", exc.HResult, exc.Message);
				return false;
				}
			}

		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>DataStatus object is retrieved if it exist, else null is retured.</returns>
		public static MappingRequirement Read(int parIDsp)
			{
			MappingRequirement result = new MappingRequirement();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<MappingRequirement>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					}
				}
			catch (Exception exc)
				{
				result = null;
				Console.WriteLine("### Exception Database reading MappingRequirement [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
				}
			return result;
			}

		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database.
		/// Specify a List of intergers containing the SharePoint ID values of all the MappingRequirement objects
		/// that need to be retrived and added to the list.
		/// </summary>
		/// <param name="parMappingServiceTowerIDsp">pass an int? of the MappingServiceTower IDsp (SharePoint ID) that need to be retrieved.
		/// If all MappingRequirements must be retrieve, pass a null value as the parameter to return all objects.</param>
		/// <returns>a List<MappingRequirement> objects are retrurned.</returns>
		public static List<MappingRequirement> ReadAll(int? parMappingServiceTowerIDsp)
			{
			List<MappingRequirement> results = new List<MappingRequirement>();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Return all Products if no product is specified
					if (parMappingServiceTowerIDsp == null)
						{
						var mappingRequirements = from thisEntry in dbSession.AllObjects<MappingRequirement>() select thisEntry;
						if (mappingRequirements.Count() > 0)
							{
							foreach (MappingRequirement item in mappingRequirements)
								{
								results.Add(item);
								}
							}
						}
					else
						{
						var mappingRequirements = from thisEntry in dbSession.AllObjects<MappingRequirement>()
												  where thisEntry.MappingServiceTowerIDsp == parMappingServiceTowerIDsp
												  select thisEntry;
						if (mappingRequirements.Count() > 0)
							{
							foreach (MappingRequirement item in mappingRequirements)
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
				Console.WriteLine("### Exception Database reading all MappingRequirement ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}
		#endregion
		}
	}
