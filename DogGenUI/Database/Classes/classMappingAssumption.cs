using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Indexing;
using VelocityDb.Session;

namespace DocGeneratorCore.Database.Classes
	{
	public class MappingAssumption : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a MappingAssumption as mapped to the SharePoint List named MappingServicePowers.
		/// </summary>
		#region Variables
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		private string _Title;
		[Index]
		private int? _MappingRequirementIDsp;
		private string _Description;
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
		public string Description {
			get { return this._Description; }
			set { UpdateNonIndexField(); this._Description= value; }
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
			string parDescription
			)

			{
			MappingAssumption newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<MappingAssumption>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();

					if (newEntry == null)
						newEntry = new MappingAssumption();
					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.MappingRequirementIDsp = parMappingRequirementIDsp;
					newEntry.Description = parDescription;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					return true;
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database persisting MappingAssumption ### - {0} - {1}", exc.HResult, exc.Message);
				return false;
				}
			}

		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>MappingAssumption object is retrieved if it exist, else null is retured.</returns>
		public static MappingAssumption Read(int parIDsp)
			{
			MappingAssumption result = new MappingAssumption();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<MappingAssumption>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					}
				}
			catch (Exception exc)
				{
				result = null;
				Console.WriteLine("### Exception Database reading MappingAssumption [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
				}
			return result;
			}

		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database.
		/// Specify an interger containing the SharePoint ID values of a MapingRequirement to retrieve all the related MappingAssumption objects.
		/// </summary>
		/// <param name="parMappingRequirementIDsp">pass an int? of the MappingRequirement IDsp (SharePoint ID) that need to be retrieved.
		/// If all MappingAssumptions must be retrieve, pass a null value as the parameter to return all objects.</param>
		/// <returns>a List<MappingAssumption> objects are retrurned.</returns>
		public static List<MappingAssumption> ReadAll(int? parMappingRequirementIDsp)
			{
			List<MappingAssumption> results = new List<MappingAssumption>();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Return all MappingAssumptions Null is specified
					if (parMappingRequirementIDsp == null)
						{
						var mappingRequirements = from thisEntry in dbSession.AllObjects<MappingAssumption>() select thisEntry;
						if (mappingRequirements.Count() > 0)
							{
							foreach (MappingAssumption item in mappingRequirements)
								{
								results.Add(item);
								}
							}
						}
					else
						{
						var mappingRequirements = from thisEntry in dbSession.AllObjects<MappingAssumption>()
												  where thisEntry.MappingRequirementIDsp == parMappingRequirementIDsp
												  select thisEntry;
						if (mappingRequirements.Count() > 0)
							{
							foreach (MappingAssumption item in mappingRequirements)
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
				Console.WriteLine("### Exception Database reading all MappingAssumption ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}
		#endregion
		}
	}
