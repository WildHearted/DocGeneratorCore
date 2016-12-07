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

		private string _Description;
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
			string parDescription)

			{
			MappingAssumption newEntry;
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
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
				catch (Exception exc)
					{
					dbSession.Abort();
					Console.WriteLine("### Exception Database persisting MappingAssumption ### - {0} - {1}", exc.HResult, exc.Message);
					return false;
					}
				}
			
			}

		//++Read
		/// <summary>
		/// Read/retrieve an  entry from the database
		/// </summary>
		/// <returns>MappingAssumption object is retrieved if it exist, else null is retured.</returns>
		public static MappingAssumption Read(int parIDsp)
			{
			MappingAssumption result = new MappingAssumption();
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<MappingAssumption>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					dbSession.Commit();
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					result = null;
					Console.WriteLine("### Exception Database reading MappingAssumption [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
					}
				}
			
			return result;
			}

		//++ReadAllForRequirement
		/// <summary>
		/// Read/retrieve all the Assumptions entries from the database, for a specific Requirement.
		/// </summary>
		/// <param name="parMappingRequirementIDsp">pass an int? of the MappingRequirement IDsp (SharePoint ID) that need to be retrieved.</param>
		/// <returns>a List of MappingAssumption objects are retrurned.</returns>
		public static List<MappingAssumption> ReadAllForRequirement(int? parMappingRequirementIDsp)
			{
			List<MappingAssumption> results = new List<MappingAssumption>();
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();
					var mappingRequirements = from thisEntry in dbSession.AllObjects<MappingAssumption>()
											  where thisEntry.MappingRequirementIDsp == parMappingRequirementIDsp
											  select thisEntry;
					
					foreach (MappingAssumption item in mappingRequirements)
						{
						results.Add(item);
						}
					dbSession.Commit();
					return results;
					}
				catch (Exception exc)
					{
					Console.WriteLine("### Exception Database reading all MappingAssumption ### - {0} - {1}", exc.HResult, exc.Message);
					dbSession.Abort();
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

					foreach (MappingAssumption entry in dbSession.AllObjects<MappingAssumption>())
						{
						dbSession.Unpersist(entry);
						}

					dbSession.Commit();
					result = true;
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					Console.WriteLine("### Exception Database deleting all Mapping Assumptions  ### - {0} - {1}", exc.HResult, exc.Message);
					result = false;
					}
				}
			return result;
			}
		#endregion
		}
	}
