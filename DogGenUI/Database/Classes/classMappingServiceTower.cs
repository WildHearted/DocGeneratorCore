using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Indexing;
using VelocityDb.Session;

namespace DocGeneratorCore.Database.Classes
	{
	public class MappingServiceTower : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a MappingServiceTower as mapped to the SharePoint List named MappingServicePowers.
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

		private int? _MappingIDsp;
		public int? MappingIDsp {
			get { return this._MappingIDsp; }
			set { UpdateNonIndexField(); this._MappingIDsp = value; }
			}
		#endregion

		//===G
		#region Methods

		//===G
		//++Store
		/// <summary>
		/// Store/Save a new Object in the database, use the same Store method for New and Updates.
		/// </summary>
		public static bool Store(
			int parIDsp,
			string parTitle,
			int? parMappingIDsp)

			{
			MappingServiceTower newEntry;

			using (ServerClientSession dbSession = new ServerClientSession(
				systemHost: Properties.Settings.Default.CurrentDatabaseHost,
				systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<MappingServiceTower>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();

					if (newEntry == null)
						newEntry = new MappingServiceTower();
					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.MappingIDsp = parMappingIDsp;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					return true;
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					Console.WriteLine("### Exception Database persisting MappingServiceTower ### - {0} - {1}", exc.HResult, exc.Message);
					return false;
					}
				}
			}

		//++DeleteAll
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

					foreach (MappingServiceTower entry in dbSession.AllObjects<MappingServiceTower>())
						{
						dbSession.Unpersist(entry);
						}

					dbSession.Commit();
					result = true;
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					Console.WriteLine("### Exception Database deleting all Mapping Service Towers  ### - {0} - {1}", exc.HResult, exc.Message);
					result = false;
					}
				}
			return result;
			}
		//===G
		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>DataStatus object is retrieved if it exist, else null is retured.</returns>
		public static MappingServiceTower Read(int parIDsp)
			{
			MappingServiceTower result = new MappingServiceTower();

			using (ServerClientSession dbSession = new ServerClientSession(
				systemHost: Properties.Settings.Default.CurrentDatabaseHost,
				systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<MappingServiceTower>()
								where thisEntry.IDsp == parIDsp
								select thisEntry).FirstOrDefault();

					dbSession.Commit();
					}
				catch (Exception exc)
					{
					result = null;
					dbSession.Abort();
					Console.WriteLine("### Exception Database reading MappingServiceTower [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
					}
				}
			
			return result;
			}

		//===G
		//++ReadMappingServiceTowersForMapping
		/// <summary>
		/// Read/retrieve all the entries from the database for a specific Mapping.
		/// Provide a Mapping IDsp for which all the MappingServiceTower objects must be returned.
		/// </summary>
		/// <param name="parMappingIDsp">pass a IDsp (SharePoint ID) that need to be retrieved and returned.</param>
		/// <returns>a List of MappingServiceTower objects are retrurned.</returns>
		public static List<MappingServiceTower> ReadMappingServiceTowersForMapping(int? parMappingIDsp)
			{
			List<MappingServiceTower> results = new List<MappingServiceTower>();

			using (ServerClientSession dbSession = new ServerClientSession(
				systemHost: Properties.Settings.Default.CurrentDatabaseHost,
				systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginUpdate();

					IEnumerable<MappingServiceTower> mappingServiceTowers = (from thisEntry in dbSession.AllObjects<MappingServiceTower>()
												where thisEntry.MappingIDsp == parMappingIDsp
												select thisEntry).AsEnumerable();
					

					foreach (MappingServiceTower item in mappingServiceTowers)
						{
						results.Add(item);
						}
					dbSession.Commit();
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					Console.WriteLine("### Exception Database reading all MappingServiceTower ### - {0} - {1}", exc.HResult, exc.Message);
					}
				}
			return results;
			}
		#endregion
		}
	}
