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
		#region Variables
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		private string _Title;
		[Index]
		private int? _MappingIDsp;
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
		public int? MappingIDsp {
			get { return this._MappingIDsp; }
			set { Update(); this._MappingIDsp = value; }
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
			int? parMappingIDsp)

			{
			MappingServiceTower newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
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
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database persisting MappingServiceTower ### - {0} - {1}", exc.HResult, exc.Message);
				return false;
				}
			}

		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>DataStatus object is retrieved if it exist, else null is retured.</returns>
		public static MappingServiceTower Read(int parIDsp)
			{
			MappingServiceTower result = new MappingServiceTower();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<MappingServiceTower>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					}
				}
			catch (Exception exc)
				{
				result = null;
				Console.WriteLine("### Exception Database reading MappingServiceTower [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
				}
			return result;
			}

		//++ReadMappingServiceTowersForMapping
		/// <summary>
		/// Read/retrieve all the entries from the database.
		/// Specify a List of intergers containing the SharePoint ID values of all the MappingServiceTower objects
		/// that need to be retrived and added to the list.
		/// </summary>
		/// <param name="parMappingIDs">pass a List<int> of all the IDsp (SharePoint ID) that need to be retrieved and returned.
		/// If all MappingServiceTowers must be retrieve, pass an empty List (with count = 0) to return all objects.</int> </param>
		/// <returns>a List<MappingServiceTower> objects are retrurned.</returns>
		public static List<MappingServiceTower> ReadMappingServiceTowersForMapping(int? parMappingIDs)
			{
			List<MappingServiceTower> results = new List<MappingServiceTower>();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Return all Products if no product is specified
					var mappingServiceTowers = from thisEntry in dbSession.AllObjects<MappingServiceTower>()
											   where thisEntry.MappingIDsp == parMappingIDs
											   select thisEntry;
					if (mappingServiceTowers.Count() > 0)
						{
						foreach (MappingServiceTower item in mappingServiceTowers)
							{
							results.Add(item);
							}
						}
					return results;
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all MappingServiceTower ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}
		#endregion
		}
	}
