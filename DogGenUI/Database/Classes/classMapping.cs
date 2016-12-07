using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Indexing;
using VelocityDb.Session;

namespace DocGeneratorCore.Database.Classes
	{
	public class Mapping : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a Mapping as mapped to the SharePoint List named Mappings.
		/// </summary>

		#region Properties
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		public int IDsp {
			get { return this._IDsp; }
			set { Update(); this._IDsp = value;}
			}

		private string _Title;
		public string Title {
			get { return this._Title; }
			set { UpdateNonIndexField(); this._Title = value; }
			}

		private string _ClientName;
		public string ClientName {
			get { return this._ClientName; }
			set { UpdateNonIndexField(); this._ClientName = value; }
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
			string parClientName)

			{
			Mapping newEntry;
				using (ServerClientSession dbSession = new ServerClientSession(
				systemHost: Properties.Settings.Default.CurrentDatabaseHost,
				systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					try
						{
						dbSession.BeginUpdate();
						newEntry = (from objEntry in dbSession.AllObjects<Mapping>()
									where objEntry.IDsp == parIDsp
									select objEntry).FirstOrDefault();

						if (newEntry == null)
							newEntry = new Mapping();
						newEntry.IDsp = parIDsp;
						newEntry.Title = parTitle;
						newEntry.ClientName = parClientName;
						dbSession.Persist(newEntry);
						dbSession.Commit();
						return true;
						}
					catch (Exception exc)
						{
						dbSession.Abort();
						Console.WriteLine("### Exception Database persisting Mapping ### - {0} - {1}", exc.HResult, exc.Message);
						return false;
						}
				}
			}

		//++Read
		/// <summary>
		/// Read/retrieve an entry from the database
		/// </summary>
		/// <returns>Mapping object is retrieved if it exist, else null is retured.</returns>
		public static Mapping Read(int parIDsp)
			{
			Mapping result = new Mapping();
			using (ServerClientSession dbSession = new ServerClientSession(
				systemHost: Properties.Settings.Default.CurrentDatabaseHost,
				systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{						
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<Mapping>()
								where thisEntry.IDsp == parIDsp
								select thisEntry).FirstOrDefault();
					dbSession.Commit();
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					result = null;
					Console.WriteLine("### Exception Database reading Mapping [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
					}
				}
			
			return result;
			}

		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database. 
		/// Specify a List of intergers containing the SharePoint ID values of all the Mapping objects
		/// that need to be retrived and added to the list.
		/// </summary>
		/// <param name="parIDs">pass a List<int> of all the IDsp (SharePoint ID) that need to be retrieved and returned.
		/// If all Mappings must be retrieve, pass an empty List (with count = 0) to return all objects.</int> </param>
		/// <returns>a List<Mapping> objects are retrurned.</returns>
		public static List<Mapping> ReadAll(List<int> parIDs)
			{
			List<Mapping> results = new List<Mapping>();
			
			using (ServerClientSession dbSession = new ServerClientSession(
				systemHost: Properties.Settings.Default.CurrentDatabaseHost,
				systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();
					//-|Return all Products if no product is specified
					if (parIDs.Count == 0)
						{
						foreach (Mapping entry in dbSession.AllObjects<Mapping>())
							{
							results.Add(entry);
							}
						}
					else //-| Specific entries were specified.
						{
						foreach (int item in parIDs)
							{
							Mapping entry = new Mapping();
							entry = (from thisEntry in dbSession.AllObjects<Mapping>()
									 where thisEntry.IDsp == item
									 select thisEntry).FirstOrDefault();
							results.Add(entry);
							}
						}
					dbSession.Commit();
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					Console.WriteLine("### Exception Database reading all Mapping ### - {0} - {1}", exc.HResult, exc.Message);
					}
				}
			return results;
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

					foreach (Mapping entry in dbSession.AllObjects<Mapping>())
						{
						dbSession.Unpersist(entry);
						}

					dbSession.Commit();
					result = true;
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					Console.WriteLine("### Exception Database deleting all Mapping ### - {0} - {1}", exc.HResult, exc.Message);
					result = false;
					}
				}
			return result;
			}

		#endregion
		}
	}
