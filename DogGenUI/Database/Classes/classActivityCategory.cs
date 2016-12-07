using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Indexing;
using VelocityDb.Session;

namespace DocGeneratorCore.Database.Classes
	{
	public class ActivityCategory : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a ActivityCategory as mapped to the SharePoint List named ActivityCategorys.
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
		#endregion

			#region Methods
			//++**Store**
			/// <summary>
			/// Store/Save a new Object in the database, use the same Store method for New and Updates.
			/// </summary>
		public static bool Store(
			int parIDsp,
			string parTitle
			)
			{
			ActivityCategory newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<ActivityCategory>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();
					if (newEntry == null)
						newEntry = new ActivityCategory();
					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					return true;
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database persisting ActivityCategory ### - {0} - {1}", exc.HResult, exc.Message);
				return false;
				}
			}

		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>DataStatus object is retrieved if it exist, else null is retured.</returns>
		public static ActivityCategory Read(int parIDsp)
			{
			ActivityCategory result = new ActivityCategory();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<ActivityCategory>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					}
				}
			catch (Exception exc)
				{
				result = null;
				Console.WriteLine("### Exception Database reading ActivityCategory [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
				}
			return result;
			}

		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database. 
		/// Specify a List of intergers containing the SharePoint ID values of all the ActivityCategory objects
		/// that need to be retrived and added to the list.
		/// </summary>
		/// <param name="parIDs">pass a List<int> of all the IDsp (SharePoint ID) that need to be retrieved and returned.
		/// If all ActivityCategorys must be retrieve, pass an empty List (with count = 0) to return all objects.</int> </param>
		/// <returns>a List<ActivityCategory> objects are retrurned.</returns>
		public static List<ActivityCategory> ReadAll(List<int> parIDs)
			{
			List<ActivityCategory> results = new List<ActivityCategory>();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Return all Products if no product is specified
					if (parIDs.Count == 0)
						{
						foreach (ActivityCategory entry in dbSession.AllObjects<ActivityCategory>())
							{
							results.Add(entry);
							}
						}
					else //-| Specific entries were specified.
						{
						foreach (int item in parIDs)
							{
							ActivityCategory entry = new ActivityCategory();
							entry = (from thisEntry in dbSession.AllObjects<ActivityCategory>()
									 where thisEntry.IDsp == item
									 select thisEntry).FirstOrDefault();
							results.Add(entry);
							}
						}
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all ActivityCategory ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}
		#endregion
		}
	}
