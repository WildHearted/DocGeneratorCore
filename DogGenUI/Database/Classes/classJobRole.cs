using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Indexing;
using VelocityDb.Session;

namespace DocGeneratorCore.Database.Classes
	{
	public class JobRole : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a JobRole as mapped to the SharePoint List named JobRoles.
		/// </summary>
		#region Variables
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		private string _Title;
		private string _DeliveryDomain;
		private string _SpecificRegion;
		private string _RelevantBusinessUnit;
		private string _OtherJobTitles;
		private string _JobFrameworkLink;
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
		public string DeliveryDomain {
			get { return this._DeliveryDomain; }
			set { UpdateNonIndexField(); this._DeliveryDomain = value; }
			}
		public string SpecificRegion {
			get { return this._SpecificRegion; }
			set { UpdateNonIndexField(); this._SpecificRegion = value; }
			}
		public string RelevantBusinessUnit {
			get { return this._RelevantBusinessUnit; }
			set { UpdateNonIndexField (); this._RelevantBusinessUnit = value;}
			}
		public string OtherJobTitles {
			get { return this._OtherJobTitles; }
			set { UpdateNonIndexField(); this._OtherJobTitles = value;
				}
			}
		public string JobFrameworkLink {
			get { return this._JobFrameworkLink; }
			set { UpdateNonIndexField(); this._JobFrameworkLink = value; }
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
			string parDeliveryDomain,
			string parSpecificRegion,
			string parRelevantBusinessUnit,
			string parOtherJobTitles,
			string parJobFrameworkLink
			)
			{
			JobRole newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<JobRole>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();
					if (newEntry == null)
						newEntry = new JobRole();
					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.DeliveryDomain = parDeliveryDomain;
					newEntry.SpecificRegion = parSpecificRegion;
					newEntry.RelevantBusinessUnit = parRelevantBusinessUnit;
					newEntry.OtherJobTitles = parOtherJobTitles;
					newEntry.JobFrameworkLink = parJobFrameworkLink;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					return true;
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database persisting JobRole ### - {0} - {1}", exc.HResult, exc.Message);
				return false;
				}
			}

		//---G
		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>DataStatus object is retrieved if it exist, else null is retured.</returns>
		public static JobRole Read(int parIDsp)
			{
			JobRole result = new JobRole();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<JobRole>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					}
				}
			catch (Exception exc)
				{
				result = null;
				Console.WriteLine("### Exception Database reading JobRole [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
				}
			return result;
			}

		//---G
		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database. 
		/// Specify a List of intergers containing the SharePoint ID values of all the JobRole objects
		/// that need to be retrived and added to the list.
		/// </summary>
		/// <param name="parIDs">pass a List<int> of all the IDsp (SharePoint ID) that need to be retrieved and returned.
		/// If all JobRoles must be retrieve, pass an empty List (with count = 0) to return all objects.</int> </param>
		/// <returns>a List<JobRole> objects are retrurned.</returns>
		public static List<JobRole> ReadAll(List<int> parIDs)
			{
			List<JobRole> results = new List<JobRole>();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Return all Products if no product is specified
					if (parIDs.Count == 0)
						{
						foreach (JobRole entry in dbSession.AllObjects<JobRole>())
							{
							results.Add(entry);
							}
						}
					else //-| Specific entries were specified.
						{
						foreach (int item in parIDs)
							{
							JobRole entry = new JobRole();
							entry = (from thisEntry in dbSession.AllObjects<JobRole>()
									 where thisEntry.IDsp == item
									 select thisEntry).FirstOrDefault();
							results.Add(entry);
							}
						}
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all JobRole ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}
		#endregion
		}
	}
