using System;
using System.Linq;
using VelocityDb;
using VelocityDb.Session;

namespace DocGeneratorCore.Database.Classes
	{
	class DataStatus : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains the date and time 
		/// when the specific database was last synchronised with the SharePoint environment
		/// </summary>

		#region Properties

		private DateTime _LastRefreshedOn;
		public DateTime LastRefreshedOn {
			get { return this._LastRefreshedOn; }
			set { UpdateNonIndexField(); this._LastRefreshedOn = value; }
			}

		private DateTime _CreatedOn;
		public DateTime CreatedOn {
			get { return this._CreatedOn; }
			set { UpdateNonIndexField(); this._CreatedOn = value; }
			}
		#endregion

		#region Methods
		//++Store
		/// <summary>
		/// Store/Save the date and time that the when the database was last refreshed or created from SharePoint. 
		/// There is either no objects or a single object in the database, therefore validate for null.
		/// </summary>
		public static bool Store(DateTime? parRefreshedOn, DateTime? parCreatedOn)
			{
			DataStatus result;
			
			using (ServerClientSession dbSession = new ServerClientSession(
				systemDir: Properties.Settings.Default.CurrentDatabaseLocation, 
				systemHost: Properties.Settings.Default.CurrentDatabaseHost))
				{
				try
					{
					dbSession.BeginUpdate();
					result = (from theDBstatus in dbSession.AllObjects<DataStatus>()
							  select theDBstatus).FirstOrDefault();

					if (result == null)
						result = new DataStatus();

					//-|Update the CreatedOn field if provided
					if (parCreatedOn != null)
						result.CreatedOn = Convert.ToDateTime(parCreatedOn);

					//-Update the LastRefreshedOn field if provided
					if (parRefreshedOn != null)
						result.LastRefreshedOn = Convert.ToDateTime(parRefreshedOn);

					dbSession.Persist(result);
					dbSession.Commit();
					return true;
					}
				catch (Exception exc)
					{
					Console.WriteLine("\n*** EXCEPTION *** {0} - {1}", exc.HResult, exc.Message);
					return false;
					}
				}
			}

		//++Read
		/// <summary>
		/// Read/retrive the DataStatus object from the database
		/// </summary>
		/// <returns>DataStatus object is retrieved if it exist, else null is retured.</returns>
		public DataStatus Read()
			{
			DataStatus result;
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();
					result = (from theStatus in dbSession.AllObjects<DataStatus>() select theStatus).FirstOrDefault();
					dbSession.Commit();
					}
				catch (Exception)
					{
					result = null;
					dbSession.Abort();
					}
				return result;
				}
			}
		#endregion
		}
	}
