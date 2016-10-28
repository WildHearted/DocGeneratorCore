using System;
using System.Collections.Generic;
using System.Linq;
using System.Transactions;
using VelocityDb;
using VelocityDb.Collection;
using VelocityDb.Session;
using System.Threading.Tasks;

namespace DocGeneratorCore.Database.Classes
	{
	class DataStatus : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains the date and time 
		/// when the specific database was last synchronised with the SharePoint environment
		/// </summary>
		#region Variables
		private DateTime? _LastRefreshedOn;
		#endregion

		#region Properties
		public DateTime? LastRefreshedOn {
			get { return this._LastRefreshedOn; }
			set { Update(); this._LastRefreshedOn = value; }
			}

		#endregion

		#region Methods
		//++Store
		/// <summary>
		/// Store/Save the date and time that the date and time when the database was last refreshed from SharePoint. 
		/// There is either no objects or a single object in the database, therefore validate for null.
		/// </summary>
		public static bool Store(DateTime parRefreshedOn)
			{
			DataStatus result;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginUpdate();
					result = (from theDBstatus in dbSession.AllObjects<DataStatus>()
							  select theDBstatus).FirstOrDefault();
					if (result == null)
						result = new DataStatus();
					
					result.LastRefreshedOn = parRefreshedOn;
					dbSession.Persist(result);
					dbSession.Commit();
					return true;
					}
				}
			catch (Exception)
				{
				return false;
				}
			}

		//++Read
		/// <summary>
		/// Read/retrive the DataStatus object from the database
		/// </summary>
		/// <returns>DataStatus object is retrieved if it exist, else null is retured.</returns>
		public static DateTime? Read()
			{
			DateTime? result;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					DataStatus entry = (from theStatus in dbSession.AllObjects<DataStatus>()
							 select theStatus).FirstOrDefault();
					result = entry.LastRefreshedOn;
					dbSession.Commit();
					}
				}
			catch (Exception)
				{
				result = null;
				}
			return result;
			}
		#endregion
		}
	}
