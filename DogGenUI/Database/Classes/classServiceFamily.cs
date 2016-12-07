using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Indexing;
using VelocityDb.Session;

namespace DocGeneratorCore.Database.Classes
	{
	public class ServiceFamily : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a ServiceFamily as mapped to the SharePoint List named ServiceFamilies.
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

		private int? _ServicePortfolioIDsp;
		public int? ServicePortfolioIDsp {
			get { return this._ServicePortfolioIDsp; }
			set { UpdateNonIndexField(); this._ServicePortfolioIDsp = value; }
			}

		private string _ISDheading;
		public string ISDheading {
			get { return this._ISDheading; }
			set {UpdateNonIndexField();this._ISDheading = value;}
			}

		private string _ISDdescription;
		public string ISDdescription {
			get {return this._ISDdescription;}
			set {UpdateNonIndexField();this._ISDdescription = value;}
			}

		private string _CSDheading;
		public string CSDheading {
			get {return this._CSDheading;}
			set {UpdateNonIndexField();this._CSDheading = value;}
			}

		private string _CSDdescription;
		public string CSDdescription {
			get {return this._CSDdescription;}
			set {UpdateNonIndexField();this._CSDdescription = value;}
			}

		private string _SOWheading;
		public string SOWheading {
			get {return this._SOWheading;}
			set {UpdateNonIndexField();this._SOWheading = value;}
			}

		private string _SOWdescription;
		public string SOWdescription {
			get {return this._SOWdescription;}
			set {UpdateNonIndexField();this._SOWdescription = value;}
			}

		#endregion

		#region Methods
		//---g
		//++Store
		/// <summary>
		/// Store/Save a new Object in the database, use the same Store method for New and Updates.
		/// </summary>
		public static bool Store(
			int parIDsp,
			string parTitle,
			int parServicePortfolioIDsp,
			string parISDheading,
			string parISDdescription,
			string parCSDheading,
			string parCSDdescription,
			string parSOWheading,
			string parSOWdescription)
			{
			bool result = false;
			ServiceFamily newEntry;
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<ServiceFamily>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();
					if (newEntry == null)
						newEntry = new ServiceFamily();
					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.ServicePortfolioIDsp = parServicePortfolioIDsp;
					newEntry.ISDheading = parISDheading;
					newEntry.ISDdescription = parISDdescription;
					newEntry.CSDheading = parCSDheading;
					newEntry.CSDdescription = parCSDdescription;
					newEntry.SOWheading = parSOWheading;
					newEntry.SOWdescription = parSOWdescription;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					result = true;
					}
				catch (Exception exc)
					{
					Console.WriteLine("### Exception Database persisting Service Family ### - {0} - {1}", exc.HResult, exc.Message);
					dbSession.Abort();
					}
				return result;
				}
			
			}

		//---g
		//++Read
		/// <summary>
		/// Read/retrieve a specific entry from the database
		/// </summary>
		/// <param name="parIDsp">SharePoint ID of the the entry to retrieve </param>
		/// <returns>object is retrieved if it exist, else null is retured.</returns>
		public static ServiceFamily Read(int parIDsp)
			{
			ServiceFamily result = new ServiceFamily();
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();
					result = (from thisEntry in dbSession.AllObjects<ServiceFamily>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					dbSession.Commit();
					}
				catch (Exception exc)
					{
					result = null;
					Console.WriteLine("### Exception Database reading Service Families [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
					dbSession.Abort();
					}
				}
			return result;
			}

		//---g
		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>DataStatus object is retrieved if it exist, else null is retured.</returns>
		public static List<ServiceFamily> ReadAll()
			{
			List<ServiceFamily> results = new List<ServiceFamily>();
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();
					foreach (ServiceFamily entry in dbSession.AllObjects<ServiceFamily>())
						{
						results.Add(entry);
						}
					dbSession.Commit();
					}
				catch (Exception exc)
					{
					Console.WriteLine("### Exception Database reading all Service Families ### - {0} - {1}", exc.HResult, exc.Message);
					dbSession.Abort();
					}
				}
			return results;
			}
		#endregion
		}
	}
