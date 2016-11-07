﻿using System;
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
		#region Variables
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		private string _Title;
		[Index]
		private ServicePortfolio _ServicePortfolio;
		private string _ISDheading;
		private string _ISDdescription;
		private string _CSDheading;
		private string _CSDdescription;
		private string _SOWheading;
		private string _SOWdescription;
		#endregion

		#region Properties
		public int IDsp {
			get { return this._IDsp; }
			set { Update(); this._IDsp = value; }
			}
		public string Title {
			get { return this._Title; }
			set { Update(); this._Title = value; }
			}
		public ServicePortfolio ServicePortfolio {
			get { return this._ServicePortfolio; }
			set { Update(); this._ServicePortfolio = value; }
			}
		public string ISDheading {
			get { return this._ISDheading; }
			set {Update();this._ISDheading = value;}
			}
		public string ISDdescription {
			get {return this._ISDdescription;}
			set {Update();this._ISDdescription = value;}
			}
		public string CSDheading {
			get {return this._CSDheading;}
			set {Update();this._CSDheading = value;}
			}
		public string CSDdescription {
			get {return this._CSDdescription;}
			set {Update();this._CSDdescription = value;}
			}
		public string SOWheading {
			get {return this._SOWheading;}
			set {Update();this._SOWheading = value;}
			}
		public string SOWdescription {
			get {return this._SOWdescription;}
			set {Update();this._SOWdescription = value;}
			}

		#endregion

		#region Methods
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
			ServiceFamily newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<ServiceFamily>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();
					if (newEntry == null)
						newEntry = new ServiceFamily();
					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					//-|Use the **ServicePortfolioIDsp** to retrieve the ServicePortfolio Object instance.
					newEntry.ServicePortfolio = ServicePortfolio.Read(parIDsp: parServicePortfolioIDsp);
					newEntry.ISDheading = parISDheading;
					newEntry.ISDdescription = parISDdescription;
					newEntry.CSDheading = parCSDheading;
					newEntry.CSDdescription = parCSDdescription;
					newEntry.SOWheading = parSOWheading;
					newEntry.SOWdescription = parSOWdescription;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					return true;
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database persisting Service Family ### - {0} - {1}", exc.HResult, exc.Message);
				return false;
				}
			}

		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>DataStatus object is retrieved if it exist, else null is retured.</returns>
		public static ServiceFamily Read(int parIDsp)
			{
			ServiceFamily result = new ServiceFamily();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<ServiceFamily>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					}
				}
			catch (Exception exc)
				{
				result = null;
				Console.WriteLine("### Exception Database reading Service Families [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
				}
			return result;
			}

		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>DataStatus object is retrieved if it exist, else null is retured.</returns>
		public static List<ServiceFamily> ReadAll()
			{
			List<ServiceFamily> results = new List<ServiceFamily>();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					foreach (ServiceFamily entry in dbSession.AllObjects<ServiceFamily>())
						{
						results.Add(entry);
						}
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all Service Families ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}
		#endregion
		}
	}
