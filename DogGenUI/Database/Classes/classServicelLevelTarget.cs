using System;
using System.Collections.Generic;
using System.Data.Services.Client;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using DocGeneratorCore.SDDPServiceReference;
using VelocityDb;
using VelocityDb.Collection;
using VelocityDb.Indexing;
using VelocityDb.Session;
using VelocityDb.TypeInfo;

namespace DocGeneratorCore.Database.Classes
	{
	public class ServiceLevelTarget : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a ServiceLevelTarget as mapped to the SharePoint List named ServiceLevelTarget.
		/// </summary>

		#region Properties

		[Index]
		[UniqueConstraint]
		private int _IDsp;
		public int IDsp {
			get { return this._IDsp; }
			set { Update(); this._IDsp = value; }
			}

		private string _Type;
		public string Type {
			get { return this._Type; }
			set { UpdateNonIndexField(); this._Type = value; }
			}

		private string _Title;
		public string Title {
			get { return this._Title; }
			set { UpdateNonIndexField(); this._Title = value; }
			}

		private string _ContentStatus;
		public string ContentStatus {
			get { return this._ContentStatus; }
			set { Update(); this._ContentStatus = value; }
			}
		#endregion

		#region Methods
		//++Store
		/// <summary>
		/// Store/Save a new Object in the database, use the same Store method for New and Updates.
		/// </summary>
		public static bool Store(
			int parIDsp,
			string parType,
			string parTitle,
			string parContentStatus)
			{
			GlossaryAcronym newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<GlossaryAcronym>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();
					if (newEntry == null)
						newEntry = new GlossaryAcronym();
					newEntry.IDsp = parIDsp;
					newEntry.Term = parType;
					newEntry.Meaning = parTitle;
					newEntry.Acronym = parContentStatus;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					return true;
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database ### - {0} - {1}", exc.HResult, exc.Message);
				return false;
				}
			}

		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>DataStatus object is retrieved if it exist, else null is retured.</returns>
		public static List<GlossaryAcronym> ReadAll()
			{
			List<GlossaryAcronym> results = new List<GlossaryAcronym>();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					foreach (GlossaryAcronym entry in dbSession.AllObjects<GlossaryAcronym>())
						{
						results.Add(entry);
						}
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}
		#endregion
		}
	}
