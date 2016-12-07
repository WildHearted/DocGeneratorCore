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
	public class GlossaryAcronym : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a Glossary&Acronym as mapped to the SharePoint List named GlossaryAcronyms.
		/// </summary>

		#region Properties
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		public int IDsp {
			get { return this._IDsp; }
			set { Update(); this._IDsp = value; }
			}

		private string _Term;
		public string Term {
			get { return this._Term; }
			set { UpdateNonIndexField(); this._Term = value; }
			}

		private string _Meaning;
		public string Meaning {
			get { return this._Meaning; }
			set { UpdateNonIndexField(); this._Meaning = value; }
			}

		private string _Acronym;
		public string Acronym {
			get { return this._Acronym; }
			set { UpdateNonIndexField(); this._Acronym = value; }
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
			string parTerm, 
			string parMeaning, 
			string parAcronym)
			{
			GlossaryAcronym newEntry;
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<GlossaryAcronym>()
							  where objEntry.IDsp == parIDsp
							  select objEntry).FirstOrDefault();
					if (newEntry == null)
						newEntry = new GlossaryAcronym();
					newEntry.IDsp = parIDsp;
					newEntry.Term = parTerm;
					newEntry.Meaning = parMeaning;
					newEntry.Acronym = parAcronym;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					return true;
					}
				catch (Exception exc)
					{
					Console.WriteLine("### Exception Database ### - {0} - {1}", exc.HResult, exc.Message);
					return false;
					}
				}
			}
		
		//---G
		//++Read
		/// <summary>
		/// Read/retrieve a specific entry from the database and return it as an object.
		/// </summary>
		/// <returns>List containing all GlossaryAcronym objects retrieved.</returns>
		public static GlossaryAcronym Read(int parIDsp)
			{
			GlossaryAcronym result = new GlossaryAcronym();

			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();
					result = (from entry in dbSession.AllObjects<GlossaryAcronym>()
							  where entry.IDsp == parIDsp
							  select entry).FirstOrDefault();

					dbSession.Commit();
					}
				catch (Exception exc)
					{
					Console.WriteLine("### Exception Database ### - {0} - {1}", exc.HResult, exc.Message);
					dbSession.Abort();
					}
				}
			return result;
			}

		//---G
		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database and return a List containing all objects.
		/// </summary>
		/// <returns>List containing all GlossaryAcronym objects retrieved.</returns>
		public static List<GlossaryAcronym> ReadAll()
			{
			List<GlossaryAcronym> results = new List<GlossaryAcronym>();
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();

					foreach (GlossaryAcronym entry in dbSession.AllObjects<GlossaryAcronym>())
						{
						results.Add(entry);
						}
					dbSession.Commit();			
					}
				catch (Exception exc)
					{
					Console.WriteLine("### Exception Database ### - {0} - {1}", exc.HResult, exc.Message);
					}
				}
			return results;
			}
		#endregion
		}
	}
