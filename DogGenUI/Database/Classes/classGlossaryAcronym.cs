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
		#region Variables
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		[Index]
		private string _Term;
		private string _Meaning;
		[Index]
		private string _Acronym;
		#endregion

		#region Properties
		public int IDsp {
			get { return this._IDsp; }
			set { Update(); this._IDsp = value; }
			}
		public string Term {
			get { return this._Term; }
			set { Update(); this._Term = value; }
			}
		public string Meaning {
			get { return this._Meaning; }
			set { UpdateNonIndexField(); this._Meaning = value; }
			}
		public string Acronym {
			get { return this._Acronym; }
			set { Update(); this._Acronym = value; }
			}
		#endregion

		#region Methods
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
					newEntry.Term = parTerm;
					newEntry.Meaning = parMeaning;
					newEntry.Acronym = parAcronym;
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
