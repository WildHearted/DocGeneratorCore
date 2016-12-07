using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Indexing;
using VelocityDb.Session;

namespace DocGeneratorCore.Database.Classes
	{
	public class MappingRisk : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a MappingRisk as mapped to the SharePoint List named MappingServicePowers.
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

		[Index]
		private int? _MappingRequirementIDsp;
		public int? MappingRequirementIDsp {
			get { return this._MappingRequirementIDsp; }
			set { Update(); this._MappingRequirementIDsp = value; }
			}

		private string _Statement;
		public string Statement {
			get { return this._Statement; }
			set { UpdateNonIndexField(); this._Statement = value; }
			}

		private string _Mittigation;
		public string Mittigation {
			get { return this._Mittigation; }
			set { UpdateNonIndexField(); this._Mittigation= value; }
			}

		private double? _ExposureValue;
		public double? ExposureValue {
			get { return this._ExposureValue; }
			set { UpdateNonIndexField(); this._ExposureValue = value; }
			}

		private string _Status;
		public string Status {
			get { return this._Status; }
			set { UpdateNonIndexField(); this._Status = value; }
			}

		private string _Exposure;
		public string Exposure {
			get { return this._Exposure; }
			set { UpdateNonIndexField(); this._Exposure = value; }
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
			int? parMappingRequirementIDsp,
			string parStatement,
			string parMittigation,
			double? parExposureValue,
			string parStatus,
			string parExposure)

			{
			MappingRisk newEntry;
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<MappingRisk>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();

					if (newEntry == null)
						newEntry = new MappingRisk();
					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.MappingRequirementIDsp = parMappingRequirementIDsp;
					newEntry.Statement = parStatement;
					newEntry.Mittigation = parMittigation;
					newEntry._ExposureValue = parExposureValue;
					newEntry.Status = parStatus;
					newEntry.Exposure = parExposure;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					return true;
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					Console.WriteLine("### Exception Database persisting MappingRisk ### - {0} - {1}", exc.HResult, exc.Message);
					return false;
					}
				}
			}

		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>MappingRisk object is retrieved if it exist, else null is retured.</returns>
		public static MappingRisk Read(int parIDsp)
			{
			MappingRisk result = new MappingRisk();
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<MappingRisk>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					dbSession.Commit();
					}
				catch (Exception exc)
					{
					result = null;
					dbSession.Abort();
					Console.WriteLine("### Exception Database reading MappingRisk [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
					}
				}
			return result;
			}

		//++ReadAllForRequirement
		/// <summary>
		/// Read/retrieve all the entries for a specific requirement from the database.
		/// Specify an interger containing the SharePoint ID values of a MappingRequirement to retrieve all the related MappingRisk objects.
		/// </summary>
		/// <param name="parMappingRequirementIDsp">pass an int? of the MappingSRequirement IDsp (SharePoint ID) that need to be retrieved.</param>
		/// <returns>a List<MappingRisk> objects are retrurned.</returns>
		public static List<MappingRisk> ReadAllForRequirement(int? parMappingRequirementIDsp)
			{
			List<MappingRisk> results = new List<MappingRisk>();
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					dbSession.BeginRead();
					//-|Return all MappingRisks if Null is specified
					var mappingRequirements = from thisEntry in dbSession.AllObjects<MappingRisk>()
											where thisEntry.MappingRequirementIDsp == parMappingRequirementIDsp
											select thisEntry;

					foreach (MappingRisk item in mappingRequirements)
						{
						results.Add(item);
						}
					dbSession.Commit();
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					Console.WriteLine("### Exception Database reading all MappingRisk ### - {0} - {1}", exc.HResult, exc.Message);
					}
				}
			return results;
			}

		//+DeleteAll
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

					foreach (MappingRisk entry in dbSession.AllObjects<MappingRisk>())
						{
						dbSession.Unpersist(entry);
						}

					dbSession.Commit();
					result = true;
					}
				catch (Exception exc)
					{
					dbSession.Abort();
					Console.WriteLine("### Exception Database deleting all Mapping Risks  ### - {0} - {1}", exc.HResult, exc.Message);
					result = false;
					}
				}
			return result;
			}
		#endregion
		}
	}
