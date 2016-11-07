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
		#region Variables
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		private string _Title;
		[Index]
		private int? _MappingRequirementIDsp;
		private string _Statement;
		private string _Mittigation;
		private double? _ExposureValue;
		private string _ComplianceComments;
		private string _Status;
		private string _Exposure;
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
		public int? MappingRequirementIDsp {
			get { return this._MappingRequirementIDsp; }
			set { Update(); this._MappingRequirementIDsp = value; }
			}
		public string Statement {
			get { return this._Statement; }
			set { UpdateNonIndexField(); this._Statement = value; }
			}
		public string Mittigation {
			get { return this._Mittigation; }
			set { UpdateNonIndexField(); this._Mittigation= value; }
			}
		public double? ExposureValue {
			get { return this._ExposureValue; }
			set { UpdateNonIndexField(); this._ExposureValue = value; }
			}
		public string Status {
			get { return this._Statement; }
			set { UpdateNonIndexField(); this._Status = value; }
			}
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
			string parExposure
			)

			{
			MappingRisk newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
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
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database persisting MappingRisk ### - {0} - {1}", exc.HResult, exc.Message);
				return false;
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
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<MappingRisk>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					}
				}
			catch (Exception exc)
				{
				result = null;
				Console.WriteLine("### Exception Database reading MappingRisk [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
				}
			return result;
			}

		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database.
		/// Specify an interger containing the SharePoint ID values of a MaapingRequirement to retrieve all the related MappingRisk objects.
		/// </summary>
		/// <param name="parMappingRequirementIDsp">pass an int? of the MappingSRequirement IDsp (SharePoint ID) that need to be retrieved.
		/// If all MappingRisks must be retrieve, pass a null value as the parameter to return all objects.</param>
		/// <returns>a List<MappingRisk> objects are retrurned.</returns>
		public static List<MappingRisk> ReadAll(int? parMappingRequirementIDsp)
			{
			List<MappingRisk> results = new List<MappingRisk>();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Return all MappingRisks if Null is specified
					if (parMappingRequirementIDsp == null)
						{
						var mappingRequirements = from thisEntry in dbSession.AllObjects<MappingRisk>() select thisEntry;
						if (mappingRequirements.Count() > 0)
							{
							foreach (MappingRisk item in mappingRequirements)
								{
								results.Add(item);
								}
							}
						}
					else
						{
						var mappingRequirements = from thisEntry in dbSession.AllObjects<MappingRisk>()
												  where thisEntry.MappingRequirementIDsp == parMappingRequirementIDsp
												  select thisEntry;
						if (mappingRequirements.Count() > 0)
							{
							foreach (MappingRisk item in mappingRequirements)
								{
								results.Add(item);
								}
							}
						}
					return results;
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all MappingRisk ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}
		#endregion
		}
	}
