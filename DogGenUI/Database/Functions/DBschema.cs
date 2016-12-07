using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Collection;
using VelocityDb.Collection.BTree;
using VelocityDb.Indexing;
using VelocityDb.Session;
using VelocityDb.TypeInfo;
using VelocityDBExtensions;
using DocGeneratorCore.Database.Classes;

namespace DocGeneratorCore.Database.Functions
	{
	class DBSchema
		{
		/// <summary>
		/// This method creates a new Database and returns the result as a boolean.
		/// </summary>
		/// <returns>TRUE if the Database was successfully created; FALSE if an error occurred</returns>
		public static bool CreateDatabaseSchema()
			{
			bool result = false;
			Console.WriteLine("\t\t + CreateDatabaseSchema +");
			
			using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
				{
				try
					{
					if (dbSession.LocateDb(dbNum: 1) != null)
						{
						Console.WriteLine("\t\t - The Database Schema already exist - no need to create it again.");
						result = true;
						}
					else
						{
						dbSession.BeginUpdate();
						dbSession.RegisterClass(type: typeof(Activity));
						dbSession.RegisterClass(type: typeof(ActivityCategory));
						dbSession.RegisterClass(type: typeof(DataStatus));
						dbSession.RegisterClass(type: typeof(Deliverable));
						dbSession.RegisterClass(type: typeof(DeliverableActivity));
						dbSession.RegisterClass(type: typeof(DeliverableServiceLevel));
						dbSession.RegisterClass(type: typeof(DeliverableTechnology));
						dbSession.RegisterClass(type: typeof(ElementDeliverable));
						dbSession.RegisterClass(type: typeof(FeatureDeliverable));
						dbSession.RegisterClass(type: typeof(GlossaryAcronym));
						dbSession.RegisterClass(type: typeof(JobRole));
						dbSession.RegisterClass(type: typeof(Mapping));
						dbSession.RegisterClass(type: typeof(MappingAssumption));
						dbSession.RegisterClass(type: typeof(MappingDeliverable));
						dbSession.RegisterClass(type: typeof(MappingRequirement));
						dbSession.RegisterClass(type: typeof(MappingRisk));
						dbSession.RegisterClass(type: typeof(MappingServiceLevel));
						dbSession.RegisterClass(type: typeof(MappingServiceTower));
						dbSession.RegisterClass(type: typeof(ServiceElement));
						dbSession.RegisterClass(type: typeof(ServiceFamily));
						dbSession.RegisterClass(type: typeof(ServiceFeature));
						dbSession.RegisterClass(type: typeof(ServiceLevel));
						dbSession.RegisterClass(type: typeof(ServiceLevelCategory));
						dbSession.RegisterClass(type: typeof(ServiceLevelTarget));
						dbSession.RegisterClass(type: typeof(ServicePortfolio));
						dbSession.RegisterClass(type: typeof(ServiceProduct));
						dbSession.RegisterClass(type: typeof(TechnologyCategory));
						dbSession.RegisterClass(type: typeof(TechnologyProduct));
						dbSession.RegisterClass(type: typeof(TechnologyVendor));

						dbSession.Commit();
						result = true;
						Console.WriteLine("\t\t = The Database Schema Successfully created.");
						}
					}
				catch (Exception exc)
					{
					Console.WriteLine("### Exception ### " + exc.HResult + " - " + exc.Message);
					dbSession.Abort();
					throw new LocalDatabaseExeption(message: "The following error occurred while creating the new Database: " + exc.HResult + " - " + exc.Message);
					}
				}
			return result;
			}
		}
	}