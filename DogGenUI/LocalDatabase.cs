using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Data.SQLite;
using System.Data.SQLite.Generic;
using System.Data.SQLite.Linq;
using System.Data.SQLite.EF6;
using System.Text;
using System.Threading.Tasks;

namespace DocGenerator
	{
	class LocalDatabase
		{

		public async Task LoadSQLiteDatabase(string parString)
			{

			string strDBfilePath = Environment.CurrentDirectory;
			string strDBfileName = Properties.AppResources.localDatabaseName;
			string strDB = Path.Combine(strDBfilePath, strDBfileName);
			//Check if the Database exist, if Not create it...
			if(!File.Exists(strDB))
				{
				SQLiteConnection.CreateFile(databaseFileName: strDB);
				}
			// Connect to the Database				
			SQLiteConnection objDBconnection = new SQLiteConnection("Data Source = " + strDB);
			// Open the Database Connection for the LocalDatabase

			await objDBconnection.OpenAsync();

			SQLiteCommand objSQLiteCommand = new SQLiteCommand();
			objSQLiteCommand.Connection = objDBconnection;
			// Begin Transaction Processing
			objSQLiteCommand.CommandText = "BEGIN TRANSACTION";
			await objSQLiteCommand.ExecuteNonQueryAsync();

			//Create all tables in the Database;
			//--- First Table is the LastUpdated table which will keep the Date and Time when the last Synchronisation was done with the SharePoint platform.
			//--- If the LastUpdated Table doesn't exisit, Create it..
			string strSQL = @"CREATE TABLE IF NOT EXISTS LastUpdated
						(UpdatedOn INTEGER PRIMARY KEY,
						 Succeeded BOOL)";
			objSQLiteCommand.CommandText = strSQL;
			await objSQLiteCommand.ExecuteNonQueryAsync();
			//--- --- Add the first entry to the table
			objSQLiteCommand.CommandText = "INSERT INTO LastUpdated (UpdatedOn, Succeeded) values (0, false)";
			await objSQLiteCommand.ExecuteNonQueryAsync();

			await InsertTable(objSQLiteCommand, "JobRoles");
			await InsertTable(objSQLiteCommand, "GlossaryAcronyms");
			await InsertTable(objSQLiteCommand, "Portfolios");
			await InsertTable(objSQLiteCommand, "Families");
			await InsertTable(objSQLiteCommand, "Products");
			await InsertTable(objSQLiteCommand, "Elements");
			await InsertTable(objSQLiteCommand, "Features");
			await InsertTable(objSQLiteCommand, "Deliverables");
			await InsertTable(objSQLiteCommand, "ElementDeliverables");
			await InsertTable(objSQLiteCommand, "FeatureDeliverables");
			await InsertTable(objSQLiteCommand, "Activities");
			await InsertTable(objSQLiteCommand, "ServiceLevels");
			await InsertTable(objSQLiteCommand, "TechnologyProducts");
			await InsertTable(objSQLiteCommand, "DeliverableActivities");
			await InsertTable(objSQLiteCommand, "DeliverableServiceLevels");
			await InsertTable(objSQLiteCommand, "Deliverabletechnologies");

			objSQLiteCommand.CommandText = "COMMIT TRANSACTION";
			await objSQLiteCommand.ExecuteNonQueryAsync();

			// done creating and initialising the local Database...

			}
		public async Task InsertTable(SQLiteCommand parSQLiteCommand, string parTableName)
			{
			string strCommandText =	@"CREATE TABLE IF NOT EXISTS " + parTableName +
								"(ID INTEGER PRIMARY KEY, OBJect BLOB)";
			parSQLiteCommand.CommandText = strCommandText;
			await parSQLiteCommand.ExecuteNonQueryAsync();

			await parSQLiteCommand

			}
		}
	}
