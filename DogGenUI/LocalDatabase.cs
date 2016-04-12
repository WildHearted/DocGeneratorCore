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

		//public async Task LoadSQLiteDatabase(string parString)
          public static bool CreateNewSQLiteDatabase(String parDBFilePath, string parDBFileName)
			{
			string strDB = Path.Combine(parDBFilePath, parDBFileName);

			try
				{
				// Connect to the Database				
				SQLiteConnection objDBconnection = new SQLiteConnection("Data Source = " + strDB);
				// Open the Database Connection for the LocalDatabase
				//await objDBconnection.OpenAsync();
				objDBconnection.Open();

				// Create the SQLiteCommand object with is used to send instructions to SQLite.
				SQLiteCommand objSQLiteCommand = new SQLiteCommand();
				objSQLiteCommand.Connection = objDBconnection;

				// Begin Transaction Processing
				objSQLiteCommand.CommandText = "BEGIN IMMEDIATE TRANSACTION";
				objSQLiteCommand.ExecuteNonQuery();
				//await objSQLiteCommand.ExecuteNonQueryAsync();

				//Create all tables that are required.
				//--- First Table is the LastUpdated table which will keep the Date and Time when the last Synchronisation was done with the SharePoint platform.
				//--- If the LastUpdated Table doesn't exisit, Create it..
				string strSQL = @"CREATE TABLE IF NOT EXISTS LastUpdated (UpdatedOn INTEGER PRIMARY KEY, Succeeded BOOL)";
				objSQLiteCommand.CommandText = strSQL;
				//await objSQLiteCommand.ExecuteNonQueryAsync();
				objSQLiteCommand.ExecuteNonQuery();
				//--- --- Add the first entry to the table
				objSQLiteCommand.Connection = objDBconnection;
				objSQLiteCommand.CommandText = "INSERT INTO LastUpdated (UpdatedOn) values (1)";
				//await objSQLiteCommand.ExecuteNonQueryAsync();
				objSQLiteCommand.ExecuteNonQuery();

				//await InsertTable(objSQLiteCommand, "JobRoles");
				InsertTable(objSQLiteCommand, "JobRoles");
				InsertTable(objSQLiteCommand, "GlossaryAcronyms");
				InsertTable(objSQLiteCommand, "Portfolios");
				InsertTable(objSQLiteCommand, "Families");
				InsertTable(objSQLiteCommand, "Products");
				InsertTable(objSQLiteCommand, "Elements");
				InsertTable(objSQLiteCommand, "Features");
				InsertTable(objSQLiteCommand, "Deliverables");
				InsertTable(objSQLiteCommand, "ElementDeliverables");
				InsertTable(objSQLiteCommand, "FeatureDeliverables");
				InsertTable(objSQLiteCommand, "Activities");
				InsertTable(objSQLiteCommand, "ServiceLevels");
				InsertTable(objSQLiteCommand, "TechnologyProducts");
				InsertTable(objSQLiteCommand, "DeliverableActivities");
				InsertTable(objSQLiteCommand, "DeliverableServiceLevels");
				InsertTable(objSQLiteCommand, "DeliverableTechnologies");

				objSQLiteCommand.CommandText = "COMMIT TRANSACTION";
				objSQLiteCommand.ExecuteNonQuery();
				//await objSQLiteCommand.ExecuteNonQueryAsync();

				// done creating and initialising the local Database...
				objDBconnection.Close();
				objDBconnection.Dispose();
				return true;
				}
			catch(SQLiteException exc)
				{
				Console.WriteLine("\n *** Error Crearing the SQLite Database! {0}\n{1}", exc.HResult, exc.Message);
				return false;
				}
			}


		//public async Task InsertTable(SQLiteCommand parSQLiteCommand, string parTableName)
          public static void InsertTable(SQLiteCommand parSQLiteCommand, string parTableName)
			{
			string strCommandText =	@"CREATE TABLE IF NOT EXISTS " + parTableName + "(ID INTEGER PRIMARY KEY, OBJect BLOB)";
			parSQLiteCommand.CommandText = strCommandText;
			//await parSQLiteCommand.ExecuteNonQueryAsync();
			parSQLiteCommand.ExecuteNonQuery();
               }
		}
	}
