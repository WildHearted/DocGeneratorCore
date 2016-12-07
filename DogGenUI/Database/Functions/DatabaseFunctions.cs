using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VelocityDb;
using VelocityDb.Session;
using VelocityDb.Collection;
using VelocityDb.TypeInfo;
using DocGeneratorCore.Database.Classes;
using DocGeneratorCore.Database.Functions;

namespace DocGeneratorCore.Database.Functions
	{

	class DatabaseFunctions
		{
		#region Methods
		//===G
		//++DoesLocalDatabaseExist
		//---G
		/// <summary>
		/// Determine if the relevant Local database exist. If it doesn't exist, the method will create and initialise it.
		/// Ensure that the Properties.Settings.Default.CurrentPlatform and ...CurrentDatabseHost and ...CurrentDatabaseLocation is set before calling this method.
		/// </summary>
		/// <returns>TRUE if the database exist, FALSE if it doesn't exit and could not be created and initialised.</returns>
		public static bool DoesLocalDatabaseExist()
			{
			bool result = false;

			Console.WriteLine("Check if the local database exist: " + Properties.Settings.Default.CurrentPlatform);

			if (string.IsNullOrEmpty(Properties.Settings.Default.CurrentPlatform)
			|| string.IsNullOrEmpty(Properties.Settings.Default.CurrentDatabaseHost)
			|| string.IsNullOrEmpty(Properties.Settings.Default.CurrentDatabaseLocation))
				{
				Console.WriteLine("\t - One of these values are blank...");
				Console.WriteLine("\t\t - CurrentDatabaseHost....: " + Properties.Settings.Default.CurrentDatabaseHost);
				Console.WriteLine("\t\t - CurrentPlatform........: " + Properties.Settings.Default.CurrentPlatform);
				Console.WriteLine("\t\t - CurrentDatabaseLocation: " + Properties.Settings.Default.CurrentDatabaseLocation);
				return result;
				}

			Console.WriteLine(new string('-', 150));
			Console.WriteLine("\t Check if Local " + Properties.Settings.Default.CurrentPlatform 
				+ " Database exist at Location " + Properties.Settings.Default.CurrentDatabaseLocation
				+ " on Host: " + Properties.Settings.Default.CurrentDatabaseHost);

			try
				{
				//-|Check if the database folder exist...
				if (Directory.Exists(Properties.Settings.Default.CurrentDatabaseLocation))
					{
					Console.WriteLine("\t + The Database Directory Exist...");
					//-|Does the database Files exist?
					int fileCount = Directory.GetFiles(path: Properties.Settings.Default.CurrentDatabaseLocation).Count();
					if (fileCount > 0)
						{
						Console.WriteLine("\t + " + fileCount + " Database files found...");
						result = true;
						}
					else //-| There are no files in the folder, meaning the database doesn't exist yet.
						{
						Console.WriteLine("\t - The local database does not exist in folder: " + Properties.Settings.Default.CurrentDatabaseLocation);
						if (CreateNewDatabase())
							{//-|Successfull!
							Console.WriteLine("\t + Database created...");
							result = true;
							}
						else
							{//-|Failed
							Console.WriteLine("\t - Database creation failed.");
							result = false;
							}
						}
					}
				else //-|The database directory does **NOT** exist, therefore crete the database in the folder
					{
					Console.WriteLine("\t + The Local Database Folder DOES NOT Exist!");
					//-|Check if the VelocityDb license file exist
					string databaseLicenseFolder = AppDomain.CurrentDomain.BaseDirectory + Properties.Settings.Default.DatabaseLocationLicense;
					if (File.Exists(databaseLicenseFolder + "\\" + "4.odb"))
						{
						Console.WriteLine("\t + Database License File EXIST in Database License Location: {0}", databaseLicenseFolder);
						}
					else
						{
						Console.WriteLine("\t ### Exception ### the Database License File DOES NOT Exist! in {0}", databaseLicenseFolder);
						throw new LocalDatabaseExeption("The Database License file could not be found in folder: " + databaseLicenseFolder
							+ " which means the Local database could not be created)");
						}
					//-|Create the Directory for the Local Database
					Directory.CreateDirectory(path: Properties.Settings.Default.CurrentDatabaseLocation);

					//-|Copy the LicenseFile to the new Database location
					File.Copy(sourceFileName: databaseLicenseFolder + "\\" + "4.odb",
						destFileName: Properties.Settings.Default.CurrentDatabaseLocation + "\\" + "4.odb", overwrite: true);
					Console.WriteLine("\t + Database License File Copied to {0}", Properties.Settings.Default.CurrentDatabaseLocation);

					if (CreateNewDatabase())
						{//-|Successfull!
						Console.WriteLine("\t + Database created...");
						result = true;
						}
					else
						{//-|Failed!
						Console.WriteLine("\t - Database creation failed.");
						result = false;
						}
					}
				}
			catch (UnauthorizedAccessException exc)
				{
				Console.Write("\t\t ### Exception ### - " + exc.Message);
				throw new LocalDatabaseExeption("You are not authorised to create files in " + Properties.Settings.Default.CurrentDatabaseLocation);
				}
			catch (PathTooLongException exc)
				{
				Console.Write("\t\t ### Exception ### - " + exc.Message);
				throw new LocalDatabaseExeption("The Database location's path: " + Properties.Settings.Default.CurrentDatabaseLocation + "is to long");
				}
			catch (DirectoryNotFoundException exc)
				{
				Console.Write("\t\t ### Exception ### - " + exc.Message);
				throw new LocalDatabaseExeption("The Database folder: " + Properties.Settings.Default.CurrentDatabaseLocation + " does not exist!");
				}
			catch (FileNotFoundException exc)
				{
				Console.Write("\t\t ### Exception ### - " + exc.Message);
				throw new LocalDatabaseExeption("Database License File could not be found in " + Properties.Settings.Default.CurrentDatabaseLocation);
				}
			catch (IOException exc)
				{
				Console.Write("\t\t ### Exception ### - " + exc.Message);
				throw new LocalDatabaseExeption("Unexpected Input|Output error occurred. This may be due to a disk or network error.");
				}
			catch (LocalDatabaseCreationExeption exc)
				{
				Console.Write("\t\t ### Exception ### - " + exc.Message);
				throw new LocalDatabaseExeption("Unable to create the Local Database. " + exc.Message);
				}
			catch (Exception exc)
				{
				Console.Write("\t\t ### Exception ### - " + exc.Message);
				throw new LocalDatabaseExeption("Unexpected error occurred: #" + exc.HResult + " - " + exc.Message);
				}
			Console.WriteLine(new string('-', 150));
			return result;
			}


		//===G
		//++CreateNewDatabase
		//===G
		/// <summary>
		/// This method creates a new blank database from scratch with the required information.
		/// The database location is determined by the value in the parDatabaseLocation, therefore ensure that the Directory exist 
		/// before calling this method. Use the DoesLocalDatabaseExist method to ensure the direcotry exist.
		/// Exceptions can occur, threfore test for the following exceptions: 
		/// - CustomeExceptions.LocalDatabaseCreationException
		/// </summary>
		public static bool CreateNewDatabase()
			{
			bool result = false;
			//-|If no parameter is passed, create the database in the default location...
			if(string.IsNullOrEmpty(Properties.Settings.Default.CurrentDatabaseLocation))
				{
				return result;
				}
			
			//+|Create the Database Schema
			Console.WriteLine("\t + Creating the Database Schema...");
			
			try
				{
				//-|Create the database schema
				if (DBSchema.CreateDatabaseSchema())
					{ //-|Database Schema successfully created
					Console.WriteLine("\t + Database Schema successfully created!");
					//-|Initialise the **DataStatus**
					Console.Write("\t + Store DataStatus...");
					DataStatus dataStatus = new DataStatus();
					if (DataStatus.Store(parCreatedOn: DateTime.UtcNow, parRefreshedOn: new DateTime(2000, 1, 1, 0, 0, 0)))
						{
						Console.Write("...Done!\n");
						result = true;
						}
					else
						{
						Console.WriteLine("\nUnable to update the DataStatus ... Failed!\n");
						result = false;
						}
					}
				else
					{ //-|Database Schama creation failed!
					Console.WriteLine("\t\t\t - Database Schema creation failed!");
					result = false;
					}
				}
			catch(Exception exc)
				{
				Console.Write("\n\t\t ### Exception ### - " + exc.HResult + "-" + exc.Message);
				throw new LocalDatabaseExeption(message: "A error occurred while initialising a new database. Details: ["
					+ exc.HResult + "] - " + exc.Message);
				}
			return result;
			}
		#endregion

		}
	}
