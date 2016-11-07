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

namespace DocGeneratorCore.Database.Functions
	{
	class DatabaseFunctions
		{
	/// <summary>
	/// This class checks if aspecific database exist and creates a new blank database from scratch with the required information for a 
	/// if it doesn't exist. The database location is determined by the parameter passed to the method and then specified in Propperties.Settings.
	/// Exceptions can occur, threfore test for the following exceptions: 
	/// - CustomeExceptions.DatabaseLocationException
	/// - CustomeExceptions.CreateDatabaseException
	/// </summary>
		public static bool CreateNewDatabase(enumPlatform parPlatform)
			{
			string databaseFolder = string.Empty;
			bool result = false;
			Debug.WriteLine(new string('=', 200));
			Debug.WriteLine("=== Creating a NEW " + parPlatform 
				+ " Database on Host: " + Properties.Settings.Default.CurrentDatabaseHost 
				+ " at Location: " + Properties.Settings.Default.CurrentDatabaseHost + " ===");
			
			//-|If no parameter is passed, create the database in the default location...
			if(string.IsNullOrEmpty(Properties.Settings.Default.CurrentDatabaseHost)
			|| string.IsNullOrEmpty(Properties.Settings.Default.CurrentDatabaseHost))
				{
				return false;
				}
			//---g
			//+| First Delete the current database and copy the license file ...
			Debug.Write("\n + Copying the Database License File...");
			try
				{
				//-|Check if the database folder exist...
				if (Directory.Exists(Properties.Settings.Default.CurrentDatabaseLocation))
					{
					Debug.Write("...Default Database Directory Exist...");
					//-|Delete all files in the Database folder...
					foreach(string file in Directory.EnumerateFiles(Properties.Settings.Default.CurrentDatabaseLocation))
						{
						File.Delete(file);
						}
					Debug.Write("...Deleted database...");
					//-|Copy the VelocityDb license file
					if (File.Exists(Properties.Settings.Default.CurrentDatabaseLocation + "\\" + "4.odb"))
						{
						Debug.Write("...License File Exist...");
						File.Copy(sourceFileName: Properties.Settings.Default.CurrentDatabaseLocation + "\\" + "4.odb",
							destFileName: Properties.Settings.Default.CurrentDatabaseLocation + "\\" + "4.odb", overwrite: true);
						Debug.Write("...License File Copied.");
						}
					else
						{
						Debug.Write("...File DOES NOT Exist!");
						throw new DocumentUploadException("addsdf");
						throw new DatabaseLocationException("The Database License file could not be found in folder: " + Properties.Settings.Default.CurrentDatabaseLocation);
						}
					}
				else
					{
					Debug.Write("...Directory DOES NOT Exist!"); ;
					throw new DatabaseLocationException("The Default Database folder: " + Properties.Settings.Default.CurrentDatabaseLocation + " does not exist!");
					}
				}
			catch (UnauthorizedAccessException exc)
				{
				Debug.Write("\t\t ### Exception ### - " + exc.Message);
				throw new CustomExceptions.DatabaseLocationException("You are not authorised to create files in " + Properties.Settings.Default.CurrentDatabaseLocation);
				}
			catch (PathTooLongException exc)
				{
				Debug.Write("\t\t ### Exception ### - " + exc.Message);
				throw new CustomExceptions.DatabaseLocationException("The Database location's path: " + Properties.Settings.Default.CurrentDatabaseLocation + "is to long");
				}
			catch(DirectoryNotFoundException exc)
				{
				Debug.Write("\t\t ### Exception ### - " + exc.Message);
				throw new CustomExceptions.DatabaseLocationException("The Database folder: " + Properties.Settings.Default.CurrentDatabaseLocation + " does not exist!");
				}
			catch(FileNotFoundException exc)
				{
				Debug.Write("\t\t ### Exception ### - " + exc.Message);
				throw new CustomExceptions.DatabaseLocationException("Database License File could not be found in " + Properties.Settings.Default.CurrentDatabaseLocation);
				}
			catch(IOException exc)
				{
				Debug.Write("\t\t ### Exception ### - " + exc.Message);
				throw new CustomExceptions.DatabaseLocationException("Unexpected Input|Output error occurred. This may be due to a disk or network error.");
				}
			catch(Exception exc)
				{
				Debug.Write("\t\t ### Exception ### - " + exc.Message);
				throw new CustomExceptions.DatabaseLocationException("Unexpected error occurred: #" + exc.HResult + " - " + exc.Message );
				}

			//---g
			//+|Create the primary objects first
			Debug.WriteLine("");
			
			try
				{
				//+|Initialise the **Permissions**
				//Debug.Write(" + Permissions...");
				//VelocityDbList<Permission> permissions = new VelocityDbList<Permission>();
				//Permission permission = new Permission();
				//permissions = permission.Initilize();
				//Debug.Write("...Done!\n");
				////-|Just list all the permissions that was created.
				//foreach (Permission prm in permissions)
				//	{
				//	Debug.WriteLine("\t\t\t - " + prm.Name + " (" + prm.Oid + ") - " + prm.Description);
				//	}

				////+|Initialise the **Users**
				//Debug.Write(" + Users...");
				//User user = new User();
				//user.Initialize(parPermissions: permissions);
				//Debug.Write("...Done!\n");

				//List<User> users = new List<User>();
				//users = User.ReadAll();

				//foreach (User usr in users)
				//	{
				//	Debug.WriteLine("\t\t\t - " + usr.UserName + " (" + usr.Oid + ") - " + usr.Name);
				//	}

				//users = null;

				}
			catch(Exception exc)
				{
				Debug.Write("\n\t\t ### Exception ### - " + exc.HResult + "-" + exc.Message);
				//throw new CustomExceptions.CreateDatabaseException(message: "A error occurred while initialising a new database. Details: ["
				//	+ exc.HResult + "] - " + exc.Message);
				}

			Debug.WriteLine("\n\n=== NEW Database created ###");

			result = true;
			return result;
			}




		}
	}
