using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGeneratorCore
	{
	class LocalDatabaseFunctions
		{

		public static void PopulateLocalDatabase()
			{
			try
				{
				//-|Define the SQLconnection to be initialised
				SqlConnection sddpSQLConnection = new SqlConnection(connectionString: Properties.Settings.Default.CurrentDatabaseLocation);
				sddpSQLConnection.Open();
				Console.WriteLine("Connected to Database: {0}", sddpSQLConnection.Database);
				Console.WriteLine("DataSource...........: {0}", sddpSQLConnection.DataSource);
				Console.WriteLine("Database State.......: {0}", sddpSQLConnection.State);

				//SDDPdatasetTableAdapters.ServicePortfoliosTableAdapter taServicePortfolio = new SDDPdatasetTableAdapters.ServicePortfoliosTableAdapter();
				//taServicePortfolio.Insert(
				//	ID: 1,
				//	Title: "Test Porfolio",
				//	PortfolioType: "Services Framework",
				//	ISDheading: "Test Portfolio",
				//	ISDdescription: string.Empty,
				//	CSDheading: "Test Portfolio",
				//	CSDdescription: string.Empty,
				//	SOWheading: "Test Portfolio",
				//	SOWdescription: string.Empty,
				//	Modified: DateTime.Now);

				////-|the SQL string will be used for the SQL instructions to retrieve data from the Database
				//string sqlString = "";

				////-|Define the SQLDataAdapter object instance which will retrieve the data from the SQL database
				//SqlDataAdapter sddpSQLDataAdapter = new SqlDataAdapter();
				//sddpSQLDataAdapter.SelectCommand = new SqlCommand(cmdText: sqlString, connection: sddpSQLConnection);

				////-|Define the **DataSet** object instance
				//DataSet glossaryDataSet = new DataSet();

				//sddpSQLDataAdapter.Fill(dataSet: glossaryDataSet, srcTable: "GlossaryAcronym");

				//sddpSQLConnection.Close();

				}
			catch(SqlException exc)
				{
				Console.WriteLine("*** SQLexception: {0 - {1}", exc.ErrorCode, exc.Message);
				}
			catch(Exception exc)
				{
				Console.WriteLine("*** Exception: {0 - {1}", exc.HResult, exc.Message);
				}
			}
		}
	}
