using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using Dapper;
using DataProcessing.Models;

namespace DataProcessing.Repositories
{

    /*
CREATE TABLE ""TimeStampBlueprint"" (
	""Id""	INTEGER NOT NULL UNIQUE,
	""Time""	NUMERIC NOT NULL,
	""State""	NUMERIC NOT NULL,
	""IsMarker""	NUMERIC NOT NULL,
	PRIMARY KEY(""Id"" AUTOINCREMENT)
);
     */
    class WorkfileRepo : BaseRepo
    {
		public void Create(Workfile workfile)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                using (SQLiteTransaction transaction = conn.BeginTransaction())
                {
                    string tableQuery = $@"
CREATE TABLE ""{workfile.Name}"" (
	""Id""	INTEGER NOT NULL UNIQUE,
	""Time""	NUMERIC NOT NULL,
	""State""	NUMERIC NOT NULL,
	""IsMarker""	NUMERIC NOT NULL,
	PRIMARY KEY(""Id"" AUTOINCREMENT)
);
";
                    conn.Execute(tableQuery, transaction: transaction);
                    conn.Execute("INSERT INTO Workfile (Name, ImportDate) VALUES (@Name, @ImportDate)", new { Name = workfile.Name, ImportDate = workfile.ImportDate }, transaction: transaction);
                    transaction.Commit();
                }
            }
        }
		public List<Workfile> Find()
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                return conn.Query<Workfile>("SELECT Id, Name, ImportDate FROM Workfile ORDER BY date(ImportDate) ASC;").ToList();
            }
        }
        public Workfile FindByName(string name)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                return conn.QuerySingleOrDefault<Workfile>("SELECT Id, Name, ImportDate FROM Workfile WHERE Name=@Name;", new { Name = name });
            }
        }
		public Workfile FindById(int id)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                return conn.QuerySingleOrDefault<Workfile>("SELECT Id, Name, ImportDate FROM Workfile WHERE Id=@Id", 
                    new { Id = id });
            }
        }
		public void Update(Workfile workfile, string oldName)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                using (SQLiteTransaction transaction = conn.BeginTransaction())
                {
                    conn.Execute("UPDATE Workfile SET Name=@Name WHERE Id=@Id",
                        new { Name = workfile.Name, Id = workfile.Id }, transaction: transaction);
                    conn.Execute($"ALTER TABLE '{oldName}' RENAME TO '{workfile.Name}'", transaction: transaction);
                    conn.Execute("UPDATE Workfile SET Name=@Name WHERE Name=@OldName", 
                        new { Name = workfile.Name, OldName = oldName }, transaction: transaction);
                    transaction.Commit();
                }
            }
        }
		public void Delete(Workfile workfile)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                using (SQLiteTransaction transaction = conn.BeginTransaction())
                {
                    conn.Execute($"DROP TABLE '{workfile.Name}'", transaction: transaction);
                    conn.Execute("DELETE FROM Workfile WHERE Id=@Id", new { Id = workfile.Id }, transaction: transaction);
                    transaction.Commit();
                }
            }
        }
    }
}
