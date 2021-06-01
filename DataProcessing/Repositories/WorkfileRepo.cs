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
CREATE TABLE ""TimeStamp"" (
	""Id""	INTEGER NOT NULL UNIQUE,
	""A""	NUMERIC,
	""B""	NUMERIC,
	""C""	NUMERIC,
	""D""	NUMERIC,
	""State""	INTEGER,
	""Counter""	INTEGER,
	""PreviousId""	INTEGER UNIQUE,
	PRIMARY KEY(""Id"" AUTOINCREMENT),
	FOREIGN KEY(""PreviousId"") REFERENCES ""TimeStamp""(""Id"") ON DELETE SET NULL
);
     */
	class WorkfileRepo : BaseRepo
    {
		public void Create(string name)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                using (SQLiteTransaction transaction = conn.BeginTransaction())
                {
                    string tableQuery = $@"
CREATE TABLE ""{name}"" (
	""Id""	INTEGER NOT NULL UNIQUE,
	""A""	NUMERIC,
	""B""	NUMERIC,
	""C""	NUMERIC,
	""D""	NUMERIC,
	""State""	INTEGER,
	""Counter""	INTEGER,
	""PreviousId""	INTEGER UNIQUE,
	PRIMARY KEY(""Id"" AUTOINCREMENT),
	FOREIGN KEY(""PreviousId"") REFERENCES ""{name}""(""Id"") ON DELETE SET NULL
);
";
                    conn.Execute(tableQuery, transaction: transaction);
                    conn.Execute("INSERT INTO Workfile (Name) VALUES (@Name)", new { Name = name }, transaction: transaction);
                    transaction.Commit();
                }
            }
        }
		public List<Workfile> Find()
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                return conn.Query<Workfile>("SELECT Id, Name FROM Workfile").ToList();
            }
        }
		public Workfile FindById(int id)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                return conn.QuerySingleOrDefault<Workfile>("SELECT Id, Name FROM Workfile WHERE Id=@Id", new { Id = id });
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
