using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using Dapper;
using DataProcessing.Models;
using DataProcessing.Classes;

namespace DataProcessing.Repositories
{
    class TimeStampRepo : BaseRepo
    {
        // Private attributes
        private string table;

        // Constructor
        public TimeStampRepo()
        {
            if (WorkfileManager.GetInstance().SelectedWorkFile == null) throw new Exception("Can not operate data without active workfile");

            table = $"'{WorkfileManager.GetInstance().SelectedWorkFile.Name}'";
        }

        // CRUD operations
        public void Create(TimeStamp record)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                long ticks = record.Time.Ticks;
                conn.Execute($"INSERT INTO {table} (Time, State, IsMarker) VALUES (@Time, @State, @IsMarker);",
                    new { Time = ticks, State = record.State, IsMarker = record.IsMarker });
            }
        }
        public void CreateMany(List<TimeStamp> records)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                using (SQLiteTransaction transaction = conn.BeginTransaction())
                {
                    long ticks;
                    foreach (TimeStamp record in records)
                    {
                        ticks = record.Time.Ticks;
                        conn.Execute($"INSERT INTO {table} (Time, State, IsMarker) VALUES (@Time, @State, @IsMarker);",
                            new { Time = ticks, State = record.State, IsMarker = record.IsMarker }, transaction: transaction);
                    }
                    transaction.Commit();
                }
            }
        }
        public void Update(TimeStamp record)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                long ticks = record.Time.Ticks;
                conn.Execute($"UPDATE {table} SET Time=@Time, State=@State WHERE Id=@Id;",
                    new { Time = ticks, State = record.State, Id = record.Id });
            }
        }
        public List<TimeStamp> Find()
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                string query = $@"SELECT Id, Time AS TimeTicks, State, IsMarker FROM {table};";
                return conn.Query<TimeStamp>(query).ToList();
            }
        }
    }
}
