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
    class SampleRepo : BaseRepo
    {
        // Private attributes
        private string table;

        // Constructor
        public SampleRepo()
        {
            if (WorkfileManager.GetInstance().SelectedWorkFile == null) throw new Exception("Can not do data operation without workfile");

            table = $"'{WorkfileManager.GetInstance().SelectedWorkFile.Name}'";
        }

        // CRUD operations
        public void Create(DataSample record)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                int counter = GetCounter(conn);
                counter++;
                long tickA = record.AT.Ticks;
                long tickB = record.BT.Ticks;
                int prevId = GetLastRecordId(conn);
                conn.Execute($"INSERT INTO {table} (A, B, C, D, State, IsMarker Counter, PreviousId) VALUES (@A, @B, @C, @D, @State, @IsMarker, @Counter, @PreviousId);",
                    new { A = tickA, B = tickB, C = record.C, D = record.D, State = record.State, IsMarker = record.IsMarker, Counter = counter, PreviousId = prevId });
            }
        }
        public void CreateMany(List<DataSample> records)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                using(SQLiteTransaction transaction = conn.BeginTransaction())
                {
                    foreach (DataSample record in records)
                    {
                        record.CalculateStats(true);
                        int counter = GetCounter(conn);
                        counter++;
                        long tickA = record.AT.Ticks;
                        long tickB = record.BT.Ticks;
                        int prevId = GetLastRecordId(conn);
                        conn.Execute($"INSERT INTO {table} (A, B, C, D, State, IsMarker, Counter, PreviousId) VALUES (@A, @B, @C, @D, @State, @IsMarker, @Counter, @PreviousId);",
                            new { A = tickA, B = tickB, C = record.C, D = record.D, State = record.State, IsMarker = record.IsMarker, Counter = counter, PreviousId = prevId }, transaction: transaction);
                    }

                    transaction.Commit();
                }
            }
        }
        public void Update(DataSample record)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                long tickA = record.AT.Ticks;
                long tickB = record.BT.Ticks;
                //int prevId = GetLastRecordId(conn);
                conn.Execute($"UPDATE {table} SET A=@A, B=@B, C=@C, D=@D, State=@State WHERE Id=@Id;",
                    new { A = tickA, B = tickB, C = record.C, D = record.D, State = record.State, Id = record.Id });
            }
        }
        public List<DataSample> Find()
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                string query = $@"SELECT t.Id, t.A, t.B, t.C, t.D, t.State, t.IsMarker, t.Counter, tt.A AS PrevA
FROM {table} AS t
LEFT JOIN {table} AS tt ON t.PreviousId = tt.Id;";
                return   conn.Query<DataSample>(query).ToList();
            }
        }

        public DataSample GetLastRecord()
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                return conn.QuerySingleOrDefault<DataSample>($"SELECT A, B, C, D, State, Counter FROM {table} ORDER BY Id DESC LIMIT 1;");
            }
        }
        public DataSample GetPreviousRecord(int counter)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                return conn.QuerySingleOrDefault<DataSample>($"SELECT A, B, C, D, State, Counter FROM {table} WHERE Counter<@Counter ORDER BY Counter DESC LIMIT 1;", 
                    new { Counter = counter });
            }
        }
        public DataSample GetLinkedRecord(DataSample record)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                return conn.QuerySingleOrDefault<DataSample>($"SELECT Id, A, B, C, D, State, Counter FROM {table} WHERE PreviousId=@Id", 
                    new { Id = record.Id });
            }
        }

        // Private helpers
        private int GetLastRecordId(SQLiteConnection conn)
        {
                string query = $"SELECT Id FROM {table} ORDER BY Id DESC LIMIT 1;";
                return conn.QuerySingleOrDefault<int>(query);
        }
        private int GetCounter(SQLiteConnection conn)
        {
            return conn.QuerySingleOrDefault<int>($"SELECT Counter FROM {table} ORDER BY Counter DESC LIMIT 1;");
        }
    }
}
