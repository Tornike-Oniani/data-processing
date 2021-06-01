using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Repositories
{
    class BaseRepo
    {
        protected string connectionString = $"Data Source={Path.Combine(Environment.CurrentDirectory, "database.sqlite3;Version=3;")}";
    }
}
