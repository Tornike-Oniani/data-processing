using DataProcessing.Repositories;
using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Models
{
    class Workfile
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string ImportDate { get; set; }
        public int Sheets { get; set; }
    }
}
