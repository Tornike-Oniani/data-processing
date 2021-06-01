using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Repositories
{
    interface IRepo<T>
    {
        void Create(T record);
        List<T> Find();
        T FindById(int id);
        void Update(T record);
        void Delete(T record);
    }
}
