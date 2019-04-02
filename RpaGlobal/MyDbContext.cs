using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaGlobal
{
    class MyDbContext : DbContext
    {
        public DbSet<RPA> Rpa { get; set; }
        public DbSet<Calcule> Calcule { get; set; }
        public MyDbContext() : base()
        {

        }
    }
}
