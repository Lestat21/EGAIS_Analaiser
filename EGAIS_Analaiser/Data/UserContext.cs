using EGAIS_Analaiser.Model.Egais;
using EGAIS_Analaiser.Model.IC;
using Microsoft.EntityFrameworkCore;

namespace EGAIS_Analaiser.Data
{
    public class UserContext : DbContext
    {
        public UserContext() => Database.EnsureCreated();

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer("data source=DY-SRV-ASUP\\MYDB;Initial Catalog=EGAIS33; Trusted_Connection=True; TrustServerCertificate=True;Integrated Security = true; MultipleActiveResultSets=true;");
        }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Selling>()
                .Property(p => p.Volume)
                .HasPrecision(18, 3);
        }

        //ЕГАИС
        public DbSet<Remains> Remains { get; set; }
        public DbSet<Selling> Sellings { get; set; }
        public DbSet<Zagotovka> Zagotovkas { get; set; }
        public DbSet<Sklad> Sklads { get; set; }
        public DbSet<TDLes> TDLes { get; set; }


        //1С
        public DbSet<Remains1C> Remains1Cs { get; set; }
        public DbSet<FullShort1C> FullShort1Cs { get; set; }

    }
}
