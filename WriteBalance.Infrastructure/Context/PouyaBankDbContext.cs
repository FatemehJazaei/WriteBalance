using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Domain.Entities;

namespace WriteBalance.Infrastructure.Context
{
    public class PouyaBankDbContext : DbContext
    {
        public DbSet<PouyaFinancialRecord> PouyaFinancialBalance { get; set; }
        public PouyaBankDbContext(DbContextOptions<PouyaBankDbContext> options)
            : base(options)
        {
        }

        public override int SaveChanges() =>
            throw new InvalidOperationException("This context is read-only.");

        public override Task<int> SaveChangesAsync(CancellationToken cancellationToken = default) =>
            throw new InvalidOperationException("This context is read-only.");


        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<PouyaFinancialRecord>().HasNoKey();
            modelBuilder.Entity<PouyaFinancialRecord>(entity =>
            {
                entity.Property(e => e.Taraz_Date)
                    .HasColumnName("trz_dt");

                //entity.Property(e => e.Code_shobeh)
                //.HasColumnName("brn_cod");

                entity.Property(e => e.Kol_Code_Markazi)
                      .HasColumnName("cntrlbmi");

                entity.Property(e => e.Kol_Title)
                      .HasColumnName("cntrlfdsc");

                entity.Property(e => e.Hesab_Code)
                    .HasColumnName("memcod");

                entity.Property(e => e.Kol_Code)
                    .HasColumnName("cntrl");

                entity.Property(e => e.Arz_Code)
                    .HasColumnName("curr");

                entity.Property(e => e.Moeen_Code)
                      .HasColumnName("Cust_kd");

                entity.Property(e => e.Moeen)
                     .HasColumnName("Cust_no");

                entity.Property(e => e.Tafzili)
                     .HasColumnName("detail");

                entity.Property(e => e.Code_Arz_Abbr)
                    .HasColumnName("abbr");

                entity.Property(e => e.Sharh_Arz)
                    .HasColumnName("abbrfdsc");

                entity.Property(e => e.Mande_Bed_arzi)
                    .HasColumnName("drbal");

                entity.Property(e => e.Mande_Bes_arzi)
                    .HasColumnName("crbal");

                entity.Property(e => e.Mande_Bed_rial)
                    .HasColumnName("drbaleq");

                entity.Property(e => e.Mande_Bes_rial)
                    .HasColumnName("crbaleq");

                entity.Property(e => e.Gardersh_Bed_rial)
                    .HasColumnName("drbsequv");

                entity.Property(e => e.Gardersh_Bes_rial)
                    .HasColumnName("crbsequv");

                entity.Property(e => e.Gardersh_Bed_arzi)
                    .HasColumnName("dramnt");

                entity.Property(e => e.Gardersh_Bes_arzi)
                    .HasColumnName("cramnt");
                
            });

            // modelBuilder.Entity<FinancialRecord>().ToTable("FinancialBalance").HasKey(x => x.Id);

        }
    }
}
