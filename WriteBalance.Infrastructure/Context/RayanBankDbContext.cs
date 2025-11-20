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
    public class RayanBankDbContext : DbContext
    {
        public DbSet<RayanFinancialRecord> RayanFinancialBalance { get; set; }
        public RayanBankDbContext(DbContextOptions<RayanBankDbContext> options)
            : base(options)
        {
        }

        public override int SaveChanges() =>
            throw new InvalidOperationException("This context is read-only.");

        public override Task<int> SaveChangesAsync(CancellationToken cancellationToken = default) =>
            throw new InvalidOperationException("This context is read-only.");


        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<RayanFinancialRecord>().HasNoKey();
            modelBuilder.Entity<RayanFinancialRecord>(entity =>
            {
                entity.Property(e => e.Group_code)
                    .HasColumnName("کد گروه");

                entity.Property(e => e.Group_Title)
                      .HasColumnName("نام گروه");

                entity.Property(e => e.Kol_Code)
                      .HasColumnName("حساب کل");

                entity.Property(e => e.Kol_Title)
                      .HasColumnName("نام حساب کل");

                entity.Property(e => e.Moeen_Code)
                      .HasColumnName("حساب معین");

                entity.Property(e => e.Moeen_Title)
                    .HasColumnName("نام حساب معین");

                entity.Property(e => e.Tafsili_Code)
                     .HasColumnName("حساب تفصیلی");

                entity.Property(e => e.Tafsili_Title)
                    .HasColumnName("نام حساب تفصیلی");

                entity.Property(e => e.joze1_Code)
                    .HasColumnName("حساب جز1");

                entity.Property(e => e.joze1_Title)
                    .HasColumnName("نام حساب جز1");

                entity.Property(e => e.joze2_Code)
                    .HasColumnName("حساب جز2");

                entity.Property(e => e.joze2_Title)
                    .HasColumnName("نام حساب جز2");

                entity.Property(e => e.Code_Markaz_Hazineh)
                    .HasColumnName("کد مرکز هزینه");

                entity.Property(e => e.Code_Vahed_Amaliyat)
                    .HasColumnName("کد واحد عملیاتی");

                entity.Property(e => e.Name_Vahed_Amaliyat)
                    .HasColumnName("نام واحد عملیاتی");

                entity.Property(e => e.Code_Parvandeh)
                    .HasColumnName("کد پرونده");

                entity.Property(e => e.Name_Parvandeh)
                    .HasColumnName("نام پرونده");

                entity.Property(e => e.Mandeh_Aval_dore)
                    .HasColumnName("مانده اول دوره");

                entity.Property(e => e.bedehkar)
                    .HasColumnName("بدهکار");

                entity.Property(e => e.bestankar)
                    .HasColumnName("بستانکار");

                entity.Property(e => e.Mande_Bed)
                    .HasColumnName("مانده بدهکار");

                entity.Property(e => e.Mande_Bes)
                    .HasColumnName("مانده بستانکار");
            });

            // modelBuilder.Entity<FinancialRecord>().ToTable("FinancialBalance").HasKey(x => x.Id);

        }
    }
}
