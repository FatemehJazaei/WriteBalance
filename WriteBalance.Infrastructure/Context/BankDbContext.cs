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
    public class BankDbContext : DbContext
    {
        public DbSet<FinancialRecord> FinancialRecord { get; set; }
        public BankDbContext(DbContextOptions<BankDbContext> options)
            : base(options)
        {
        }

        public override int SaveChanges() =>
            throw new InvalidOperationException("This context is read-only.");

        public override Task<int> SaveChangesAsync(CancellationToken cancellationToken = default) =>
            throw new InvalidOperationException("This context is read-only.");


        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<FinancialRecord>().HasNoKey();
            // modelBuilder.Entity<FinancialRecord>().ToTable("FinancialBalance").HasKey(x => x.Id);
        }
    }
}
