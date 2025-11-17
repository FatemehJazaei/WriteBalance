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
    public class AppDbContext : DbContext
    {
        public DbSet<Period> Periods { get; set; }
        public AppDbContext(DbContextOptions<AppDbContext> options)
            : base(options)
        {
        }

        public override int SaveChanges() =>
            throw new InvalidOperationException("This context is read-only.");

        public override Task<int> SaveChangesAsync(CancellationToken cancellationToken = default) =>
            throw new InvalidOperationException("This context is read-only.");


        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {

            modelBuilder.Entity<Period>()
                        .ToTable("Periods")
                        .HasKey(x => x.Id);

            modelBuilder.Entity<Period>()
                        .Property(x => x.CompanyId)
                        .HasColumnName("CategoryId")
                        .IsRequired();

            modelBuilder.Entity<Period>()
            .Property(x => x.StartDate)
            .HasColumnName("StartDate")
            .IsRequired();

            modelBuilder.Entity<Period>()
            .Property(x => x.TimeEnd)
            .HasColumnName("TimeEnd")
            .IsRequired();
        }
    }
}
