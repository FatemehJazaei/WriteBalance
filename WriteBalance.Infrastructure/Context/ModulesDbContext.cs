using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Domain.Entities;

namespace WriteBalance.Infrastructure.Context
{
    public class ModulesDbContext : DbContext
    {
        public DbSet<PooyaCoding> PooyaCodings { get; set; }
        public ModulesDbContext(DbContextOptions<ModulesDbContext> options)
            : base(options)
        {
        }

        public override int SaveChanges() =>
            throw new InvalidOperationException("This context is read-only.");

        public override Task<int> SaveChangesAsync(CancellationToken cancellationToken = default) =>
            throw new InvalidOperationException("This context is read-only.");


        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<PooyaCoding>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("Refah_PooyaCoding", "dbo");

                entity.Property(x => x.CodeKol)
                    .HasColumnName("CodeKol")
                    .IsRequired();

                entity.Property(x => x.CodeArz)
                    .HasColumnName("CodeArz")
                    .IsRequired();

                entity.Property(x => x.GroupMoein)
                    .HasColumnName("GroupMoein")
                    .IsRequired();

                entity.Property(x => x.CodeOmoorMali)
                    .HasColumnName("CodeOmoorMali")
                    .IsRequired();
            });
        }
    }
}
