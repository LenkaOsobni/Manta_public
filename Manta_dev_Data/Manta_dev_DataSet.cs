namespace Manta_dev_Data
{
    using System.Data.Entity;

    public partial class Manta_dev_DataSet : DbContext
    {
        public Manta_dev_DataSet()
            : base("name=Manta_dev_DataSet")
        {
        }

        public virtual DbSet<Archiv> Archiv { get; set; }
        public virtual DbSet<Calendar> Calendar { get; set; }
        public virtual DbSet<Setting_Group> Setting_Group { get; set; }
        public virtual DbSet<Settings_Filter> Settings_Filter { get; set; }
        public virtual DbSet<Settings_Name_Columns> Settings_Name_Columns { get; set; }
        public virtual DbSet<Aktive_WP> Aktive_WP { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Aktive_WP>()
                .Property(e => e.RowVersion)
                .IsFixedLength();
        }
    }
}
