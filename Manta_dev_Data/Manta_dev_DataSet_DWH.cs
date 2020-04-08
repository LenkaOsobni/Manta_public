namespace Manta_dev_Data
{
    using System.Data.Entity;
    using System.Data.Entity.ModelConfiguration.Conventions;

    public partial class Manta_dev_DataSet_DWH : DbContext
    {

        public Manta_dev_DataSet_DWH()
            : base("name=Manta_dev_DataSet_DWH")
        {
        }

        public virtual DbSet<WP_DWH> WPs_DWH { get; set; }
        public virtual DbSet<Milestone_DWH> Milestones_DWH { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();
            Database.SetInitializer<Manta_dev_DataSet_DWH>(null);
            base.OnModelCreating(modelBuilder);
        }
    }
    
}
