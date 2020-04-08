namespace Manta_dev_Data
{
    using System.ComponentModel.DataAnnotations.Schema;
    using System.ComponentModel.DataAnnotations;

    [Table("WP", Schema = "CLOONEY_MANTA_EXPORT")]
    public class WP_DWH
    {
        [Key]
        [Column("WP")]
        public string WP { get; set; }

        [Column("WP_NAME")]
        public string WP_NAME { get; set; }

        [Column("WP_STEP_NAME")]
        public string WP_STEP_NAME { get; set; }

        [Column("PROJECT")]
        public string PROJECT { get; set; }

        [Column("PROJECT_NAME")]
        public string PROJECT_NAME { get; set; }

        [Column("PROJECT_MANAGER_NAME")]
        public string PROJECT_MANAGER_NAME { get; set; }

        [Column("RELEASE_NAME")]
        public string RELEASE_NAME { get; set; }
    }
}
