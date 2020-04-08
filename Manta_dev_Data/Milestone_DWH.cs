namespace Manta_dev_Data
{
    using System;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.ComponentModel.DataAnnotations;

    [Table("MILESTONE", Schema = "CLOONEY_MANTA_EXPORT")]
    public class Milestone_DWH
    {
        [Key]
        [Column("GUID")]
        public int Guid { get; set; }

        [Column("MILESTONE_NAME")]
        public string MILESTONE_NAME { get; set; }

        [Column("MILESTONE_STARTTIME")]
        public DateTime MILESTONE_STARTTIME { get; set; }

        [Column("MILESTONE_ENDTIME")]
        public DateTime? MILESTONE_ENDTIME { get; set; }

        [Column("RELEASE_NAME")]
        public string RELEASE_NAME { get; set; }
    }
}
