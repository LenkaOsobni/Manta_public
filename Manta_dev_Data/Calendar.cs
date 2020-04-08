namespace Manta_dev_Data
{
    using System;
    using System.ComponentModel.DataAnnotations.Schema;

    [Table("Calendar")]
    public partial class Calendar
    {
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { get; set; }

        public string CAPTION { get; set; }

        public string RELEASE { get; set; }

        [Column(TypeName = "datetime2")]
        public DateTime? ENDTIME { get; set; }

        [Column(TypeName = "datetime2")]
        public DateTime? STARTTIME { get; set; }
    }
}

