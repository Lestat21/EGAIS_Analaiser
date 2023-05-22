using System.ComponentModel.DataAnnotations;

namespace EGAIS_Analaiser.Model.IC
{
    public class Remains1C
    {
        [Key]
        public int ID { get; set; }
        public string WarehouseOwner { get; set; }
        public string? Product { get; set; }
        public decimal Remainder { get; set; }
    }


}
