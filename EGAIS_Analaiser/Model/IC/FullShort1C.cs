using System.ComponentModel.DataAnnotations;

namespace EGAIS_Analaiser.Model.IC
{
    public class FullShort1C
    {
        [Key]
        public int ID { get; set; }
        public string Subdivision { get; set; }
        public decimal Procurement { get; set; }
        public decimal SelfConsumption { get; set; }
        public decimal Sale { get; set; }
        public decimal Processing { get; set; }
        public decimal Balance { get; set; }
    }
}
