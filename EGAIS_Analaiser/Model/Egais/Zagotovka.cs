using System.ComponentModel.DataAnnotations;

namespace EGAIS_Analaiser.Model.Egais
{
    public class Zagotovka
    {
        [Key]
        public int ID { get; set; }
        public decimal OperationalAccountingTotal { get; set; }
        public string LoggingMethod { get; set; }
        public string Forestry { get; set; }
    }
}
