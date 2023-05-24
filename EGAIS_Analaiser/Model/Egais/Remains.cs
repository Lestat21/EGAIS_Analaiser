using System.ComponentModel.DataAnnotations;

namespace EGAIS_Analaiser.Model.Egais
{
    public class Remains
    {
        [Key]
        public int ID { get; set; }
        public string? LesHoz { get; set; }
        public string? WarehouseOwner { get; set; }
        public string? MinistryLevel { get; set; }
        public string? GPLHO_Level { get; set; }
        public string? Warehouse { get; set; }
        public string? ForestQuartalNumber { get; set; }
        public string? TaxDivisionNumber { get; set; }
        public string? Nomenclature { get; set; }
        public string? VolumeDeterminationMethod { get; set; }
        public int Quantity { get; set; }
        public decimal Volume { get; set; }
        public string? TreeSpecies { get; set; }
        public string? Assortment { get; set; }
        public string? DiameterGroup { get; set; }
        public string? Diameter { get; set; }
        public string? Length { get; set; }
    }
}
