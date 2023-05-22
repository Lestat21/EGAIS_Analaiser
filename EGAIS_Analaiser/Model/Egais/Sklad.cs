using System.ComponentModel.DataAnnotations;

namespace EGAIS_Analaiser.Model.Egais
{
    public class Sklad
    {
        [Key]
        public int ID { get; set; }
        public string Name { get; set; } // наименование
        public string WarehouseOwner { get; set; } //владелец
        public string ForestQuarterNumber { get; set; }
        public string TaxationPlotNumber { get; set; }
        public string WarehouseAddress { get; set; }
        public string WarehouseType { get; set; }
        public string ActivityType { get; set; }
        public string LoggingSite { get; set; }
        public double Latitude { get; set; }
        public double Longitude { get; set; }
        public DateTime OpenDate { get; set; }
        public DateTime? CloseDate { get; set; }
        public string Status { get; set; }
        public int SkladId { get; set; }
        public DateTime CreateDate { get; set; }
        public string CreateUser { get; set; }
        public DateTime? ModifyDate { get; set; }
        public string ModifyUser { get; set; }
        public bool WoodHarvestedInRadioactiveContaminationZone { get; set; }
        public string StatusCode { get; set; }
    }
}
