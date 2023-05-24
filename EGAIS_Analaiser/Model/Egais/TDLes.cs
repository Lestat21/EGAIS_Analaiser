using System.ComponentModel.DataAnnotations;

namespace EGAIS_Analaiser.Model.Egais
{
    public class TDLes
    {
        [Key]
        public int ID { get; set; }
        public string? WarehouseOwner { get; set; } // Владелец склада
        public string? DocumentType { get; set; } // Тип документа
        public string? DocumentNumber { get; set; } // Номер документа
        public string? RelatedDocumentNumber { get; set; } // Номер связанного документа
        public DateTime DocumentDate { get; set; } // Дата документа
        public string? Employee { get; set; } // Сотрудник
        public string? OperationWarehouse { get; set; } // Склад операции
        public string? ForestQuarterNumber { get; set; } // Номер лесного квартала
        public string? TaxationPlotNumber { get; set; } // Номер таксационного выдела
        public string? CounterpartyWarehouse { get; set; } // Склад-контрагент
        public string? BasisDocument { get; set; } // Документ основание
        public string? BasisDocumentNumber { get; set; } // Номер основания
        public DateTime? BasisDocumentDate { get; set; } // Дата основания
        public string? Status { get; set; } // Статус
        public string? Shipper { get; set; } // Грузоотправитель
        public string? Carrier { get; set; } // Грузоперевозчик
        public string? Consignee { get; set; } // Грузополучатель
        public string? ConsigneeFlag { get; set; } // Признак грузополучателя
        public string? CreatedByUser { get; set; } // Пользователь создания
        public string? ModifiedByUser { get; set; } // Пользователь изменения
        public DateTime ServerProcessingDateTime { get; set; } // Дата и время обработки на сервере
        public DateTime? ModifiedDate { get; set; } // Дата изменения
        public string? ChangeBasis { get; set; } // Основания внесения изменений
        public string? AdjustmentBasisDocumentNumber { get; set; } // № документа основания корректировки
        public DateTime? AdjustmentBasisDocumentDate { get; set; } // Дата документа основания корректировки
    }
}
