using System.ComponentModel.DataAnnotations;

namespace EGAIS_Analaiser.Model.Egais
{
    public class Selling
    {
        [Key]
        public int ID { get; set; }
        public string Division { get; set; } // Структурное подразделение
        public string WarehouseOwner { get; set; } // Владелец склада
        public string ForestQuartNumber { get; set; } // Номер лесного квартала
        public string TaxationAreaNumber { get; set; } // Номер таксационного выдела
        public string OperationWarehouse { get; set; } // Склад операции
        public string Nomenclature { get; set; } // Номенклатура
        public decimal Quantity { get; set; } // Кол-во
        public decimal Volume { get; set; } // Объем
        public string DocumentType { get; set; } // Тип документа
        public string DocumentNumber { get; set; } // Номер документа
        public DateTime DocumentDate { get; set; } // Дата документа
        public string Shipper { get; set; } // Грузоотправитель
        public string Consignee { get; set; } // Грузополучатель
        public string CounterpartyWarehouse { get; set; } // Склад контрагента
        public string Employee { get; set; } // Сотрудник
        public string Reason { get; set; } // Основание
        public string ReasonNumber { get; set; } // Номер основания
        public DateTime ReasonDate { get; set; } // Дата основания
        public string Status { get; set; } // Статус
        public DateTime ServerProcessingDateTime { get; set; } // Дата и время обработки на сервере
        public string CreationUser { get; set; } // Пользователь создания
        public string Transport { get; set; } // Транспорт
        public string Trailer { get; set; } // Прицеп
        public string VolumeDeterminationMethod { get; set; } // Метод определения объема
    }
}
