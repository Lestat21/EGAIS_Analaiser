using EGAIS_Analaiser.Data;

namespace EGAIS_Analaiser.View
{
    public class ZogotovkaToConsole
    {
        public static void ZagotovkaToCW()
        {
            using (var dbContext = new UserContext())
            {
                var zagotovkaEGAIS = dbContext.Zagotovkas
                         .GroupBy(record => record.Forestry)
                         .Select(group => new
                         {
                             WarehouseOwner = group.Key,
                             TotalVolume = group.Sum(record => record.OperationalAccountingTotal)
                         })
                         .OrderBy(result => result.WarehouseOwner);

                var result1C = dbContext.FullShort1Cs
                          .Select(record => new
                          {
                              Subdivision = record.Subdivision,
                              Zagotovka = record.Procurement,
                              Balance = record.Balance,
                              Sale = record.Sale,
                              SelfConsumption = record.SelfConsumption,
                              Processing = record.Processing
                          }).OrderBy(result => result.Subdivision).ToList();

                Console.WriteLine(new string('=', 38) + " ЗАГОТОВКА " + new string('=', 39));
                Console.WriteLine("{0,-40} {1,15} {2,15} {3,15}", "Владелец склада", "Заготовка ЕГАИС", "Заготовка 1C", "Расхождение");
                Console.WriteLine(new string('-', 88));

                foreach (var result in zagotovkaEGAIS)
                {
                    var r1c = result1C.FirstOrDefault(p => p.Subdivision == result.WarehouseOwner)?.Zagotovka ?? 0;

                    Console.WriteLine("{0,-40} {1,15} {2,15} {3,15}", result.WarehouseOwner, result.TotalVolume, r1c, r1c - result.TotalVolume);
                }
                Console.WriteLine(new string('-', 88));
                Console.WriteLine("{0,-40} {1,15} {2,15} {3,15}", "", zagotovkaEGAIS.Sum(r => r.TotalVolume), result1C.Sum(r => r.Zagotovka), result1C.Sum(r => r.Zagotovka) - zagotovkaEGAIS.Sum(r => r.TotalVolume));



            }
        }
    }
}
