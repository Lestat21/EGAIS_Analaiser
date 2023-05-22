using EGAIS_Analaiser.Data;

namespace EGAIS_Analaiser.View
{
    public static class RemainsToConsole
    {
        public static void RemainsToCW()
        {
            using (var dbContext = new UserContext())
            {
                var resultsEGAIS = dbContext.Remains
                          .GroupBy(record => record.WarehouseOwner)
                          .Select(group => new
                          {
                              WarehouseOwner = group.Key,
                              TotalVolume = group.Sum(record => record.Volume)
                          })
                          .OrderBy(result => result.WarehouseOwner);

                var result1C = dbContext.FullShort1Cs
                          .Select(record => new
                          {
                              Subdivision = record.Subdivision,
                              Balance = record.Balance
                          }).OrderBy(result => result.Subdivision).ToList();


                //var result1C = dbContext.Remains1Cs // по данным более точным
                //    .GroupBy(record => record.WarehouseOwner)
                //    .Select(group => new
                //    {
                //        WarehouseOwner = group.Key,
                //        TotelVolume = group.Sum(record => record.Remainder)
                //    }
                //    ).OrderBy(result => result.WarehouseOwner).ToList();
                Console.WriteLine(new string('=', 38) + "  ОСТАТКИ  " + new string('=', 39));
                Console.WriteLine("{0,-40} {1,15} {2,15} {3,15}", "Владелец склада", "Объем ЕГАИС", "Объем 1C", "Расхождение");
                Console.WriteLine(new string('-', 88));

                foreach (var result in resultsEGAIS)
                {
                    var r1c = result1C.FirstOrDefault(p => p.Subdivision == result.WarehouseOwner)?.Balance ?? 0;

                    Console.WriteLine("{0,-40} {1,15} {2,15} {3,15}", result.WarehouseOwner, result.TotalVolume, r1c, r1c - result.TotalVolume);
                }
                Console.WriteLine(new string('-', 88));
                Console.WriteLine("{0,-40} {1,15} {2,15} {3,15}", "", resultsEGAIS.Sum(r => r.TotalVolume), result1C.Sum(r => r.Balance), result1C.Sum(r => r.Balance) - resultsEGAIS.Sum(r => r.TotalVolume));
                Console.WriteLine(new string('-', 88) + "\n");

            }
        }
    }
}
