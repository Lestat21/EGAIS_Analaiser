using EGAIS_Analaiser.Data;

namespace EGAIS_Analaiser.View
{
    public static class SellingsToConsole
    {
        public static void SellingToCW()
        {
            using (var dbContext = new UserContext())
            {

                var groupedData = dbContext.Sellings
                    .GroupBy(x => new { x.Division, x.DocumentType })
                                  .Select(group => new
                                  {
                                      Division = group.Key.Division,
                                      DocumentType = group.Key.DocumentType,
                                      VolumeSum = group.Sum(x => x.Volume)
                                  })
                                  .OrderBy(result => result.Division);

                var sell = groupedData.Where(p => p.DocumentType == "Расход при реализации потребителю").OrderBy(s => s.Division);
                var pererabotka = groupedData.Where(p => p.DocumentType == "Расход для переработки").OrderBy(s => s.Division);
                var sobstvennoe = groupedData.Where(p => p.DocumentType == "Расход для собственного потребления").OrderBy(s => s.Division);


                var _sell = dbContext.Sellings
                    .GroupBy(record => record.Division)
                                      .Select(group => new
                                      {
                                          Division = group.Key,
                                          TotalVolume = group.Sum(record => record.Volume)
                                      })
                                      .OrderBy(result => result.Division);

                var result1C = dbContext.FullShort1Cs
                        .Select(record => new
                        {
                            Subdivision = record.Subdivision,
                            Sale = record.Sale,
                            SelfConsumption = record.SelfConsumption,
                            Processing = record.Processing,
                        }).OrderBy(result => result.Subdivision).ToList();


                decimal _rashod = 0;
                decimal _rashod1c = 0;

                //Console.WriteLine(new string('=', 5));
                //Console.WriteLine("{0,-40} {1,15} {2,12} {3,13} {4,13} ", "Владелец склада", "Расход ЕГАИС |","Реализация","переработка","собственное |");
                //Console.WriteLine(new string('-', 100));
                //foreach (var item in _sell)
                //{
                //    var sellVolume = sell.FirstOrDefault(s => s.Division == item.Division)?.VolumeSum ?? 0;
                //    var pererabotkaVolume = pererabotka.FirstOrDefault(p => p.Division == item.Division)?.VolumeSum ?? 0;
                //    var sobstvennoeVolume = sobstvennoe.FirstOrDefault(s => s.Division == item.Division)?.VolumeSum ?? 0;

                //    Console.WriteLine("{0,-40} {1,15} {2,12} {3,13} {4,13} ", item.Division, item.TotalVolume + " |", sellVolume, pererabotkaVolume, sobstvennoeVolume + " |");
                //    _rashod += item.TotalVolume;
                //    _finSell += sellVolume;
                //    _finPererab += pererabotkaVolume;
                //    _finSobstv += sobstvennoeVolume;
                //}
                //Console.WriteLine(new string('-', 100));
                //Console.WriteLine("{0,-40} {1,15} {2,13} {3,12} {4,13} ", "", _sell.Sum(r => r.TotalVolume) + " |",  _finSell, _finPererab, _finSobstv + " |");
                //Console.WriteLine(new string('-', 100));


                Console.WriteLine(new string('=', 38) + " РЕАЛИЗАЦИЯ " + new string('=', 38));
                Console.WriteLine("{0,-40} {1,15} {2,15} {3,15} ", "Владелец склада", "Расход ЕГАИС |", "Расход 1С |", "Расхождение");
                Console.WriteLine(new string('-', 88));
                foreach (var item in _sell)
                {

                    var r1c = (result1C.FirstOrDefault(p => p.Subdivision == item.Division)?.Sale ?? 0) + (result1C.FirstOrDefault(p => p.Subdivision == item.Division)?.SelfConsumption ?? 0) + (result1C.FirstOrDefault(p => p.Subdivision == item.Division)?.Processing ?? 0);
                    Console.WriteLine("{0,-40} {1,15} {2,15} {3,15}  ", item.Division, item.TotalVolume + " |", r1c + " |", r1c - item.TotalVolume);
                    _rashod += item.TotalVolume;
                    _rashod1c += r1c;
                }
                Console.WriteLine(new string('-', 88));
                Console.WriteLine("{0,-40} {1,15} {2,15} {3,15}", "", _sell.Sum(r => r.TotalVolume) + " |", _rashod1c + " |", _rashod1c - _rashod);
                Console.WriteLine(new string('-', 88) + "\n");
            }
        }
    }
}
