using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Common.Logging;
using WriteBalance.Domain.Entities;

namespace WriteBalance.Infrastructure.Services
{
    public class BalanceMerge
    {

        // مرج کردن رکورد ها با کد یکسان 
        // در صوتری که هم بستانکار و هم بدهکار داشته باشد 
        // اختلاف قدر مطلق بستانکار و بدهکار 
        // اگر مثبت باشد در ستون بدهکار 
        // اگر منفی باشد در ستون بستانکار قرار میگیرد
        public List<ExcelRow> MergeDuplicateRows(List<ExcelRow> rows)
        {
            try
            {
                var merged = rows
                                .GroupBy(r => r.Col1)
                                .Select(g =>
                                {
                                    var first = g.First();
                                    var bed = g.Sum(x => x.Col3);
                                    var bes = g.Sum(x => x.Col4);

                                    var Mande = bed - bes;
                                    if (Mande >= 0)
                                    {
                                        bed = Mande;
                                        bes = 0;
                                    }
                                    else if (Mande < 0)
                                    {
                                        bed = 0;
                                        bes = Math.Abs(Mande);
                                    }

                                    return new ExcelRow
                                    {
                                        Col1 = first.Col1,
                                        Col2 = first.Col2,
                                        Col3 = bed,
                                        Col4 = bes,
                                        Col5 = 0,
                                        Col6 = 0

                                    };
                                }).ToList();
                return merged;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"BalanceGenerator:MergeDuplicateRows --typeReport:Error");
                throw;
            }

        }


        public List<ExcelRow> MergeDuplicateGardeshRows(List<ExcelRow> rows)
        {
            try
            {
                var merged = rows
                    .GroupBy(r => r.Col1)
                    .SelectMany(g =>
                    {
                        var first = g.First();
                        var bed = g.Sum(x => x.Col5);
                        var bes = g.Sum(x => x.Col6);

                        if (bed < 0)
                        {
                            bes += Math.Abs(bed??0);
                        }
                        if (bes < 0)
                        {
                            bed += Math.Abs(bes ?? 0);
                        }

                        var rowList = new List<ExcelRow>();
                        if (bed > 0)
                        {
                            rowList.Add(new ExcelRow
                            {
                                Col1 = $"{first.Col1}BED",
                                Col2 = first.Col2,
                                Col3 = 0,
                                Col4 = 0,
                                Col5 = bed,
                                Col6 = 0
                            });
                        }
                        if (bes > 0)
                        {
                            rowList.Add(new ExcelRow
                            {
                                Col1 = $"{first.Col1}BES",
                                Col2 = first.Col2,
                                Col3 = 0,
                                Col4 = 0,
                                Col5 = 0,
                                Col6 = bes
                            });
                        }
                        return rowList;
                    }).ToList();
                return merged;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"BalanceGenerator:MergeDuplicateRows --typeReport:Error");
                throw;
            }

        }


        public List<ExcelRow> MergeDuplicateGardeshPouyaRows(List<ExcelRow> rows)
        {
            try
            {
                var merged = rows
                    .GroupBy(r => r.Col1)
                    .SelectMany(g =>
                    {
                        var first = g.First();
                        var bed = g.Sum(x => x.Col5);
                        var bes = g.Sum(x => x.Col6);

                        if (bed < 0)
                        {
                            bes += Math.Abs(bed ?? 0);
                        }
                        if (bes < 0)
                        {
                            bed += Math.Abs(bes ?? 0);
                        }

                        var rowList = new List<ExcelRow>();
                        if (bed > 0)
                        {
                            rowList.Add(new ExcelRow
                            {
                                Col1 = $"{first.Col1}_BED",
                                Col2 = first.Col2,
                                Col3 = 0,
                                Col4 = 0,
                                Col5 = bed,
                                Col6 = 0
                            });
                        }
                        if (bes > 0)
                        {
                            rowList.Add(new ExcelRow
                            {
                                Col1 = $"{first.Col1}_BES",
                                Col2 = first.Col2,
                                Col3 = 0,
                                Col4 = 0,
                                Col5 = 0,
                                Col6 = bes
                            });
                        }
                        return rowList;
                    }).ToList();
                return merged;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"BalanceGenerator:MergeDuplicateRows --typeReport:Error");
                throw;
            }

        }
    }
}
