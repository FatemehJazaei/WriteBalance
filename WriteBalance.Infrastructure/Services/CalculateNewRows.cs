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
    public class CalculateNewRows
    {
        // محاسبه مقادیر بستانکار و بدهکار با در نظر گرفتن گردش ها 
        public async Task<List<ExcelRow>> Calculate_New_rows(List<ExcelRow> Rows)
        {
            try
            {
                foreach (ExcelRow row in Rows)
                {
                    decimal bed = 0;
                    decimal bes = 0;

                    // mandeh bedehkar
                    if (row.Col3 < 0)
                    {
                        bes += Math.Abs(row.Col3);
                    }
                    else if (row.Col3 >= 0)
                    {
                        bed += Math.Abs(row.Col3);
                    }

                    // mandeh bestankar
                    if (row.Col4 < 0)
                    {
                        bed += Math.Abs(row.Col4);
                    }
                    else if (row.Col4 >= 0)
                    {
                        bes += Math.Abs(row.Col4);
                    }

                    // gardesh bedehkar
                    if (row.Col5 < 0)
                    {
                        bes += Math.Abs(row.Col5 ?? decimal.Zero);
                    }
                    if (row.Col5 >= 0)
                    {
                        bed += Math.Abs(row.Col5 ?? decimal.Zero);
                    }

                    //gardesh bestankar
                    if (row.Col6 < 0)
                    {
                        bed += Math.Abs(row.Col6 ?? decimal.Zero);
                    }
                    else if (row.Col6 >= 0)
                    {
                        bes += Math.Abs(row.Col6 ?? decimal.Zero);
                    }

                    // mandeh
                    if (bed - bes >= 0)
                    {
                        row.Col3 = Math.Abs(bed - bes);
                        row.Col4 = 0;
                        row.Col5 = 0;
                        row.Col6 = 0;
                    }
                    else if (bed - bes < 0)
                    {
                        row.Col3 = 0;
                        row.Col4 = Math.Abs(bed - bes);
                        row.Col5 = 0;
                        row.Col6 = 0;
                    }
                }

                return await Task.FromResult(Rows);
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"BalanceGenerator:Calculate_New_rows --typeReport:Error");
                throw;
            }
        }
    }
}
