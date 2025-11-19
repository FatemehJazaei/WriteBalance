using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Domain.Entities;
using WriteBalance.Application.Interfaces;
using WriteBalance.Infrastructure.Context;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Exceptions;
using System.Linq.Expressions;
using Microsoft.EntityFrameworkCore.Query.SqlExpressions;
using WriteBalance.Infrastructure.Services;
using Newtonsoft.Json;
using WriteBalance.Common.Logging;

namespace WriteBalance.Infrastructure.Repositories
{
    public class FinancialRepository : IFinancialRepository
    {
        private readonly BankDbContext _context;
        private readonly RayanBankDbContext _rayanContext;
        private readonly CheckInput _checkInput;
        private readonly bool _IsTest;

        public FinancialRepository(BankDbContext context, RayanBankDbContext rayanContext)
        {
            _context = context;
            _rayanContext = rayanContext;
            _IsTest = false;
        }

        public List<FinancialRecord> ExecuteSPList(DBRequestDto requestDB, DateTime startTime, DateTime endTime)
        {
            Logger.WriteEntry(JsonConvert.SerializeObject($"Starting ExecuteSPList ."), $"FinancialRepository: ExecuteSPList--typeReport:Info");
            var tarazName = "";
            switch (requestDB.TarazType)
            {
                case "1":
                    tarazName = "سما";
                    break;
                case "4":
                    tarazName = "همراه";
                    break;
                case "3":
                    tarazName = "کاربردی";
                    break;
            }

            if (_context == null)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"context is null"), $"FinancialRepository:ExecuteSPList --typeReport:Error");
                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { $" {tarazName} خطا در ارتباط با پایگاه داده" }
                    },
                    requestDB.FolderPath
                );
            }


            if (_context.FinancialRecord == null)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"_context.FinancialRecord is null"), $"FinancialRepository:ExecuteSPList --typeReport:Error");
                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { $" {tarazName} خطا در ارتباط با جدول " }
                    },
                    requestDB.FolderPath
                );
            }

            try
            {
                _context.Database.SetCommandTimeout(300);

                string startTimePersian = DateTimeExtentions.ToPersianDate(startTime);
                string endTimePersian = DateTimeExtentions.ToPersianDate(endTime);

                if (requestDB.FromDateDB != "")
                {
                    startTimePersian = requestDB.FromDateDB;
                }
                if (requestDB.ToDateDB != "")
                {
                    endTimePersian = requestDB.ToDateDB;
                }

                bool correctDate = _checkInput.CheckDateInput(requestDB, startTimePersian, endTimePersian);

                Logger.WriteEntry(JsonConvert.SerializeObject($"startTimePersian:{startTimePersian}, endTimePersian:{endTimePersian}"), $"FinancialRepository:ExecuteRayanSPList --typeReport:Debug");

                if (_IsTest)
                {
                    var result = _context.FinancialRecord
                                .FromSqlRaw(
                                    @"EXEC dbo.MainProc 
                                        @username = {0}, 
                                        @ptoken = {1}, 
                                        @objecttoken = {2}, 
                                        @parameterslist = {3}, 
                                        @OrginalClientAddress = {4}",
                                    requestDB.UserNameDB,
                                    requestDB.PtokenDB,
                                    requestDB.ObjecttokenDB,
                                   $"{startTimePersian},{endTimePersian},{requestDB.TarazType}",
                                    requestDB.OrginalClientAddressDB
                                )
                                .ToList();
                    return result;
                }
                else
                {
                    var result = _context.FinancialRecord
                                .FromSqlRaw(
                                    @"EXEC  [10.15.43.83].DWProxyDB.dbo.MainProc
                                                            @username = {0}, 
                                                            @ptoken = {1}, 
                                                            @objecttoken = {2}, 
                                                            @parameterslist = {3}, 
                                                            @OrginalClientAddress = {4}",
                                    requestDB.UserNameDB,
                                    requestDB.PtokenDB,
                                    requestDB.ObjecttokenDB,
                                   $"{startTimePersian},{endTimePersian},{requestDB.TarazType}",
                                    requestDB.OrginalClientAddressDB
                                )
                                .ToList();



                    return result;
                }

            }
            catch( Exception ex ) 
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($" {tarazName} خطا در بارگیری اطلاعات "), $"FinancialRepository:ExecuteSPList --typeReport:Error");
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"FinancialRepository:ExecuteSPList --typeReport:Error");
                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { $" {tarazName} خطا در بارگیری اطلاعات " }
                    },
                    requestDB.FolderPath
                );

            }

            /*
            var result = _context.FinancialBalance
                .FromSqlRaw(
                    "EXEC dbo.MainProc"
                )
                .ToList();
            */

        }

        public List<RayanFinancialRecord> ExecuteRayanSPList(DBRequestDto requestDB, DateTime startTime, DateTime endTime)
        {
            Logger.WriteEntry(JsonConvert.SerializeObject($"Starting ExecuteRayanSPList "), $"FinancialRepository:ExecuteRayanSPList --typeReport:Info");

            var tarazName = "رایان";
            requestDB.TarazType = "2";


            if (_rayanContext == null)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"context is null"), $"FinancialRepository:ExecuteRayanSPList --typeReport:Error");
                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { $" {tarazName} خطا در ارتباط با پایگاه داده" }
                    },
                    requestDB.FolderPath
                );
            }


            if (_rayanContext.RayanFinancialBalance == null)
            {

                Logger.WriteEntry(JsonConvert.SerializeObject($"_rayanContext.RayanFinancialBalance is null"), $"FinancialRepository:ExecuteRayanSPList --typeReport:Error");
                throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $" {tarazName} خطا در ارتباط با جدول " }
                        },
                        requestDB.FolderPath
                    );
            }

            try
            {
                string startTimePersian = DateTimeExtentions.ToPersianDate(startTime);
                string endTimePersian = DateTimeExtentions.ToPersianDate(endTime);

                if (requestDB.FromDateDB != "")
                {
                    startTimePersian = requestDB.FromDateDB;
                }
                if (requestDB.ToDateDB != "")
                {
                    endTimePersian = requestDB.ToDateDB;
                }

                bool correctDate = _checkInput.CheckDateInput(requestDB, startTimePersian, endTimePersian);
                Logger.WriteEntry(JsonConvert.SerializeObject($"startTimePersian:{startTimePersian}, endTimePersian:{endTimePersian}"), $"FinancialRepository:ExecuteRayanSPList --typeReport:Debug");

                _rayanContext.Database.SetCommandTimeout(300);

                if (_IsTest)
                {
                    var result = _rayanContext.RayanFinancialBalance
                    .FromSqlRaw(
                        @"EXEC dbo.SouratMali 
		            @FromDate = {0}, 
		            @ToDate = {1},
		            @FromVoucherNum = {2},
		            @ToVoucherNum = {3}",
                       startTimePersian,
                        endTimePersian,
                        requestDB.FromVoucherNum,
                        requestDB.ToVoucherNum
                    )
                    .ToList();
                    return result;
                }
                else
                {
                    var result = _rayanContext.RayanFinancialBalance
                .FromSqlRaw(
                @"EXEC [10.15.7.87].[AccountingDB].[dbo].[SouratMali]
		                                    @FromDate = {0}, 
		                                    @ToDate = {1},
		                                    @FromVoucherNum = {2},
		                                    @ToVoucherNum = {3}",
                int.Parse(startTimePersian),
                int.Parse(endTimePersian),
                 //requestDB.FromVoucherNum,
                 //requestDB.ToVoucherNum
                 DBNull.Value,
                 DBNull.Value
            )
            .ToList();

                    Logger.WriteEntry(JsonConvert.SerializeObject($"rayan list count:{result.Count}"), $"BalanceGenerator:GenerateRayanTablesAsync --typeReport:Info");
                    return result;
                }

            }
            catch ( Exception ex )
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($" {tarazName} خطا در بارگیری اطلاعات "), $"FinancialRepository:ExecuteRayanSPList --typeReport:Error");
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"FinancialRepository:ExecuteRayanSPList --typeReport:Error");
                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { $" {tarazName} خطا در بارگیری اطلاعات " }
                    },
                    requestDB.FolderPath
                );

            }

        }

        public List<FinancialRecord> ExecutePoyaSPList(DBRequestDto requestDB, DateTime startTime, DateTime endTime)
        {
            Logger.WriteEntry(JsonConvert.SerializeObject($"Starting ExecutePoyaSPList "), $"FinancialRepository:ExecutePoyaSPList --typeReport:Info");

            var tarazName = "پویا";
            requestDB.TarazType = "5";

            if (_context == null)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"context is null"), $"FinancialRepository:ExecutePoyaSPList --typeReport:Error");
                throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $" {tarazName} خطا در ارتباط با پایگاه داده" }
                        },
                        requestDB.FolderPath
                    );
            }


            if (_context.FinancialRecord == null)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"context.FinancialRecord is null"), $"FinancialRepository:ExecutePoyaSPList --typeReport:Error");
                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { $" {tarazName} خطا در ارتباط با جدول " }
                    },
                    requestDB.FolderPath
                );

            }

            try
            {
                string startTimePersian = DateTimeExtentions.ToPersianDate(startTime);
                string endTimePersian = DateTimeExtentions.ToPersianDate(endTime);

                if (requestDB.FromDateDB != "")
                {
                    startTimePersian = requestDB.FromDateDB;
                }
                if (requestDB.ToDateDB != "")
                {
                    endTimePersian = requestDB.ToDateDB;
                }

                bool correctDate = _checkInput.CheckDateInput(requestDB, startTimePersian, endTimePersian);

                Logger.WriteEntry(JsonConvert.SerializeObject($"startTimePersian:{startTimePersian}, endTimePersian:{endTimePersian}"), $"FinancialRepository:ExecuteRayanSPList --typeReport:Debug");
                _context.Database.SetCommandTimeout(300);

                if (_IsTest) 
                {
                    var result = _context.FinancialRecord
                .FromSqlRaw(
                    @"EXEC dbo.MainProc 
                                        @username = {0}, 
                                        @ptoken = {1}, 
                                        @objecttoken = {2}, 
                                        @parameterslist = {3}, 
                                        @OrginalClientAddress = {4}",
                    requestDB.UserNameDB,
                    requestDB.PtokenDB,
                    requestDB.ObjecttokenDB,
                   $"{startTimePersian},{endTimePersian},{requestDB.TarazType}",
                    requestDB.OrginalClientAddressDB
                )
                .ToList();
                    return result;
                }
                else
                {
                    var result = _context.FinancialRecord
                        .FromSqlRaw(
                            @"EXEC [10.15.43.83].DWProxyDB.dbo.MainProc
                            @username = {0}, 
                            @ptoken = {1}, 
                            @objecttoken = {2}, 
                            @parameterslist = {3}, 
                            @OrginalClientAddress = {4}",
                        requestDB.UserNameDB,
                        requestDB.PtokenDB,
                        requestDB.ObjecttokenDB,
                       $"{startTimePersian},{endTimePersian},{requestDB.TarazType}",
                        requestDB.OrginalClientAddressDB
                    )
                    .ToList();

                    return result;
                }

            }
            catch( Exception ex )
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($" {tarazName} خطا در بارگیری اطلاعات "), $"FinancialRepository:ExecutePoyaSPList --typeReport:Error");
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"FinancialRepository:ExecutePoyaSPList --typeReport:Error");
                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { $" {tarazName} خطا در بارگیری اطلاعات " }
                    },
                    requestDB.FolderPath
                );

            }
        }

    }
}
