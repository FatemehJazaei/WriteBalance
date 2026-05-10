using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Query.SqlExpressions;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Exceptions;
using WriteBalance.Application.Interfaces;
using WriteBalance.Common.Logging;
using WriteBalance.Domain.Entities;
using WriteBalance.Infrastructure.Context;
using WriteBalance.Infrastructure.Services;

namespace WriteBalance.Infrastructure.Repositories
{
    public class FinancialRepository : IFinancialRepository
    {
        private readonly BankDbContext _context;
        private readonly GLBankDbContext _glContext;
        private readonly RayanBankDbContext _rayanContext;
        private readonly PouyaBankDbContext _pouyaContext;
        private readonly ICheckInput _checkInput;
        private readonly bool _IsTest;

        public FinancialRepository(BankDbContext context, RayanBankDbContext rayanContext, PouyaBankDbContext pouyaContext, GLBankDbContext gLBankDbContext, ICheckInput checkInput)
        {
            _context = context;
            _rayanContext = rayanContext;
            _checkInput = checkInput;
            _pouyaContext = pouyaContext;
            _glContext = gLBankDbContext;
            _IsTest = false;
        }


        public List<GLFinancialRecord> ExecuteGLList(APIRequestDto request, DBRequestDto requestDB, string startTimePersian, string endTimePersian)
        {
            Logger.WriteEntry(JsonConvert.SerializeObject($"Starting ExecuteGLList ."), $"FinancialRepository: ExecuteGLList--typeReport:Info");
            var tarazName = "جی ال";
            request.tarazNameLatin = "GL";

            if (_glContext == null)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"_glContext is null"), $"FinancialRepository:ExecuteGLList --typeReport:Error");
                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { $" {tarazName} خطا در ارتباط با پایگاه داده" }
                    },
                    requestDB.FolderPath
                );
            }


            if (_glContext.GLFinancialRecords == null)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"_glContext.GLFinancialRecords is null"), $"FinancialRepository:ExecuteGLList --typeReport:Error");
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

                var pc = new PersianCalendar();
                var now = DateTime.Now;

                string timestamp = $"{pc.GetDayOfMonth(now):00}_{pc.GetMonth(now):00}_{pc.GetYear(now):0000}";
                requestDB.FileName = $"تراز {tarazName} دریافت شده در تاریخ {timestamp} برای {endTimePersian}.xlsx";

                Logger.WriteEntry(JsonConvert.SerializeObject($"startTimePersian:{startTimePersian}, endTimePersian:{endTimePersian}"), $"FinancialRepository:ExecuteGLList --typeReport:Debug");

                if (_IsTest)
                {
                    // Gl Taraz type = 5
                    var result = _glContext.GLFinancialRecords
                                .FromSqlRaw(
                                    @"EXEC dbo.MainProcGL 
                                        @username = {0}, 
                                        @ptoken = {1}, 
                                        @objecttoken = {2}, 
                                        @parameterslist = {3}, 
                                        @OrginalClientAddress = {4}",
                                    requestDB.UserNameDB,
                                    requestDB.PtokenDB,
                                    requestDB.ObjecttokenGL,
                                   //$"{startTimePersian},{endTimePersian},{5}",
                                   $"{endTimePersian},{5}",
                                    requestDB.OrginalClientAddressDB
                                )
                                .ToList();

                    if (result != null && result.Count == 0)
                    {
                        Logger.WriteEntry(JsonConvert.SerializeObject($"result.Count = {result.Count} "), $"FinancialRepository:ExecuteGLList --typeReport:Error");
                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $"اطلاعاتی برای این تواریخ در {tarazName} وجود ندارد." }
                            },
                            requestDB.FolderPath
                        );
                    }
                    else if (result == null)
                    {
                        throw new Exception("result is null");
                    }

                    return result;
                }
                else
                {
                    var result = _glContext.GLFinancialRecords
                                .FromSqlRaw(
                                    @"EXEC  [10.15.43.83].DWProxyDB.dbo.MainProc
                                                            @username = {0}, 
                                                            @ptoken = {1}, 
                                                            @objecttoken = {2}, 
                                                            @parameterslist = {3}, 
                                                            @OrginalClientAddress = {4}",
                                    requestDB.UserNameDB,
                                    requestDB.PtokenDB,
                                    requestDB.ObjecttokenGL,
                                   //$"{startTimePersian},{endTimePersian},{5}",
                                   $"{endTimePersian},{5}",
                                    requestDB.OrginalClientAddressDB
                                )
                                .ToList();

                    if (result != null && result.Count == 0)
                    {
                        Logger.WriteEntry(JsonConvert.SerializeObject($"result.Count = {result.Count} "), $"FinancialRepository:ExecuteGLList --typeReport:Error");
                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $"اطلاعاتی برای این تواریخ در {tarazName} وجود ندارد." }
                            },
                            requestDB.FolderPath
                        );
                    }
                    else if (result == null)
                    {
                        throw new Exception("result is null");
                    }
                    return result;

                }

            }
            catch (ConnectionMessageException)
            {
                throw;

            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($" {tarazName} خطا در بارگیری اطلاعات "), $"FinancialRepository:ExecuteGLList --typeReport:Error");
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"FinancialRepository:ExecuteGLList --typeReport:Error");
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


        public List<FinancialRecord> ExecuteSPList(APIRequestDto request, DBRequestDto requestDB, string startTimePersian, string endTimePersian)
        {
            Logger.WriteEntry(JsonConvert.SerializeObject($"Starting ExecuteSPList ."), $"FinancialRepository: ExecuteSPList--typeReport:Info");
            var tarazName = "";
            switch (requestDB.TarazType)
            {
                case "1":
                    tarazName = "سما";
                    request.tarazNameLatin = "Sama";
                    break;
                case "4":
                    tarazName = "همراه";
                    request.tarazNameLatin = "Hamrah";
                    break;
                case "3":
                    tarazName = "کاربردی";
                    request.tarazNameLatin = "Karbordi";
                    break;
                case "2":
                    tarazName = "پویا";
                    request.tarazNameLatin = "pouya";
                    break;
                case "5":
                    tarazName = "رایان";
                    request.tarazNameLatin = "Rayan";
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

                var pc = new PersianCalendar();
                var now = DateTime.Now;

                string timestamp = $"{pc.GetDayOfMonth(now):00}_{pc.GetMonth(now):00}_{pc.GetYear(now):0000}";

                requestDB.FileName = $"تراز {tarazName} دریافت شده در تاریخ {timestamp} برای {endTimePersian}.xlsx";

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

                    if (result != null && result.Count == 0)
                    {
                        Logger.WriteEntry(JsonConvert.SerializeObject($"result.Count = {result.Count} "), $"FinancialRepository:ExecuteSPList --typeReport:Error");
                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $"اطلاعاتی برای این تواریخ در {tarazName} وجود ندارد." }
                            },
                            requestDB.FolderPath
                        );
                    }
                    else if(result == null)
                    {
                        throw new Exception("result is null");
                    }

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
                                   $"{startTimePersian},{endTimePersian},{requestDB.TarazType},{requestDB.BeforeClose}",
                                    requestDB.OrginalClientAddressDB
                                )
                                .ToList();

                    if (result != null && result.Count == 0)
                    {
                        Logger.WriteEntry(JsonConvert.SerializeObject($"result.Count = {result.Count} "), $"FinancialRepository:ExecuteSPList --typeReport:Error");
                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $"اطلاعاتی برای این تواریخ در {tarazName} وجود ندارد." }
                            },
                            requestDB.FolderPath
                        );
                    }
                    else if (result == null)
                    {
                        throw new Exception("result is null");
                    }
                    return result;

                }

            }
            catch (ConnectionMessageException)
            {
                throw;

            }
            catch ( Exception ex ) 
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

        public List<RayanFinancialRecord> ExecuteRayanSPList(APIRequestDto request, DBRequestDto requestDB, string startTimePersian, string endTimePersian, bool IsExcept)
        {
            Logger.WriteEntry(JsonConvert.SerializeObject($"Starting ExecuteRayanSPList "), $"FinancialRepository:ExecuteRayanSPList --typeReport:Info");

            var tarazName = "رایان";
            requestDB.TarazType = "2";
            request.tarazNameLatin = "Rayan";

            if (_rayanContext == null)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"_rayanContext is null"), $"FinancialRepository:ExecuteRayanSPList --typeReport:Error");
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

                var pc = new PersianCalendar();
                var now = DateTime.Now;
                string timestamp = $"{pc.GetDayOfMonth(now):00}_{pc.GetMonth(now):00}_{pc.GetYear(now):0000}";

                requestDB.FileName = $"تراز {tarazName} دریافت شده در تاریخ {timestamp} برای {endTimePersian}.xlsx";

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

                    if (result != null && result.Count == 0 && !IsExcept)
                    {
                        Logger.WriteEntry(JsonConvert.SerializeObject($"result.Count = {result.Count} "), $"FinancialRepository:ExecuteRayanSPList --typeReport:Error");
                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $"اطلاعاتی برای این تواریخ در {tarazName} وجود ندارد." }
                            },
                            requestDB.FolderPath
                        );
                    }
                    else if (result == null)
                    {
                        throw new Exception("result is null");
                    }
                    return result;
                }
                else
                {

                    object toVoucherParam = int.TryParse(requestDB.ToVoucherNum, out var toVoucher)
                    ? toVoucher
                    : DBNull.Value;

                    object FromVoucherParam = int.TryParse(requestDB.FromVoucherNum, out var fromVoucher)
                    ? fromVoucher
                    : DBNull.Value;

                    Logger.WriteEntry(JsonConvert.SerializeObject($"FromVoucherParam: {FromVoucherParam},toVoucherParam = {toVoucherParam} "), $"FinancialRepository:ExecuteRayanSPList --typeReport:Error");


                    var result = _rayanContext.RayanFinancialBalance
                                .FromSqlRaw(
                                @"EXEC [10.15.7.87].[AccountingDB].[dbo].[SouratMali]
		                                                    @FromDate = {0}, 
		                                                    @ToDate = {1},
		                                                    @FromVoucherNum = {2},
		                                                    @ToVoucherNum = {3}",
                                int.Parse(startTimePersian),
                                int.Parse(endTimePersian),
                                FromVoucherParam,
                                toVoucherParam
                            //DBNull.Value,
                            //DBNull.Value
                            )
                            .ToList();

                    if (result != null && result.Count == 0)
                    {
                        Logger.WriteEntry(JsonConvert.SerializeObject($"result.Count = {result.Count} "), $"FinancialRepository:ExecuteRayanSPList --typeReport:Error");
                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $"اطلاعاتی برای این تواریخ در {tarazName} وجود ندارد." }
                            },
                            requestDB.FolderPath
                        );
                    }
                    else if (result == null)
                    {
                        throw new Exception("result is null");
                    }
                    return result;
                }

            }
            catch (ConnectionMessageException)
            {
                throw;

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

        public List<PouyaFinancialRecord> ExecutePoyaSPList(APIRequestDto request, DBRequestDto requestDB, string startTimePersian, string endTimePersian)
        {
            Logger.WriteEntry(JsonConvert.SerializeObject($"Starting ExecutePoyaSPList."), $"FinancialRepository:ExecutePoyaSPList--typeReport:Info");
            var tarazName = "پویا";
            request.tarazNameLatin = "Pouya";

            if (_pouyaContext == null)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"_pouyaContext is null"), $"FinancialRepository:ExecutePoyaSPList--typeReport:Error");
                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { $" {tarazName} خطا در ارتباط با پایگاه داده" }
                    },
                    requestDB.FolderPath
                );
            }


            if (_pouyaContext.PouyaFinancialBalance == null)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"_pouyaContext.PouyaFinancialBalance is null"), $"FinancialRepository:ExecutePoyaSPList --typeReport:Error");
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
                _pouyaContext.Database.SetCommandTimeout(300);

                var pc = new PersianCalendar();
                var now = DateTime.Now;

                string timestamp = $"{pc.GetDayOfMonth(now):00}_{pc.GetMonth(now):00}_{pc.GetYear(now):0000}";

                requestDB.FileName = $"تراز  {tarazName} دریافت شده در تاریخ {timestamp} برای {endTimePersian}.xlsx";

                Logger.WriteEntry(JsonConvert.SerializeObject($"startTimePersian:{startTimePersian}, endTimePersian:{endTimePersian}"), $"FinancialRepository:ExecutePoyaSPList --typeReport:Debug");

                if (_IsTest)
                {
                    var result = _pouyaContext.PouyaFinancialBalance
                                .FromSqlRaw(
                                    @"EXEC dbo.usp_in5GetArziBalance
                                        @intBalkd  = {0}, 
                                        @intRptKd ={1}, 
                                        @fromDate = {2},
                                        @ToDate  ={3}",
                                    int.Parse(requestDB.TarazTypePouya),
                                    3,
                                    $"{startTimePersian}",
                                   $"{endTimePersian}"
                                )
                                .ToList();

                    if (result != null && result.Count == 0)
                    {
                        Logger.WriteEntry(JsonConvert.SerializeObject($"result.Count = {result.Count} "), $"FinancialRepository:ExecuteSPList --typeReport:Error");
                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $"اطلاعاتی برای این تواریخ در {tarazName} وجود ندارد." }
                            },
                            requestDB.FolderPath
                        );
                    }
                    else if (result == null)
                    {
                        throw new Exception("result is null");
                    }
                    return result;
                }
                else
                {

                    string sql =$"EXEC  [10.15.43.52].[arzi].[dbo].[usp_in5GetArziBalance] {int.Parse(requestDB.TarazTypePouya)},3,'{startTimePersian}','{endTimePersian}'";

                     var result = _pouyaContext.PouyaFinancialBalance
                                .FromSqlRaw(sql)
                                .ToList();

                    Logger.WriteEntry(JsonConvert.SerializeObject($"sql = {sql} "), $"FinancialRepository:ExecutePoyaSPList --typeReport:Error");
                    if (result != null && result.Count == 0)
                    {
                        Logger.WriteEntry(JsonConvert.SerializeObject($"result.Count = {result.Count} "), $"FinancialRepository:ExecuteSPList --typeReport:Error");
                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $"اطلاعاتی برای این تواریخ در {tarazName} وجود ندارد." }
                            },
                            requestDB.FolderPath
                        );
                    }
                    else if (result == null)
                    {
                        throw new Exception("result is null");
                    }
                    return result;
                }

            }
            catch (ConnectionMessageException)
            {
                throw;

            }
            catch (Exception ex)
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
