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

namespace WriteBalance.Infrastructure.Repositories
{
    public class FinancialRepository : IFinancialRepository
    {
        private readonly BankDbContext _context;
        private readonly RayanBankDbContext _rayanContext;

        public FinancialRepository(BankDbContext context, RayanBankDbContext rayanContext)
        {
            _context = context;
            _rayanContext = rayanContext;
        }

        public List<FinancialRecord> ExecuteSPList(DBRequestDto requestDB)
        {
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
                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { $" {tarazName} خطا در ارتباط با پایگاه داده" }
                    },
                    requestDB.FolderPath
                );

            if (_context.FinancialRecord == null)
                throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $" {tarazName} خطا در ارتباط با جدول " }
                        },
                        requestDB.FolderPath
                    );
            try
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
                               $"{requestDB.FromDateDB},{requestDB.ToDate},{requestDB.TarazType}",
                                requestDB.OrginalClientAddressDB
                            )
                            .ToList();
                return result;
            }
            catch
            {
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

        public List<RayanFinancialRecord> ExecuteRayanSPList(DBRequestDto requestDB)
        {
            var tarazName = "رایان";
            requestDB.TarazType = "2";

            if (_context == null)
                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { $" {tarazName} خطا در ارتباط با پایگاه داده" }
                    },
                    requestDB.FolderPath
                );

            if (_rayanContext.RayanFinancialBalance == null)
                throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $" {tarazName} خطا در ارتباط با جدول " }
                        },
                        requestDB.FolderPath
                    );
            try
            {
                var result = _rayanContext.RayanFinancialBalance
                    .FromSqlRaw(
                        @"EXEC dbo.SouratMali 
		            @FromDate = {0}, 
		            @ToDate = {1},
		            @FromVoucherNum = {3},
		            @ToVoucherNum = {4}",
                        requestDB.FromDateDB,
                        requestDB.ToDate,
                        requestDB.FromDateDB,
                        requestDB.ToVoucherNum
                    )
                    .ToList();

                return result;
            }
            catch
            {
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

        public List<FinancialRecord> ExecutePoyaSPList(DBRequestDto requestDB)
        {
            var tarazName = "پویا";
            requestDB.TarazType = "5";

            if (_context == null)
                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { $" {tarazName} خطا در ارتباط با پایگاه داده" }
                    },
                    requestDB.FolderPath
                );

            if (_context.FinancialRecord == null)
                throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $" {tarazName} خطا در ارتباط با جدول " }
                        },
                        requestDB.FolderPath
                    );

            try
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
                               $"{requestDB.FromDateDB},{requestDB.ToDate},{requestDB.TarazType}",
                                requestDB.OrginalClientAddressDB
                            )
                            .ToList();

                return result;
            }
            catch
            {
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
