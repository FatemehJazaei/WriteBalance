using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Vml;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Exceptions;
using WriteBalance.Application.Interfaces;
using WriteBalance.Common.Logging;
using WriteBalance.Domain.Entities;
using WriteBalance.Infrastructure.Repositories;
using Path = System.IO.Path;

namespace WriteBalance.Infrastructure.Services
{
    public class CheckInput: ICheckInput
    {
        public (string, string) CheckDateInput(DBRequestDto requestDB, DateTime startDateTime, DateTime endDateTime)
        {
            try 
            {
                string startFinancialPeriod = DateTimeExtentions.ToPersianDate(startDateTime);
                string endFinancialPeriod = DateTimeExtentions.ToPersianDate(endDateTime);

                if (requestDB.FromDateDB == "")
                {
                    requestDB.FromDateDB = startFinancialPeriod;
                }
                else if (!int.TryParse(requestDB.FromDateDB, out var number1) || requestDB.FromDateDB.Length != 8 || int.Parse(requestDB.FromDateDB) < 0 || !DateTimeExtentions.IValidDate(requestDB.FromDateDB))
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"Date is invalid. FromDateDB: {requestDB.FromDateDB} "), $"CheckDateInput--typeReport:Error");
                    throw new ConnectionMessageException(
                       new ConnectionMessage
                       {
                           MessageType = MessageType.Error,
                           Messages = new List<string> { $" تاریخ شروع نامعتبر" }
                       },
                   requestDB.FolderPath
                   );

                }

                if (requestDB.ToDateDB == "")
                {
                    requestDB.ToDateDB = endFinancialPeriod;
                }
                else if (!int.TryParse(requestDB.ToDateDB, out var number1) || requestDB.ToDateDB.Length != 8 || int.Parse(requestDB.ToDateDB) < 0 || !DateTimeExtentions.IValidDate(requestDB.ToDateDB))
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"Date is invalid. ToDateDB: {requestDB.ToDateDB} "), $"CheckDateInput--typeReport:Error");
                    throw new ConnectionMessageException(
                       new ConnectionMessage
                       {
                           MessageType = MessageType.Error,
                           Messages = new List<string> { $"تاریخ پایان نامعتبر " }
                       },
                   requestDB.FolderPath
                   );

                }

                if (int.Parse(requestDB.ToDateDB) < int.Parse(requestDB.FromDateDB) || int.Parse(endFinancialPeriod) < int.Parse(startFinancialPeriod))
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"Date is invalid. (FromDateDB {requestDB.FromDateDB} , ToDateDB {requestDB.ToDateDB}) "), $"CheckDateInput--typeReport:Error");
                    Logger.WriteEntry(JsonConvert.SerializeObject($"Date is invalid. (startFinancialPeriod {startFinancialPeriod} , endFinancialPeriod {endFinancialPeriod}) "), $"CheckDateInput--typeReport:Error");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $" تاریخ شروع نمیتواند از تاریخ پایان بزرگتر باشد. " }
                        },
                    requestDB.FolderPath
                    );
                }

                if (int.Parse(requestDB.FromDateDB) < int.Parse(startFinancialPeriod) || int.Parse(endFinancialPeriod) < int.Parse(requestDB.ToDateDB))
                {

                    Logger.WriteEntry(JsonConvert.SerializeObject("Date is invalid. FromDateDB or ToDateDB is not in Financial Period Range!"), $"CheckDateInput--typeReport:Error");
                    Logger.WriteEntry(JsonConvert.SerializeObject($"Date is invalid. (FromDateDB {requestDB.FromDateDB} , ToDateDB {requestDB.ToDateDB}) "), $"CheckDateInput--typeReport:Error");
                    Logger.WriteEntry(JsonConvert.SerializeObject($"Date is invalid. (startFinancialPeriod {startFinancialPeriod} , endFinancialPeriod {endFinancialPeriod}) "), $"CheckDateInput--typeReport:Error");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $" تاریخ ورودی در بازه دوره مالی قرار ندارد. " }
                        },
                    requestDB.FolderPath
                    );
                }

                return (requestDB.FromDateDB, requestDB.ToDateDB);
            }
            catch (ConnectionMessageException ex)
            {
                throw;
            }
        }

        public bool CheckUserInput(Dictionary<string, string> config)
        {

            try
            {
                /*

                if (config["OnlyVoucherNum"] != "" ||  config["ToVoucherNum"]  != ""  ||  config["OnlyVoucherNum"] != "" ||  config["ExceptVoucherNum"] != "")
                {
                    if (config["OnlyVoucherNum"] != "" && config["ToVoucherNum"] != "" && config["OnlyVoucherNum"] != "")
                    {
                        Logger.WriteEntry(JsonConvert.SerializeObject("VoucherNum is invalid"), $"CheckInput--typeReport:Error");
                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $" ورودی نامعتبر " }
                            },
                        Path.Combine(config["op"], config["of"])
                        );
                    }


                    if (int.Parse(config["OnlyVoucherNum"]) < 0 || int.Parse(config["ExceptVoucherNum"]) < 0 || int.Parse(config["ToVoucherNum"]) < 0 || int.Parse(config["FromVoucherNum"]) < 0)
                    {

                        Logger.WriteEntry(JsonConvert.SerializeObject("VoucherNum is negetive"), $"CheckInput--typeReport:Error");

                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $" ورودی نامعتبر " }
                            },
                        Path.Combine(config["op"], config["of"])
                        );
                    }


                    if (int.Parse(config["ToVoucherNum"]) < int.Parse(config["ExceptVoucherNum"]) || int.Parse(config["ToVoucherNum"]) < int.Parse(config["FromVoucherNum"]) || int.Parse(config["FromVoucherNum"]) > int.Parse(config["ExceptVoucherNum"]))
                    {
                        
                        Logger.WriteEntry(JsonConvert.SerializeObject("VoucherNum is invalid"), $"CheckInput--typeReport:Error");

                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $" ورودی نامعتبر " }
                            },
                        Path.Combine(config["op"], config["of"])
                        );
                    }

                    if (config["ToVoucherNum"] == "" && config["FromVoucherNum"] == "" && config["ExceptVoucherNum"] != "")
                    {
                        Logger.WriteEntry(JsonConvert.SerializeObject("VoucherNum is invalid"), $"CheckInput--typeReport:Error");

                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $" ورودی نامعتبر " }
                            },
                        Path.Combine(config["op"], config["of"])
                        );
                    }
                }
                */
                if (config["tarazType"] != "-1" && config["tarazType"] != "1" && config["tarazType"] != "2" && config["tarazType"] != "3" && config["tarazType"] != "4" && config["tarazType"] != "5")
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject("tarazType is invalid"), $"CheckInput--typeReport:Error");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $" ورودی نامعتبر " }
                        },
                    Path.Combine(config["op"], config["of"])
                    );
                }


                if (config["PrintOrReport"] != "1" && config["PrintOrReport"] != "2")
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject("PrintOrReport is invalid"), $"CheckInput--typeReport:Error");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $" ورودی نامعتبر " }
                        },
                    Path.Combine(config["op"], config["of"])
                    );
                }


                if (config["BalanceName"] == "")
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject("BalanceName is empty"), $"CheckInput--typeReport:Error");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $" نام تراز خالی است. " }
                        },
                    Path.Combine(config["op"], config["of"])
                    );
                }

                if (config["tarazType"] == "5")
                {
                    if (config["tarazTypePouya"] != "1" && config["tarazTypePouya"] != "2" && config["tarazTypePouya"] != "3")
                    {
                        Logger.WriteEntry(JsonConvert.SerializeObject("tarazTypePouya is empty"), $"CheckInput--typeReport:Error");

                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $" نوع تراز پویا نامشخص است. " }
                            },
                        Path.Combine(config["op"], config["of"])
                        );
                    }
                }



                return true;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"{ex}"), $"CheckInput--typeReport:Error");
                throw;
            }
        }

        public List<ExceptCode> CheckExceptCode(Dictionary<string, string> config)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting CheckExceptCode ..."), $"CheckInput:CheckExceptCode--typeReport:Info");
                List<ExceptCode> RayanExceptCode = new List<ExceptCode>();
                if (config["ExceptCode"] != "")
                {
                    string[] Codes = config["ExceptCode"].Split('|');
                    foreach (string Code in Codes) {
                        var parts = Code.Split('_',2);
                        if (parts.Length == 2)
                        {
                            if (parts[0].Length == 4 && parts[1].Length == 3 && int.TryParse(parts[0], out var number0) && int.TryParse(parts[1], out var number1))
                            {
                                RayanExceptCode.Add(new ExceptCode
                                {
                                    Kol_Code = parts[0],
                                    Moeen_Code = parts[1]
                                });
                            }
                            else 
                            {
                                throw new Exception($"code is wrong! parts[0]: {parts[0]} , parts[1]: {parts[1]}");
                            }
                        }
                        else
                        {
                            throw new Exception($"code doesn't has two parts! Code: {Code}");
                        }
                    }

                }
                return RayanExceptCode;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"{ex}"), $"CheckInput:CheckExceptCode--typeReport:Error");
                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { $"کدها به صورت صحیح وارد نشده اند." }
                    },
                Path.Combine(config["op"], config["of"])
                );
            }
        }
    }
}
