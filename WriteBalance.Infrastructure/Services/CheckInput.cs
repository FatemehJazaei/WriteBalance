using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Exceptions;
using WriteBalance.Common.Logging;
using WriteBalance.Infrastructure.Repositories;

namespace WriteBalance.Infrastructure.Services
{
    public class CheckInput
    {

        public bool CheckDateInput(DBRequestDto requestDB, string startFinancialPeriod, string endFinancialPeriod )
        {
            try {
 
                    if (int.Parse(requestDB.ToDateDB) < 0 || int.Parse(requestDB.FromDateDB) < 0)
                    {

                        Logger.WriteEntry(JsonConvert.SerializeObject("Date is negetive"), $"CheckDateInput--typeReport:Error");

                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $" ورودی نامعتبر " }
                            },
                        requestDB.FolderPath
                        );
                    }


                    if (int.Parse(requestDB.ToDateDB) < int.Parse(requestDB.FromDateDB) || int.Parse(endFinancialPeriod) < int.Parse(startFinancialPeriod) )
                {

                        Logger.WriteEntry(JsonConvert.SerializeObject("Date is invalid"), $"CheckDateInput--typeReport:Error");

                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $" ورودی نامعتبر " }
                            },
                        requestDB.FolderPath
                        );
                    }

                if (int.Parse(requestDB.ToDateDB) < int.Parse(requestDB.FromDateDB) || int.Parse(endFinancialPeriod) < int.Parse(startFinancialPeriod))
                {

                    Logger.WriteEntry(JsonConvert.SerializeObject("Date is invalid"), $"CheckDateInput--typeReport:Error");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $" ورودی نامعتبر " }
                        },
                    requestDB.FolderPath
                    );
                }

                return true;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public async Task<bool> CheckUserInput(Dictionary<string, string> config)
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

                return await Task.FromResult(true);
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"{ex}"), $"CheckInput--typeReport:Error");

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


    }
}
