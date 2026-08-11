using Azure.Core;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Exceptions;
using WriteBalance.Application.Interfaces.Repository;
using WriteBalance.Common.Logging;
using WriteBalance.Domain.Entities;
using WriteBalance.Infrastructure.Context;
using WriteBalance.Infrastructure.Repositories;

namespace WriteBalance.Infrastructure.Services
{
    public class GetPooyaCoding
    {

        private readonly IPooyaCodingRepository _pooyaCodingRepository;

        public GetPooyaCoding( IPooyaCodingRepository pooyaCodingRepository)
        {
            _pooyaCodingRepository = pooyaCodingRepository;
        }

        public List<EquivalentCodePouya> ExecuteAsync(string FolderPath)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Start ExecuteAsync"), $"GetPooyaCoding:ExecuteAsync --typeReport:Debug");

                var Coding = _pooyaCodingRepository.GetPooyaCodingAsync(FolderPath);
                var DicCode = CreateDicFromnCode(Coding, FolderPath);
                return DicCode;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"GetPooyaCoding:ExecuteAsync --typeReport:Error");
                throw;
            }
        }

        public List<EquivalentCodePouya> CreateDicFromnCode(List<PooyaCoding> pooyaCodings, string FolderPath)
        {
            try
            {
                Logger.WriteEntry(
                    JsonConvert.SerializeObject("Start CreateDicFromnCode"),
                    "GetPooyaCoding:CreateDicFromnCode --typeReport:Debug");

                var equivalentCodes = pooyaCodings
                    .Select(x => new EquivalentCodePouya
                    {
                        SourceCode = $"{x.CodeKol}_{x.CodeArz}_{x.GroupMoein}",
                        EquivalentCode = $"{x.CodeKol}_{int.Parse(x.CodeOmoorMali):D4}"
                    })
                    .ToList();

                return equivalentCodes;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(
                    JsonConvert.SerializeObject(ex),
                    "GetPooyaCoding:CreateDicFromnCode --typeReport:Error");

                throw;
            }
        }


        //public Dictionary<string, string> CreateDicFromnCode(List<PooyaCoding> pooyaCodings, string FolderPath)
        //{
        //    try
        //    {

        //        Logger.WriteEntry(JsonConvert.SerializeObject("Start CreateDicFromnCode"), $"GetPooyaCoding:CreateDicFromnCode --typeReport:Debug");

        //        var dictionary = pooyaCodings.ToDictionary(
        //            x => $"{x.CodeKol}_{x.CodeArz}_{x.GroupMoein}",
        //            x => $"{x.CodeKol}_{int.Parse(x.CodeOmoorMali):D4}"
        //        );
        //        return dictionary;
        //    }
        //    catch (Exception ex) 
        //    {
        //        Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"GetPooyaCoding:ExecuteAsync --typeReport:Error");
        //        throw;
        //    }

        //}
    }
}
