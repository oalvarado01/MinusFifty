using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Discord;
using Discord.Commands;
using Discord.WebSocket;
using Google.Apis.Sheets.v4.Data;
using JsonConfig;
using MinusFifty;

namespace MinusFifty.Commands
{
    [Group("store")]
    public class StoreCommand : ModuleBase<SocketCommandContext>
    {
        [Command]
        [Summary("Displays information about the DKP store")]
        public async Task StoreAsync([RequiredRoleParameter(388954821551456256)] string announce = null)
        {
            ValueRange itemResult = await GoogleSheetsHelper.Instance.GetAsync(Config.Global.DKPStoreTab);
            string message = "--- DKP Store ---\n";
            foreach (IList<object> row in itemResult.Values)
            {
                message += $"{row[0]}: {row[1]} (Stock: {row[2]})\n";
            }

            if (!string.IsNullOrEmpty(announce) && (announce.Contains("announce") || announce.Contains("channel")))
            {
                await ReplyAsync(message);
            }
            else
            {
                await Context.User.SendMessageAsync(message);
            }
        }

        [Command("auto")]
        [Summary("Automatically purchases from the requests list until there are no requests able to be fulfilled")]
        [RequireOwner]
        public async Task StoreAutoAsync(bool verbose = false)
        {
            int _purchases = 0;
            bool _purchased = false;
            var storeWatch = System.Diagnostics.Stopwatch.StartNew();
            GridRange range = Program.GetTabRange(Program.EDKPTabs.Main);
            while (true)
            {
                List<Request> sortRequest = new List<Request>
                {
                    new Request
                    {
                        SortRange = new SortRangeRequest
                        {
                            Range = range,
                            SortSpecs = new List<SortSpec>
                            {
                                new SortSpec
                                {
                                    DimensionIndex = 1,
                                    SortOrder = "DESCENDING"
                                }
                            }
                        }
                    }
                };
                await GoogleSheetsHelper.Instance.TransactionAsync(sortRequest);

                string _name;
                _purchased = false;
                List<string> ranges = new List<string>
                {
                    Config.Global.DKPMainTab,
                    Config.Global.DKPRequestsTab,
                    Config.Global.DKPStoreTab
                };

                BatchGetValuesResponse response = await GoogleSheetsHelper.Instance.BatchGetAsync(ranges);
                if (response.ValueRanges == null || response.ValueRanges.Count < ranges.Count)
                {
                    await ReplyAsync("Error reading spreadsheet, please restart.");
                    break;
                }
                ValueRange dkpResult = response.ValueRanges[0];
                ValueRange reqResult = response.ValueRanges[1];
                ValueRange itemResult = response.ValueRanges[2];
                if (dkpResult.Values != null && dkpResult.Values.Count > 0 && reqResult.Values != null && reqResult.Values.Count > 0)
                {
                    foreach (IList<object> player in dkpResult.Values)
                    {
                        _name = player[0].ToString();
                        foreach (IList<object> request in reqResult.Values)
                        {
                            if (request[0].ToString().Equals(_name))
                            {
                                var result = await Program.ProcessPurchase(request[1].ToString(), _name, 1, itemResult, dkpResult, reqResult);
                                if (result.Item1)
                                {
                                    await ReplyAsync(result.Item2);
                                    _purchased = true;
                                    ++_purchases;
                                    break;
                                }
                                else if (verbose)
                                {
                                    await ReplyAsync(result.Item2);
                                }
                            }
                        }
                        if (_purchased)
                        {
                            break;
                        }
                    }
                }

                if (!_purchased)
                {
                    break;
                }
                await Task.Delay(Config.Global.Commands.Store.DelayMS);
            }
            storeWatch.Stop();

            await ReplyAsync($"Loot distribution has been processed with a total of {_purchases} purchases made over {storeWatch.ElapsedMilliseconds/1000} seconds.");
        }
    }
}