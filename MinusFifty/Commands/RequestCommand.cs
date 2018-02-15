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
    [Group("request")]
    public class RequestCommand : ModuleBase<SocketCommandContext>
    {
        [Command("list")]
        public async Task RequestListAsync([Summary("The user to request the item for"), RequiredRoleParameter(388954821551456256)] IUser user = null)
        {
            IUser userInfo = user ?? Context.Message.Author;
            string _name = await Program.GetIGNFromUser(userInfo);

            string response = $"{_name} is currently requesting ";
            bool found = false;
            ValueRange reqResult = await GoogleSheetsHelper.Instance.GetAsync(Config.Global.DKPRequestsTab);
            if (reqResult.Values != null && reqResult.Values.Count > 0)
            {
                foreach (IList<object> row in reqResult.Values)
                {
                    if (row[0].ToString().Equals(_name))
                    {
                        response += $"\n{row[2].ToString()}x {row[1].ToString()}";
                        found = true;
                    }
                }
            }

            if (!found)
            {
                response += "nothing.";
            }

            await ReplyAsync(response);
        }

        [Command]
        [Summary("Requests <qty> of <item> from clan storage [for <player> if specified, otherwise for the current user]")]
        public async Task RequestAsync([Summary("The item to request")] string item = "", [Summary("The amount to request")] int qty = 1, [Summary("The user to request the item for"), RequiredRoleParameter(388954821551456256)] IUser user = null)
        {
            IUser userInfo = user ?? Context.Message.Author;
            string _name = await Program.GetIGNFromUser(userInfo);

            // find the actual item name from the store
            string _item = "";
            ValueRange itemResult = await GoogleSheetsHelper.Instance.GetAsync(Config.Global.DKPStoreTab);
            if (itemResult.Values != null && itemResult.Values.Count > 0)
            {
                int idx = GoogleSheetsHelper.Instance.IndexInRange(itemResult, item);
                if (idx >= 0)
                {
                    _item = itemResult.Values[idx][0].ToString();
                }
            }

            if (string.IsNullOrEmpty(_item))
            {
                await ReplyAsync($"Unable to find {item} in the DKP store! Use !store to list available items.");
                return;
            }

            // find any existing requests from this user for this item
            int reqIdx = -1;
            ValueRange reqResult = await GoogleSheetsHelper.Instance.GetAsync(Config.Global.DKPRequestsTab);
            if (reqResult.Values != null && reqResult.Values.Count > 0)
            {
                reqIdx = GoogleSheetsHelper.Instance.SubIndexInRange(reqResult, _name, 1, _item);
            }

            int? result = -1;
            // remove items from request
            if (qty < 0)
            {
                if (reqIdx < 0)
                {
                    await ReplyAsync($"{_name} is no longer requesting {_item} (and wasn't in the first place)");
                    return;
                }

                int.TryParse(reqResult.Values[reqIdx][2].ToString(), out int currentQty);
                currentQty += qty;

                // if removing items requests in no items remaining, delete this request row
                if (currentQty <= 0)
                {
                    GridRange gridRange = Program.GetTabRange(Program.EDKPTabs.Requests);
                    
                    BatchUpdateSpreadsheetResponse deleteResult = await GoogleSheetsHelper.Instance.DeleteRowsAsync(gridRange.SheetId ?? 0, reqIdx);
                    result = deleteResult.Replies.Count;
                    if (result < 0)
                    {
                        await ReplyAsync($"Error deleting {_name}'s request!");
                    }
                    else
                    {
                        await ReplyAsync($"{_name} is no longer requesting {_item}");
                    }
                }
                else // otherwise just deduct from the existing request
                {
                    reqResult.Values[reqIdx][2] = currentQty.ToString();
                    UpdateValuesResponse writeResult = await GoogleSheetsHelper.Instance.UpdateAsync(reqResult.Range, reqResult);
                    result = writeResult.UpdatedCells;

                    if (result < 0)
                    {
                        await ReplyAsync($"Error removing {_item} from request for {_name}");
                    }
                    else
                    {
                        await ReplyAsync($"{_name} is now requesting {currentQty.ToString()} {_item}");
                    }
                }
            }
            else // adding items to request
            {
                int currentQty = 0;
                if (reqIdx > 0)
                {
                    int.TryParse(reqResult.Values[reqIdx][2].ToString(), out currentQty);
                }
                
                if (currentQty > 0) // updating request
                {
                    currentQty += qty;
                    reqResult.Values[reqIdx][2] = currentQty.ToString();

                    UpdateValuesResponse writeResult = await GoogleSheetsHelper.Instance.UpdateAsync(reqResult.Range, reqResult);
                    result = writeResult.UpdatedCells;

                    if (result < 0)
                    {
                        await ReplyAsync($"Error adding {_item} to request for {_name}");
                    }
                    else
                    {
                        await ReplyAsync($"{_name} is now requesting {currentQty.ToString()} {_item}");
                    }
                }
                else // new request
                {
                    IList<IList<object>> writeValues = new List<IList<object>> { new List<object> { _name, _item, qty.ToString() } };
                    ValueRange writeBody = new ValueRange
                    {
                        Values = writeValues
                    };
                    AppendValuesResponse writeResult = await GoogleSheetsHelper.Instance.AppendAsync(Config.Global.DKPRequestsTab, writeBody);
                    result = writeResult.Updates.UpdatedCells;

                    if (result < 0)
                    {
                        await ReplyAsync($"Error adding {_item} to request for {_name}");
                    }
                    else
                    {
                        await ReplyAsync($"{_name} is now requesting {qty.ToString()} {_item}");
                    }
                }
            }
        }
    }
}