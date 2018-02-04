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

    public class StoreCommand : ModuleBase<SocketCommandContext>
    {
        [Command("store")]
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
    }
}