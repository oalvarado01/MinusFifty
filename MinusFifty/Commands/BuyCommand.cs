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

    public class BuyCommand : ModuleBase<SocketCommandContext>
    {
        [Command("buy")]
        [Summary("Buys <qty> of <item> from clan storage [for <player> if specified, otherwise for the current user]")]
        [RequireOwner]
        public async Task RequestAsync([Summary("The item to buy")] string item = "", [Summary("The amount to buy")] int qty = 1, [Summary("The user to request the item for"), RequiredRoleParameter(388954821551456256)] IUser user = null)
        {
            IUser userInfo = user ?? Context.Message.Author;
            string _name = await Program.GetIGNFromUser(userInfo);

            string result = await Program.ProcessPurchase(item, _name, qty);
            await ReplyAsync(result);
        }
    }
}