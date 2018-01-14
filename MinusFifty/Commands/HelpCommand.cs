using System;
using System.Threading.Tasks;
using Discord.Commands;

namespace MinusFifty.Commands
{
    [Group("help")]
    public class HelpCommand : ModuleBase<SocketCommandContext>
    {
        [Command]
        public async Task HelpBase()
        {
            await ReplyAsync("This is where you would get help with this bot, if it did anything yet, which it doesn't.");
        }
    }
}