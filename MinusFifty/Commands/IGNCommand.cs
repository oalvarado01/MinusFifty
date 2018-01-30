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

    public class IGNCommand : ModuleBase<SocketCommandContext>
    {
        [Command("ign")]
        [Summary("Sets the current player's in-game name to <IGN>, or returns current known IGN if none specified.")]
        [Alias("name", "nickname", "nick")]
        public async Task IGNAsync([Summary("The name to set")] string name = "", [Summary("The user to set the name on"), RequiredRoleParameter(388954821551456256)] IUser user = null)
        {
            var userInfo = user ?? Context.Message.Author;
            string _discordUser = userInfo.ToString();

            string _name = "";
            ValueRange readResult = await GoogleSheetsHelper.Instance.GetAsync(Config.Global.Commands.IGN.NameMapSheet);
            if (readResult.Values != null && readResult.Values.Count > 0)
            {
                int idx = GoogleSheetsHelper.Instance.IndexInRange(readResult, _discordUser);
                if (idx >= 0)
                {
                    _name = readResult.Values[idx][1].ToString();
                    readResult.Values[idx][1] = name;
                }
            }

            // If no name is passed, the user wants to know their current name
            if (string.IsNullOrEmpty(name))
            {
                // and if the bot doesn't know their current name, it assumes it's their discord username
                if (string.IsNullOrEmpty(_name))
                {
                    _name = userInfo.Username;
                }
                await ReplyAsync($"I think {userInfo.Username}'s IGN is {_name}.");
            }
            else
            {
                int? result = -1;
                // if the bot doesn't know their current IGN, we're writing a new line
                if (string.IsNullOrEmpty(_name))
                {
                    IList<IList<object>> writeValues = new List<IList<object>> { new List<object> { _discordUser, name } };
                    ValueRange writeBody = new ValueRange
                    {
                        Values = writeValues
                    };
                    AppendValuesResponse writeResult = await GoogleSheetsHelper.Instance.AppendAsync(Config.Global.Commands.IGN.NameMapSheet, writeBody);
                    result = writeResult.Updates.UpdatedCells;
                }
                else // otherwise we're overwriting an existing line
                {
                    UpdateValuesResponse writeResult = await GoogleSheetsHelper.Instance.UpdateAsync(readResult.Range, readResult);
                    result = writeResult.UpdatedCells;
                }

                if (result <= 0)
                {
                    await ReplyAsync($"Error updating {userInfo.Username}'s IGN");
                }
                else
                {
                    await Context.Message.AddReactionAsync(new Emoji(Config.Global.AcknowledgeEmoji));
                }
            }
        }
    }
}