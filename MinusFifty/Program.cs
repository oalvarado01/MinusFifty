using Discord;
using Discord.WebSocket;
using Discord.Commands;
using System;
using System.Threading.Tasks;
using System.Reflection;
using Microsoft.Extensions.DependencyInjection;
using JsonConfig;
using Google.Apis.Sheets.v4.Data;

namespace MinusFifty
{
    class Program
    {
        private DiscordSocketClient _client;
        private CommandService _commands;
        private IServiceProvider _services;
        public static SAttendanceEvent AttendanceEvent { get; set; }

        public static void Main(string[] args)
            => new Program().MainAsync().GetAwaiter().GetResult();

        public async Task MainAsync()
        {

            _client = new DiscordSocketClient(new DiscordSocketConfig
            {
                LogLevel = LogSeverity.Info,
                MessageCacheSize = 200
            });

            _commands = new CommandService();

            _client.Log += Log;

            _services = new ServiceCollection()
                .AddSingleton(_client)
                .AddSingleton(_commands)
                .BuildServiceProvider();

            await InstallCommandsAsync();

            await _client.LoginAsync(TokenType.Bot, Config.Global.DiscordToken);
            await _client.StartAsync();

            //GoogleSheetsHelper.Instance.TestSheets();

            await Task.Delay(-1);
        }

        private Task Log(LogMessage msg)
        {
            Console.WriteLine(msg.ToString());
            return Task.CompletedTask;
        }

        public async Task InstallCommandsAsync()
        {
            _client.MessageReceived += HandleCommandAsync;
            _client.ReactionAdded += HandleReactionsAsync;
            await _commands.AddModulesAsync(Assembly.GetEntryAssembly());
        }

        private async Task HandleReactionsAsync(Cacheable<IUserMessage, ulong> before, ISocketMessageChannel channel, SocketReaction reactionParam)
        {
            if (!AttendanceEvent.active)
                return;

            if (reactionParam.User.ToString() == _client.CurrentUser.ToString())
                return;

            if (reactionParam.Emote.Name != new Emoji(Config.Global.Commands.Attendance.HereEmoji).Name)
                return;

            string _name = await GetIGNFromUser(reactionParam.User.Value);

            if (!AttendanceEvent.members.Contains(_name))
            {
                AttendanceEvent.members.Add(_name);
                await channel.SendMessageAsync($"Added {_name} to the {AttendanceEvent.type} event {AttendanceEvent.note}");
            }
        }

        private async Task HandleCommandAsync(SocketMessage messageParam)
        {
            var message = messageParam as SocketUserMessage;
            if (message == null)
                return;

            int argPos = 0;
            if (AttendanceEvent.active && message.HasCharPrefix('x', ref argPos) || message.HasCharPrefix('X', ref argPos))
            {
                string _name = await GetIGNFromUser(messageParam.Author);

                if (!AttendanceEvent.members.Contains(_name))
                {
                    AttendanceEvent.members.Add(_name);
                    await messageParam.Channel.SendMessageAsync($"Added {_name} to the {AttendanceEvent.type} event {AttendanceEvent.note}");
                    return;
                }                
            }
            if (!(message.HasCharPrefix((char)Config.Global.CommandPrefix[0], ref argPos) || message.HasMentionPrefix(_client.CurrentUser, ref argPos)))
                return;

            var context = new SocketCommandContext(_client, message);
            var result = await _commands.ExecuteAsync(context, argPos, _services);
            if (!result.IsSuccess)
                await context.Channel.SendMessageAsync(result.ErrorReason);
        }

        public static async Task<string> GetIGNFromUser(IUser user)
        {
            ValueRange readResult = await GoogleSheetsHelper.Instance.GetAsync(Config.Global.Commands.IGN.NameMapSheet);
            if (readResult.Values != null && readResult.Values.Count > 0)
            {
                int idx = GoogleSheetsHelper.Instance.IndexInRange(readResult, user.ToString());
                if (idx >= 0)
                {
                    return readResult.Values[idx][1].ToString();
                }
            }

            return user.Username;
        }
    }
}
