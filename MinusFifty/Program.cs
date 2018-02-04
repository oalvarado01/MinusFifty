using Discord;
using Discord.WebSocket;
using Discord.Commands;
using System;
using System.Threading.Tasks;
using System.Reflection;
using Microsoft.Extensions.DependencyInjection;
using JsonConfig;
using Google.Apis.Sheets.v4.Data;
using System.Collections.Generic;

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

        public static async Task<string> ProcessPurchase(string item, string name, int qty)
        {
            // find the actual item from the store
            string _item = "";
            int _cost = 0;
            int _stock = 0;
            int _itemIdx = -1;
            ValueRange itemResult = await GoogleSheetsHelper.Instance.GetAsync(Config.Global.DKPStoreTab);
            if (itemResult.Values != null && itemResult.Values.Count > 0)
            {
                _itemIdx = GoogleSheetsHelper.Instance.IndexInRange(itemResult, item);
                if (_itemIdx >= 0)
                {
                    _item = itemResult.Values[_itemIdx][0].ToString();
                    int.TryParse(itemResult.Values[_itemIdx][1].ToString(), out _cost);
                    int.TryParse(itemResult.Values[_itemIdx][2].ToString(), out _stock);
                }
            }

            if (string.IsNullOrEmpty(_item))
            {
                return $"Unable to find {item} in the DKP store! Use !store to list available items.";
            }

            // check available clan stock of item
            qty = Math.Min(qty, _stock);
            if (qty <= 0)
            {
                return $"The clan doesn't have any {item} available to buy!";
            }

            // check available balance of player DKP
            int _dkp = 0;
            int _bonusPriority = 0;
            ValueRange dkpResult = await GoogleSheetsHelper.Instance.GetAsync(Config.Global.DKPMainTab);
            if (dkpResult.Values != null && dkpResult.Values.Count > 0)
            {
                int idx = GoogleSheetsHelper.Instance.IndexInRange(dkpResult, name);
                if (idx >= 0)
                {
                    int.TryParse(dkpResult.Values[idx][2].ToString(), out _dkp);
                    int.TryParse(dkpResult.Values[idx][6].ToString(), out _bonusPriority);
                }
            }

            if (_dkp < _cost * qty)
            {
                return $"Insufficient DKP [{_dkp}] for {name} to cover the cost [{_cost * qty}] of {qty} {_item}";
            }

            IList<Request> transaction = new List<Request>();

            // find any existing requests from this user for this item
            ValueRange reqResult = await GoogleSheetsHelper.Instance.GetAsync(Config.Global.DKPRequestsTab);
            if (reqResult.Values != null && reqResult.Values.Count > 0)
            {
                int reqIdx = GoogleSheetsHelper.Instance.SubIndexInRange(reqResult, name, 1, _item);
                if (reqIdx >= 0)
                {
                    GridRange reqRange = await GoogleSheetsHelper.Instance.GetGridRange(Config.Global.DKPRequestsTab);
                    int.TryParse(reqResult.Values[reqIdx][2].ToString(), out int currentQty);
                    currentQty -= qty;
                    reqRange.StartRowIndex = reqIdx;
                    reqRange.EndRowIndex = reqRange.StartRowIndex + 1;
                    reqRange.StartColumnIndex = reqRange.EndColumnIndex;
                    reqRange.EndColumnIndex = reqRange.StartColumnIndex + 1;

                    if (currentQty > 0)
                    {
                        transaction.Add(new Request
                        {
                            UpdateCells = new UpdateCellsRequest
                            {
                                Range = reqRange,
                                Fields = "*",
                                Rows = new List<RowData>
                                {
                                    new RowData
                                    {
                                        Values = new List<CellData>
                                        {
                                            new CellData
                                            {
                                                UserEnteredValue = new ExtendedValue
                                                {
                                                    NumberValue = currentQty
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        });
                    }
                    else
                    {
                        transaction.Add(new Request
                        {
                            DeleteDimension = new DeleteDimensionRequest
                            {
                                Range = new DimensionRange
                                {
                                    Dimension = "ROWS",
                                    SheetId = reqRange.SheetId,
                                    StartIndex = reqRange.StartRowIndex,
                                    EndIndex = reqRange.EndRowIndex
                                }
                            }
                        });
                    }
                }
            }

            // log the purchase
            GridRange logRange = await GoogleSheetsHelper.Instance.GetGridRange(Config.Global.DKPLogTab);
            transaction.Add(new Request
            {
                AppendCells = new AppendCellsRequest
                {
                    SheetId = logRange.SheetId,
                    Fields = "*",
                    Rows = new List<RowData>
                    {
                        new RowData
                        {
                            Values = new List<CellData>
                            {
                                new CellData
                                {
                                    UserEnteredValue = new ExtendedValue
                                    {
                                        StringValue = name
                                    }
                                },
                                new CellData
                                {
                                    UserEnteredValue = new ExtendedValue
                                    {
                                        NumberValue = _cost * qty * -1
                                    }
                                },
                                new CellData
                                {
                                    UserEnteredValue = new ExtendedValue
                                    {
                                        NumberValue = DateTime.Now.ToOADate()
                                    }
                                },
                                new CellData
                                {
                                    UserEnteredValue = new ExtendedValue
                                    {
                                        StringValue = Config.Global.Commands.Buy.PurchaseReason
                                    }
                                },
                                new CellData
                                {
                                    UserEnteredValue = new ExtendedValue
                                    {
                                        StringValue = _item
                                    }
                                }
                            }
                        }
                    }
                }
            });

            // log the priority (if necessary)
            if (_bonusPriority >= 0)
            {
                GridRange priorityRange = await GoogleSheetsHelper.Instance.GetGridRange(Config.Global.DKPPriorityLogTab);
                transaction.Add(new Request
                {
                    AppendCells = new AppendCellsRequest
                    {
                        SheetId = priorityRange.SheetId,
                        Fields = "*",
                        Rows = new List<RowData>
                        {
                            new RowData
                            {
                                Values = new List<CellData>
                                {
                                    new CellData
                                    {
                                        UserEnteredValue = new ExtendedValue
                                        {
                                            StringValue = name
                                        }
                                    },
                                    new CellData
                                    {
                                        UserEnteredValue = new ExtendedValue
                                        {
                                            NumberValue = Math.Min(_bonusPriority, _cost * qty) * -1
                                        }
                                    },
                                    new CellData
                                    {
                                        UserEnteredValue = new ExtendedValue
                                        {
                                            NumberValue = DateTime.Now.ToOADate()
                                        }
                                    },
                                    new CellData
                                    {
                                        UserEnteredValue = new ExtendedValue
                                        {
                                            StringValue = _item
                                        }
                                    }
                                }
                            }
                        }
                    }
                });
            }

            // update the stock
            GridRange storeRange = await GoogleSheetsHelper.Instance.GetGridRange(Config.Global.DKPStoreTab);
            storeRange.StartRowIndex += _itemIdx;
            storeRange.EndRowIndex = storeRange.StartRowIndex + 1;
            storeRange.StartColumnIndex += 2;
            storeRange.EndColumnIndex = storeRange.StartColumnIndex + 1;
            transaction.Add(new Request
            {
                UpdateCells = new UpdateCellsRequest
                {
                    Range = storeRange,
                    Fields = "*",
                    Rows = new List<RowData>
                    {
                        new RowData
                        {
                            Values = new List<CellData>
                            {
                                new CellData
                                {
                                    UserEnteredValue = new ExtendedValue
                                    {
                                        NumberValue = _stock - qty
                                    }
                                }
                            }
                        }
                    }
                }
            });


            BatchUpdateSpreadsheetResponse txResponse = await GoogleSheetsHelper.Instance.TransactionAsync(transaction);
            return $"{name} purchased {qty} {_item} for {_cost * qty} DKP";
        }
    }
}
