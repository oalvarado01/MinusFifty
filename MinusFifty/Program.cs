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
        public enum EDKPTabs : byte
        {
            Main,
            Store,
            Reasons,
            Requests,
            Log,
            PriorityLog,
            DKPTabsCount,
        };

        private static GridRange[] _tabRanges = new GridRange[(byte)EDKPTabs.DKPTabsCount];
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

            await LoadTabRanges();

            await Task.Delay(-1);
        }

        private async Task LoadTabRanges()
        {
            _tabRanges[(byte)EDKPTabs.Main] = await GoogleSheetsHelper.Instance.GetGridRange(Config.Global.DKPMainTab);
            _tabRanges[(byte)EDKPTabs.Store] = await GoogleSheetsHelper.Instance.GetGridRange(Config.Global.DKPStoreTab);
            _tabRanges[(byte)EDKPTabs.Reasons] = await GoogleSheetsHelper.Instance.GetGridRange(Config.Global.DKPReasonsTab);
            _tabRanges[(byte)EDKPTabs.Requests] = await GoogleSheetsHelper.Instance.GetGridRange(Config.Global.DKPRequestsTab);
            _tabRanges[(byte)EDKPTabs.Log] = await GoogleSheetsHelper.Instance.GetGridRange(Config.Global.DKPLogTab);
            _tabRanges[(byte)EDKPTabs.PriorityLog] = await GoogleSheetsHelper.Instance.GetGridRange(Config.Global.DKPPriorityLogTab);
        }

        public static GridRange GetTabRange(EDKPTabs tab)
        {
            if (tab >= EDKPTabs.DKPTabsCount)
            {
                return null;
            }

            return _tabRanges[(byte)tab];
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

            if (AttendanceEvent.excused.Contains(_name))
            {
                AttendanceEvent.excused.Remove(_name);
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

                if (AttendanceEvent.excused.Contains(_name))
                {
                    AttendanceEvent.excused.Remove(_name);
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

        public static async Task<Tuple<bool, string>> ProcessPurchase(string item, string name, int qty, ValueRange itemResult = null, ValueRange dkpResult = null, ValueRange reqResult = null)
        {
            // find the actual item from the store
            string _item = "";
            int _cost = 0;
            int _stock = 0;
            int _lotSize = 1;
            int _itemIdx = -1;
            if (itemResult == null)
            {
                itemResult = await GoogleSheetsHelper.Instance.GetAsync(Config.Global.DKPStoreTab);
            }
            if (itemResult.Values != null && itemResult.Values.Count > 0)
            {
                _itemIdx = GoogleSheetsHelper.Instance.IndexInRange(itemResult, item);
                if (_itemIdx >= 0)
                {
                    _item = itemResult.Values[_itemIdx][0].ToString();
                    int.TryParse(itemResult.Values[_itemIdx][1].ToString(), out _cost);
                    int.TryParse(itemResult.Values[_itemIdx][2].ToString(), out _stock);
                    int.TryParse(itemResult.Values[_itemIdx][3].ToString(), out _lotSize);
                }
            }

            if (string.IsNullOrEmpty(_item))
            {
                return new Tuple<bool,string>(false, $"Unable to find {item} in the DKP store! Use !store to list available items.");
            }

            // check available clan stock of item
            if (qty * _lotSize > _stock)
            {
                return new Tuple<bool,string>(false, $"Unable to purchase {item} for {name}: Out of stock");
            }

            // check available balance of player DKP
            int _dkp = 0;
            int _bonusPriority = 0;
            int _limit = 0;
            int _playerIdx = -1;
            if (dkpResult == null)
            {
                dkpResult = await GoogleSheetsHelper.Instance.GetAsync(Config.Global.DKPMainTab);
            }
            if (dkpResult.Values != null && dkpResult.Values.Count > 0)
            {
                _playerIdx = GoogleSheetsHelper.Instance.IndexInRange(dkpResult, name);
                if (_playerIdx >= 0)
                {
                    int.TryParse(dkpResult.Values[_playerIdx][2].ToString(), out _dkp);
                    int.TryParse(dkpResult.Values[_playerIdx][6].ToString(), out _bonusPriority);
                    int.TryParse(dkpResult.Values[_playerIdx][7].ToString(), out _limit);
                }
            }

            if (_dkp < _cost * qty)
            {
                return new Tuple<bool,string>(false, $"Insufficient DKP [{_dkp}] for {name} to cover the cost [{_cost * qty}] of {qty} {_item}");
            }

            if (_limit + qty * _lotSize >= Config.Global.Commands.Buy.WeeklyLimit)
            {
                return new Tuple<bool, string>(false, $"{name} is unable to purchase {_item} due to weekly limit of {_limit}/{Config.Global.Commands.Buy.WeeklyLimit}");
            }

            IList<Request> transaction = new List<Request>();

            // find any existing requests from this user for this item
            if (reqResult == null)
            {
                reqResult = await GoogleSheetsHelper.Instance.GetAsync(Config.Global.DKPRequestsTab);
            }
            if (reqResult.Values != null && reqResult.Values.Count > 0)
            {
                int reqIdx = GoogleSheetsHelper.Instance.SubIndexInRange(reqResult, name, 1, _item);
                if (reqIdx >= 0)
                {
                    int.TryParse(reqResult.Values[reqIdx][2].ToString(), out int currentQty);
                    currentQty -= qty;

                    GridRange staticReqRange = GetTabRange(EDKPTabs.Requests);
                    GridRange reqRange = new GridRange
                    {
                        SheetId = staticReqRange.SheetId,
                        StartRowIndex = reqIdx,
                        EndRowIndex = reqIdx + 1,
                        StartColumnIndex = staticReqRange.EndColumnIndex - 1,
                        EndColumnIndex = staticReqRange.EndColumnIndex
                    };

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
            GridRange logRange = GetTabRange(EDKPTabs.Log);
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
            if (_bonusPriority > 0)
            {
                GridRange priorityRange = GetTabRange(EDKPTabs.PriorityLog);
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
            GridRange staticStoreRange = GetTabRange(EDKPTabs.Store);
            GridRange storeRange = new GridRange
            {
                SheetId = staticStoreRange.SheetId,
                StartRowIndex = staticStoreRange.StartRowIndex + _itemIdx,
                EndRowIndex = staticStoreRange.StartRowIndex + _itemIdx + 1,
                StartColumnIndex = staticStoreRange.StartColumnIndex + 2,
                EndColumnIndex = staticStoreRange.StartColumnIndex + 3
            };
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
                                        NumberValue = _stock - (qty * _lotSize)
                                    }
                                }
                            }
                        }
                    }
                }
            });

            // update the weekly limit
            GridRange staticMainRange = GetTabRange(EDKPTabs.Main);
            GridRange dkpRange = new GridRange
            {
                SheetId = staticMainRange.SheetId,
                StartRowIndex = staticMainRange.StartRowIndex + _playerIdx,
                EndRowIndex = staticMainRange.StartRowIndex + _playerIdx + 1,
                StartColumnIndex = staticMainRange.StartColumnIndex + 7,
                EndColumnIndex = staticMainRange.EndColumnIndex
            };
            transaction.Add(new Request
            {
                UpdateCells = new UpdateCellsRequest
                {
                    Range = dkpRange,
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
                                        NumberValue = _limit + qty * _lotSize
                                    }
                                }
                            }
                        }
                    }
                }
            });

            BatchUpdateSpreadsheetResponse txResponse = await GoogleSheetsHelper.Instance.TransactionAsync(transaction);
            return new Tuple<bool,string>(true,$"{name} purchased {qty} {_item} for {_cost * qty} DKP");
        }
    }
}
