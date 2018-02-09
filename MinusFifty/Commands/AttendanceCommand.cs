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
    [Group("attendance")]
    public class AttendanceCommand : ModuleBase<SocketCommandContext>
    {
        [Command("start")]
        [Summary("Starts tracking attendance for an event of <type> and <note>")]
        [RequiredRole(388954821551456256)]
        public async Task StartAsync([Summary("Event type")] string type, [Summary("Note for the DKP Log")] [Remainder] string note)
        {
            if (Program.AttendanceEvent.active)
            {
                await ReplyAsync($"Unable to start event, another event is still running: {Program.AttendanceEvent.type} {Program.AttendanceEvent.note}");
                return;
            }

            string[] types = Config.Global.Commands.Attendance.ValidAttendanceTypes;
            bool valid = false;
            foreach (string validType in types)
            {
                if (type == validType)
                {
                    valid = true;
                    break;
                }
            }

            if (!valid)
            {
                await ReplyAsync($"I don't recognize attendance type {type}");
                return;
            }

            Discord.Rest.RestUserMessage message = await Context.Channel.SendMessageAsync($"<@&{Config.Global.Commands.Attendance.MemberRole}> Attendance tracking for {type} {note} has begun, click the check mark to be counted!");
            Program.AttendanceEvent = new SAttendanceEvent(type, note, message.Id);
            await message.AddReactionAsync(new Emoji(Config.Global.Commands.Attendance.HereEmoji));
            await message.PinAsync();
        }

        [Command("add")]
        [Summary("Adds <user> to the current event so they get full credit")]
        [RequiredRole(388954821551456256)]
        public async Task AddAsync([Summary("The user to add to this event")] IUser user = null)
        {
            if (user == null)
            {
                await ReplyAsync("Unable to add user: no user provided!");
                return;
            }

            string _name = await Program.GetIGNFromUser(user);
            if (!Program.AttendanceEvent.members.Contains(_name))
                Program.AttendanceEvent.members.Add(_name);

            await Context.Message.AddReactionAsync(new Emoji(Config.Global.AcknowledgeEmoji));
        }

        [Command("excuse")]
        [Summary("Excuses <user> from the current event so they still get partial credit")]
        [RequiredRole(388954821551456256)]
        public async Task ExcuseAsync([Summary("The user to excuse from this event")] IUser user = null)
        {
            if (user == null)
            {
                await ReplyAsync("Unable to excuse user: no user provided!");
                return;
            }

            string _name = await Program.GetIGNFromUser(user);
            if (!Program.AttendanceEvent.excused.Contains(_name))
                Program.AttendanceEvent.excused.Add(_name);

            await Context.Message.AddReactionAsync(new Emoji(Config.Global.AcknowledgeEmoji));
        }

        [Command("list")]
        [Summary("Shows who is currently participating in the active event")]
        public async Task ListAsync()
        {
            if (!Program.AttendanceEvent.active)
            {
                await ReplyAsync("There is no attendance event currently running!");
                return;
            }

            string message = $"Event {Program.AttendanceEvent.type} {Program.AttendanceEvent.note} currently has the following {Program.AttendanceEvent.members.Count} members attending: ";
            foreach (string member in Program.AttendanceEvent.members)
                message += $"{member}, ";

            if (Program.AttendanceEvent.excused.Count > 0)
                message += $"and the following {Program.AttendanceEvent.excused.Count} member(s) excused: ";
            foreach (string excused in Program.AttendanceEvent.excused)
                message += $"{excused}, ";

            message += "and that's all, folks!";

            await ReplyAsync(message);
        }

        [Command("stop")]
        [Summary("Freezes tracking attendance for the current event and posts DKP")]
        [RequiredRole(388954821551456256)]
        public async Task StopAsync()
        {
            if (!Program.AttendanceEvent.active)
            {
                await ReplyAsync("Unable to stop event, no event is running!");
                return;
            }

            string _attendanceName = Program.AttendanceEvent.type + Config.Global.Commands.Attendance.AttendanceSuffix;
            string _excusedName = Program.AttendanceEvent.type + Config.Global.Commands.Attendance.ExcusedSuffix;
            int attendanceWeight = 0;
            int excusedWeight = 0;

            ValueRange readResult = await GoogleSheetsHelper.Instance.GetAsync(Config.Global.DKPReasonsTab);
            if (readResult.Values != null && readResult.Values.Count > 0)
            {
                int idx = GoogleSheetsHelper.Instance.IndexInRange(readResult, _attendanceName);
                if (idx >= 0)
                {
                     int.TryParse(readResult.Values[idx][1].ToString(), out attendanceWeight);
                }

                idx = GoogleSheetsHelper.Instance.IndexInRange(readResult, _excusedName);
                if (idx >= 0)
                {
                    int.TryParse(readResult.Values[idx][1].ToString(), out excusedWeight);
                }
            }

            string timestamp = DateTime.Now.ToString();
            IList<IList<object>> writeValues = new List<IList<object>>();
            foreach (string member in Program.AttendanceEvent.members)
            {
                writeValues.Add(new List<object> { member, attendanceWeight, timestamp, _attendanceName, Program.AttendanceEvent.note });
            }
            foreach (string excused in Program.AttendanceEvent.excused)
            {
                writeValues.Add(new List<object> { excused, excusedWeight, timestamp, _excusedName, Program.AttendanceEvent.note });
            }
            ValueRange writeBody = new ValueRange
            {
                Values = writeValues
            };
            AppendValuesResponse writeResult = await GoogleSheetsHelper.Instance.AppendAsync(Config.Global.DKPLogTab, writeBody);
            int? result = writeResult.Updates.UpdatedCells;

            if (result <= 0)
            {
                await ReplyAsync($"Error writing DKP for event!");
            }

            string message = $"Ended event {Program.AttendanceEvent.type} {Program.AttendanceEvent.note} with the following {Program.AttendanceEvent.members.Count} members attending: ";
            foreach (string member in Program.AttendanceEvent.members)
                message += $"{member}, ";

            if (Program.AttendanceEvent.excused.Count > 0)
                message += $"and the following {Program.AttendanceEvent.excused.Count} member(s) excused: ";
            foreach (string excused in Program.AttendanceEvent.excused)
                message += $"{excused}, ";

            message += "and that's all, folks!";

            IUserMessage pin = await Context.Channel.GetMessageAsync(Program.AttendanceEvent.messageId) as IUserMessage;
            if (pin != null)
            {
                await pin.UnpinAsync();
            }

            Program.AttendanceEvent = default(SAttendanceEvent);
            await Context.Channel.SendMessageAsync(message);
        }
    }
}