using Discord;
using Discord.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MinusFifty.Commands
{
    class RequiredPermissionParameterAttribute : ParameterPreconditionAttribute
    {
        private GuildPermission _permission;
        public RequiredPermissionParameterAttribute(GuildPermission permission)
        {
            _permission = permission;
        }

        public async override Task<PreconditionResult> CheckPermissions(ICommandContext context, ParameterInfo parameter, object value, IServiceProvider map)
        {
            if (!(context is SocketCommandContext))
                return PreconditionResult.FromError("null");

            if (value == parameter.DefaultValue)
                return PreconditionResult.FromSuccess();

            IGuildUser guildUser = await context.Guild.GetUserAsync(context.User.Id);
            if (guildUser.GuildPermissions.Has(_permission))
                return PreconditionResult.FromSuccess();

            return PreconditionResult.FromError($"You must have the {_permission.ToString()} permission to use this command.");
        }
    }
}
