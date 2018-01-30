﻿using Discord;
using Discord.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MinusFifty.Commands
{
    class RequiredRoleParameterAttribute : ParameterPreconditionAttribute
    {
        private ulong _roleId;
        public RequiredRoleParameterAttribute(ulong roleId)
        {
            _roleId = roleId;
        }

        public async override Task<PreconditionResult> CheckPermissions(ICommandContext context, ParameterInfo parameter, object value, IServiceProvider map)
        {
            if (!(context is SocketCommandContext))
                return PreconditionResult.FromError("null");

            if (value == parameter.DefaultValue)
                return PreconditionResult.FromSuccess();

            var roles = context.Guild.Roles;
            IRole role = context.Guild.GetRole(_roleId);
            if (role == null)
                return PreconditionResult.FromError("The role required to use this command was not found.");

            IGuildUser guildUser = await context.Guild.GetUserAsync(context.User.Id);
            if (guildUser.RoleIds.Contains<ulong>(_roleId))
                return PreconditionResult.FromSuccess();

            return PreconditionResult.FromError($"You must have the {role.Name} permission to use this command.");
        }
    }
}
