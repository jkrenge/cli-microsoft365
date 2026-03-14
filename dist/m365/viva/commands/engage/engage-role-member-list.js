import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    roleId: z.uuid().optional().alias('i'),
    roleName: z.string().optional().alias('n')
});
class VivaEngageRoleMemberListCommand extends GraphCommand {
    get name() {
        return commands.ENGAGE_ROLE_MEMBER_LIST;
    }
    get description() {
        return 'Lists all users assigned to a Viva Engage role';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => [options.roleId, options.roleName].filter(x => x !== undefined).length === 1, {
            error: 'Specify either roleId, or roleName, but not both.'
        });
    }
    defaultProperties() {
        return ['id', 'userId'];
    }
    async commandAction(logger, args) {
        let roleId = args.options.roleId;
        try {
            if (args.options.roleName) {
                if (this.verbose) {
                    await logger.logToStderr(`Retrieving Viva Engage role ID for role name '${args.options.roleName}'...`);
                }
                roleId = await vivaEngage.getRoleIdByName(args.options.roleName);
            }
            if (this.verbose) {
                await logger.logToStderr(`Getting all users assigned to a Viva Engage role ${roleId}...`);
            }
            const results = await odata.getAllItems(`${this.resource}/beta/employeeExperience/roles/${roleId}/members`);
            await logger.log(results);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new VivaEngageRoleMemberListCommand();
//# sourceMappingURL=engage-role-member-list.js.map