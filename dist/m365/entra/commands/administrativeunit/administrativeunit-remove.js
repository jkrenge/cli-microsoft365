import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    id: z.uuid().optional().alias('i'),
    displayName: z.string().optional().alias('n'),
    force: z.boolean().optional().alias('f')
});
class EntraAdministrativeUnitRemoveCommand extends GraphCommand {
    get name() {
        return commands.ADMINISTRATIVEUNIT_REMOVE;
    }
    get description() {
        return 'Removes an administrative unit';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => options.id || options.displayName, {
            error: 'Specify either id or displayName'
        })
            .refine(options => !(options.id && options.displayName), {
            error: 'Specify either id or displayName but not both'
        });
    }
    async commandAction(logger, args) {
        const removeAdministrativeUnit = async () => {
            try {
                let administrativeUnitId = args.options.id;
                if (args.options.displayName) {
                    const administrativeUnit = await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.displayName);
                    administrativeUnitId = administrativeUnit.id;
                }
                const requestOptions = {
                    url: `${this.resource}/v1.0/directory/administrativeUnits/${administrativeUnitId}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    }
                };
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeAdministrativeUnit();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove administrative unit '${args.options.id || args.options.displayName}'?` });
            if (result) {
                await removeAdministrativeUnit();
            }
        }
    }
}
export default new EntraAdministrativeUnitRemoveCommand();
//# sourceMappingURL=administrativeunit-remove.js.map