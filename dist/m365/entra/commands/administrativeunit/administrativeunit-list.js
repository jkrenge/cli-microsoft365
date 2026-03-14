import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    properties: z.string().optional().alias('p')
});
class EntraAdministrativeUnitListCommand extends GraphCommand {
    get name() {
        return commands.ADMINISTRATIVEUNIT_LIST;
    }
    get description() {
        return 'Retrieves a list of administrative units';
    }
    get schema() {
        return options;
    }
    defaultProperties() {
        return ['id', 'displayName', 'visibility'];
    }
    async commandAction(logger, args) {
        const queryParameters = [];
        if (args.options.properties) {
            const allProperties = args.options.properties.split(',');
            const selectProperties = allProperties.filter(prop => !prop.includes('/'));
            if (selectProperties.length > 0) {
                queryParameters.push(`$select=${selectProperties}`);
            }
        }
        const queryString = queryParameters.length > 0
            ? `?${queryParameters.join('&')}`
            : '';
        try {
            const results = await odata.getAllItems(`${this.resource}/v1.0/directory/administrativeUnits${queryString}`);
            await logger.log(results);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new EntraAdministrativeUnitListCommand();
//# sourceMappingURL=administrativeunit-list.js.map