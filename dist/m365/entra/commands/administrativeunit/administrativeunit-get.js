import { z } from 'zod';
import request from "../../../../request.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";
import { entraAdministrativeUnit } from "../../../../utils/entraAdministrativeUnit.js";
import { globalOptionsZod } from "../../../../Command.js";
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    id: z.uuid().optional().alias('i'),
    displayName: z.string().optional().alias('n'),
    properties: z.string().optional().alias('p')
});
class EntraAdministrativeUnitGetCommand extends GraphCommand {
    get name() {
        return commands.ADMINISTRATIVEUNIT_GET;
    }
    get description() {
        return 'Gets information about a specific administrative unit';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => [options.id, options.displayName].filter(Boolean).length === 1, {
            error: 'Specify either id or displayName'
        });
    }
    async commandAction(logger, args) {
        let administrativeUnit;
        try {
            if (args.options.id) {
                administrativeUnit = await this.getAdministrativeUnitById(args.options.id, args.options.properties);
            }
            else {
                administrativeUnit = await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.displayName);
            }
            await logger.log(administrativeUnit);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getAdministrativeUnitById(id, properties) {
        const queryParameters = [];
        if (properties) {
            const allProperties = properties.split(',');
            const selectProperties = allProperties.filter(prop => !prop.includes('/'));
            if (selectProperties.length > 0) {
                queryParameters.push(`$select=${selectProperties}`);
            }
        }
        const queryString = queryParameters.length > 0
            ? `?${queryParameters.join('&')}`
            : '';
        const requestOptions = {
            url: `${this.resource}/v1.0/directory/administrativeUnits/${id}${queryString}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        return await request.get(requestOptions);
    }
}
export default new EntraAdministrativeUnitGetCommand();
//# sourceMappingURL=administrativeunit-get.js.map