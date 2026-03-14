import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request from "../../../../request.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";
const options = z.looseObject({
    ...globalOptionsZod.shape,
    displayName: z.string().alias('n'),
    description: z.string().optional().alias('d'),
    hiddenMembership: z.boolean().optional()
});
class EntraAdministrativeUnitAddCommand extends GraphCommand {
    get name() {
        return commands.ADMINISTRATIVEUNIT_ADD;
    }
    get description() {
        return 'Creates an administrative unit';
    }
    allowUnknownOptions() {
        return true;
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        const requestBody = {
            description: args.options.description,
            displayName: args.options.displayName,
            visibility: args.options.hiddenMembership ? 'HiddenMembership' : null
        };
        this.addUnknownOptionsToPayloadZod(requestBody, args.options);
        const requestOptions = {
            url: `${this.resource}/v1.0/directory/administrativeUnits`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: requestBody
        };
        try {
            const administrativeUnit = await request.post(requestOptions);
            await logger.log(administrativeUnit);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new EntraAdministrativeUnitAddCommand();
//# sourceMappingURL=administrativeunit-add.js.map