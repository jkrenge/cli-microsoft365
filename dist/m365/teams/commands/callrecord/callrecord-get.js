import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import commands from '../../commands.js';
import GraphApplicationCommand from '../../../base/GraphApplicationCommand.js';
import request from '../../../../request.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    id: z.uuid().alias('i')
});
class TeamsCallRecordGetCommand extends GraphApplicationCommand {
    get name() {
        return commands.CALLRECORD_GET;
    }
    get description() {
        return 'Gets a specific Teams call';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        try {
            const callRecordId = args.options.id;
            if (this.verbose) {
                await logger.logToStderr(`Retrieving call record ${callRecordId}...`);
            }
            // only one relationship can be expanded at a time
            let requestOptions = {
                url: `${this.resource}/v1.0/communications/callRecords/${callRecordId}?$expand=sessions($expand=segments)`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const callRecordPart1 = await request.get(requestOptions);
            requestOptions = {
                url: `${this.resource}/v1.0/communications/callRecords/${callRecordId}?$select=participants_v2&$expand=participants_v2`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const callRecordPart2 = await request.get(requestOptions);
            const callRecord = { ...callRecordPart1, ...callRecordPart2 };
            await logger.log(callRecord);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new TeamsCallRecordGetCommand();
//# sourceMappingURL=callrecord-get.js.map