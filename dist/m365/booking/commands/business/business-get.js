import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    id: z.string().optional().alias('i'),
    name: z.string().optional().alias('n')
});
class BookingBusinessGetCommand extends GraphCommand {
    get name() {
        return commands.BUSINESS_GET;
    }
    get description() {
        return 'Retrieve the specified Microsoft Bookings business.';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => options.id || options.name, {
            error: 'Specify either id or name'
        });
    }
    async commandAction(logger, args) {
        try {
            const businessId = await this.getBusinessId(args.options);
            const requestOptions = {
                url: `${this.resource}/v1.0/solutions/bookingBusinesses/${formatting.encodeQueryParameter(businessId)}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const business = await request.get(requestOptions);
            await logger.log(business);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getBusinessId(options) {
        if (options.id) {
            return options.id;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/solutions/bookingBusinesses`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        const name = options.name;
        const bookingBusinesses = response.value.filter(val => val.displayName?.toLocaleLowerCase() === name.toLocaleLowerCase());
        if (!bookingBusinesses.length) {
            throw `The specified business with name ${options.name} does not exist.`;
        }
        if (bookingBusinesses.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', bookingBusinesses);
            const result = await cli.handleMultipleResultsFound(`Multiple businesses with name '${options.name}' found.`, resultAsKeyValuePair);
            return result.id;
        }
        return bookingBusinesses[0].id;
    }
}
export default new BookingBusinessGetCommand();
//# sourceMappingURL=business-get.js.map