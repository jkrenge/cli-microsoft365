import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';
import Auth from '../../../../Auth.js';
const consumingTenantOverridablesOptions = ['urlTemplate', 'isDiscoverabilityEnabled', 'isSearchEnabled', 'isItemVersioningEnabled', 'itemMajorVersionLimit', 'maxStoragePerContainerInBytes'];
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    name: z.string().alias('n'),
    appId: z.uuid().optional(),
    billingType: z.enum(['standard', 'trial', 'directToCustomer']).default('standard'),
    consumingTenantOverridables: z.string()
        .refine(values => values.split(',').every(v => consumingTenantOverridablesOptions.includes(v.trim())), {
        error: e => `'${e.input}' is not a valid value. Valid options are: ${consumingTenantOverridablesOptions.join(', ')}.`
    }).optional(),
    isDiscoverabilityEnabled: z.boolean().optional(),
    isItemVersioningEnabled: z.boolean().optional(),
    isSearchEnabled: z.boolean().optional(),
    isSharingRestricted: z.boolean().optional(),
    itemMajorVersionLimit: z.number()
        .refine(n => validation.isValidPositiveInteger(n), {
        error: e => `'${e.input}' is not a valid positive integer.`
    }).optional(),
    maxStoragePerContainerInBytes: z.number()
        .refine(n => validation.isValidPositiveInteger(n), {
        error: e => `'${e.input}' is not a valid positive integer.`
    }).optional(),
    sharingCapability: z.enum(['disabled', 'externalUserSharingOnly', 'externalUserAndGuestSharing', 'existingExternalUserSharingOnly']).optional(),
    urlTemplate: z.string().optional()
});
class SpeContainerTypeAddCommand extends GraphDelegatedCommand {
    get name() {
        return commands.CONTAINERTYPE_ADD;
    }
    get description() {
        return 'Creates a new container type';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => !options.itemMajorVersionLimit || options.isItemVersioningEnabled !== false, {
            error: `Cannot set itemMajorVersionLimit when isItemVersioningEnabled is false.`
        });
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Creating a new Container Type for your app with name '${args.options.name}'.`);
            }
            const appId = args.options.appId ?? Auth.connection.appId;
            const requestOptions = {
                url: `${this.resource}/beta/storage/fileStorage/containerTypes`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    name: args.options.name,
                    owningAppId: appId,
                    billingClassification: args.options.billingType,
                    settings: {
                        consumingTenantOverridables: args.options.consumingTenantOverridables?.split(',').map(s => s.trim()).join(','),
                        isDiscoverabilityEnabled: args.options.isDiscoverabilityEnabled,
                        isItemVersioningEnabled: args.options.isItemVersioningEnabled,
                        isSearchEnabled: args.options.isSearchEnabled,
                        isSharingRestricted: args.options.isSharingRestricted,
                        itemMajorVersionLimit: args.options.itemMajorVersionLimit,
                        maxStoragePerContainerInBytes: args.options.maxStoragePerContainerInBytes,
                        sharingCapability: args.options.sharingCapability,
                        urlTemplate: args.options.urlTemplate
                    }
                }
            };
            const response = await request.post(requestOptions);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpeContainerTypeAddCommand();
//# sourceMappingURL=containertype-add.js.map