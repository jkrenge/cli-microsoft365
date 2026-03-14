import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import request from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    id: z.uuid().optional().alias('i'),
    displayName: z.string().optional().alias('d'),
    marketingNotificationEmails: z.string().refine(emails => validation.isValidUserPrincipalNameArray(emails) === true, {
        error: e => `The following marketing notification emails are invalid: ${e.input}.`
    }).transform((value) => value.split(',')).optional(),
    securityComplianceNotificationMails: z.string().refine(emails => validation.isValidUserPrincipalNameArray(emails) === true, {
        error: e => `The following security compliance notification emails are invalid: ${e.input}.`
    }).transform((value) => value.split(',')).optional(),
    securityComplianceNotificationPhones: z.string().transform((value) => value.split(',')).optional(),
    technicalNotificationMails: z.string().refine(emails => validation.isValidUserPrincipalNameArray(emails) === true, {
        error: e => `The following technical notification emails are invalid: ${e.input}.`
    }).transform((value) => value.split(',')).optional(),
    contactEmail: z.string().refine(id => validation.isValidUserPrincipalName(id), {
        error: e => `'${e.input}' is not a valid email.`
    }).optional(),
    statementUrl: z.string().optional()
});
class EntraOrganizationSetCommand extends GraphCommand {
    get name() {
        return commands.ORGANIZATION_SET;
    }
    get description() {
        return 'Updates info about the organization';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => !(options.id && options.displayName), {
            error: 'Specify either id or displayName, but not both'
        })
            .refine(options => options.id || options.displayName, {
            error: 'Specify either id or displayName'
        })
            .refine(options => [
            options.contactEmail, options.marketingNotificationEmails, options.securityComplianceNotificationMails, options.securityComplianceNotificationPhones,
            options.statementUrl, options.technicalNotificationMails
        ].filter(o => o !== undefined).length > 0, {
            error: 'Specify at least one of the following options: contactEmail, marketingNotificationEmails, securityComplianceNotificationMails, securityComplianceNotificationPhones, statementUrl, or technicalNotificationMails'
        });
    }
    async commandAction(logger, args) {
        try {
            let organizationId = args.options.id;
            if (args.options.displayName) {
                organizationId = await this.getOrganizationIdByDisplayName(args.options.displayName);
            }
            if (args.options.verbose) {
                await logger.logToStderr(`Updating organization with ID ${organizationId}...`);
            }
            const data = {
                marketingNotificationEmails: args.options.marketingNotificationEmails,
                securityComplianceNotificationMails: args.options.securityComplianceNotificationMails,
                securityComplianceNotificationPhones: args.options.securityComplianceNotificationPhones,
                technicalNotificationMails: args.options.technicalNotificationMails
            };
            if (args.options.contactEmail || args.options.statementUrl) {
                data.privacyProfile = {};
            }
            if (args.options.contactEmail) {
                data.privacyProfile.contactEmail = args.options.contactEmail;
            }
            if (args.options.statementUrl) {
                data.privacyProfile.statementUrl = args.options.statementUrl;
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/organization/${organizationId}`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json'
                },
                data: data,
                responseType: 'json'
            };
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getOrganizationIdByDisplayName(displayName) {
        const url = `${this.resource}/v1.0/organization?$select=id,displayName`;
        // the endpoint always returns one item
        const organizations = await odata.getAllItems(url);
        if (organizations[0].displayName !== displayName) {
            throw `The specified organization '${displayName}' does not exist.`;
        }
        return organizations[0].id;
    }
}
export default new EntraOrganizationSetCommand();
//# sourceMappingURL=organization-set.js.map