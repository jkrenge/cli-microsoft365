import commands from '../../commands.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { validation } from '../../../../utils/validation.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    webUrl: z.string()
        .refine(url => validation.isValidSharePointUrl(url) === true, {
        error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
        .alias('u'),
    fileUrl: z.string().optional(),
    fileId: z.uuid().optional().alias('i'),
    label: z.string()
});
class SpoFileVersionKeepCommand extends SpoCommand {
    get name() {
        return commands.FILE_VERSION_KEEP;
    }
    get description() {
        return 'Ensure that a specific file version will never expire';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => [options.fileUrl, options.fileId].filter(o => o !== undefined).length === 1, {
            error: `Specify 'fileUrl' or 'fileId', but not both.`
        });
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Ensuring version '${args.options.label}' of file '${args.options.fileUrl || args.options.fileId}' at site '${args.options.webUrl}' will never expire...`);
        }
        try {
            const baseApiUrl = this.getBaseApiUrl(args.options.webUrl, args.options.fileUrl, args.options.fileId);
            const response = await odata.getAllItems(`${baseApiUrl}/versions?$filter=VersionLabel eq '${formatting.encodeQueryParameter(args.options.label)}'&$select=ID`);
            if (response.length === 0) {
                throw `Version with label '${args.options.label}' not found.`;
            }
            const requestExpirationOptions = {
                url: `${baseApiUrl}/versions(${response[0].ID})/SetExpirationDate()`,
                headers: {
                    accept: 'application/json;odata=nometadata',
                    'content-type': 'application/json'
                },
                responseType: 'json'
            };
            await request.post(requestExpirationOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getBaseApiUrl(webUrl, fileUrl, fileId) {
        let requestUrl;
        if (fileUrl) {
            const serverRelUrl = urlUtil.getServerRelativePath(webUrl, fileUrl);
            requestUrl = `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelUrl)}')`;
        }
        else {
            requestUrl = `${webUrl}/_api/web/GetFileById('${fileId}')`;
        }
        return requestUrl;
    }
}
export default new SpoFileVersionKeepCommand();
//# sourceMappingURL=file-version-keep.js.map