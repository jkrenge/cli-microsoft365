import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    webUrl: z.string()
        .refine(url => validation.isValidSharePointUrl(url) === true, {
        error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
        .alias('u'),
    name: z.string().optional().alias('n'),
    default: z.boolean().optional(),
    metadataOnly: z.boolean().optional()
});
class SpoPageGetCommand extends SpoCommand {
    get name() {
        return commands.PAGE_GET;
    }
    get description() {
        return 'Gets information about the specific modern page';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => [options.name, options.default].filter(x => x !== undefined).length === 1, {
            error: `Specify either name or default, but not both.`
        });
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving information about the page...`);
        }
        let pageName = '';
        try {
            if (args.options.name) {
                pageName = args.options.name.endsWith('.aspx')
                    ? args.options.name
                    : `${args.options.name}.aspx`;
            }
            else if (args.options.default) {
                const requestOptions = {
                    url: `${args.options.webUrl}/_api/Web/RootFolder?$select=WelcomePage`,
                    headers: {
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                const { WelcomePage } = await request.get(requestOptions);
                pageName = WelcomePage.split('/').pop();
            }
            let requestOptions = {
                url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/SitePages/${formatting.encodeQueryParameter(pageName)}')?$expand=ListItemAllFields/ClientSideApplicationId,ListItemAllFields/PageLayoutType,ListItemAllFields/CommentsDisabled`,
                headers: {
                    'content-type': 'application/json;charset=utf-8',
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const page = await request.get(requestOptions);
            if (page.ListItemAllFields.ClientSideApplicationId !== 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec') {
                throw `Page ${pageName} is not a modern page.`;
            }
            let pageItemData = {};
            pageItemData = Object.assign({}, page);
            pageItemData.commentsDisabled = page.ListItemAllFields.CommentsDisabled;
            pageItemData.title = page.ListItemAllFields.Title;
            if (page.ListItemAllFields.PageLayoutType) {
                pageItemData.layoutType = page.ListItemAllFields.PageLayoutType;
            }
            if (!args.options.metadataOnly) {
                requestOptions = {
                    url: `${args.options.webUrl}/_api/SitePages/Pages(${page.ListItemAllFields.Id})`,
                    headers: {
                        'content-type': 'application/json;charset=utf-8',
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                const res = await request.get(requestOptions);
                const canvasData = JSON.parse(res.CanvasContent1);
                pageItemData.canvasContentJson = res.CanvasContent1;
                if (canvasData && canvasData.length > 0) {
                    pageItemData.numControls = canvasData.length;
                    const sections = [...new Set(canvasData.filter(c => c.position).map(c => c.position.zoneIndex))];
                    pageItemData.numSections = sections.length;
                }
            }
            delete pageItemData.ListItemAllFields.ID;
            await logger.log(pageItemData);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoPageGetCommand();
//# sourceMappingURL=page-get.js.map