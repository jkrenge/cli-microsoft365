import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { Page } from './Page.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    webUrl: z.string()
        .refine(url => validation.isValidSharePointUrl(url) === true, {
        error: e => `'${e.input}' is not a valid SharePoint URL.`
    })
        .alias('u'),
    pageName: z.string().alias('n'),
    id: z.uuid().alias('i'),
    draft: z.boolean().optional(),
    force: z.boolean().optional().alias('f')
});
class SpoPageControlRemoveCommand extends SpoCommand {
    get name() {
        return commands.PAGE_CONTROL_REMOVE;
    }
    get description() {
        return 'Removes a control from a modern page';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        if (!args.options.force) {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to delete control '${args.options.id}' on page '${args.options.pageName}'?` });
            if (!result) {
                return;
            }
        }
        try {
            if (this.verbose) {
                await logger.logToStderr(`Getting page properties for page '${args.options.pageName}'...`);
            }
            const pageName = urlUtil.removeLeadingSlashes(args.options.pageName.toLowerCase().endsWith('.aspx') ? args.options.pageName : `${args.options.pageName}.aspx`);
            let requestOptions = {
                url: `${args.options.webUrl}/_api/SitePages/Pages/GetByUrl('SitePages/${formatting.encodeQueryParameter(pageName)}')?$select=CanvasContent1`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const pageProps = await request.get(requestOptions);
            if (!pageProps.CanvasContent1) {
                throw `Page '${pageName}' doesn't contain canvas control '${args.options.id}'.`;
            }
            const pageControls = JSON.parse(pageProps.CanvasContent1);
            const hasControl = pageControls.some(control => control.id?.toLowerCase() === args.options.id.toLowerCase());
            if (!hasControl) {
                throw `Control with ID '${args.options.id}' was not found on page '${pageName}'.`;
            }
            if (this.verbose) {
                await logger.logToStderr('Checking out page...');
            }
            const page = await Page.checkout(pageName, args.options.webUrl, logger, this.verbose);
            const canvasContent = JSON.parse(page.CanvasContent1);
            if (this.verbose) {
                await logger.logToStderr(`Removing control with ID '${args.options.id}' from page...`);
            }
            const pageContent = canvasContent.filter(control => !control.id || control.id.toLowerCase() !== args.options.id.toLowerCase());
            requestOptions = {
                url: `${args.options.webUrl}/_api/SitePages/Pages/GetByUrl('SitePages/${formatting.encodeQueryParameter(pageName)}')/SavePageAsDraft`,
                headers: {
                    'content-type': 'application/json;odata=nometadata',
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json',
                data: {
                    CanvasContent1: JSON.stringify(pageContent)
                }
            };
            await request.patch(requestOptions);
            if (!args.options.draft) {
                if (this.verbose) {
                    await logger.logToStderr(`Republishing page...`);
                }
                await Page.publishPage(args.options.webUrl, pageName);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoPageControlRemoveCommand();
//# sourceMappingURL=page-control-remove.js.map