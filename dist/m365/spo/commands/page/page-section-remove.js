import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { Page } from './Page.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    webUrl: z.string()
        .refine(url => validation.isValidSharePointUrl(url) === true, {
        error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
        .alias('u'),
    pageName: z.string().alias('n'),
    section: z.number().alias('s'),
    force: z.boolean().optional().alias('f')
});
class SpoPageSectionRemoveCommand extends SpoCommand {
    get name() {
        return commands.PAGE_SECTION_REMOVE;
    }
    get description() {
        return 'Removes the specified section from the modern page';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeSection(logger, args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove section ${args.options.section} from '${args.options.pageName}'?` });
            if (result) {
                await this.removeSection(logger, args);
            }
        }
    }
    async removeSection(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Removing modern page section ${args.options.pageName} - ${args.options.section}...`);
            }
            const reqDigest = await spo.getRequestDigest(args.options.webUrl);
            const clientSidePage = await Page.getPage(args.options.pageName, args.options.webUrl, logger, this.debug, this.verbose);
            const sectionToDelete = clientSidePage.sections
                .findIndex(section => section.order === args.options.section);
            if (sectionToDelete === -1) {
                throw new Error(`Section ${args.options.section} not found`);
            }
            clientSidePage.sections.splice(sectionToDelete, 1);
            const updatedContent = clientSidePage.toHtml();
            const requestOptions = {
                url: `${args.options
                    .webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/sitepages/${args.options.pageName}')/ListItemAllFields`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue,
                    'content-type': 'application/json;odata=nometadata',
                    'X-HTTP-Method': 'MERGE',
                    'IF-MATCH': '*',
                    accept: 'application/json;odata=nometadata'
                },
                data: {
                    CanvasContent1: updatedContent
                },
                responseType: 'json'
            };
            return request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoPageSectionRemoveCommand();
//# sourceMappingURL=page-section-remove.js.map