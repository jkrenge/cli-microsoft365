var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoPageListCommand_instances, _SpoPageListCommand_initOptions, _SpoPageListCommand_initValidators;
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoPageListCommand extends SpoCommand {
    get name() {
        return commands.PAGE_LIST;
    }
    get description() {
        return 'Lists all modern pages in the given site';
    }
    defaultProperties() {
        return ['Name', 'Title'];
    }
    constructor() {
        super();
        _SpoPageListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoPageListCommand_instances, "m", _SpoPageListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoPageListCommand_instances, "m", _SpoPageListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving client-side pages...`);
            }
            let pages = [];
            const pagesList = await odata.getAllItems(`${args.options.webUrl}/_api/sitepages/pages?$orderby=Title`);
            if (pagesList && pagesList.length > 0) {
                pages = pagesList;
            }
            const files = await odata.getAllItems(`${args.options.webUrl}/_api/web/lists/SitePages/rootfolder/files?$expand=ListItemAllFields/ClientSideApplicationId&$orderby=Name`);
            if (files?.length > 0) {
                const clientSidePages = files.filter(f => f.ListItemAllFields.ClientSideApplicationId === 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec');
                pages = pages.map(p => {
                    const clientSidePage = clientSidePages.find(cp => cp && cp.ListItemAllFields && cp.ListItemAllFields.Id === p.Id);
                    if (clientSidePage) {
                        return {
                            ...clientSidePage,
                            ...p
                        };
                    }
                    return p;
                });
                pages.filter(p => p.ListItemAllFields).forEach(page => delete page.ListItemAllFields.ID);
            }
            await logger.log(pages);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoPageListCommand_instances = new WeakSet(), _SpoPageListCommand_initOptions = function _SpoPageListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    });
}, _SpoPageListCommand_initValidators = function _SpoPageListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new SpoPageListCommand();
//# sourceMappingURL=page-list.js.map