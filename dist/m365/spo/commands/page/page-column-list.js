var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoPageColumnListCommand_instances, _SpoPageColumnListCommand_initOptions, _SpoPageColumnListCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { Page } from './Page.js';
class SpoPageColumnListCommand extends SpoCommand {
    get name() {
        return commands.PAGE_COLUMN_LIST;
    }
    get description() {
        return 'Lists columns in the specific section of a modern page';
    }
    constructor() {
        super();
        _SpoPageColumnListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoPageColumnListCommand_instances, "m", _SpoPageColumnListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoPageColumnListCommand_instances, "m", _SpoPageColumnListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const clientSidePage = await Page.getPage(args.options.pageName, args.options.webUrl, logger, this.debug, this.verbose);
            const sections = clientSidePage.sections
                .filter(section => section.order === args.options.section);
            if (sections.length) {
                const isJSONOutput = !cli.shouldTrimOutput(args.options.output);
                await logger.log(sections[0].columns.map(c => {
                    const column = Page.getColumnsInformation(c, isJSONOutput);
                    column.controls = c.controls.length;
                    return column;
                }));
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoPageColumnListCommand_instances = new WeakSet(), _SpoPageColumnListCommand_initOptions = function _SpoPageColumnListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-n, --pageName <pageName>'
    }, {
        option: '-s, --section <sectionId>'
    });
}, _SpoPageColumnListCommand_initValidators = function _SpoPageColumnListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (isNaN(args.options.section)) {
            return `${args.options.section} is not a number`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
};
export default new SpoPageColumnListCommand();
//# sourceMappingURL=page-column-list.js.map