var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoPageSectionGetCommand_instances, _SpoPageSectionGetCommand_initOptions, _SpoPageSectionGetCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { Page } from './Page.js';
class SpoPageSectionGetCommand extends SpoCommand {
    get name() {
        return commands.PAGE_SECTION_GET;
    }
    get description() {
        return 'Get information about the specified modern page section';
    }
    constructor() {
        super();
        _SpoPageSectionGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoPageSectionGetCommand_instances, "m", _SpoPageSectionGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoPageSectionGetCommand_instances, "m", _SpoPageSectionGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const clientSidePage = await Page.getPage(args.options.pageName, args.options.webUrl, logger, this.debug, this.verbose);
            const sections = clientSidePage.sections
                .filter(section => section.order === args.options.section);
            const isJSONOutput = !cli.shouldTrimOutput(args.options.output);
            if (sections.length) {
                await logger.log(Page.getSectionInformation(sections[0], isJSONOutput));
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoPageSectionGetCommand_instances = new WeakSet(), _SpoPageSectionGetCommand_initOptions = function _SpoPageSectionGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-n, --pageName <pageName>'
    }, {
        option: '-s, --section <section>'
    });
}, _SpoPageSectionGetCommand_initValidators = function _SpoPageSectionGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (isNaN(args.options.section)) {
            return `${args.options.section} is not a number`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
};
export default new SpoPageSectionGetCommand();
//# sourceMappingURL=page-section-get.js.map