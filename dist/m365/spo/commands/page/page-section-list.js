var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoPageSectionListCommand_instances, _SpoPageSectionListCommand_initOptions, _SpoPageSectionListCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { Page } from './Page.js';
class SpoPageSectionListCommand extends SpoCommand {
    get name() {
        return commands.PAGE_SECTION_LIST;
    }
    get description() {
        return 'List sections in the specific modern page';
    }
    constructor() {
        super();
        _SpoPageSectionListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoPageSectionListCommand_instances, "m", _SpoPageSectionListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoPageSectionListCommand_instances, "m", _SpoPageSectionListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const clientSidePage = await Page.getPage(args.options.pageName, args.options.webUrl, logger, this.debug, this.verbose);
            const sections = clientSidePage.sections;
            const isJSONOutput = !cli.shouldTrimOutput(args.options.output);
            if (sections.length) {
                const output = sections.map(section => Page.getSectionInformation(section, isJSONOutput));
                if (isJSONOutput) {
                    await logger.log(output);
                }
                else {
                    await logger.log(output.map(s => {
                        const sectionOutput = {
                            order: s.order,
                            columns: s.columns.length
                        };
                        if (s.isVertical) {
                            sectionOutput.isVertical = s.isVertical;
                        }
                        return sectionOutput;
                    }));
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoPageSectionListCommand_instances = new WeakSet(), _SpoPageSectionListCommand_initOptions = function _SpoPageSectionListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-n, --pageName <pageName>'
    });
}, _SpoPageSectionListCommand_initValidators = function _SpoPageSectionListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new SpoPageSectionListCommand();
//# sourceMappingURL=page-section-list.js.map