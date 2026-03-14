var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoPageColumnGetCommand_instances, _SpoPageColumnGetCommand_initOptions, _SpoPageColumnGetCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { Page } from './Page.js';
class SpoPageColumnGetCommand extends SpoCommand {
    get name() {
        return commands.PAGE_COLUMN_GET;
    }
    get description() {
        return 'Get information about a specific column of a modern page';
    }
    constructor() {
        super();
        _SpoPageColumnGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoPageColumnGetCommand_instances, "m", _SpoPageColumnGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoPageColumnGetCommand_instances, "m", _SpoPageColumnGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const clientSidePage = await Page.getPage(args.options.pageName, args.options.webUrl, logger, this.debug, this.verbose);
            const sections = clientSidePage.sections
                .filter(section => section.order === args.options.section);
            if (sections.length) {
                const isJSONOutput = !cli.shouldTrimOutput(args.options.output);
                const columns = sections[0].columns.filter(col => col.order === args.options.column);
                if (columns.length) {
                    const column = Page.getColumnsInformation(columns[0], isJSONOutput);
                    column.controls = columns[0].controls
                        .map(control => Page.getControlsInformation(control, isJSONOutput));
                    if (!isJSONOutput) {
                        column.controls = column.controls
                            .map(control => `${control.id} (${control.title})`)
                            .join(', ');
                    }
                    await logger.log(column);
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoPageColumnGetCommand_instances = new WeakSet(), _SpoPageColumnGetCommand_initOptions = function _SpoPageColumnGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-n, --pageName <pageName>'
    }, {
        option: '-s, --section <section>'
    }, {
        option: '-c, --column <column>'
    });
}, _SpoPageColumnGetCommand_initValidators = function _SpoPageColumnGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (isNaN(args.options.section)) {
            return `${args.options.section} is not a number`;
        }
        if (isNaN(args.options.column)) {
            return `${args.options.column} is not a number`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
};
export default new SpoPageColumnGetCommand();
//# sourceMappingURL=page-column-get.js.map