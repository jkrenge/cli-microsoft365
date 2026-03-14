var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteRecycleBinItemClearCommand_instances, _SpoSiteRecycleBinItemClearCommand_initTelemetry, _SpoSiteRecycleBinItemClearCommand_initOptions, _SpoSiteRecycleBinItemClearCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteRecycleBinItemClearCommand extends SpoCommand {
    get name() {
        return commands.SITE_RECYCLEBINITEM_CLEAR;
    }
    get description() {
        return 'Permanently removes all items in a site recycle bin';
    }
    constructor() {
        super();
        _SpoSiteRecycleBinItemClearCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteRecycleBinItemClearCommand_instances, "m", _SpoSiteRecycleBinItemClearCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteRecycleBinItemClearCommand_instances, "m", _SpoSiteRecycleBinItemClearCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteRecycleBinItemClearCommand_instances, "m", _SpoSiteRecycleBinItemClearCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.clearRecycleBin(args, logger);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to clear the recycle bin of site ${args.options.siteUrl}?` });
            if (result) {
                await this.clearRecycleBin(args, logger);
            }
        }
    }
    async clearRecycleBin(args, logger) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Permanently removing all items in recycle bin of site ${args.options.siteUrl}...`);
            }
            const requestOptions = {
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            if (args.options.secondary) {
                if (this.verbose) {
                    await logger.logToStderr('Removing all items from the second-stage recycle bin');
                }
                requestOptions.url = `${args.options.siteUrl}/_api/site/RecycleBin/DeleteAllSecondStageItems`;
            }
            else {
                if (this.verbose) {
                    await logger.logToStderr('Removing all items from the first-stage recycle bin');
                }
                requestOptions.url = `${args.options.siteUrl}/_api/web/RecycleBin/DeleteAll`;
            }
            const result = await request.post(requestOptions);
            if (result['odata.null'] !== true) {
                throw result;
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteRecycleBinItemClearCommand_instances = new WeakSet(), _SpoSiteRecycleBinItemClearCommand_initTelemetry = function _SpoSiteRecycleBinItemClearCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            secondary: !!args.options.secondary,
            force: !!args.options.force
        });
    });
}, _SpoSiteRecycleBinItemClearCommand_initOptions = function _SpoSiteRecycleBinItemClearCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '--secondary'
    }, {
        option: '-f, --force'
    });
}, _SpoSiteRecycleBinItemClearCommand_initValidators = function _SpoSiteRecycleBinItemClearCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.siteUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        return true;
    });
};
export default new SpoSiteRecycleBinItemClearCommand();
//# sourceMappingURL=site-recyclebinitem-clear.js.map