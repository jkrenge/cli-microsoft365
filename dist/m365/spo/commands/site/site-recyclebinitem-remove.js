var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteRecycleBinItemRemoveCommand_instances, _SpoSiteRecycleBinItemRemoveCommand_initTelemetry, _SpoSiteRecycleBinItemRemoveCommand_initOptions, _SpoSiteRecycleBinItemRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteRecycleBinItemRemoveCommand extends SpoCommand {
    get name() {
        return commands.SITE_RECYCLEBINITEM_REMOVE;
    }
    get description() {
        return 'Permanently deletes specific items from the site recycle bin';
    }
    constructor() {
        super();
        _SpoSiteRecycleBinItemRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteRecycleBinItemRemoveCommand_instances, "m", _SpoSiteRecycleBinItemRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteRecycleBinItemRemoveCommand_instances, "m", _SpoSiteRecycleBinItemRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteRecycleBinItemRemoveCommand_instances, "m", _SpoSiteRecycleBinItemRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeRecycleBinItem(args, logger);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to permanently delete ${args.options.ids.split(',').length} item(s) from the site recycle bin?` });
            if (result) {
                await this.removeRecycleBinItem(args, logger);
            }
        }
    }
    async removeRecycleBinItem(args, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Permanently deleting items from the site recycle bin at site ${args.options.siteUrl}...`);
        }
        try {
            const requestOptions = {
                url: `${args.options.siteUrl}/_api/site/RecycleBin/DeleteByIds`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json',
                data: {
                    ids: args.options.ids.split(',')
                }
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteRecycleBinItemRemoveCommand_instances = new WeakSet(), _SpoSiteRecycleBinItemRemoveCommand_initTelemetry = function _SpoSiteRecycleBinItemRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force
        });
    });
}, _SpoSiteRecycleBinItemRemoveCommand_initOptions = function _SpoSiteRecycleBinItemRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '-i, --ids <ids>'
    }, {
        option: '-f, --force'
    });
}, _SpoSiteRecycleBinItemRemoveCommand_initValidators = function _SpoSiteRecycleBinItemRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.siteUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.ids);
        if (isValidGUIDArrayResult !== true) {
            return `The following GUIDs are invalid for the option 'ids': ${isValidGUIDArrayResult}.`;
        }
        return true;
    });
};
export default new SpoSiteRecycleBinItemRemoveCommand();
//# sourceMappingURL=site-recyclebinitem-remove.js.map