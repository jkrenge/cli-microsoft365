var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteRecycleBinItemMoveCommand_instances, _SpoSiteRecycleBinItemMoveCommand_initTelemetry, _SpoSiteRecycleBinItemMoveCommand_initOptions, _SpoSiteRecycleBinItemMoveCommand_initOptionSets, _SpoSiteRecycleBinItemMoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteRecycleBinItemMoveCommand extends SpoCommand {
    get name() {
        return commands.SITE_RECYCLEBINITEM_MOVE;
    }
    get description() {
        return 'Moves items from the first-stage recycle bin to the second-stage recycle bin';
    }
    constructor() {
        super();
        _SpoSiteRecycleBinItemMoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteRecycleBinItemMoveCommand_instances, "m", _SpoSiteRecycleBinItemMoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteRecycleBinItemMoveCommand_instances, "m", _SpoSiteRecycleBinItemMoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteRecycleBinItemMoveCommand_instances, "m", _SpoSiteRecycleBinItemMoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoSiteRecycleBinItemMoveCommand_instances, "m", _SpoSiteRecycleBinItemMoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.moveRecycleBinItem(args, logger);
        }
        else {
            const result = await cli.promptForConfirmation({ message: 'Are you sure you want to move these items to the second-stage recycle bin?' });
            if (result) {
                await this.moveRecycleBinItem(args, logger);
            }
        }
    }
    async moveRecycleBinItem(args, logger) {
        try {
            const requestOptions = {
                url: `${args.options.siteUrl}/_api/web/recycleBin/MoveAllToSecondStage`,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            if (args.options.ids !== undefined) {
                const ids = args.options.ids.split(',');
                if (this.verbose) {
                    await logger.logToStderr(`Moving ${ids.length} items to the second-stage recycle bin.`);
                }
                requestOptions.data = {
                    ids: ids
                };
            }
            else if (this.verbose) {
                await logger.logToStderr('Moving all items to the second-stage recycle bin.');
            }
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteRecycleBinItemMoveCommand_instances = new WeakSet(), _SpoSiteRecycleBinItemMoveCommand_initTelemetry = function _SpoSiteRecycleBinItemMoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            ids: typeof args.options.ids !== 'undefined',
            all: !!args.options.all,
            force: !!args.options.force
        });
    });
}, _SpoSiteRecycleBinItemMoveCommand_initOptions = function _SpoSiteRecycleBinItemMoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '-i, --ids [ids]'
    }, {
        option: '--all'
    }, {
        option: '-f, --force'
    });
}, _SpoSiteRecycleBinItemMoveCommand_initOptionSets = function _SpoSiteRecycleBinItemMoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['ids', 'all'] });
}, _SpoSiteRecycleBinItemMoveCommand_initValidators = function _SpoSiteRecycleBinItemMoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.siteUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.ids) {
            const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.ids);
            if (isValidGUIDArrayResult !== true) {
                return `The following GUIDs are invalid for the option 'ids': ${isValidGUIDArrayResult}.`;
            }
        }
        return true;
    });
};
export default new SpoSiteRecycleBinItemMoveCommand();
//# sourceMappingURL=site-recyclebinitem-move.js.map