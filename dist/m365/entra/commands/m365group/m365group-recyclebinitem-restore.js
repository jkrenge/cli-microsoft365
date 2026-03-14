var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraM365GroupRecycleBinItemRestoreCommand_instances, _EntraM365GroupRecycleBinItemRestoreCommand_initTelemetry, _EntraM365GroupRecycleBinItemRestoreCommand_initOptions, _EntraM365GroupRecycleBinItemRestoreCommand_initValidators, _EntraM365GroupRecycleBinItemRestoreCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
class EntraM365GroupRecycleBinItemRestoreCommand extends GraphCommand {
    get name() {
        return commands.M365GROUP_RECYCLEBINITEM_RESTORE;
    }
    get description() {
        return 'Restores a deleted Microsoft 365 Group';
    }
    constructor() {
        super();
        _EntraM365GroupRecycleBinItemRestoreCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraM365GroupRecycleBinItemRestoreCommand_instances, "m", _EntraM365GroupRecycleBinItemRestoreCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupRecycleBinItemRestoreCommand_instances, "m", _EntraM365GroupRecycleBinItemRestoreCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupRecycleBinItemRestoreCommand_instances, "m", _EntraM365GroupRecycleBinItemRestoreCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupRecycleBinItemRestoreCommand_instances, "m", _EntraM365GroupRecycleBinItemRestoreCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Restoring Microsoft 365 Group: ${args.options.id || args.options.displayName || args.options.mailNickname}...`);
        }
        try {
            const groupId = await this.getGroupId(args.options);
            const requestOptions = {
                url: `${this.resource}/v1.0/directory/deleteditems/${groupId}/restore`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json'
                },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getGroupId(options) {
        const { id, displayName, mailNickname } = options;
        if (id) {
            return id;
        }
        let filterValue = '';
        if (displayName) {
            filterValue = `displayName eq '${formatting.encodeQueryParameter(displayName)}'`;
        }
        if (mailNickname) {
            filterValue = `mailNickname eq '${formatting.encodeQueryParameter(mailNickname)}'`;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=${filterValue}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        const groups = response.value;
        if (groups.length === 0) {
            throw `The specified group '${displayName || mailNickname}' does not exist.`;
        }
        if (groups.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', groups);
            const result = await cli.handleMultipleResultsFound(`Multiple groups with name '${displayName || mailNickname}' found.`, resultAsKeyValuePair);
            return result.id;
        }
        return groups[0].id;
    }
}
_EntraM365GroupRecycleBinItemRestoreCommand_instances = new WeakSet(), _EntraM365GroupRecycleBinItemRestoreCommand_initTelemetry = function _EntraM365GroupRecycleBinItemRestoreCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            displayName: typeof args.options.displayName !== 'undefined',
            mailNickname: typeof args.options.mailNickname !== 'undefined'
        });
    });
}, _EntraM365GroupRecycleBinItemRestoreCommand_initOptions = function _EntraM365GroupRecycleBinItemRestoreCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-d, --displayName [displayName]'
    }, {
        option: '-m, --mailNickname [mailNickname]'
    });
}, _EntraM365GroupRecycleBinItemRestoreCommand_initValidators = function _EntraM365GroupRecycleBinItemRestoreCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
}, _EntraM365GroupRecycleBinItemRestoreCommand_initOptionSets = function _EntraM365GroupRecycleBinItemRestoreCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'displayName', 'mailNickname'] });
};
export default new EntraM365GroupRecycleBinItemRestoreCommand();
//# sourceMappingURL=m365group-recyclebinitem-restore.js.map