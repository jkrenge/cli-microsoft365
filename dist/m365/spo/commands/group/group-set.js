var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoGroupSetCommand_instances, _SpoGroupSetCommand_initTelemetry, _SpoGroupSetCommand_initOptions, _SpoGroupSetCommand_initTypes, _SpoGroupSetCommand_initValidators, _SpoGroupSetCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import entraUserGetCommand from '../../../entra/commands/user/user-get.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoGroupSetCommand extends SpoCommand {
    get name() {
        return commands.GROUP_SET;
    }
    get description() {
        return 'Updates a group in the specified site';
    }
    constructor() {
        super();
        _SpoGroupSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoGroupSetCommand_instances, "m", _SpoGroupSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoGroupSetCommand_instances, "m", _SpoGroupSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoGroupSetCommand_instances, "m", _SpoGroupSetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoGroupSetCommand_instances, "m", _SpoGroupSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoGroupSetCommand_instances, "m", _SpoGroupSetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Setting properties for group ${args.options.id || args.options.name}`);
        }
        const requestOptions = {
            url: `${args.options.webUrl}/_api/web/sitegroups/${args.options.id ? `GetById(${args.options.id})` : `GetByName('${args.options.name}')`}`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json'
            },
            responseType: 'json',
            data: {
                Title: args.options.newName,
                Description: args.options.description,
                AllowMembersEditMembership: args.options.allowMembersEditMembership,
                OnlyAllowMembersViewMembership: args.options.onlyAllowMembersViewMembership,
                AllowRequestToJoinLeave: args.options.allowRequestToJoinLeave,
                AutoAcceptRequestToJoinLeave: args.options.autoAcceptRequestToJoinLeave,
                RequestToJoinLeaveEmailSetting: args.options.requestToJoinLeaveEmailSetting
            }
        };
        try {
            await request.patch(requestOptions);
            if (args.options.ownerEmail || args.options.ownerUserName) {
                await this.setGroupOwner(args.options);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async setGroupOwner(options) {
        const ownerId = await this.getOwnerId(options);
        const requestOptions = {
            url: `${options.webUrl}/_api/web/sitegroups/${options.id ? `GetById(${options.id})` : `GetByName('${options.name}')`}/SetUserAsOwner(${ownerId})`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json'
            },
            responseType: 'json'
        };
        return request.post(requestOptions);
    }
    async getOwnerId(options) {
        const cmdOptions = {
            userName: options.ownerUserName,
            email: options.ownerEmail,
            output: 'json',
            debug: options.debug,
            verbose: options.verbose
        };
        const output = await cli.executeCommandWithOutput(entraUserGetCommand, { options: { ...cmdOptions, _: [] } });
        const getUserOutput = JSON.parse(output.stdout);
        const requestOptions = {
            url: `${options.webUrl}/_api/web/ensureUser('${getUserOutput.userPrincipalName}')?$select=Id`,
            headers: {
                accept: 'application/json',
                'content-type': 'application/json'
            },
            responseType: 'json'
        };
        const response = await request.post(requestOptions);
        return response.Id;
    }
}
_SpoGroupSetCommand_instances = new WeakSet(), _SpoGroupSetCommand_initTelemetry = function _SpoGroupSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            newName: typeof args.options.newName !== 'undefined',
            description: typeof args.options.description !== 'undefined',
            allowMembersEditMembership: args.options.allowMembersEditMembership,
            onlyAllowMembersViewMembership: args.options.onlyAllowMembersViewMembership,
            allowRequestToJoinLeave: args.options.allowRequestToJoinLeave,
            autoAcceptRequestToJoinLeave: args.options.autoAcceptRequestToJoinLeave,
            requestToJoinLeaveEmailSetting: typeof args.options.requestToJoinLeaveEmailSetting !== 'undefined',
            ownerEmail: typeof args.options.ownerEmail !== 'undefined',
            ownerUserName: typeof args.options.ownerUserName !== 'undefined'
        });
    });
}, _SpoGroupSetCommand_initOptions = function _SpoGroupSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '--newName [newName]'
    }, {
        option: '--description [description]'
    }, {
        option: '--allowMembersEditMembership [allowMembersEditMembership]',
        autocomplete: ['true', 'false']
    }, {
        option: '--onlyAllowMembersViewMembership [onlyAllowMembersViewMembership]',
        autocomplete: ['true', 'false']
    }, {
        option: '--allowRequestToJoinLeave [allowRequestToJoinLeave]',
        autocomplete: ['true', 'false']
    }, {
        option: '--autoAcceptRequestToJoinLeave [autoAcceptRequestToJoinLeave]',
        autocomplete: ['true', 'false']
    }, {
        option: '--requestToJoinLeaveEmailSetting [requestToJoinLeaveEmailSetting]'
    }, {
        option: '--ownerEmail [ownerEmail]'
    }, {
        option: '--ownerUserName [ownerUserName]'
    });
}, _SpoGroupSetCommand_initTypes = function _SpoGroupSetCommand_initTypes() {
    this.types.boolean.push('allowMembersEditMembership', 'onlyAllowMembersViewMembership', 'allowRequestToJoinLeave', 'autoAcceptRequestToJoinLeave');
}, _SpoGroupSetCommand_initValidators = function _SpoGroupSetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.id && isNaN(args.options.id)) {
            return `Specified id ${args.options.id} is not a number`;
        }
        if (args.options.ownerEmail && args.options.ownerUserName) {
            return 'Specify either ownerEmail or ownerUserName but not both';
        }
        return true;
    });
}, _SpoGroupSetCommand_initOptionSets = function _SpoGroupSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new SpoGroupSetCommand();
//# sourceMappingURL=group-set.js.map