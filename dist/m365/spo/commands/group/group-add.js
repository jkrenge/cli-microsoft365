var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoGroupAddCommand_instances, _SpoGroupAddCommand_initTelemetry, _SpoGroupAddCommand_initOptions, _SpoGroupAddCommand_initTypes, _SpoGroupAddCommand_initValidators, _SpoGroupAddCommand_initOptionSets;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoGroupAddCommand extends SpoCommand {
    get name() {
        return commands.GROUP_ADD;
    }
    get description() {
        return 'Creates group in the specified site';
    }
    constructor() {
        super();
        _SpoGroupAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoGroupAddCommand_instances, "m", _SpoGroupAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoGroupAddCommand_instances, "m", _SpoGroupAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoGroupAddCommand_instances, "m", _SpoGroupAddCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoGroupAddCommand_instances, "m", _SpoGroupAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoGroupAddCommand_instances, "m", _SpoGroupAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const requestOptions = {
            url: `${args.options.webUrl}/_api/web/sitegroups`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json',
            data: {
                Title: args.options.name,
                Description: args.options.description,
                AllowMembersEditMembership: args.options.allowMembersEditMembership,
                OnlyAllowMembersViewMembership: args.options.onlyAllowMembersViewMembership,
                AllowRequestToJoinLeave: args.options.allowRequestToJoinLeave,
                AutoAcceptRequestToJoinLeave: args.options.autoAcceptRequestToJoinLeave,
                RequestToJoinLeaveEmailSetting: args.options.requestToJoinLeaveEmailSetting
            }
        };
        try {
            const response = await request.post(requestOptions);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoGroupAddCommand_instances = new WeakSet(), _SpoGroupAddCommand_initTelemetry = function _SpoGroupAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            description: typeof args.options.description !== 'undefined',
            allowMembersEditMembership: args.options.allowMembersEditMembership,
            onlyAllowMembersViewMembership: args.options.onlyAllowMembersViewMembership,
            allowRequestToJoinLeave: args.options.allowRequestToJoinLeave,
            autoAcceptRequestToJoinLeave: args.options.autoAcceptRequestToJoinLeave,
            requestToJoinLeaveEmailSetting: typeof args.options.requestToJoinLeaveEmailSetting !== 'undefined'
        });
    });
}, _SpoGroupAddCommand_initOptions = function _SpoGroupAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-n, --name <name>'
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
    });
}, _SpoGroupAddCommand_initTypes = function _SpoGroupAddCommand_initTypes() {
    this.types.boolean.push('allowMembersEditMembership', 'onlyAllowMembersViewMembership', 'allowRequestToJoinLeave', 'autoAcceptRequestToJoinLeave');
}, _SpoGroupAddCommand_initValidators = function _SpoGroupAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        return true;
    });
}, _SpoGroupAddCommand_initOptionSets = function _SpoGroupAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new SpoGroupAddCommand();
//# sourceMappingURL=group-add.js.map