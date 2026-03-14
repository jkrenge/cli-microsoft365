var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoUserEnsureCommand_instances, _SpoUserEnsureCommand_initTelemetry, _SpoUserEnsureCommand_initOptions, _SpoUserEnsureCommand_initValidators, _SpoUserEnsureCommand_initOptionSets, _SpoUserEnsureCommand_initTypes;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { entraUser } from '../../../../utils/entraUser.js';
class SpoUserEnsureCommand extends SpoCommand {
    get name() {
        return commands.USER_ENSURE;
    }
    get description() {
        return 'Ensures that a user is available on a specific site';
    }
    constructor() {
        super();
        _SpoUserEnsureCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoUserEnsureCommand_instances, "m", _SpoUserEnsureCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoUserEnsureCommand_instances, "m", _SpoUserEnsureCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoUserEnsureCommand_instances, "m", _SpoUserEnsureCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoUserEnsureCommand_instances, "m", _SpoUserEnsureCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoUserEnsureCommand_instances, "m", _SpoUserEnsureCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Ensuring user ${args.options.entraId || args.options.userName || args.options.loginName || args.options.entraGroupId || args.options.entraGroupName} at site ${args.options.webUrl}`);
        }
        try {
            const requestBody = {
                logonName: await this.getUpn(args.options)
            };
            const requestOptions = {
                url: `${args.options.webUrl}/_api/web/ensureuser`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                data: requestBody,
                responseType: 'json'
            };
            const response = await request.post(requestOptions);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getUpn(options) {
        if (options.userName) {
            return options.userName;
        }
        if (options.entraId) {
            return entraUser.getUpnByUserId(options.entraId);
        }
        if (options.loginName) {
            return options.loginName;
        }
        let upn = '';
        if (options.entraGroupId || options.entraGroupName) {
            const entraGroup = await this.getEntraGroup(options.entraGroupId, options.entraGroupName);
            upn = entraGroup.mailEnabled ? `c:0o.c|federateddirectoryclaimprovider|${entraGroup.id}` : `c:0t.c|tenant|${entraGroup.id}`;
        }
        return upn;
    }
    async getEntraGroup(entraGroupId, entraGroupName) {
        if (entraGroupId) {
            return entraGroup.getGroupById(entraGroupId);
        }
        return entraGroup.getGroupByDisplayName(entraGroupName);
    }
}
_SpoUserEnsureCommand_instances = new WeakSet(), _SpoUserEnsureCommand_initTelemetry = function _SpoUserEnsureCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            entraId: typeof args.options.entraId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            loginName: typeof args.options.loginName !== 'undefined',
            entraGroupId: typeof args.options.entraGroupId !== 'undefined',
            entraGroupName: typeof args.options.entraGroupName !== 'undefined'
        });
    });
}, _SpoUserEnsureCommand_initOptions = function _SpoUserEnsureCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--entraId [entraId]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '--loginName [loginName]'
    }, {
        option: '--entraGroupId [entraGroupId]'
    }, {
        option: '--entraGroupName [entraGroupName]'
    });
}, _SpoUserEnsureCommand_initValidators = function _SpoUserEnsureCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.entraId && !validation.isValidGuid(args.options.entraId)) {
            return `${args.options.entraId} is not a valid GUID.`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid userName.`;
        }
        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
            return `${args.options.entraGroupId} is not a valid GUID for option 'entraGroupId'.`;
        }
        return true;
    });
}, _SpoUserEnsureCommand_initOptionSets = function _SpoUserEnsureCommand_initOptionSets() {
    this.optionSets.push({ options: ['entraId', 'userName', 'loginName', 'entraGroupId', 'entraGroupName'] });
}, _SpoUserEnsureCommand_initTypes = function _SpoUserEnsureCommand_initTypes() {
    this.types.string.push('webUrl', 'entraId', 'userName', 'loginName', 'entraGroupId', 'entraGroupName');
};
export default new SpoUserEnsureCommand();
//# sourceMappingURL=user-ensure.js.map