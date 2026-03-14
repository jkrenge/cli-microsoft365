var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoUserGetCommand_instances, _SpoUserGetCommand_initTelemetry, _SpoUserGetCommand_initOptions, _SpoUserGetCommand_initTypes, _SpoUserGetCommand_initValidators, _SpoUserGetCommand_initOptionSets;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoUserGetCommand extends SpoCommand {
    get name() {
        return commands.USER_GET;
    }
    get description() {
        return 'Gets a site user within specific web';
    }
    constructor() {
        super();
        _SpoUserGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoUserGetCommand_instances, "m", _SpoUserGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoUserGetCommand_instances, "m", _SpoUserGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoUserGetCommand_instances, "m", _SpoUserGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoUserGetCommand_instances, "m", _SpoUserGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoUserGetCommand_instances, "m", _SpoUserGetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving information for user in site '${args.options.webUrl}'...`);
        }
        let requestUrl = `${args.options.webUrl}/_api/web/`;
        if (args.options.id) {
            requestUrl += `siteusers/GetById('${formatting.encodeQueryParameter(args.options.id.toString())}')`;
        }
        else if (args.options.email) {
            requestUrl += `siteusers/GetByEmail('${formatting.encodeQueryParameter(args.options.email)}')`;
        }
        else if (args.options.loginName) {
            requestUrl += `siteusers/GetByLoginName('${formatting.encodeQueryParameter(args.options.loginName)}')`;
        }
        else if (args.options.userName) {
            const user = await this.getUser(requestUrl, args.options.userName);
            requestUrl += `siteusers/GetById('${formatting.encodeQueryParameter(user.Id.toString())}')`;
        }
        else if (args.options.entraGroupId || args.options.entraGroupName) {
            const entraGroup = await this.getEntraGroup(args.options.entraGroupId, args.options.entraGroupName);
            // For entra groups, M365 groups have an associated email and security groups don't
            if (entraGroup?.mail) {
                requestUrl += `siteusers/GetByEmail('${formatting.encodeQueryParameter(entraGroup.mail)}')`;
            }
            else {
                requestUrl += `siteusers/GetByLoginName('c:0t.c|tenant|${entraGroup?.id}')`;
            }
        }
        else {
            requestUrl += `currentuser`;
        }
        const requestOptions = {
            url: requestUrl,
            method: 'GET',
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const userInstance = await request.get(requestOptions);
            await logger.log(userInstance);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getUser(baseUrl, userName) {
        const requestUrl = `${baseUrl}siteusers?$filter=UserPrincipalName eq ('${formatting.encodeQueryParameter(userName)}')`;
        const requestOptions = {
            url: requestUrl,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const userInstance = await request.get(requestOptions);
        const userInstanceValue = userInstance.value[0];
        if (!userInstanceValue) {
            throw `User not found: ${userName}`;
        }
        return userInstanceValue;
    }
    async getEntraGroup(entraGroupId, entraGroupName) {
        if (entraGroupId) {
            return entraGroup.getGroupById(entraGroupId);
        }
        return entraGroup.getGroupByDisplayName(entraGroupName);
    }
}
_SpoUserGetCommand_instances = new WeakSet(), _SpoUserGetCommand_initTelemetry = function _SpoUserGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            email: typeof args.options.email !== 'undefined',
            loginName: typeof args.options.loginName !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            entraGroupId: typeof args.options.entraGroupId !== 'undefined',
            entraGroupName: typeof args.options.entraGroupName !== 'undefined'
        });
    });
}, _SpoUserGetCommand_initOptions = function _SpoUserGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '--email [email]'
    }, {
        option: '--loginName [loginName]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '--entraGroupId [entraGroupId]'
    }, {
        option: '--entraGroupName [entraGroupName]'
    });
}, _SpoUserGetCommand_initTypes = function _SpoUserGetCommand_initTypes() {
    this.types.string.push('webUrl', 'id', 'email', 'loginName', 'userName', 'entraGroupId', 'entraGroupName');
}, _SpoUserGetCommand_initValidators = function _SpoUserGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id &&
            typeof args.options.id !== 'number') {
            return `Specified id ${args.options.id} is not a number`;
        }
        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
            return `${args.options.entraGroupId} is not a valid GUID.`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid userName.`;
        }
        if (args.options.email && !validation.isValidUserPrincipalName(args.options.email)) {
            return `${args.options.email} is not a valid email.`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoUserGetCommand_initOptionSets = function _SpoUserGetCommand_initOptionSets() {
    this.optionSets.push({
        options: ['id', 'email', 'loginName', 'userName', 'entraGroupId', 'entraGroupName'],
        runsWhen: (args) => args.options.id || args.options.email || args.options.loginName || args.options.userName || args.options.entraGroupId || args.options.entraGroupName
    });
};
export default new SpoUserGetCommand();
//# sourceMappingURL=user-get.js.map