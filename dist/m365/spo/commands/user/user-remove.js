var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoUserRemoveCommand_instances, _SpoUserRemoveCommand_initTelemetry, _SpoUserRemoveCommand_initOptions, _SpoUserRemoveCommand_initValidators, _SpoUserRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import { spo } from '../../../../utils/spo.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
;
class SpoUserRemoveCommand extends SpoCommand {
    get name() {
        return commands.USER_REMOVE;
    }
    get description() {
        return 'Removes user from specific web';
    }
    constructor() {
        super();
        _SpoUserRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoUserRemoveCommand_instances, "m", _SpoUserRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoUserRemoveCommand_instances, "m", _SpoUserRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoUserRemoveCommand_instances, "m", _SpoUserRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoUserRemoveCommand_instances, "m", _SpoUserRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeUser(logger, args.options);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove specified user from the site ${args.options.webUrl}?` });
            if (result) {
                await this.removeUser(logger, args.options);
            }
        }
    }
    async removeUser(logger, options) {
        if (this.verbose) {
            await logger.logToStderr(`Removing user from  subsite ${options.webUrl} ...`);
        }
        try {
            let requestUrl = `${encodeURI(options.webUrl)}/_api/web/siteusers/`;
            if (options.id) {
                requestUrl += `removebyid(${options.id})`;
            }
            else if (options.loginName) {
                requestUrl += `removeByLoginName('${formatting.encodeQueryParameter(options.loginName)}')`;
            }
            else if (options.email) {
                const user = await spo.getUserByEmail(options.webUrl, options.email, logger, this.verbose);
                requestUrl += `removebyid(${user.Id})`;
            }
            else if (options.userName) {
                const user = await this.getUser(options);
                if (!user) {
                    throw new Error(`User not found: ${options.userName}`);
                }
                if (this.verbose) {
                    await logger.logToStderr(`Removing user ${user.Title} ...`);
                }
                requestUrl += `removebyid(${user.Id})`;
            }
            else if (options.entraGroupId || options.entraGroupName) {
                const entraGroup = await this.getEntraGroup(options);
                if (this.verbose) {
                    await logger.logToStderr(`Removing entra group ${entraGroup?.displayName} ...`);
                }
                //for entra groups, M365 groups have an associated email and security groups don't
                if (entraGroup?.mail) {
                    //M365 group is prefixed with c:0o.c|federateddirectoryclaimprovider
                    requestUrl += `removeByLoginName('c:0o.c|federateddirectoryclaimprovider|${entraGroup.id}')`;
                }
                else {
                    //security group is prefixed with c:0t.c|tenant
                    requestUrl += `removeByLoginName('c:0t.c|tenant|${entraGroup?.id}')`;
                }
            }
            const requestOptions = {
                url: requestUrl,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getUser(options) {
        const requestUrl = `${options.webUrl}/_api/web/siteusers?$filter=UserPrincipalName eq ('${formatting.encodeQueryParameter(options.userName)}')`;
        const requestOptions = {
            url: requestUrl,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const userInstance = await request.get(requestOptions);
        return userInstance.value[0];
    }
    async getEntraGroup(options) {
        return options.entraGroupId ? await entraGroup.getGroupById(options.entraGroupId) : await entraGroup.getGroupByDisplayName(options.entraGroupName);
    }
}
_SpoUserRemoveCommand_instances = new WeakSet(), _SpoUserRemoveCommand_initTelemetry = function _SpoUserRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            loginName: typeof args.options.loginName !== 'undefined',
            email: typeof args.options.email !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            entraGroupId: typeof args.options.entraGroupId !== 'undefined',
            entraGroupName: typeof args.options.entraGroupName !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoUserRemoveCommand_initOptions = function _SpoUserRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '--loginName [loginName]'
    }, {
        option: '--email [email]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '--entraGroupId [entraGroupId]'
    }, {
        option: '--entraGroupName [entraGroupName]'
    }, {
        option: '-f, --force'
    });
}, _SpoUserRemoveCommand_initValidators = function _SpoUserRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.id && isNaN(parseInt(args.options.id))) {
            return `Specified id ${args.options.id} is not a number`;
        }
        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
            return `${args.options.entraId} is not a valid GUID.`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid userName.`;
        }
        if (args.options.email && !validation.isValidUserPrincipalName(args.options.email)) {
            return `${args.options.email} is not a valid email.`;
        }
        return true;
    });
}, _SpoUserRemoveCommand_initOptionSets = function _SpoUserRemoveCommand_initOptionSets() {
    this.optionSets.push({
        options: ['id', 'loginName', 'email', 'userName', 'entraGroupId', 'entraGroupName']
    });
};
export default new SpoUserRemoveCommand();
//# sourceMappingURL=user-remove.js.map