var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraM365GroupListCommand_instances, _EntraM365GroupListCommand_initTelemetry, _EntraM365GroupListCommand_initOptions;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraM365GroupListCommand extends GraphCommand {
    get name() {
        return commands.M365GROUP_LIST;
    }
    get description() {
        return 'Lists Microsoft 365 Groups in the current tenant';
    }
    constructor() {
        super();
        _EntraM365GroupListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraM365GroupListCommand_instances, "m", _EntraM365GroupListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupListCommand_instances, "m", _EntraM365GroupListCommand_initOptions).call(this);
    }
    defaultProperties() {
        return ['id', 'displayName', 'mailNickname', 'siteUrl'];
    }
    async commandAction(logger, args) {
        const groupFilter = `?$filter=groupTypes/any(c:c+eq+'Unified')`;
        const displayNameFilter = args.options.displayName ? ` and startswith(DisplayName,'${formatting.encodeQueryParameter(args.options.displayName)}')` : '';
        const mailNicknameFilter = args.options.mailNickname ? ` and startswith(MailNickname,'${formatting.encodeQueryParameter(args.options.mailNickname)}')` : '';
        const expandOwners = args.options.orphaned ? '&$expand=owners' : '';
        const topCount = '&$top=100';
        try {
            let groups = [];
            groups = await odata.getAllItems(`${this.resource}/v1.0/groups${groupFilter}${displayNameFilter}${mailNicknameFilter}${expandOwners}${topCount}`);
            if (args.options.orphaned) {
                const orphanedGroups = [];
                groups.forEach((group) => {
                    if (!group.owners || group.owners.length === 0) {
                        orphanedGroups.push(group);
                    }
                });
                groups = orphanedGroups;
            }
            if (args.options.withSiteUrl) {
                const res = await Promise.all(groups.map(g => this.getGroupSiteUrl(g.id)));
                res.forEach(r => {
                    for (let i = 0; i < groups.length; i++) {
                        if (groups[i].id !== r.id) {
                            continue;
                        }
                        groups[i].siteUrl = r.url;
                        break;
                    }
                });
            }
            await logger.log(groups);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getGroupSiteUrl(groupId) {
        const requestOptions = {
            url: `${this.resource}/v1.0/groups/${groupId}/drive?$select=webUrl`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const res = await request.get(requestOptions);
        return {
            id: groupId,
            url: res.webUrl ? res.webUrl.substring(0, res.webUrl.lastIndexOf('/')) : ''
        };
    }
}
_EntraM365GroupListCommand_instances = new WeakSet(), _EntraM365GroupListCommand_initTelemetry = function _EntraM365GroupListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            displayName: typeof args.options.displayName !== 'undefined',
            mailNickname: typeof args.options.mailNickname !== 'undefined',
            withSiteUrl: !!args.options.withSiteUrl,
            orphaned: !!args.options.orphaned
        });
    });
}, _EntraM365GroupListCommand_initOptions = function _EntraM365GroupListCommand_initOptions() {
    this.options.unshift({
        option: '-d, --displayName [displayName]'
    }, {
        option: '-m, --mailNickname [displayName]'
    }, {
        option: '--withSiteUrl'
    }, {
        option: '--orphaned'
    });
};
export default new EntraM365GroupListCommand();
//# sourceMappingURL=m365group-list.js.map