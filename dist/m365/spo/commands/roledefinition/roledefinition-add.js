var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoRoleDefinitionAddCommand_instances, _SpoRoleDefinitionAddCommand_initTelemetry, _SpoRoleDefinitionAddCommand_initOptions, _SpoRoleDefinitionAddCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { BasePermissions, PermissionKind } from '../../base-permissions.js';
import commands from '../../commands.js';
class SpoRoleDefinitionAddCommand extends SpoCommand {
    get name() {
        return commands.ROLEDEFINITION_ADD;
    }
    get description() {
        return 'Adds a new roledefinition to web';
    }
    get permissionsKindMap() {
        const result = [];
        for (const kind in PermissionKind) {
            if (typeof PermissionKind[kind] === 'number') {
                result.push(kind);
            }
        }
        return result;
    }
    constructor() {
        super();
        _SpoRoleDefinitionAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoRoleDefinitionAddCommand_instances, "m", _SpoRoleDefinitionAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoRoleDefinitionAddCommand_instances, "m", _SpoRoleDefinitionAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoRoleDefinitionAddCommand_instances, "m", _SpoRoleDefinitionAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Adding role definition to ${args.options.webUrl}...`);
        }
        const description = args.options.description || '';
        const permissions = new BasePermissions();
        if (args.options.rights) {
            const rights = args.options.rights.split(',');
            for (const item of rights) {
                const kind = PermissionKind[item.trim()];
                permissions.set(kind);
            }
        }
        const requestOptions = {
            url: `${args.options.webUrl}/_api/web/roledefinitions`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json',
            data: {
                'BasePermissions': {
                    'High': permissions.high.toString(),
                    'Low': permissions.low.toString()
                },
                'Description': `${description}`,
                'Name': `${args.options.name}`
            }
        };
        try {
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoRoleDefinitionAddCommand_instances = new WeakSet(), _SpoRoleDefinitionAddCommand_initTelemetry = function _SpoRoleDefinitionAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            rights: args.options.rights,
            description: (!(!args.options.description)).toString()
        });
    });
}, _SpoRoleDefinitionAddCommand_initOptions = function _SpoRoleDefinitionAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-n, --name <name>'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '--rights [rights]',
        autocomplete: this.permissionsKindMap
    });
}, _SpoRoleDefinitionAddCommand_initValidators = function _SpoRoleDefinitionAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.rights) {
            const rights = args.options.rights.split(',');
            for (const item of rights) {
                const kind = PermissionKind[item.trim()];
                if (!kind) {
                    return `Rights option '${item}' is not recognized as valid PermissionKind choice. Please note it is case-sensitive. Allowed values are ${this.permissionsKindMap.join('|')}.`;
                }
            }
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
};
export default new SpoRoleDefinitionAddCommand();
//# sourceMappingURL=roledefinition-add.js.map