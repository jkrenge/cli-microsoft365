var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraMultitenantRemoveCommand_instances, _EntraMultitenantRemoveCommand_initTelemetry, _EntraMultitenantRemoveCommand_initOptions;
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { odata } from '../../../../utils/odata.js';
import { cli } from '../../../../cli/cli.js';
class EntraMultitenantRemoveCommand extends GraphCommand {
    get name() {
        return commands.MULTITENANT_REMOVE;
    }
    get description() {
        return 'Removes a multitenant organization';
    }
    constructor() {
        super();
        _EntraMultitenantRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraMultitenantRemoveCommand_instances, "m", _EntraMultitenantRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraMultitenantRemoveCommand_instances, "m", _EntraMultitenantRemoveCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        const removeMultitenantOrg = async () => {
            try {
                const tenantId = await this.getCurrentTenantId();
                let tenantsId = await this.getAllTenantsIds();
                const tenantsCount = tenantsId.length;
                if (tenantsCount > 0) {
                    const tasks = tenantsId
                        .filter(x => x !== tenantId)
                        .map(t => this.removeTenant(logger, t));
                    await Promise.all(tasks);
                    do {
                        if (this.verbose) {
                            await logger.logToStderr(`Waiting 30 seconds...`);
                        }
                        await new Promise(resolve => setTimeout(resolve, 30000));
                        // from current behavior, removing tenant can take a few seconds
                        // current tenant must be removed once all previous ones were removed
                        if (this.verbose) {
                            await logger.logToStderr(`Checking all tenants were removed...`);
                        }
                        tenantsId = await this.getAllTenantsIds();
                        if (this.verbose) {
                            await logger.logToStderr(`Number of removed tenants: ${tenantsCount - tenantsId.length}`);
                        }
                    } while (tenantsId.length !== 1);
                    // current tenant must be removed as the last one
                    await this.removeTenant(logger, tenantId);
                    await logger.logToStderr('Your Multi-Tenant organization is being removed. This can take up to 2 hours.');
                }
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeMultitenantOrg();
        }
        else {
            const result = await cli.promptForConfirmation({ message: 'Are you sure you want to remove multitenant organization?' });
            if (result) {
                await removeMultitenantOrg();
            }
        }
    }
    async getAllTenantsIds() {
        const requestOptions = {
            url: `${this.resource}/v1.0/tenantRelationships/multiTenantOrganization/tenants?$select=tenantId`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const tenants = await odata.getAllItems(requestOptions);
        return tenants.map(x => x.tenantId);
    }
    async getCurrentTenantId() {
        const requestOptions = {
            url: `${this.resource}/v1.0/organization?$select=id`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const organizations = await odata.getAllItems(requestOptions);
        return organizations[0].id;
    }
    async removeTenant(logger, tenantId) {
        if (this.verbose) {
            await logger.logToStderr(`Removing tenant: ${tenantId}`);
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/tenantRelationships/multiTenantOrganization/tenants/${tenantId}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        await request.delete(requestOptions);
    }
}
_EntraMultitenantRemoveCommand_instances = new WeakSet(), _EntraMultitenantRemoveCommand_initTelemetry = function _EntraMultitenantRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force
        });
    });
}, _EntraMultitenantRemoveCommand_initOptions = function _EntraMultitenantRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-f, --force'
    });
};
export default new EntraMultitenantRemoveCommand();
//# sourceMappingURL=multitenant-remove.js.map