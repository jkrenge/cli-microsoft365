var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _ExternalConnectionRemoveCommand_instances, _ExternalConnectionRemoveCommand_initTelemetry, _ExternalConnectionRemoveCommand_initOptions, _ExternalConnectionRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class ExternalConnectionRemoveCommand extends GraphCommand {
    get name() {
        return commands.CONNECTION_REMOVE;
    }
    get description() {
        return 'Removes a specific external connection';
    }
    alias() {
        return [commands.EXTERNALCONNECTION_REMOVE];
    }
    constructor() {
        super();
        _ExternalConnectionRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _ExternalConnectionRemoveCommand_instances, "m", _ExternalConnectionRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _ExternalConnectionRemoveCommand_instances, "m", _ExternalConnectionRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _ExternalConnectionRemoveCommand_instances, "m", _ExternalConnectionRemoveCommand_initOptionSets).call(this);
    }
    async getExternalConnectionId(args) {
        if (args.options.id) {
            return args.options.id;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/external/connections?$filter=name eq '${formatting.encodeQueryParameter(args.options.name)}'&$select=id`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const res = await request.get(requestOptions);
        if (res.value.length === 1) {
            return res.value[0].id;
        }
        if (res.value.length === 0) {
            throw `The specified connection does not exist`;
        }
        const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', res.value);
        const result = await cli.handleMultipleResultsFound(`Multiple external connections with name ${args.options.name} found.`, resultAsKeyValuePair);
        return result.id;
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeExternalConnection(args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the external connection '${args.options.id || args.options.name}'?` });
            if (result) {
                await this.removeExternalConnection(args);
            }
        }
    }
    async removeExternalConnection(args) {
        try {
            const externalConnectionId = await this.getExternalConnectionId(args);
            const requestOptions = {
                url: `${this.resource}/v1.0/external/connections/${formatting.encodeQueryParameter(externalConnectionId)}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            await request.delete(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_ExternalConnectionRemoveCommand_instances = new WeakSet(), _ExternalConnectionRemoveCommand_initTelemetry = function _ExternalConnectionRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            force: (!(!args.options.force)).toString()
        });
    });
}, _ExternalConnectionRemoveCommand_initOptions = function _ExternalConnectionRemoveCommand_initOptions() {
    this.options.unshift({ option: '--id [id]' }, { option: '--name [name]' }, { option: '-f, --force' });
}, _ExternalConnectionRemoveCommand_initOptionSets = function _ExternalConnectionRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new ExternalConnectionRemoveCommand();
//# sourceMappingURL=connection-remove.js.map