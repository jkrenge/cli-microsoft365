var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _ExternalConnectionAddCommand_instances, _ExternalConnectionAddCommand_initTelemetry, _ExternalConnectionAddCommand_initOptions, _ExternalConnectionAddCommand_initValidators;
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class ExternalConnectionAddCommand extends GraphCommand {
    get name() {
        return commands.CONNECTION_ADD;
    }
    get description() {
        return 'Adds a new external connection';
    }
    alias() {
        return [commands.EXTERNALCONNECTION_ADD];
    }
    constructor() {
        super();
        _ExternalConnectionAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _ExternalConnectionAddCommand_instances, "m", _ExternalConnectionAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _ExternalConnectionAddCommand_instances, "m", _ExternalConnectionAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _ExternalConnectionAddCommand_instances, "m", _ExternalConnectionAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let appIds = [];
        if (args.options.authorizedAppIds !== undefined &&
            args.options.authorizedAppIds !== '') {
            appIds = args.options.authorizedAppIds.split(',');
        }
        const commandData = {
            id: args.options.id,
            name: args.options.name,
            description: args.options.description,
            configuration: {
                authorizedAppIds: appIds
            }
        };
        const requestOptions = {
            url: `${this.resource}/v1.0/external/connections`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: commandData
        };
        try {
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_ExternalConnectionAddCommand_instances = new WeakSet(), _ExternalConnectionAddCommand_initTelemetry = function _ExternalConnectionAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            authorizedAppIds: typeof args.options.authorizedAppIds !== 'undefined'
        });
    });
}, _ExternalConnectionAddCommand_initOptions = function _ExternalConnectionAddCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-n, --name <name>'
    }, {
        option: '-d, --description <description>'
    }, {
        option: '--authorizedAppIds [authorizedAppIds]'
    });
}, _ExternalConnectionAddCommand_initValidators = function _ExternalConnectionAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const id = args.options.id;
        if (id.length < 3 || id.length > 32) {
            return 'ID must be between 3 and 32 characters in length.';
        }
        const alphaNumericRegEx = /[^\w]|_/g;
        if (alphaNumericRegEx.test(id)) {
            return 'ID must only contain alphanumeric characters.';
        }
        if (id.length > 9 &&
            id.startsWith('Microsoft')) {
            return 'ID cannot begin with Microsoft';
        }
        const invalidIds = ['None',
            'Directory',
            'Exchange',
            'ExchangeArchive',
            'LinkedIn',
            'Mailbox',
            'OneDriveBusiness',
            'SharePoint',
            'Teams',
            'Yammer',
            'Connectors',
            'TaskFabric',
            'PowerBI',
            'Assistant',
            'TopicEngine',
            'MSFT_All_Connectors'
        ];
        if (invalidIds.indexOf(id) > -1) {
            return `ID cannot be one of the following values: ${invalidIds.join(', ')}.`;
        }
        return true;
    });
};
export default new ExternalConnectionAddCommand();
//# sourceMappingURL=connection-add.js.map