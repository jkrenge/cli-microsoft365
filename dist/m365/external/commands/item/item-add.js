var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _ExternalItemAddCommand_instances, _a, _ExternalItemAddCommand_initTelemetry, _ExternalItemAddCommand_initOptions, _ExternalItemAddCommand_initValidators;
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class ExternalItemAddCommand extends GraphCommand {
    get name() {
        return commands.ITEM_ADD;
    }
    get description() {
        return 'Creates external item';
    }
    constructor() {
        super();
        _ExternalItemAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _ExternalItemAddCommand_instances, "m", _ExternalItemAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _ExternalItemAddCommand_instances, "m", _ExternalItemAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _ExternalItemAddCommand_instances, "m", _ExternalItemAddCommand_initValidators).call(this);
    }
    allowUnknownOptions() {
        return true;
    }
    async commandAction(logger, args) {
        const acls = args.options.acls
            .split(';')
            .map(acl => {
            const aclParts = acl.split(',');
            return {
                accessType: aclParts[0],
                type: aclParts[1],
                value: aclParts[2]
            };
        });
        const requestBody = {
            id: args.options.id,
            content: {
                value: args.options.content,
                type: args.options.contentType ?? 'text'
            },
            acl: acls,
            properties: {}
        };
        // we need to rewrite the @odata properties to the correct format
        // to extract multiple values for collections into arrays
        this.rewriteCollectionProperties(args.options);
        this.addUnknownOptionsToPayload(requestBody.properties, args.options);
        const requestOptions = {
            url: `${this.resource}/v1.0/external/connections/${args.options.externalConnectionId}/items/${args.options.id}`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json'
            },
            responseType: 'json',
            data: requestBody
        };
        try {
            const externalItem = await request.put(requestOptions);
            if (args.options.output === 'csv' || args.options.output === 'md') {
                // for CSV and md, we need to bring the properties to the main object
                // and convert arrays to comma-separated strings or they will be dropped
                // from the output
                Object.getOwnPropertyNames(externalItem.properties).forEach(name => {
                    if (Array.isArray(externalItem.properties[name])) {
                        externalItem[name] = externalItem.properties[name].join(', ');
                    }
                    else {
                        externalItem[name] = externalItem.properties[name];
                    }
                });
            }
            await logger.log(externalItem);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    rewriteCollectionProperties(options) {
        Object.getOwnPropertyNames(options).forEach(name => {
            if (!name.includes('@odata')) {
                return;
            }
            // convert the value of a collection to an array
            const nameWithoutOData = name.substring(0, name.indexOf('@odata'));
            if (options[nameWithoutOData]) {
                options[nameWithoutOData] = options[nameWithoutOData].split(';#');
            }
        });
    }
}
_a = ExternalItemAddCommand, _ExternalItemAddCommand_instances = new WeakSet(), _ExternalItemAddCommand_initTelemetry = function _ExternalItemAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            contentType: typeof args.options.contentType
        });
        this.trackUnknownOptions(this.telemetryProperties, args.options);
    });
}, _ExternalItemAddCommand_initOptions = function _ExternalItemAddCommand_initOptions() {
    this.options.unshift({
        option: '--id <id>'
    }, {
        option: '--externalConnectionId <externalConnectionId>'
    }, {
        option: '--content <content>'
    }, {
        option: '--contentType [contentType]',
        autocomplete: _a.contentType
    }, {
        option: '--acls <acls>'
    });
}, _ExternalItemAddCommand_initValidators = function _ExternalItemAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.contentType &&
            _a.contentType.indexOf(args.options.contentType) < 0) {
            return `${args.options.contentType} is not a valid value for contentType. Allowed values are ${_a.contentType.join(', ')}`;
        }
        // verify that each value for ACLs consists of three parts
        // and that the values are correct
        const acls = args.options.acls.split(';');
        for (let i = 0; i < acls.length; i++) {
            const acl = acls[i].split(',');
            if (acl.length !== 3) {
                return `The value ${acls[i]} for option acls is not in the correct format. The correct format is "accessType,type,value", eg. "grant,everyone,everyone"`;
            }
            const accessTypeValues = ['grant', 'deny'];
            if (accessTypeValues.indexOf(acl[0]) < 0) {
                return `The value ${acl[0]} for option acls is not valid. Allowed values are ${accessTypeValues.join(', ')}}`;
            }
            const aclTypeValues = ['user', 'group', 'everyone', 'everyoneExceptGuests', 'externalGroup'];
            if (aclTypeValues.indexOf(acl[1]) < 0) {
                return `The value ${acl[1]} for option acls is not valid. Allowed values are ${aclTypeValues.join(', ')}}`;
            }
        }
        return true;
    });
};
ExternalItemAddCommand.contentType = ['text', 'html'];
export default new ExternalItemAddCommand();
//# sourceMappingURL=item-add.js.map