var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PurviewRetentionEventAddCommand_instances, _PurviewRetentionEventAddCommand_initOptions, _PurviewRetentionEventAddCommand_initValidators, _PurviewRetentionEventAddCommand_initTelemetry, _PurviewRetentionEventAddCommand_initOptionSets;
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { odata } from '../../../../utils/odata.js';
class PurviewRetentionEventAddCommand extends GraphCommand {
    get name() {
        return commands.RETENTIONEVENT_ADD;
    }
    get description() {
        return 'Create a retention event';
    }
    constructor() {
        super();
        _PurviewRetentionEventAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventAddCommand_instances, "m", _PurviewRetentionEventAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventAddCommand_instances, "m", _PurviewRetentionEventAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventAddCommand_instances, "m", _PurviewRetentionEventAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventAddCommand_instances, "m", _PurviewRetentionEventAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Creating retention event...`);
        }
        const eventQueries = [];
        if (args.options.assetIds) {
            eventQueries.push({ queryType: "files", query: args.options.assetIds });
        }
        if (args.options.keywords) {
            eventQueries.push({ queryType: "messages", query: args.options.keywords });
        }
        const eventTypeId = await this.getEventTypeId(logger, args);
        const data = {
            'retentionEventType@odata.bind': `https://graph.microsoft.com/v1.0/security/triggerTypes/retentionEventTypes/${eventTypeId}`,
            displayName: args.options.displayName,
            description: args.options.description,
            eventQueries: eventQueries,
            eventTriggerDateTime: args.options.triggerDateTime
        };
        try {
            const requestOptions = {
                url: `${this.resource}/v1.0/security/triggers/retentionEvents`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: data
            };
            const res = await request.post(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getEventTypeId(logger, args) {
        if (args.options.eventTypeId) {
            return args.options.eventTypeId;
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving the event type id for event type ${args.options.eventTypeName}`);
        }
        const items = await odata.getAllItems(`${this.resource}/v1.0/security/triggerTypes/retentionEventTypes`);
        const eventTypes = items.filter((x) => x.displayName === args.options.eventTypeName);
        if (eventTypes.length === 0) {
            throw `The specified event type '${args.options.eventTypeName}' does not exist.`;
        }
        return eventTypes[0].id;
    }
}
_PurviewRetentionEventAddCommand_instances = new WeakSet(), _PurviewRetentionEventAddCommand_initOptions = function _PurviewRetentionEventAddCommand_initOptions() {
    this.options.unshift({
        option: '-n, --displayName <displayName>'
    }, {
        option: '-i, --eventTypeId [eventTypeId]'
    }, {
        option: '-e, --eventTypeName [eventTypeName]'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '--triggerDateTime [triggerDateTime]'
    }, {
        option: '-a, --assetIds [assetIds]'
    }, {
        option: '-k, --keywords [keywords]'
    });
}, _PurviewRetentionEventAddCommand_initValidators = function _PurviewRetentionEventAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.triggerDateTime && !validation.isValidISODateTime(args.options.triggerDateTime)) {
            return 'The triggerDateTime is not a valid ISO date string';
        }
        if (!args.options.assetIds && !args.options.keywords) {
            return 'Specify assetIds and/or keywords, but at least one.';
        }
        return true;
    });
}, _PurviewRetentionEventAddCommand_initTelemetry = function _PurviewRetentionEventAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            description: typeof args.options.description !== 'undefined',
            triggerDateTime: typeof args.options.triggerDateTime !== 'undefined',
            eventTypeId: typeof args.options.eventTypeId !== 'undefined',
            eventTypeName: typeof args.options.eventTypeName !== 'undefined',
            assetIds: typeof args.options.assetIds !== 'undefined',
            keywords: typeof args.options.keywords !== 'undefined'
        });
    });
}, _PurviewRetentionEventAddCommand_initOptionSets = function _PurviewRetentionEventAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['eventTypeId', 'eventTypeName'] });
};
export default new PurviewRetentionEventAddCommand();
//# sourceMappingURL=retentionevent-add.js.map