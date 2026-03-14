var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PurviewRetentionLabelAddCommand_instances, _a, _PurviewRetentionLabelAddCommand_initTelemetry, _PurviewRetentionLabelAddCommand_initOptions, _PurviewRetentionLabelAddCommand_initValidators, _PurviewRetentionLabelAddCommand_initOptionSets;
import request from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PurviewRetentionLabelAddCommand extends GraphCommand {
    get name() {
        return commands.RETENTIONLABEL_ADD;
    }
    get description() {
        return 'Create a retention label';
    }
    constructor() {
        super();
        _PurviewRetentionLabelAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _PurviewRetentionLabelAddCommand_instances, "m", _PurviewRetentionLabelAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionLabelAddCommand_instances, "m", _PurviewRetentionLabelAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionLabelAddCommand_instances, "m", _PurviewRetentionLabelAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionLabelAddCommand_instances, "m", _PurviewRetentionLabelAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const retentionTrigger = args.options.retentionTrigger ? args.options.retentionTrigger : 'dateLabeled';
        const defaultRecordBehavior = args.options.defaultRecordBehavior ? args.options.defaultRecordBehavior : 'startLocked';
        const requestBody = {
            displayName: args.options.displayName,
            behaviorDuringRetentionPeriod: args.options.behaviorDuringRetentionPeriod,
            actionAfterRetentionPeriod: args.options.actionAfterRetentionPeriod,
            retentionTrigger: retentionTrigger,
            retentionDuration: {
                '@odata.type': '#microsoft.graph.security.retentionDurationInDays',
                days: args.options.retentionDuration
            },
            defaultRecordBehavior: defaultRecordBehavior
        };
        if (args.options.retentionTrigger === 'dateOfEvent') {
            const eventTypeId = await this.getEventTypeId(args, logger);
            requestBody['retentionEventType@odata.bind'] = `https://graph.microsoft.com/beta/security/triggerTypes/retentionEventTypes/${eventTypeId}`;
        }
        if (args.options.descriptionForAdmins) {
            if (this.verbose) {
                await logger.logToStderr(`Using '${args.options.descriptionForAdmins}' as descriptionForAdmins`);
            }
            requestBody.descriptionForAdmins = args.options.descriptionForAdmins;
        }
        if (args.options.descriptionForUsers) {
            if (this.verbose) {
                await logger.logToStderr(`Using '${args.options.descriptionForUsers}' as descriptionForUsers`);
            }
            requestBody.descriptionForUsers = args.options.descriptionForUsers;
        }
        if (args.options.labelToBeApplied) {
            if (this.verbose) {
                await logger.logToStderr(`Using '${args.options.labelToBeApplied}' as labelToBeApplied...`);
            }
            requestBody.labelToBeApplied = args.options.labelToBeApplied;
        }
        const requestOptions = {
            url: `${this.resource}/beta/security/labels/retentionLabels`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            data: requestBody,
            responseType: 'json'
        };
        try {
            const response = await request.post(requestOptions);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataPromise(err);
        }
    }
    async getEventTypeId(args, logger) {
        if (args.options.eventTypeId) {
            return args.options.eventTypeId;
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving the event type id for event type ${args.options.eventTypeName}`);
        }
        const eventTypes = await odata.getAllItems(`${this.resource}/beta/security/triggerTypes/retentionEventTypes`);
        const filteredEventTypes = eventTypes.filter((eventType) => eventType.displayName === args.options.eventTypeName);
        if (filteredEventTypes.length === 0) {
            throw `The specified retention event type '${args.options.eventTypeName}' does not exist.`;
        }
        return filteredEventTypes[0].id;
    }
}
_a = PurviewRetentionLabelAddCommand, _PurviewRetentionLabelAddCommand_instances = new WeakSet(), _PurviewRetentionLabelAddCommand_initTelemetry = function _PurviewRetentionLabelAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            retentionTrigger: typeof args.options.retentionTrigger !== 'undefined',
            defaultRecordBehavior: typeof args.options.defaultRecordBehavior !== 'undefined',
            descriptionForUsers: typeof args.options.descriptionForUsers !== 'undefined',
            descriptionForAdmins: typeof args.options.descriptionForAdmins !== 'undefined',
            labelToBeApplied: typeof args.options.labelToBeApplied !== 'undefined',
            eventTypeId: typeof args.options.eventTypeId !== 'undefined',
            eventTypeName: typeof args.options.eventTypeName !== 'undefined'
        });
    });
}, _PurviewRetentionLabelAddCommand_initOptions = function _PurviewRetentionLabelAddCommand_initOptions() {
    this.options.unshift({
        option: '-n, --displayName <displayName>'
    }, {
        option: '--behaviorDuringRetentionPeriod <behaviorDuringRetentionPeriod>',
        autocomplete: _a.behaviorDuringRetentionPeriods
    }, {
        option: '--actionAfterRetentionPeriod <actionAfterRetentionPeriod>',
        autocomplete: _a.actionAfterRetentionPeriods
    }, {
        option: '--retentionDuration <retentionDuration>'
    }, {
        option: '-t, --retentionTrigger [retentionTrigger]',
        autocomplete: _a.retentionTriggers
    }, {
        option: '--defaultRecordBehavior [defaultRecordBehavior]',
        autocomplete: _a.defaultRecordBehavior
    }, {
        option: '--descriptionForUsers [descriptionForUsers]'
    }, {
        option: '--descriptionForAdmins [descriptionForAdmins]'
    }, {
        option: '--labelToBeApplied [labelToBeApplied]'
    }, {
        option: '--eventTypeId [eventTypeId]'
    }, {
        option: '--eventTypeName [eventTypeName]'
    });
}, _PurviewRetentionLabelAddCommand_initValidators = function _PurviewRetentionLabelAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (isNaN(args.options.retentionDuration)) {
            return `Specified retentionDuration ${args.options.retentionDuration} is not a number`;
        }
        if (_a.behaviorDuringRetentionPeriods.indexOf(args.options.behaviorDuringRetentionPeriod) === -1) {
            return `${args.options.behaviorDuringRetentionPeriod} is not a valid behavior of a document with the label. Allowed values are ${_a.behaviorDuringRetentionPeriods.join(', ')}`;
        }
        if (_a.actionAfterRetentionPeriods.indexOf(args.options.actionAfterRetentionPeriod) === -1) {
            return `${args.options.actionAfterRetentionPeriod} is not a valid action to take on a document with the label. Allowed values are ${_a.actionAfterRetentionPeriods.join(', ')}`;
        }
        if (args.options.retentionTrigger &&
            _a.retentionTriggers.indexOf(args.options.retentionTrigger) === -1) {
            return `${args.options.retentionTrigger} is not a valid action retention duration calculation. Allowed values are ${_a.retentionTriggers.join(', ')}`;
        }
        if (args.options.defaultRecordBehavior &&
            _a.defaultRecordBehavior.indexOf(args.options.defaultRecordBehavior) === -1) {
            return `${args.options.defaultRecordBehavior} is not a valid state of a record label. Allowed values are ${_a.defaultRecordBehavior.join(', ')}`;
        }
        return true;
    });
}, _PurviewRetentionLabelAddCommand_initOptionSets = function _PurviewRetentionLabelAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['eventTypeId', 'eventTypeName'], runsWhen(args) { return args.options.retentionTrigger === 'dateOfEvent'; } });
};
PurviewRetentionLabelAddCommand.behaviorDuringRetentionPeriods = ['doNotRetain', 'retain', 'retainAsRecord', 'retainAsRegulatoryRecord'];
PurviewRetentionLabelAddCommand.actionAfterRetentionPeriods = ['none', 'delete', 'startDispositionReview'];
PurviewRetentionLabelAddCommand.retentionTriggers = ['dateLabeled', 'dateCreated', 'dateModified', 'dateOfEvent'];
PurviewRetentionLabelAddCommand.defaultRecordBehavior = ['startLocked', 'startUnlocked'];
export default new PurviewRetentionLabelAddCommand();
//# sourceMappingURL=retentionlabel-add.js.map