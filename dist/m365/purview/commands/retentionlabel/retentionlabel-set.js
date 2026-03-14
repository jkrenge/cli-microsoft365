var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PurviewRetentionLabelSetCommand_instances, _PurviewRetentionLabelSetCommand_initTelemetry, _PurviewRetentionLabelSetCommand_initOptions, _PurviewRetentionLabelSetCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PurviewRetentionLabelSetCommand extends GraphCommand {
    get name() {
        return commands.RETENTIONLABEL_SET;
    }
    get description() {
        return 'Update a retention label';
    }
    constructor() {
        super();
        _PurviewRetentionLabelSetCommand_instances.add(this);
        this.allowedBehaviorDuringRetentionPeriodValues = ['doNotRetain', 'retain', 'retainAsRecord', 'retainAsRegulatoryRecord'];
        this.allowedActionAfterRetentionPeriodValues = ['none', 'delete', 'startDispositionReview'];
        this.allowedRetentionTriggerValues = ['dateLabeled', 'dateCreated', 'dateModified', 'dateOfEvent'];
        this.allowedDefaultRecordBehaviorValues = ['startLocked', 'startUnlocked'];
        __classPrivateFieldGet(this, _PurviewRetentionLabelSetCommand_instances, "m", _PurviewRetentionLabelSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionLabelSetCommand_instances, "m", _PurviewRetentionLabelSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionLabelSetCommand_instances, "m", _PurviewRetentionLabelSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.log(`Starting to update retention label with id ${args.options.id}`);
        }
        const requestBody = this.mapRequestBody(args.options);
        const requestOptions = {
            url: `${this.resource}/beta/security/labels/retentionLabels/${args.options.id}`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json',
            data: requestBody
        };
        try {
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    mapRequestBody(options) {
        const requestBody = {};
        const excludeOptions = [
            'debug',
            'verbose',
            'output',
            'id',
            'retentionDuration'
        ];
        Object.keys(options).forEach(key => {
            if (excludeOptions.indexOf(key) === -1) {
                requestBody[key] = `${options[key]}`;
            }
        });
        if (options.retentionDuration) {
            requestBody['retentionDuration'] = {
                '@odata.type': 'microsoft.graph.security.retentionDurationInDays',
                'days': options.retentionDuration
            };
        }
        return requestBody;
    }
}
_PurviewRetentionLabelSetCommand_instances = new WeakSet(), _PurviewRetentionLabelSetCommand_initTelemetry = function _PurviewRetentionLabelSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            behaviorDuringRetentionPeriod: typeof args.options.behaviorDuringRetentionPeriod !== 'undefined',
            actionAfterRetentionPeriod: typeof args.options.actionAfterRetentionPeriod !== 'undefined',
            retentionDuration: typeof args.options.retentionDuration !== 'undefined',
            retentionTrigger: typeof args.options.retentionTrigger !== 'undefined',
            defaultRecordBehavior: typeof args.options.defaultRecordBehavior !== 'undefined',
            descriptionForUsers: typeof args.options.descriptionForUsers !== 'undefined',
            descriptionForAdmins: typeof args.options.descriptionForAdmins !== 'undefined',
            labelToBeApplied: typeof args.options.labelToBeApplied !== 'undefined'
        });
    });
}, _PurviewRetentionLabelSetCommand_initOptions = function _PurviewRetentionLabelSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '--behaviorDuringRetentionPeriod [behaviorDuringRetentionPeriod]',
        autocomplete: this.allowedBehaviorDuringRetentionPeriodValues
    }, {
        option: '--actionAfterRetentionPeriod [actionAfterRetentionPeriod]',
        autocomplete: this.allowedActionAfterRetentionPeriodValues
    }, {
        option: '--retentionDuration [retentionDuration]'
    }, {
        option: '-t, --retentionTrigger [retentionTrigger]',
        autocomplete: this.allowedRetentionTriggerValues
    }, {
        option: '--defaultRecordBehavior [defaultRecordBehavior]',
        autocomplete: this.allowedDefaultRecordBehaviorValues
    }, {
        option: '--descriptionForUsers [descriptionForUsers]'
    }, {
        option: '--descriptionForAdmins [descriptionForAdmins]'
    }, {
        option: '--labelToBeApplied [labelToBeApplied]'
    });
}, _PurviewRetentionLabelSetCommand_initValidators = function _PurviewRetentionLabelSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `'${args.options.id}' is not a valid GUID.`;
        }
        const { actionAfterRetentionPeriod, behaviorDuringRetentionPeriod, defaultRecordBehavior, descriptionForAdmins, descriptionForUsers, labelToBeApplied, retentionDuration, retentionTrigger } = args.options;
        if ([actionAfterRetentionPeriod, behaviorDuringRetentionPeriod, defaultRecordBehavior, descriptionForAdmins, descriptionForUsers, labelToBeApplied, retentionDuration, retentionTrigger].every(i => typeof i === 'undefined')) {
            return `Specify at least one property to update.`;
        }
        if (behaviorDuringRetentionPeriod && this.allowedBehaviorDuringRetentionPeriodValues.indexOf(behaviorDuringRetentionPeriod) === -1) {
            return `'${behaviorDuringRetentionPeriod}' is not a valid value for the behaviorDuringRetentionPeriod option. Allowed values are ${this.allowedBehaviorDuringRetentionPeriodValues.join('|')}`;
        }
        if (actionAfterRetentionPeriod && this.allowedActionAfterRetentionPeriodValues.indexOf(actionAfterRetentionPeriod) === -1) {
            return `'${actionAfterRetentionPeriod}' is not a valid value for the actionAfterRetentionPeriod option. Allowed values are ${this.allowedActionAfterRetentionPeriodValues.join('|')}`;
        }
        if (retentionTrigger && this.allowedRetentionTriggerValues.indexOf(retentionTrigger) === -1) {
            return `'${retentionTrigger}' is not a valid value for the retentionTrigger option. Allowed values are ${this.allowedRetentionTriggerValues.join('|')}`;
        }
        if (defaultRecordBehavior && this.allowedDefaultRecordBehaviorValues.indexOf(defaultRecordBehavior) === -1) {
            return `'${defaultRecordBehavior}' is not a valid value for the defaultRecordBehavior option. Allowed values are ${this.allowedDefaultRecordBehaviorValues.join('|')}`;
        }
        return true;
    });
};
export default new PurviewRetentionLabelSetCommand();
//# sourceMappingURL=retentionlabel-set.js.map