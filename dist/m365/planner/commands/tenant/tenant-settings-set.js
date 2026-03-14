var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerTenantSettingsSetCommand_instances, _PlannerTenantSettingsSetCommand_initTelemetry, _PlannerTenantSettingsSetCommand_initOptions, _PlannerTenantSettingsSetCommand_initTypes, _PlannerTenantSettingsSetCommand_initValidators;
import request from '../../../../request.js';
import PlannerCommand from '../../../base/PlannerCommand.js';
import commands from '../../commands.js';
class PlannerTenantSettingsSetCommand extends PlannerCommand {
    get name() {
        return commands.TENANT_SETTINGS_SET;
    }
    get description() {
        return 'Sets Microsoft Planner configuration of the tenant';
    }
    constructor() {
        super();
        _PlannerTenantSettingsSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerTenantSettingsSetCommand_instances, "m", _PlannerTenantSettingsSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerTenantSettingsSetCommand_instances, "m", _PlannerTenantSettingsSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerTenantSettingsSetCommand_instances, "m", _PlannerTenantSettingsSetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _PlannerTenantSettingsSetCommand_instances, "m", _PlannerTenantSettingsSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const requestOptions = {
            url: `${this.resource}/taskAPI/tenantAdminSettings/Settings`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                prefer: 'return=representation'
            },
            responseType: 'json',
            data: {
                isPlannerAllowed: args.options.isPlannerAllowed,
                allowCalendarSharing: args.options.allowCalendarSharing,
                allowTenantMoveWithDataLoss: args.options.allowTenantMoveWithDataLoss,
                allowTenantMoveWithDataMigration: args.options.allowTenantMoveWithDataMigration,
                allowRosterCreation: args.options.allowRosterCreation,
                allowPlannerMobilePushNotifications: args.options.allowPlannerMobilePushNotifications
            }
        };
        try {
            const result = await request.patch(requestOptions);
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PlannerTenantSettingsSetCommand_instances = new WeakSet(), _PlannerTenantSettingsSetCommand_initTelemetry = function _PlannerTenantSettingsSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            isPlannerAllowed: typeof args.options.isPlannerAllowed !== 'undefined',
            allowCalendarSharing: typeof args.options.allowCalendarSharing !== 'undefined',
            allowTenantMoveWithDataLoss: typeof args.options.allowTenantMoveWithDataLoss !== 'undefined',
            allowTenantMoveWithDataMigration: typeof args.options.allowTenantMoveWithDataMigration !== 'undefined',
            allowRosterCreation: typeof args.options.allowRosterCreation !== 'undefined',
            allowPlannerMobilePushNotifications: typeof args.options.allowPlannerMobilePushNotifications !== 'undefined'
        });
    });
}, _PlannerTenantSettingsSetCommand_initOptions = function _PlannerTenantSettingsSetCommand_initOptions() {
    this.options.unshift({
        option: '--isPlannerAllowed [isPlannerAllowed]',
        autocomplete: ['true', 'false']
    }, {
        option: '--allowCalendarSharing [allowCalendarSharing]',
        autocomplete: ['true', 'false']
    }, {
        option: '--allowTenantMoveWithDataLoss [allowTenantMoveWithDataLoss]',
        autocomplete: ['true', 'false']
    }, {
        option: '--allowTenantMoveWithDataMigration [allowTenantMoveWithDataMigration]',
        autocomplete: ['true', 'false']
    }, {
        option: '--allowRosterCreation [allowRosterCreation]',
        autocomplete: ['true', 'false']
    }, {
        option: '--allowPlannerMobilePushNotifications [allowPlannerMobilePushNotifications]',
        autocomplete: ['true', 'false']
    });
}, _PlannerTenantSettingsSetCommand_initTypes = function _PlannerTenantSettingsSetCommand_initTypes() {
    this.types.boolean.push('isPlannerAllowed', 'allowCalendarSharing', 'allowTenantMoveWithDataLoss', 'allowTenantMoveWithDataMigration', 'allowRosterCreation', 'allowPlannerMobilePushNotifications');
}, _PlannerTenantSettingsSetCommand_initValidators = function _PlannerTenantSettingsSetCommand_initValidators() {
    this.validators.push(async (args) => {
        const optionsArray = [
            args.options.isPlannerAllowed, args.options.allowCalendarSharing, args.options.allowTenantMoveWithDataLoss,
            args.options.allowTenantMoveWithDataMigration, args.options.allowRosterCreation, args.options.allowPlannerMobilePushNotifications
        ];
        if (optionsArray.every(o => typeof o === 'undefined')) {
            return 'You must specify at least one option';
        }
        return true;
    });
};
export default new PlannerTenantSettingsSetCommand();
//# sourceMappingURL=tenant-settings-set.js.map