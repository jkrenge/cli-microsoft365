var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListAddCommand_instances, _a, _SpoListAddCommand_initTelemetry, _SpoListAddCommand_initOptions, _SpoListAddCommand_initValidators, _SpoListAddCommand_initTypes;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { DraftVisibilityType } from './DraftVisibilityType.js';
import { ListExperience } from './ListExperience.js';
import { ListTemplateType } from './ListTemplateType.js';
class SpoListAddCommand extends SpoCommand {
    get name() {
        return commands.LIST_ADD;
    }
    get description() {
        return 'Creates list in the specified site';
    }
    /**
     * Maps the base ListTemplateType enum to string array so it can
     * more easily be used in validation or descriptions.
     */
    get listTemplateTypeMap() {
        const result = [];
        for (const template in ListTemplateType) {
            if (typeof ListTemplateType[template] === 'number') {
                result.push(template);
            }
        }
        return result;
    }
    /**
     * Maps the base DraftVisibilityType enum to string array so it can
     * more easily be used in validation or descriptions.
     */
    get draftVisibilityTypeMap() {
        const result = [];
        for (const draftType in DraftVisibilityType) {
            if (typeof DraftVisibilityType[draftType] === 'number') {
                result.push(draftType);
            }
        }
        return result;
    }
    /**
     * Maps the base ListExperience enum to string array so it can
     * more easily be used in validation or descriptions.
     */
    get listExperienceMap() {
        const result = [];
        for (const experience in ListExperience) {
            if (typeof ListExperience[experience] === 'number') {
                result.push(experience);
            }
        }
        return result;
    }
    constructor() {
        super();
        _SpoListAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListAddCommand_instances, "m", _SpoListAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListAddCommand_instances, "m", _SpoListAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListAddCommand_instances, "m", _SpoListAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListAddCommand_instances, "m", _SpoListAddCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Creating list in site at ${args.options.webUrl}...`);
        }
        const requestBody = this.mapRequestBody(args.options);
        const requestOptions = {
            url: `${args.options.webUrl}/_api/web/lists`,
            method: 'POST',
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            data: requestBody,
            responseType: 'json'
        };
        try {
            const listInstance = await request.post(requestOptions);
            await logger.log(listInstance);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    mapRequestBody(options) {
        const requestBody = {
            Title: options.title,
            BaseTemplate: options.baseTemplate ? ListTemplateType[options.baseTemplate.trim()].valueOf() : ListTemplateType.GenericList
        };
        if (options.description) {
            requestBody.Description = options.description;
        }
        if (options.templateFeatureId) {
            requestBody.TemplateFeatureId = options.templateFeatureId;
        }
        if (options.allowDeletion !== undefined) {
            requestBody.AllowDeletion = options.allowDeletion;
        }
        if (options.allowEveryoneViewItems !== undefined) {
            requestBody.AllowEveryoneViewItems = options.allowEveryoneViewItems;
        }
        if (options.allowMultiResponses !== undefined) {
            requestBody.AllowMultiResponses = options.allowMultiResponses;
        }
        if (options.contentTypesEnabled !== undefined) {
            requestBody.ContentTypesEnabled = options.contentTypesEnabled;
        }
        if (options.crawlNonDefaultViews !== undefined) {
            requestBody.CrawlNonDefaultViews = options.crawlNonDefaultViews;
        }
        if (options.defaultContentApprovalWorkflowId) {
            requestBody.DefaultContentApprovalWorkflowId = options.defaultContentApprovalWorkflowId;
        }
        if (options.defaultDisplayFormUrl) {
            requestBody.DefaultDisplayFormUrl = options.defaultDisplayFormUrl;
        }
        if (options.defaultEditFormUrl) {
            requestBody.DefaultEditFormUrl = options.defaultEditFormUrl;
        }
        if (options.direction) {
            requestBody.Direction = options.direction;
        }
        if (options.disableCommenting !== undefined) {
            requestBody.DisableCommenting = options.disableCommenting;
        }
        if (options.disableGridEditing !== undefined) {
            requestBody.DisableGridEditing = options.disableGridEditing;
        }
        if (options.draftVersionVisibility) {
            requestBody.DraftVersionVisibility = DraftVisibilityType[options.draftVersionVisibility.trim()];
        }
        if (options.emailAlias) {
            requestBody.EmailAlias = options.emailAlias;
        }
        if (options.enableAssignToEmail !== undefined) {
            requestBody.EnableAssignToEmail = options.enableAssignToEmail;
        }
        if (options.enableAttachments !== undefined) {
            requestBody.EnableAttachments = options.enableAttachments;
        }
        if (options.enableDeployWithDependentList !== undefined) {
            requestBody.EnableDeployWithDependentList = options.enableDeployWithDependentList;
        }
        if (options.enableFolderCreation !== undefined) {
            requestBody.EnableFolderCreation = options.enableFolderCreation;
        }
        if (options.enableMinorVersions !== undefined) {
            requestBody.EnableMinorVersions = options.enableMinorVersions;
        }
        if (options.enableModeration !== undefined) {
            requestBody.EnableModeration = options.enableModeration;
        }
        if (options.enablePeopleSelector !== undefined) {
            requestBody.EnablePeopleSelector = options.enablePeopleSelector;
        }
        if (options.enableResourceSelector !== undefined) {
            requestBody.EnableResourceSelector = options.enableResourceSelector;
        }
        if (options.enableSchemaCaching !== undefined) {
            requestBody.EnableSchemaCaching = options.enableSchemaCaching;
        }
        if (options.enableSyndication !== undefined) {
            requestBody.EnableSyndication = options.enableSyndication;
        }
        if (options.enableThrottling !== undefined) {
            requestBody.EnableThrottling = options.enableThrottling;
        }
        if (options.enableVersioning !== undefined) {
            requestBody.EnableVersioning = options.enableVersioning;
        }
        if (options.enforceDataValidation !== undefined) {
            requestBody.EnforceDataValidation = options.enforceDataValidation;
        }
        if (options.excludeFromOfflineClient !== undefined) {
            requestBody.ExcludeFromOfflineClient = options.excludeFromOfflineClient;
        }
        if (options.fetchPropertyBagForListView !== undefined) {
            requestBody.FetchPropertyBagForListView = options.fetchPropertyBagForListView;
        }
        if (options.followable !== undefined) {
            requestBody.Followable = options.followable;
        }
        if (options.forceCheckout !== undefined) {
            requestBody.ForceCheckout = options.forceCheckout;
        }
        if (options.forceDefaultContentType !== undefined) {
            requestBody.ForceDefaultContentType = options.forceDefaultContentType;
        }
        if (options.hidden !== undefined) {
            requestBody.Hidden = options.hidden;
        }
        if (options.includedInMyFilesScope !== undefined) {
            requestBody.IncludedInMyFilesScope = options.includedInMyFilesScope;
        }
        if (options.irmEnabled !== undefined) {
            requestBody.IrmEnabled = options.irmEnabled;
        }
        if (options.irmExpire !== undefined) {
            requestBody.IrmExpire = options.irmExpire;
        }
        if (options.irmReject !== undefined) {
            requestBody.IrmReject = options.irmReject;
        }
        if (options.isApplicationList !== undefined) {
            requestBody.IsApplicationList = options.isApplicationList;
        }
        if (options.listExperienceOptions) {
            requestBody.ListExperienceOptions = ListExperience[options.listExperienceOptions.trim()];
        }
        if (options.majorVersionLimit) {
            requestBody.MajorVersionLimit = options.majorVersionLimit;
        }
        if (options.majorWithMinorVersionsLimit) {
            requestBody.MajorWithMinorVersionsLimit = options.majorWithMinorVersionsLimit;
        }
        if (options.multipleDataList !== undefined) {
            requestBody.MultipleDataList = options.multipleDataList;
        }
        if (options.navigateForFormsPages !== undefined) {
            requestBody.NavigateForFormsPages = options.navigateForFormsPages;
        }
        if (options.needUpdateSiteClientTag !== undefined) {
            requestBody.NeedUpdateSiteClientTag = options.needUpdateSiteClientTag;
        }
        if (options.noCrawl !== undefined) {
            requestBody.NoCrawl = options.noCrawl;
        }
        if (options.onQuickLaunch !== undefined) {
            requestBody.OnQuickLaunch = options.onQuickLaunch;
        }
        if (options.ordered !== undefined) {
            requestBody.Ordered = options.ordered;
        }
        if (options.parserDisabled !== undefined) {
            requestBody.ParserDisabled = options.parserDisabled;
        }
        if (options.readOnlyUI !== undefined) {
            requestBody.ReadOnlyUI = options.readOnlyUI;
        }
        if (options.readSecurity) {
            requestBody.ReadSecurity = options.readSecurity;
        }
        if (options.requestAccessEnabled !== undefined) {
            requestBody.RequestAccessEnabled = options.requestAccessEnabled;
        }
        if (options.restrictUserUpdates !== undefined) {
            requestBody.RestrictUserUpdates = options.restrictUserUpdates;
        }
        if (options.sendToLocationName) {
            requestBody.SendToLocationName = options.sendToLocationName;
        }
        if (options.sendToLocationUrl) {
            requestBody.SendToLocationUrl = options.sendToLocationUrl;
        }
        if (options.showUser !== undefined) {
            requestBody.ShowUser = options.showUser;
        }
        if (options.useFormsForDisplay !== undefined) {
            requestBody.UseFormsForDisplay = options.useFormsForDisplay;
        }
        if (options.validationFormula) {
            requestBody.ValidationFormula = options.validationFormula;
        }
        if (options.validationMessage) {
            requestBody.ValidationMessage = options.validationMessage;
        }
        if (options.writeSecurity) {
            requestBody.WriteSecurity = options.writeSecurity;
        }
        return requestBody;
    }
}
_a = SpoListAddCommand, _SpoListAddCommand_instances = new WeakSet(), _SpoListAddCommand_initTelemetry = function _SpoListAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        const telemetryProps = {};
        // add properties with identifiable data
        [
            'baseTemplate',
            'description',
            'templateFeatureId',
            'defaultContentApprovalWorkflowId',
            'defaultDisplayFormUrl',
            'defaultEditFormUrl',
            'emailAlias',
            'sendToLocationName',
            'sendToLocationUrl',
            'validationFormula',
            'validationMessage'
        ].forEach(o => {
            const value = args.options[o];
            if (value) {
                telemetryProps[o] = (typeof value !== 'undefined').toString();
            }
        });
        // add boolean values
        _a.booleanOptions.forEach(o => {
            const value = args.options[o];
            if (value !== undefined) {
                telemetryProps[o] = value.toString();
            }
        });
        // add properties with non-identifiable data
        [
            'baseTemplate',
            'direction',
            'draftVersionVisibility',
            'listExperienceOptions',
            'majorVersionLimit',
            'majorWithMinorVersionsLimit',
            'readSecurity',
            'writeSecurity'
        ].forEach(o => {
            const value = args.options[o];
            if (value) {
                telemetryProps[o] = value;
            }
        });
        Object.assign(this.telemetryProperties, telemetryProps);
    });
}, _SpoListAddCommand_initOptions = function _SpoListAddCommand_initOptions() {
    this.options.unshift({
        option: '-t, --title <title>'
    }, {
        option: '--baseTemplate [baseTemplate]',
        autocomplete: this.listTemplateTypeMap
    }, {
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--description [description]'
    }, {
        option: '--templateFeatureId [templateFeatureId]'
    }, {
        option: '--allowDeletion [allowDeletion]',
        autocomplete: ['true', 'false']
    }, {
        option: '--allowEveryoneViewItems [allowEveryoneViewItems]',
        autocomplete: ['true', 'false']
    }, {
        option: '--allowMultiResponses [allowMultiResponses]',
        autocomplete: ['true', 'false']
    }, {
        option: '--contentTypesEnabled [contentTypesEnabled]',
        autocomplete: ['true', 'false']
    }, {
        option: '--crawlNonDefaultViews [crawlNonDefaultViews]',
        autocomplete: ['true', 'false']
    }, {
        option: '--defaultContentApprovalWorkflowId [defaultContentApprovalWorkflowId]'
    }, {
        option: '--defaultDisplayFormUrl [defaultDisplayFormUrl]'
    }, {
        option: '--defaultEditFormUrl [defaultEditFormUrl]'
    }, {
        option: '--direction [direction]',
        autocomplete: ['NONE', 'LTR', 'RTL']
    }, {
        option: '--disableCommenting [disableCommenting]',
        autocomplete: ['true', 'false']
    }, {
        option: '--disableGridEditing [disableGridEditing]',
        autocomplete: ['true', 'false']
    }, {
        option: '--draftVersionVisibility [draftVersionVisibility]',
        autocomplete: this.draftVisibilityTypeMap
    }, {
        option: '--emailAlias [emailAlias]'
    }, {
        option: '--enableAssignToEmail [enableAssignToEmail]',
        autocomplete: ['true', 'false']
    }, {
        option: '--enableAttachments [enableAttachments]',
        autocomplete: ['true', 'false']
    }, {
        option: '--enableDeployWithDependentList [enableDeployWithDependentList]',
        autocomplete: ['true', 'false']
    }, {
        option: '--enableFolderCreation [enableFolderCreation]',
        autocomplete: ['true', 'false']
    }, {
        option: '--enableMinorVersions [enableMinorVersions]',
        autocomplete: ['true', 'false']
    }, {
        option: '--enableModeration [enableModeration]',
        autocomplete: ['true', 'false']
    }, {
        option: '--enablePeopleSelector [enablePeopleSelector]',
        autocomplete: ['true', 'false']
    }, {
        option: '--enableResourceSelector [enableResourceSelector]',
        autocomplete: ['true', 'false']
    }, {
        option: '--enableSchemaCaching [enableSchemaCaching]',
        autocomplete: ['true', 'false']
    }, {
        option: '--enableSyndication [enableSyndication]',
        autocomplete: ['true', 'false']
    }, {
        option: '--enableThrottling [enableThrottling]',
        autocomplete: ['true', 'false']
    }, {
        option: '--enableVersioning [enableVersioning]',
        autocomplete: ['true', 'false']
    }, {
        option: '--enforceDataValidation [enforceDataValidation]',
        autocomplete: ['true', 'false']
    }, {
        option: '--excludeFromOfflineClient [excludeFromOfflineClient]',
        autocomplete: ['true', 'false']
    }, {
        option: '--fetchPropertyBagForListView [fetchPropertyBagForListView]',
        autocomplete: ['true', 'false']
    }, {
        option: '--followable [followable]',
        autocomplete: ['true', 'false']
    }, {
        option: '--forceCheckout [forceCheckout]',
        autocomplete: ['true', 'false']
    }, {
        option: '--forceDefaultContentType [forceDefaultContentType]',
        autocomplete: ['true', 'false']
    }, {
        option: '--hidden [hidden]',
        autocomplete: ['true', 'false']
    }, {
        option: '--includedInMyFilesScope [includedInMyFilesScope]',
        autocomplete: ['true', 'false']
    }, {
        option: '--irmEnabled [irmEnabled]',
        autocomplete: ['true', 'false']
    }, {
        option: '--irmExpire [irmExpire]',
        autocomplete: ['true', 'false']
    }, {
        option: '--irmReject [irmReject]',
        autocomplete: ['true', 'false']
    }, {
        option: '--isApplicationList [isApplicationList]',
        autocomplete: ['true', 'false']
    }, {
        option: '--listExperienceOptions [listExperienceOptions]',
        autocomplete: this.listExperienceMap
    }, {
        option: '--majorVersionLimit [majorVersionLimit]'
    }, {
        option: '--majorWithMinorVersionsLimit [majorWithMinorVersionsLimit]'
    }, {
        option: '--multipleDataList [multipleDataList]',
        autocomplete: ['true', 'false']
    }, {
        option: '--navigateForFormsPages [navigateForFormsPages]',
        autocomplete: ['true', 'false']
    }, {
        option: '--needUpdateSiteClientTag [needUpdateSiteClientTag]',
        autocomplete: ['true', 'false']
    }, {
        option: '--noCrawl [noCrawl]',
        autocomplete: ['true', 'false']
    }, {
        option: '--onQuickLaunch [onQuickLaunch]',
        autocomplete: ['true', 'false']
    }, {
        option: '--ordered [ordered]',
        autocomplete: ['true', 'false']
    }, {
        option: '--parserDisabled [parserDisabled]',
        autocomplete: ['true', 'false']
    }, {
        option: '--readOnlyUI [readOnlyUI]',
        autocomplete: ['true', 'false']
    }, {
        option: '--readSecurity [readSecurity]',
        autocomplete: ['1', '2']
    }, {
        option: '--requestAccessEnabled [requestAccessEnabled]',
        autocomplete: ['true', 'false']
    }, {
        option: '--restrictUserUpdates [restrictUserUpdates]',
        autocomplete: ['true', 'false']
    }, {
        option: '--sendToLocationName [sendToLocationName]'
    }, {
        option: '--sendToLocationUrl [sendToLocationUrl]'
    }, {
        option: '--showUser [showUser]',
        autocomplete: ['true', 'false']
    }, {
        option: '--useFormsForDisplay [useFormsForDisplay]',
        autocomplete: ['true', 'false']
    }, {
        option: '--validationFormula [validationFormula]'
    }, {
        option: '--validationMessage [validationMessage]'
    }, {
        option: '--writeSecurity [writeSecurity]',
        autocomplete: ['1', '2', '4']
    });
}, _SpoListAddCommand_initValidators = function _SpoListAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.baseTemplate) {
            const template = ListTemplateType[args.options.baseTemplate.trim()];
            if (!template) {
                return `${args.options.baseTemplate} is not a valid baseTemplate value`;
            }
        }
        if (args.options.templateFeatureId &&
            !validation.isValidGuid(args.options.templateFeatureId)) {
            return `${args.options.templateFeatureId} in option templateFeatureId is not a valid GUID`;
        }
        if (args.options.defaultContentApprovalWorkflowId &&
            !validation.isValidGuid(args.options.defaultContentApprovalWorkflowId)) {
            return `${args.options.defaultContentApprovalWorkflowId} in option defaultContentApprovalWorkflowId is not a valid GUID`;
        }
        if (args.options.direction &&
            ['NONE', 'LTR', 'RTL'].indexOf(args.options.direction) === -1) {
            return `${args.options.direction} is not a valid direction value. Allowed values are NONE|LTR|RTL`;
        }
        if (args.options.draftVersionVisibility) {
            const draftType = DraftVisibilityType[args.options.draftVersionVisibility.trim()];
            if (draftType === undefined) {
                return `${args.options.draftVersionVisibility} is not a valid draftVisibilityType value`;
            }
        }
        if (args.options.emailAlias && args.options.enableAssignToEmail !== true) {
            return `emailAlias could not be set if enableAssignToEmail is not set to true. Please set enableAssignToEmail.`;
        }
        if (args.options.listExperienceOptions) {
            const experience = ListExperience[args.options.listExperienceOptions.trim()];
            if (!experience) {
                return `${args.options.listExperienceOptions} is not a valid listExperienceOptions value`;
            }
        }
        if (args.options.majorVersionLimit && args.options.enableVersioning !== true) {
            return `majorVersionLimit option is only valid in combination with enableVersioning.`;
        }
        if (args.options.majorWithMinorVersionsLimit &&
            args.options.enableMinorVersions !== true &&
            args.options.enableModeration !== true) {
            return `majorWithMinorVersionsLimit option is only valid in combination with enableMinorVersions or enableModeration.`;
        }
        if (args.options.readSecurity &&
            args.options.readSecurity !== 1 &&
            args.options.readSecurity !== 2) {
            return `${args.options.readSecurity} is not a valid readSecurity value. Allowed values are 1|2`;
        }
        if (args.options.writeSecurity &&
            args.options.writeSecurity !== 1 &&
            args.options.writeSecurity !== 2 &&
            args.options.writeSecurity !== 4) {
            return `${args.options.writeSecurity} is not a valid writeSecurity value. Allowed values are 1|2|4`;
        }
        return true;
    });
}, _SpoListAddCommand_initTypes = function _SpoListAddCommand_initTypes() {
    this.types.string.push('title', 'baseTemplate', 'webUrl', 'templateFeatureId', 'defaultContentApprovalWorkflowId', 'draftVersionVisibility', 'listExperienceOptions');
    this.types.boolean.push(..._a.booleanOptions);
};
SpoListAddCommand.booleanOptions = [
    'allowDeletion',
    'allowEveryoneViewItems',
    'allowMultiResponses',
    'contentTypesEnabled',
    'crawlNonDefaultViews',
    'disableCommenting',
    'disableGridEditing',
    'enableAssignToEmail',
    'enableAttachments',
    'enableDeployWithDependentList',
    'enableFolderCreation',
    'enableMinorVersions',
    'enableModeration',
    'enablePeopleSelector',
    'enableResourceSelector',
    'enableSchemaCaching',
    'enableSyndication',
    'enableThrottling',
    'enableVersioning',
    'enforceDataValidation',
    'excludeFromOfflineClient',
    'fetchPropertyBagForListView',
    'followable',
    'forceCheckout',
    'forceDefaultContentType',
    'hidden',
    'includedInMyFilesScope',
    'irmEnabled',
    'irmExpire',
    'irmReject',
    'isApplicationList',
    'multipleDataList',
    'navigateForFormsPages',
    'needUpdateSiteClientTag',
    'noCrawl',
    'onQuickLaunch',
    'ordered',
    'parserDisabled',
    'readOnlyUI',
    'requestAccessEnabled',
    'restrictUserUpdates',
    'showUser',
    'useFormsForDisplay'
];
export default new SpoListAddCommand();
//# sourceMappingURL=list-add.js.map