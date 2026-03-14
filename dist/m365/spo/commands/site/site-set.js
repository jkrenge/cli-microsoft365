var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteSetCommand_instances, _SpoSiteSetCommand_initTelemetry, _SpoSiteSetCommand_initOptions, _SpoSiteSetCommand_initValidators, _SpoSiteSetCommand_initTypes;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { SharingCapabilities } from '../site/SharingCapabilities.js';
import { FlowsPolicy } from './FlowsPolicy.js';
import { setTimeout } from 'timers/promises';
class SpoSiteSetCommand extends SpoCommand {
    get name() {
        return commands.SITE_SET;
    }
    get description() {
        return 'Updates properties of the specified site';
    }
    constructor() {
        super();
        _SpoSiteSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteSetCommand_instances, "m", _SpoSiteSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteSetCommand_instances, "m", _SpoSiteSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteSetCommand_instances, "m", _SpoSiteSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoSiteSetCommand_instances, "m", _SpoSiteSetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            this.tenantId = await spo.getTenantId(logger, this.debug);
            this.spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            this.context = await spo.ensureFormDigest(this.spoAdminUrl, logger, this.context, this.debug);
            await this.loadSiteIds(args.options.url, logger);
            if (this.isGroupConnectedSite()) {
                await this.updateGroupConnectedSite(logger, args);
            }
            else {
                await this.updateSharePointOnlySite(logger, args);
            }
            const siteProps = await this.updateSiteProperties(logger, args);
            await this.waitForSiteUpdateCompletion(logger, args, siteProps);
            await this.applySiteDesign(logger, args);
            await this.setLogo(logger, args);
            await this.setThumbnail(logger, args);
            const lockState = await this.updateSiteLockState(logger, args);
            await this.waitForSiteUpdateCompletion(logger, args, lockState);
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
    async setLogo(logger, args) {
        if (typeof args.options.siteLogoUrl === 'undefined') {
            return Promise.resolve();
        }
        if (this.debug) {
            await logger.logToStderr(`Setting the site its logo...`);
        }
        const logoUrl = args.options.siteLogoUrl ? urlUtil.getUrlRelativePath(args.options.siteLogoUrl) : "";
        const requestOptions = {
            url: `${args.options.url}/_api/siteiconmanager/setsitelogo`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            data: {
                aspect: 1,
                relativeLogoUrl: logoUrl,
                type: 0
            },
            responseType: 'json'
        };
        return request.post(requestOptions);
    }
    async setThumbnail(logger, args) {
        if (typeof args.options.siteThumbnailUrl === 'undefined') {
            return;
        }
        if (this.debug) {
            await logger.logToStderr(`Setting the site thumbnail...`);
        }
        const thumbnailUrl = args.options.siteThumbnailUrl ? urlUtil.getUrlRelativePath(args.options.siteThumbnailUrl) : "";
        const requestOptions = {
            url: `${args.options.url}/_api/siteiconmanager/setsitelogo`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            data: {
                aspect: 0,
                relativeLogoUrl: thumbnailUrl,
                type: 0
            },
            responseType: 'json'
        };
        return request.post(requestOptions);
    }
    async updateSharePointOnlySite(logger, args) {
        if (this.debug) {
            await logger.logToStderr('Site is not group connected');
        }
        if (typeof args.options.isPublic !== 'undefined') {
            throw `The isPublic option can't be set on a site that is not groupified`;
        }
        await this.updateSiteDescription(logger, args);
        await this.updateSiteOwners(logger, args);
    }
    async waitForSiteUpdateCompletion(logger, args, response) {
        if (!response) {
            return;
        }
        const json = JSON.parse(response);
        const responseContent = json[0];
        if (responseContent.ErrorInfo) {
            throw responseContent.ErrorInfo.ErrorMessage;
        }
        const operation = json[json.length - 1];
        const isComplete = operation.IsComplete;
        if (!args.options.wait || isComplete) {
            return;
        }
        await setTimeout(operation.PollingInterval);
        await spo.waitUntilFinished({
            operationId: JSON.stringify(operation._ObjectIdentity_),
            siteUrl: this.spoAdminUrl,
            logger,
            currentContext: this.context,
            debug: this.debug,
            verbose: this.verbose
        });
    }
    async updateSiteOwners(logger, args) {
        if (!args.options.owners) {
            return;
        }
        if (this.verbose) {
            await logger.logToStderr(`Updating site owners ${args.options.url}...`);
        }
        await Promise.all(args.options.owners.split(',').map(o => {
            return this.setAdmin(args.options.url, o.trim());
        }));
    }
    async setAdmin(siteUrl, principal) {
        const requestOptions = {
            url: `${this.spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': this.context.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Method Id="47" ParentId="34" Name="SetSiteAdmin"><Parameters><Parameter Type="String">${formatting.escapeXml(siteUrl)}</Parameter><Parameter Type="String">${formatting.escapeXml(principal)}</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Constructor Id="34" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
        };
        const response = await request.post(requestOptions);
        const json = JSON.parse(response);
        const responseContent = json[0];
        if (responseContent.ErrorInfo) {
            throw responseContent.ErrorInfo.ErrorMessage;
        }
    }
    async updateSiteDescription(logger, args) {
        if (!args.options.description) {
            return;
        }
        if (this.verbose) {
            await logger.logToStderr(`Setting site description ${args.options.url}...`);
        }
        const requestOptions = {
            url: `${args.options.url}/_api/web`,
            headers: {
                'IF-MATCH': '*',
                'Accept': 'application/json;odata=nometadata',
                'content-type': 'application/json;odata=nometadata',
                'X-RequestDigest': this.context.FormDigestValue,
                'X-HTTP-Method': 'MERGE'
            },
            data: {
                Description: args.options.description
            },
            json: true
        };
        await request.post(requestOptions);
    }
    async updateSiteLockState(logger, args) {
        if (!args.options.lockState) {
            return;
        }
        if (this.verbose) {
            await logger.logToStderr(`Setting site lock state ${args.options.url}...`);
        }
        const requestOptions = {
            url: `${this.spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': this.context.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="7" ObjectPathId="5" Name="LockState"><Parameter Type="String">${formatting.escapeXml(args.options.lockState)}</Parameter></SetProperty><ObjectPath Id="9" ObjectPathId="8" /><ObjectIdentityQuery Id="10" ObjectPathId="5" /><Query Id="11" ObjectPathId="8"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="5" ParentId="3" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.url)}</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Id="8" ParentId="5" Name="Update" /><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
        };
        return request.post(requestOptions);
    }
    async updateGroupConnectedSite(logger, args) {
        if (this.debug) {
            await logger.logToStderr(`Site attached to group ${this.groupId}`);
        }
        if (typeof args.options.title === 'undefined' &&
            typeof args.options.description === 'undefined' &&
            typeof args.options.isPublic === 'undefined' &&
            typeof args.options.owners === 'undefined') {
            return;
        }
        const promises = [];
        if (typeof args.options.title !== 'undefined') {
            const requestOptions = {
                url: `${this.spoAdminUrl}/_api/SPOGroup/UpdateGroupPropertiesBySiteId`,
                headers: {
                    accept: 'application/json;odata=nometadata',
                    'content-type': 'application/json;charset=utf-8',
                    'X-RequestDigest': this.context.FormDigestValue
                },
                data: {
                    groupId: this.groupId,
                    siteId: this.siteId,
                    displayName: args.options.title
                },
                responseType: 'json'
            };
            promises.push(request.post(requestOptions));
        }
        if (typeof args.options.isPublic !== 'undefined') {
            promises.push(entraGroup.setGroup(this.groupId, (args.options.isPublic === false), undefined, undefined, logger, this.verbose));
        }
        if (args.options.description) {
            promises.push(this.setGroupifiedSiteDescription(args.options.description));
        }
        promises.push(this.setGroupifiedSiteOwners(logger, args));
        await Promise.all(promises);
    }
    setGroupifiedSiteDescription(description) {
        const requestOptions = {
            url: `https://graph.microsoft.com/v1.0/groups/${this.groupId}`,
            headers: {
                'content-type': 'application/json;odata.metadata=none'
            },
            data: {
                description: description
            }
        };
        return request.patch(requestOptions);
    }
    async setGroupifiedSiteOwners(logger, args) {
        if (typeof args.options.owners === 'undefined') {
            return;
        }
        const owners = args.options.owners.split(',').map(o => o.trim());
        if (this.verbose) {
            await logger.logToStderr('Retrieving user information to set group owners...');
        }
        const requestOptions = {
            url: `https://graph.microsoft.com/v1.0/users?$filter=${owners.map(o => `userPrincipalName eq '${o}'`).join(' or ')}&$select=id`,
            headers: {
                'content-type': 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        if (response.value.length === 0) {
            return;
        }
        await Promise.all(response.value.map(user => {
            const requestOptions = {
                url: `${this.spoAdminUrl}/_api/SP.Directory.DirectorySession/Group('${this.groupId}')/Owners/Add(objectId='${user.id}', principalName='')`,
                headers: {
                    'content-type': 'application/json;odata=verbose'
                }
            };
            return request.post(requestOptions);
        }));
    }
    async updateSiteProperties(logger, args) {
        const isGroupConnectedSite = this.isGroupConnectedSite();
        const sharedProperties = ['classification', 'disableFlows', 'socialBarOnSitePagesDisabled', 'shareByEmailEnabled', 'sharingCapability', 'noScriptSite'];
        const siteProperties = ['title', 'resourceQuota', 'resourceQuotaWarningLevel', 'storageQuota', 'storageQuotaWarningLevel', 'allowSelfServiceUpgrade'];
        let properties = sharedProperties;
        if (!isGroupConnectedSite) {
            properties = properties.concat(siteProperties);
        }
        let updatedProperties = false;
        for (let i = 0; i < properties.length; i++) {
            if (typeof args.options[properties[i]] !== 'undefined') {
                updatedProperties = true;
                break;
            }
        }
        if (!updatedProperties) {
            return;
        }
        this.context = await spo.ensureFormDigest(this.spoAdminUrl, logger, this.context, this.debug);
        if (this.verbose) {
            await logger.logToStderr(`Updating site ${args.options.url} properties...`);
        }
        let propertyId = 27;
        const payload = [];
        const sitePropertiesPayload = [];
        if (!isGroupConnectedSite) {
            if (args.options.title) {
                payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="Title"><Parameter Type="String">${formatting.escapeXml(args.options.title)}</Parameter></SetProperty>`);
            }
            if (args.options.resourceQuota) {
                payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="UserCodeMaximumLevel"><Parameter Type="Double">${args.options.resourceQuota}</Parameter></SetProperty>`);
            }
            if (args.options.resourceQuotaWarningLevel) {
                payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="UserCodeWarningLevel"><Parameter Type="Double">${args.options.resourceQuotaWarningLevel}</Parameter></SetProperty>`);
            }
            if (args.options.storageQuota) {
                payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="StorageMaximumLevel"><Parameter Type="Int64">${args.options.storageQuota}</Parameter></SetProperty>`);
            }
            if (args.options.storageQuotaWarningLevel) {
                payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="StorageWarningLevel"><Parameter Type="Int64">${args.options.storageQuotaWarningLevel}</Parameter></SetProperty>`);
            }
            if (args.options.allowSelfServiceUpgrade !== undefined) {
                payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="AllowSelfServiceUpgrade"><Parameter Type="Boolean">${args.options.allowSelfServiceUpgrade}</Parameter></SetProperty>`);
            }
        }
        if (typeof args.options.classification === 'string') {
            sitePropertiesPayload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="Classification"><Parameter Type="String">${formatting.escapeXml(args.options.classification)}</Parameter></SetProperty>`);
        }
        if (args.options.disableFlows !== undefined) {
            const flowsPolicy = args.options.disableFlows ? FlowsPolicy.Disabled : FlowsPolicy.Enabled;
            payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="DisableFlows"><Parameter Type="Enum">${flowsPolicy}</Parameter></SetProperty>`);
        }
        if (typeof args.options.socialBarOnSitePagesDisabled !== 'undefined') {
            payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="SocialBarOnSitePagesDisabled"><Parameter Type="Boolean">${args.options.socialBarOnSitePagesDisabled}</Parameter></SetProperty>`);
        }
        if (args.options.shareByEmailEnabled !== undefined) {
            sitePropertiesPayload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="ShareByEmailEnabled"><Parameter Type="Boolean">${args.options.shareByEmailEnabled}</Parameter></SetProperty>`);
        }
        if (args.options.sharingCapability) {
            const sharingCapability = SharingCapabilities[args.options.sharingCapability];
            payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="SharingCapability"><Parameter Type="Enum">${sharingCapability}</Parameter></SetProperty>`);
        }
        if (args.options.noScriptSite !== undefined) {
            const noScriptSite = args.options.noScriptSite ? 2 : 1;
            payload.push(`<SetProperty Id="${propertyId}" ObjectPathId="5" Name="DenyAddAndCustomizePages"><Parameter Type="Enum">${noScriptSite}</Parameter></SetProperty>`);
        }
        let response;
        let sitePropertiesResponse;
        if (sitePropertiesPayload.length > 0) {
            const requestOptions = {
                url: `${args.options.url}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': this.context.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${sitePropertiesPayload.join('')}</Actions><ObjectPaths><StaticProperty Id="1" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /><Property Id="5" ParentId="1" Name="Site" /></ObjectPaths></Request>`
            };
            sitePropertiesResponse = await request.post(requestOptions);
        }
        if (payload.length > 0) {
            const pos = this.tenantId.indexOf('|') + 1;
            const requestOptions = {
                url: `${this.spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': this.context.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${payload.join('')}<ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|${this.tenantId.substring(pos, this.tenantId.indexOf('&'))}&#xA;SiteProperties&#xA;${formatting.encodeQueryParameter(args.options.url)}" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`
            };
            response = await request.post(requestOptions);
        }
        return response || sitePropertiesResponse;
    }
    async applySiteDesign(logger, args) {
        if (typeof args.options.siteDesignId === 'undefined') {
            return;
        }
        return spo.applySiteDesign(args.options.url, args.options.siteDesignId, logger, this.verbose);
    }
    async loadSiteIds(siteUrl, logger) {
        if (this.debug) {
            await logger.logToStderr('Loading site IDs...');
        }
        const requestOptions = {
            url: `${siteUrl}/_api/site?$select=GroupId,Id`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const siteInfo = await request.get(requestOptions);
        this.groupId = siteInfo.GroupId;
        this.siteId = siteInfo.Id;
        if (this.debug) {
            await logger.logToStderr(`Retrieved site IDs. siteId: ${this.siteId}, groupId: ${this.groupId}`);
        }
    }
    isGroupConnectedSite() {
        return this.groupId !== '00000000-0000-0000-0000-000000000000';
    }
    /**
     * Maps the base sharingCapability enum to string array so it can
     * more easily be used in validation or descriptions.
     */
    get sharingCapabilities() {
        const result = [];
        for (const sharingCapability in SharingCapabilities) {
            if (typeof SharingCapabilities[sharingCapability] === 'number') {
                result.push(sharingCapability);
            }
        }
        return result;
    }
}
_SpoSiteSetCommand_instances = new WeakSet(), _SpoSiteSetCommand_initTelemetry = function _SpoSiteSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            classification: typeof args.options.classification !== 'undefined',
            disableFlows: !!args.options.disableFlows,
            socialBarOnSitePagesDisabled: !!args.options.socialBarOnSitePagesDisabled,
            isPublic: !!args.options.isPublic,
            owners: typeof args.options.owners !== 'undefined',
            shareByEmailEnabled: !!args.options.shareByEmailEnabled,
            title: typeof args.options.title !== 'undefined',
            description: typeof args.options.description !== 'undefined',
            siteDesignId: typeof args.options.siteDesignId !== 'undefined',
            sharingCapabilities: typeof args.options.sharingCapability !== 'undefined',
            siteLogoUrl: typeof args.options.siteLogoUrl !== 'undefined',
            siteThumbnailUrl: typeof args.options.siteThumbnailUrl !== 'undefined',
            resourceQuota: typeof args.options.resourceQuota !== 'undefined',
            resourceQuotaWarningLevel: typeof args.options.resourceQuotaWarningLevel !== 'undefined',
            storageQuota: typeof args.options.storageQuota !== 'undefined',
            storageQuotaWarningLevel: typeof args.options.storageQuotaWarningLevel !== 'undefined',
            allowSelfServiceUpgrade: !!args.options.allowSelfServiceUpgrade,
            lockState: typeof args.options.lockState !== 'undefined',
            noScriptSite: !!args.options.noScriptSite,
            wait: !!args.options.wait
        });
    });
}, _SpoSiteSetCommand_initOptions = function _SpoSiteSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --url <url>'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '--classification [classification]'
    }, {
        option: '--disableFlows [disableFlows]',
        autocomplete: ['true', 'false']
    }, {
        option: '--socialBarOnSitePagesDisabled [socialBarOnSitePagesDisabled]',
        autocomplete: ['true', 'false']
    }, {
        option: '--isPublic [isPublic]',
        autocomplete: ['true', 'false']
    }, {
        option: '--owners [owners]'
    }, {
        option: '--shareByEmailEnabled [shareByEmailEnabled]',
        autocomplete: ['true', 'false']
    }, {
        option: '--siteDesignId [siteDesignId]'
    }, {
        option: '--sharingCapability [sharingCapability]',
        autocomplete: this.sharingCapabilities
    }, {
        option: '--siteLogoUrl [siteLogoUrl]'
    }, {
        option: '--siteThumbnailUrl [siteThumbnailUrl]'
    }, {
        option: '--resourceQuota [resourceQuota]'
    }, {
        option: '--resourceQuotaWarningLevel [resourceQuotaWarningLevel]'
    }, {
        option: '--storageQuota [storageQuota]'
    }, {
        option: '--storageQuotaWarningLevel [storageQuotaWarningLevel]'
    }, {
        option: '--allowSelfServiceUpgrade [allowSelfServiceUpgrade]',
        autocomplete: ['true', 'false']
    }, {
        option: '--lockState [lockState]',
        autocomplete: ['Unlock', 'NoAdditions', 'ReadOnly', 'NoAccess']
    }, {
        option: '--noScriptSite [noScriptSite]',
        autocomplete: ['true', 'false']
    }, {
        option: '--wait'
    });
}, _SpoSiteSetCommand_initValidators = function _SpoSiteSetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.url);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (typeof args.options.classification === 'undefined' &&
            typeof args.options.disableFlows === 'undefined' &&
            typeof args.options.socialBarOnSitePagesDisabled === 'undefined' &&
            typeof args.options.title === 'undefined' &&
            typeof args.options.description === 'undefined' &&
            typeof args.options.isPublic === 'undefined' &&
            typeof args.options.owners === 'undefined' &&
            typeof args.options.shareByEmailEnabled === 'undefined' &&
            typeof args.options.siteDesignId === 'undefined' &&
            typeof args.options.sharingCapability === 'undefined' &&
            typeof args.options.siteLogoUrl === 'undefined' &&
            typeof args.options.siteThumbnailUrl === 'undefined' &&
            typeof args.options.resourceQuota === 'undefined' &&
            typeof args.options.resourceQuotaWarningLevel === 'undefined' &&
            typeof args.options.storageQuota === 'undefined' &&
            typeof args.options.storageQuotaWarningLevel === 'undefined' &&
            typeof args.options.noScriptSite === 'undefined' &&
            typeof args.options.allowSelfServiceUpgrade === 'undefined' &&
            typeof args.options.lockState === 'undefined') {
            return 'Specify at least one property to update';
        }
        if (typeof args.options.siteLogoUrl !== 'undefined' && typeof args.options.siteLogoUrl !== 'string') {
            return `${args.options.siteLogoUrl} is not a valid value for the siteLogoUrl option. Specify the logo URL or an empty string "" to unset the logo.`;
        }
        if (typeof args.options.siteThumbnailUrl !== 'undefined' && typeof args.options.siteThumbnailUrl !== 'string') {
            return `${args.options.siteThumbnailUrl} is not a valid value for the siteThumbnailUrl option. Specify the logo URL or an empty string "" to unset the logo.`;
        }
        if (args.options.siteDesignId) {
            if (!validation.isValidGuid(args.options.siteDesignId)) {
                return `${args.options.siteDesignId} is not a valid GUID`;
            }
        }
        if (args.options.sharingCapability &&
            this.sharingCapabilities.indexOf(args.options.sharingCapability) < 0) {
            return `${args.options.sharingCapability} is not a valid value for the sharingCapability option. Allowed values are ${this.sharingCapabilities.join('|')}`;
        }
        if (args.options.resourceQuota &&
            typeof args.options.resourceQuota !== 'number') {
            return `${args.options.resourceQuota} is not a number`;
        }
        if (args.options.resourceQuotaWarningLevel &&
            typeof args.options.resourceQuotaWarningLevel !== 'number') {
            return `${args.options.resourceQuotaWarningLevel} is not a number`;
        }
        if (args.options.resourceQuota &&
            args.options.resourceQuotaWarningLevel &&
            args.options.resourceQuotaWarningLevel > args.options.resourceQuota) {
            return `resourceQuotaWarningLevel must not exceed the resourceQuota`;
        }
        if (args.options.storageQuota &&
            typeof args.options.storageQuota !== 'number') {
            return `${args.options.storageQuota} is not a number`;
        }
        if (args.options.storageQuotaWarningLevel &&
            typeof args.options.storageQuotaWarningLevel !== 'number') {
            return `${args.options.storageQuotaWarningLevel} is not a number`;
        }
        if (args.options.storageQuota &&
            args.options.storageQuotaWarningLevel &&
            args.options.storageQuotaWarningLevel > args.options.storageQuota) {
            return `storageQuotaWarningLevel must not exceed the storageQuota`;
        }
        if (args.options.lockState &&
            ['Unlock', 'NoAdditions', 'ReadOnly', 'NoAccess'].indexOf(args.options.lockState) === -1) {
            return `${args.options.lockState} is not a valid value for the lockState option. Allowed values Unlock|NoAdditions|ReadOnly|NoAccess`;
        }
        return true;
    });
}, _SpoSiteSetCommand_initTypes = function _SpoSiteSetCommand_initTypes() {
    this.types.string.push('classification');
    this.types.boolean.push('isPublic', 'disableFlows', 'socialBarOnSitePagesDisabled', 'shareByEmailEnabled', 'allowSelfServiceUpgrade', 'noScriptSite');
};
export default new SpoSiteSetCommand();
//# sourceMappingURL=site-set.js.map