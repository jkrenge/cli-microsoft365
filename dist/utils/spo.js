import os from 'os';
import { urlUtil } from "./urlUtil.js";
import { validation } from "./validation.js";
import auth from '../Auth.js';
import config from "../config.js";
import { BasePermissions } from '../m365/spo/base-permissions.js';
import request from "../request.js";
import { formatting } from './formatting.js';
import { odata } from './odata.js';
import { RoleType } from '../m365/spo/commands/roledefinition/RoleType.js';
import { entraGroup } from './entraGroup.js';
import { SharingCapabilities } from '../m365/spo/commands/site/SharingCapabilities.js';
import { timersUtil } from './timersUtil.js';
import { DOMParser } from '@xmldom/xmldom';
export var CreateFileCopyJobsNameConflictBehavior;
(function (CreateFileCopyJobsNameConflictBehavior) {
    CreateFileCopyJobsNameConflictBehavior[CreateFileCopyJobsNameConflictBehavior["Fail"] = 0] = "Fail";
    CreateFileCopyJobsNameConflictBehavior[CreateFileCopyJobsNameConflictBehavior["Replace"] = 1] = "Replace";
    CreateFileCopyJobsNameConflictBehavior[CreateFileCopyJobsNameConflictBehavior["Rename"] = 2] = "Rename";
})(CreateFileCopyJobsNameConflictBehavior || (CreateFileCopyJobsNameConflictBehavior = {}));
export var CreateFolderCopyJobsNameConflictBehavior;
(function (CreateFolderCopyJobsNameConflictBehavior) {
    CreateFolderCopyJobsNameConflictBehavior[CreateFolderCopyJobsNameConflictBehavior["Fail"] = 0] = "Fail";
    CreateFolderCopyJobsNameConflictBehavior[CreateFolderCopyJobsNameConflictBehavior["Rename"] = 2] = "Rename";
})(CreateFolderCopyJobsNameConflictBehavior || (CreateFolderCopyJobsNameConflictBehavior = {}));
// Wrapping this into a settings object so we can alter the values in tests
const pollingInterval = 3000;
export const spo = {
    async getRequestDigest(siteUrl) {
        const requestOptions = {
            url: `${siteUrl}/_api/contextinfo`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        return request.post(requestOptions);
    },
    async ensureFormDigest(siteUrl, logger, context, debug) {
        if (validation.isValidFormDigest(context)) {
            if (debug) {
                await logger.logToStderr('Existing form digest still valid');
            }
            return context;
        }
        const res = await spo.getRequestDigest(siteUrl);
        const now = new Date();
        now.setSeconds(now.getSeconds() + res.FormDigestTimeoutSeconds - 5);
        context = {
            FormDigestValue: res.FormDigestValue,
            FormDigestTimeoutSeconds: res.FormDigestTimeoutSeconds,
            FormDigestExpiresAt: now,
            WebFullUrl: res.WebFullUrl
        };
        return context;
    },
    async waitUntilFinished({ operationId, siteUrl, logger, currentContext, debug, verbose }) {
        const resFormDigest = await spo.ensureFormDigest(siteUrl, logger, currentContext, debug);
        currentContext = resFormDigest;
        if (debug) {
            await logger.logToStderr(`Checking if operation ${operationId} completed...`);
        }
        const requestOptions = {
            url: `${siteUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': currentContext.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="${operationId.replace(/\\n/g, '&#xA;').replace(/"/g, '')}" /></ObjectPaths></Request>`
        };
        const res = await request.post(requestOptions);
        const json = JSON.parse(res);
        const response = json[0];
        if (response.ErrorInfo) {
            throw new Error(response.ErrorInfo.ErrorMessage);
        }
        else {
            const operation = json[json.length - 1];
            const isComplete = operation.IsComplete;
            if (isComplete) {
                if (!debug && verbose) {
                    process.stdout.write('\n');
                }
                return;
            }
            await timersUtil.setTimeout(pollingInterval);
            await spo.waitUntilFinished({
                operationId: JSON.stringify(operation._ObjectIdentity_),
                siteUrl,
                logger,
                currentContext,
                debug,
                verbose
            });
        }
    },
    async getSpoUrl(logger, debug) {
        if (auth.connection.spoUrl) {
            if (debug) {
                await logger.logToStderr(`SPO URL previously retrieved ${auth.connection.spoUrl}. Returning...`);
            }
            return auth.connection.spoUrl;
        }
        if (debug) {
            await logger.logToStderr(`No SPO URL available. Retrieving from MS Graph...`);
        }
        const requestOptions = {
            url: `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`,
            headers: {
                'accept': 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const res = await request.get(requestOptions);
        auth.connection.spoUrl = res.webUrl;
        try {
            await auth.storeConnectionInfo();
        }
        catch {
            if (debug) {
                await logger.logToStderr('Error while storing connection info');
            }
        }
        return auth.connection.spoUrl;
    },
    async getSpoAdminUrl(logger, debug) {
        const spoUrl = await spo.getSpoUrl(logger, debug);
        return (spoUrl.replace(/(https:\/\/)([^.]+)(.*)/, '$1$2-admin$3'));
    },
    async getTenantId(logger, debug) {
        if (auth.connection.spoTenantId) {
            if (debug) {
                await logger.logToStderr(`SPO Tenant ID previously retrieved ${auth.connection.spoTenantId}. Returning...`);
            }
            return auth.connection.spoTenantId;
        }
        if (debug) {
            await logger.logToStderr(`No SPO Tenant ID available. Retrieving...`);
        }
        const spoAdminUrl = await spo.getSpoAdminUrl(logger, debug);
        const contextInfo = await spo.getRequestDigest(spoAdminUrl);
        const tenantInfoRequestOptions = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': contextInfo.FormDigestValue,
                accept: 'application/json;odata=nometadata'
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
        };
        const res = await request.post(tenantInfoRequestOptions);
        const json = JSON.parse(res);
        auth.connection.spoTenantId = json[json.length - 1]._ObjectIdentity_.replace('\n', '&#xA;');
        try {
            await auth.storeConnectionInfo();
        }
        catch {
            if (debug) {
                await logger.logToStderr('Error while storing connection info');
            }
        }
        return auth.connection.spoTenantId;
    },
    /**
     * Returns the Graph id of a site
     * @param webUrl web url e.g. https://contoso.sharepoint.com/sites/site1
     */
    async getSpoGraphSiteId(webUrl) {
        const url = new URL(webUrl);
        const requestOptions = {
            url: `https://graph.microsoft.com/v1.0/sites/${url.hostname}:${url.pathname}?$select=id`,
            headers: {
                'accept': 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const result = await request.get(requestOptions);
        return result.id;
    },
    /**
     * Ensures the folder path exists
     * @param webFullUrl web full url e.g. https://contoso.sharepoint.com/sites/site1
     * @param folderToEnsure web relative or server relative folder path e.g. /Documents/MyFolder or /sites/site1/Documents/MyFolder
     * @param siteAccessToken a valid access token for the site specified in the webFullUrl param
     */
    async ensureFolder(webFullUrl, folderToEnsure, logger, debug) {
        try {
            new URL(webFullUrl);
        }
        catch {
            throw new Error('webFullUrl is not a valid URL');
        }
        if (!folderToEnsure) {
            throw new Error('folderToEnsure cannot be empty');
        }
        // remove last '/' of webFullUrl if exists
        const webFullUrlLastCharPos = webFullUrl.length - 1;
        if (webFullUrl.length > 1 &&
            webFullUrl[webFullUrlLastCharPos] === '/') {
            webFullUrl = webFullUrl.substring(0, webFullUrlLastCharPos);
        }
        folderToEnsure = urlUtil.getWebRelativePath(webFullUrl, folderToEnsure);
        if (debug) {
            await logger.log(`folderToEnsure`);
            await logger.log(folderToEnsure);
            await logger.log('');
        }
        let nextFolder = '';
        let prevFolder = '';
        let folderIndex = 0;
        // build array of folders e.g. ["Shared%20Documents","22","54","55"]
        const folders = folderToEnsure.substring(1).split('/');
        if (debug) {
            await logger.log('folders to process');
            await logger.log(JSON.stringify(folders));
            await logger.log('');
        }
        // recursive function
        async function checkOrAddFolder() {
            if (folderIndex === folders.length) {
                if (debug) {
                    await logger.log(`All sub-folders exist`);
                }
                return;
            }
            // append the next sub-folder to the folder path and check if it exists
            prevFolder = nextFolder;
            nextFolder += `/${folders[folderIndex]}`;
            const folderServerRelativeUrl = urlUtil.getServerRelativePath(webFullUrl, nextFolder);
            const requestOptions = {
                url: `${webFullUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folderServerRelativeUrl)}')`,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                }
            };
            try {
                await request.get(requestOptions);
                folderIndex++;
                await checkOrAddFolder();
            }
            catch {
                const prevFolderServerRelativeUrl = urlUtil.getServerRelativePath(webFullUrl, prevFolder);
                const requestOptions = {
                    url: `${webFullUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27${formatting.encodeQueryParameter(prevFolderServerRelativeUrl)}%27&@a2=%27${formatting.encodeQueryParameter(folders[folderIndex])}%27`,
                    headers: {
                        'accept': 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                try {
                    await request.post(requestOptions);
                    folderIndex++;
                    await checkOrAddFolder();
                }
                catch (err) {
                    if (debug) {
                        await logger.log(`Could not create sub-folder ${folderServerRelativeUrl}`);
                    }
                    throw err;
                }
            }
        }
        ;
        return checkOrAddFolder();
    },
    /**
     * Requests web object identity for the current web.
     * That request is something similar to _contextinfo in REST.
     * The response data looks like:
     * _ObjectIdentity_=<GUID>|<GUID>:site:<GUID>:web:<GUID>
     * _ObjectType_=SP.Web
     * ServerRelativeUrl=/sites/contoso
     * @param webUrl web url
     * @param formDigestValue formDigestValue
     */
    async getCurrentWebIdentity(webUrl, formDigestValue) {
        const requestOptions = {
            url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': formDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="1" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`
        };
        const res = await request.post(requestOptions);
        const json = JSON.parse(res);
        const contents = json.find(x => { return x.ErrorInfo; });
        if (contents && contents.ErrorInfo) {
            throw contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error';
        }
        const identityObject = json.find(x => { return x._ObjectIdentity_; });
        if (identityObject) {
            return {
                objectIdentity: identityObject._ObjectIdentity_,
                serverRelativeUrl: identityObject.ServerRelativeUrl
            };
        }
        throw 'Cannot proceed. _ObjectIdentity_ not found';
    },
    /**
     * Gets EffectiveBasePermissions for web return type is "_ObjectType_\":\"SP.Web\".
     * @param webObjectIdentity ObjectIdentity. Has format _ObjectIdentity_=<GUID>|<GUID>:site:<GUID>:web:<GUID>
     * @param webUrl web url
     * @param siteAccessToken site access token
     * @param formDigestValue formDigestValue
     */
    async getEffectiveBasePermissions(webObjectIdentity, webUrl, formDigestValue, logger, debug) {
        const basePermissionsResult = new BasePermissions();
        const requestOptions = {
            url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': formDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="11" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="EffectiveBasePermissions" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="${webObjectIdentity}" /></ObjectPaths></Request>`
        };
        const res = await request.post(requestOptions);
        if (debug) {
            await logger.log('Attempt to get the web EffectiveBasePermissions');
        }
        const json = JSON.parse(res);
        const contents = json.find(x => { return x.ErrorInfo; });
        if (contents && contents.ErrorInfo) {
            throw contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error';
        }
        const permissionsObj = json.find(x => { return x.EffectiveBasePermissions; });
        if (permissionsObj) {
            basePermissionsResult.high = permissionsObj.EffectiveBasePermissions.High;
            basePermissionsResult.low = permissionsObj.EffectiveBasePermissions.Low;
            return basePermissionsResult;
        }
        throw ('Cannot proceed. EffectiveBasePermissions not found'); // this is not supposed to happen
    },
    /**
      * Gets folder by server relative url (GetFolderByServerRelativeUrl in REST)
      * The response data looks like:
      * _ObjectIdentity_=<GUID>|<GUID>:site:<GUID>:web:<GUID>:folder:<GUID>
      * _ObjectType_=SP.Folder
      * @param webObjectIdentity ObjectIdentity. Has format _ObjectIdentity_=<GUID>|<GUID>:site:<GUID>:web:<GUID>
      * @param webUrl web url
      * @param siteRelativeUrl site relative url e.g. /Shared Documents/Folder1
      * @param formDigestValue formDigestValue
      */
    async getFolderIdentity(webObjectIdentity, webUrl, siteRelativeUrl, formDigestValue) {
        const serverRelativePath = urlUtil.getServerRelativePath(webUrl, siteRelativeUrl);
        const requestOptions = {
            url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': formDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><Query Id="12" ObjectPathId="9"><Query SelectAllProperties="false"><Properties><Property Name="Properties" SelectAll="true"><Query SelectAllProperties="false"><Properties /></Query></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="9" ParentId="5" Name="GetFolderByServerRelativeUrl"><Parameters><Parameter Type="String">${serverRelativePath}</Parameter></Parameters></Method><Identity Id="5" Name="${webObjectIdentity}" /></ObjectPaths></Request>`
        };
        const res = await request.post(requestOptions);
        const json = JSON.parse(res);
        const contents = json.find(x => { return x.ErrorInfo; });
        if (contents && contents.ErrorInfo) {
            throw contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error';
        }
        const objectIdentity = json.find(x => { return x._ObjectIdentity_; });
        if (objectIdentity) {
            return {
                objectIdentity: objectIdentity._ObjectIdentity_,
                serverRelativeUrl: serverRelativePath
            };
        }
        throw 'Cannot proceed. Folder _ObjectIdentity_ not found';
    },
    /**
     * Retrieves the SiteId, VroomItemId and VroomDriveId from a specific file.
     * @param webUrl Web url
     * @param fileId GUID ID of the file
     * @param fileUrl Decoded site-relative or server-relative URL of the file
     */
    async getVroomFileDetails(webUrl, fileId, fileUrl) {
        let requestUrl = `${webUrl}/_api/web/`;
        if (fileUrl) {
            const fileServerRelativeUrl = urlUtil.getServerRelativePath(webUrl, fileUrl);
            requestUrl += `GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(fileServerRelativeUrl)}')`;
        }
        else {
            requestUrl += `GetFileById('${fileId}')`;
        }
        requestUrl += '?$select=SiteId,VroomItemId,VroomDriveId';
        const requestOptions = {
            url: requestUrl,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const res = await request.get(requestOptions);
        return res;
    },
    /**
     * Retrieves a list of Custom Actions from a SharePoint site.
     * @param webUrl Web url
     * @param scope The scope of custom actions to retrieve, allowed values "Site", "Web" or "All".
     * @param filter An OData filter query to limit the results.
     */
    async getCustomActions(webUrl, scope, filter) {
        if (scope && scope !== "All" && scope !== "Site" && scope !== "Web") {
            throw `Invalid scope '${scope}'. Allowed values are 'Site', 'Web' or 'All'.`;
        }
        const queryString = filter ? `?$filter=${filter}` : "";
        if (scope && scope !== "All") {
            return await odata.getAllItems(`${webUrl}/_api/${scope}/UserCustomActions${queryString}`);
        }
        const customActions = [
            ...await odata.getAllItems(`${webUrl}/_api/Site/UserCustomActions${queryString}`),
            ...await odata.getAllItems(`${webUrl}/_api/Web/UserCustomActions${queryString}`)
        ];
        return customActions;
    },
    /**
     * Retrieves a Custom Actions from a SharePoint site by Id.
     * @param webUrl Web url
     * @param id The Id of the Custom Action
     * @param scope The scope of custom actions to retrieve, allowed values "Site", "Web" or "All".
     */
    async getCustomActionById(webUrl, id, scope) {
        if (scope && scope !== "All" && scope !== "Site" && scope !== "Web") {
            throw `Invalid scope '${scope}'. Allowed values are 'Site', 'Web' or 'All'.`;
        }
        async function getById(webUrl, id, scope) {
            const requestOptions = {
                url: `${webUrl}/_api/${scope}/UserCustomActions(guid'${id}')`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const result = await request.get(requestOptions);
            if (result["odata.null"] === true) {
                return undefined;
            }
            return result;
        }
        if (scope && scope !== "All") {
            return await getById(webUrl, id, scope);
        }
        const customActionOnWeb = await getById(webUrl, id, "Web");
        if (customActionOnWeb) {
            return customActionOnWeb;
        }
        const customActionOnSite = await getById(webUrl, id, "Site");
        return customActionOnSite;
    },
    async getTenantAppCatalogUrl(logger, debug) {
        const spoUrl = await spo.getSpoUrl(logger, debug);
        const requestOptions = {
            url: `${spoUrl}/_api/SP_TenantSettings_Current`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const result = await request.get(requestOptions);
        return result.CorporateCatalogUrl;
    },
    /**
     * Retrieves the Microsoft Entra ID from a SP user.
     * @param webUrl Web url
     * @param id The Id of the user
     */
    async getUserAzureIdBySpoId(webUrl, id) {
        const requestOptions = {
            url: `${webUrl}/_api/web/siteusers/GetById('${formatting.encodeQueryParameter(id)}')?$select=AadObjectId`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const res = await request.get(requestOptions);
        return res.AadObjectId.NameId;
    },
    /**
     * Ensure a user exists on a specific SharePoint site.
     * @param webUrl URL of the SharePoint site.
     * @param logonName Logon name of the user to ensure on the SharePoint site.
     * @returns SharePoint user object.
     */
    async ensureUser(webUrl, logonName) {
        const requestOptions = {
            url: `${webUrl}/_api/web/EnsureUser`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json',
            data: {
                logonName: logonName
            }
        };
        return request.post(requestOptions);
    },
    /**
     * Ensure a Microsoft Entra ID group exists on a specific SharePoint site.
     * @param webUrl URL of the SharePoint site.
     * @param group Microsoft Entra ID group.
     * @returns SharePoint user object.
     */
    async ensureEntraGroup(webUrl, group) {
        if (!group.securityEnabled) {
            throw new Error('Cannot ensure a Microsoft Entra ID group that is not security enabled.');
        }
        return this.ensureUser(webUrl, group.mailEnabled ? `c:0o.c|federateddirectoryclaimprovider|${group.id}` : `c:0t.c|tenant|${group.id}`);
    },
    /**
   * Retrieves the spo user by email.
   * Returns a user object
   * @param webUrl Web url
   * @param email The email of the user
   * @param logger the Logger object
   * @param verbose Set for verbose logging
   */
    async getUserByEmail(webUrl, email, logger, verbose) {
        if (verbose && logger) {
            await logger.logToStderr(`Retrieving the spo user by email ${email}`);
        }
        const requestUrl = `${webUrl}/_api/web/siteusers/GetByEmail('${formatting.encodeQueryParameter(email)}')`;
        const requestOptions = {
            url: requestUrl,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const userInstance = await request.get(requestOptions);
        return userInstance;
    },
    /**
     * Retrieves the menu state for the quick launch.
     * @param webUrl Web url
     */
    async getQuickLaunchMenuState(webUrl) {
        return this.getMenuState(webUrl);
    },
    /**
     * Retrieves the menu state for the top navigation.
     * @param webUrl Web url
     */
    async getTopNavigationMenuState(webUrl) {
        return this.getMenuState(webUrl, '1002');
    },
    /**
     * Retrieves the menu state.
     * @param webUrl Web url
     * @param menuNodeKey Menu node key
     */
    async getMenuState(webUrl, menuNodeKey) {
        const requestBody = {
            customProperties: null,
            depth: 10,
            mapProviderName: null,
            menuNodeKey: menuNodeKey || null
        };
        const requestOptions = {
            url: `${webUrl}/_api/navigation/MenuState`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            data: requestBody,
            responseType: 'json'
        };
        return request.post(requestOptions);
    },
    /**
    * Saves the menu state.
    * @param webUrl Web url
    * @param menuState Updated menu state
    */
    async saveMenuState(webUrl, menuState) {
        const requestOptions = {
            url: `${webUrl}/_api/navigation/SaveMenuState`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            data: { menuState: menuState },
            responseType: 'json'
        };
        return request.post(requestOptions);
    },
    /**
    * Retrieves the spo group by name.
    * Returns a group object
    * @param webUrl Web url
    * @param name The name of the group
    * @param logger the Logger object
    * @param verbose Set for verbose logging
    */
    async getGroupByName(webUrl, name, logger, verbose) {
        if (verbose && logger) {
            await logger.logToStderr(`Retrieving the group by name ${name}`);
        }
        const requestUrl = `${webUrl}/_api/web/sitegroups/GetByName('${formatting.encodeQueryParameter(name)}')`;
        const requestOptions = {
            url: requestUrl,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const groupInstance = await request.get(requestOptions);
        return groupInstance;
    },
    /**
    * Retrieves the role definition by name.
    * Returns a RoleDefinition object
    * Returns a RoleDefinition object
    * @param webUrl Web url
    * @param name the name of the role definition
    * @param logger the Logger object
    * @param verbose Set for verbose logging
    * @param verbose Set for verbose logging
    */
    async getRoleDefinitionByName(webUrl, name, logger, verbose) {
        if (verbose && logger) {
            await logger.logToStderr(`Retrieving the role definitions for ${name}`);
        }
        const roledefinitions = await odata.getAllItems(`${webUrl}/_api/web/roledefinitions`);
        const roledefinition = roledefinitions.find((role) => role.Name === name);
        if (!roledefinition) {
            throw `No roledefinition is found for ${name}`;
        }
        const permissions = new BasePermissions();
        permissions.high = roledefinition.BasePermissions.High;
        permissions.low = roledefinition.BasePermissions.Low;
        roledefinition.BasePermissionsValue = permissions.parse();
        roledefinition.RoleTypeKindValue = RoleType[roledefinition.RoleTypeKind];
        return roledefinition;
    },
    /**
    * Adds a SharePoint site.
    * @param type Type of sites to add. Allowed values TeamSite, CommunicationSite, ClassicSite, default TeamSite
    * @param title Site title
    * @param alias Site alias, used in the URL and in the team site group e-mail (applies to type TeamSite)
    * @param url Site URL (applies to type CommunicationSite, ClassicSite)
    * @param timeZone Integer representing time zone to use for the site (applies to type ClassicSite)
    * @param description Site description
    * @param lcid Site language in the LCID format
    * @param owners Comma-separated list of users to set as site owners (applies to type TeamSite, ClassicSite)
    * @param isPublic Determines if the associated group is public or not (applies to type TeamSite)
    * @param classification Site classification (applies to type TeamSite, CommunicationSite)
    * @param siteDesignType of communication site to create. Allowed values Topic, Showcase, Blank, default Topic. When creating a communication site, specify either siteDesign or siteDesignId (applies to type CommunicationSite)
    * @param siteDesignId Id of the custom site design to use to create the site. When creating a communication site, specify either siteDesign or siteDesignId (applies to type CommunicationSite)
    * @param shareByEmailEnabled Determines whether it's allowed to share file with guests (applies to type CommunicationSite)
    * @param webTemplate Template to use for creating the site. Default `STS#0` (applies to type ClassicSite)
    * @param resourceQuota The quota for this site collection in Sandboxed Solutions units. Default 0 (applies to type ClassicSite)
    * @param resourceQuotaWarningLevel The warning level for the resource quota. Default 0 (applies to type ClassicSite)
    * @param storageQuota The storage quota for this site collection in megabytes. Default 100 (applies to type ClassicSite)
    * @param storageQuotaWarningLevel The warning level for the storage quota in megabytes. Default 100 (applies to type ClassicSite)
    * @param removeDeletedSite Set, to remove existing deleted site with the same URL from the Recycle Bin (applies to type ClassicSite)
    * @param wait Wait for the site to be provisioned before completing the command (applies to type ClassicSite)
    * @param logger the Logger object
    * @param verbose set if verbose logging should be logged
    */
    async addSite(title, logger, verbose, wait, type, alias, description, owners, shareByEmailEnabled, removeDeletedSite, classification, isPublic, lcid, url, siteDesign, siteDesignId, timeZone, webTemplate, resourceQuota, resourceQuotaWarningLevel, storageQuota, storageQuotaWarningLevel) {
        if (type === 'ClassicSite') {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, verbose);
            let context = await spo.ensureFormDigest(spoAdminUrl, logger, undefined, verbose);
            let exists;
            if (removeDeletedSite) {
                exists = await spo.siteExists(url, logger, verbose);
            }
            else {
                // assume site doesn't exist
                exists = false;
            }
            if (exists) {
                if (verbose) {
                    await logger.logToStderr('Site exists in the recycle bin');
                }
                await spo.deleteSiteFromTheRecycleBin(url, logger, verbose, wait);
            }
            else {
                if (verbose) {
                    await logger.logToStderr('Site not found');
                }
            }
            context = await spo.ensureFormDigest(spoAdminUrl, logger, context, verbose);
            if (verbose) {
                await logger.logToStderr(`Creating site collection ${url}...`);
            }
            const lcidOption = typeof lcid === 'number' ? lcid : 1033;
            const storageQuotaOption = typeof storageQuota === 'number' ? storageQuota : 100;
            const storageQuotaWarningLevelOption = typeof storageQuotaWarningLevel === 'number' ? storageQuotaWarningLevel : 100;
            const resourceQuotaOption = typeof resourceQuota === 'number' ? resourceQuota : 0;
            const resourceQuotaWarningLevelOption = typeof resourceQuotaWarningLevel === 'number' ? resourceQuotaWarningLevel : 0;
            const webTemplateOption = webTemplate || 'STS#0';
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': context.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">${lcidOption}</Property><Property Name="Owner" Type="String">${formatting.escapeXml(owners)}</Property><Property Name="StorageMaximumLevel" Type="Int64">${storageQuotaOption}</Property><Property Name="StorageWarningLevel" Type="Int64">${storageQuotaWarningLevelOption}</Property><Property Name="Template" Type="String">${formatting.escapeXml(webTemplateOption)}</Property><Property Name="TimeZoneId" Type="Int32">${timeZone}</Property><Property Name="Title" Type="String">${formatting.escapeXml(title)}</Property><Property Name="Url" Type="String">${formatting.escapeXml(url)}</Property><Property Name="UserCodeMaximumLevel" Type="Double">${resourceQuotaOption}</Property><Property Name="UserCodeWarningLevel" Type="Double">${resourceQuotaWarningLevelOption}</Property></Parameter></Parameters></Method></ObjectPaths></Request>`
            };
            const res = await request.post(requestOptions);
            const json = JSON.parse(res);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            else {
                const operation = json[json.length - 1];
                const isComplete = operation.IsComplete;
                if (!wait || isComplete) {
                    return;
                }
                await timersUtil.setTimeout(pollingInterval);
                await spo.waitUntilFinished({
                    operationId: JSON.stringify(operation._ObjectIdentity_),
                    siteUrl: spoAdminUrl,
                    logger,
                    currentContext: context,
                    verbose: verbose,
                    debug: verbose
                });
            }
        }
        else {
            const isTeamSite = type !== 'CommunicationSite';
            const spoUrl = await spo.getSpoUrl(logger, verbose);
            if (verbose) {
                await logger.logToStderr(`Creating new site...`);
            }
            let requestOptions;
            if (isTeamSite) {
                requestOptions = {
                    url: `${spoUrl}/_api/GroupSiteManager/CreateGroupEx`,
                    headers: {
                        'content-type': 'application/json; odata=verbose; charset=utf-8',
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json',
                    data: {
                        displayName: title,
                        alias: alias,
                        isPublic: isPublic,
                        optionalParams: {
                            Description: description || '',
                            CreationOptions: {
                                results: [],
                                Classification: classification || ''
                            }
                        }
                    }
                };
                if (lcid) {
                    requestOptions.data.optionalParams.CreationOptions.results.push(`SPSiteLanguage:${lcid}`);
                }
                if (owners) {
                    requestOptions.data.optionalParams.Owners = {
                        results: owners.split(',').map(o => o.trim())
                    };
                }
            }
            else {
                if (!siteDesignId) {
                    if (siteDesign) {
                        switch (siteDesign) {
                            case 'Topic':
                                siteDesignId = '00000000-0000-0000-0000-000000000000';
                                break;
                            case 'Showcase':
                                siteDesignId = '6142d2a0-63a5-4ba0-aede-d9fefca2c767';
                                break;
                            case 'Blank':
                                siteDesignId = 'f6cc5403-0d63-442e-96c0-285923709ffc';
                                break;
                        }
                    }
                    else {
                        siteDesignId = '00000000-0000-0000-0000-000000000000';
                    }
                }
                requestOptions = {
                    url: `${spoUrl}/_api/SPSiteManager/Create`,
                    headers: {
                        'content-type': 'application/json;odata=nometadata',
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json',
                    data: {
                        request: {
                            Title: title,
                            Url: url,
                            ShareByEmailEnabled: shareByEmailEnabled,
                            Description: description || '',
                            Classification: classification || '',
                            WebTemplate: 'SITEPAGEPUBLISHING#0',
                            SiteDesignId: siteDesignId
                        }
                    }
                };
                if (lcid) {
                    requestOptions.data.request.Lcid = lcid;
                }
                if (owners) {
                    requestOptions.data.request.Owner = owners;
                }
            }
            const res = await request.post(requestOptions);
            if (isTeamSite) {
                if (res.ErrorMessage !== null) {
                    throw res.ErrorMessage;
                }
                else {
                    return res.SiteUrl;
                }
            }
            else {
                if (res.SiteStatus === 2) {
                    return res.SiteUrl;
                }
                else {
                    throw 'An error has occurred while creating the site';
                }
            }
        }
    },
    /**
    * Checks if a site exists
    * Returns a boolean
    * @param url The url of the site
    * @param logger the Logger object
    * @param verbose set if verbose logging should be logged
    */
    async siteExists(url, logger, verbose) {
        const spoAdminUrl = await spo.getSpoAdminUrl(logger, verbose);
        const context = await spo.ensureFormDigest(spoAdminUrl, logger, undefined, verbose);
        if (verbose) {
            await logger.logToStderr(`Checking if the site ${url} exists...`);
        }
        const requestOptions = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': context.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`
        };
        const res1 = await request.post(requestOptions);
        const json = JSON.parse(res1);
        const response = json[0];
        if (response.ErrorInfo) {
            if (response.ErrorInfo.ErrorTypeName === 'Microsoft.Online.SharePoint.Common.SpoNoSiteException') {
                return await this.siteExistsInTheRecycleBin(url, logger, verbose);
            }
            else {
                throw response.ErrorInfo.ErrorMessage;
            }
        }
        else {
            const site = json[json.length - 1];
            return site.Status === 'Recycled';
        }
    },
    /**
    * Checks if a site exists in the recycle bin
    * Returns a boolean
    * @param url The url of the site
    * @param logger the Logger object
    * @param verbose set if verbose logging should be logged
    */
    async siteExistsInTheRecycleBin(url, logger, verbose) {
        if (verbose) {
            await logger.logToStderr(`Site doesn't exist. Checking if the site ${url} exists in the recycle bin...`);
        }
        const spoAdminUrl = await spo.getSpoAdminUrl(logger, verbose);
        const context = await spo.ensureFormDigest(spoAdminUrl, logger, undefined, verbose);
        const requestOptions = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': context.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
        };
        const res = await request.post(requestOptions);
        const json = JSON.parse(res);
        const response = json[0];
        if (response.ErrorInfo) {
            if (response.ErrorInfo.ErrorTypeName === 'Microsoft.SharePoint.Client.UnknownError') {
                return false;
            }
            throw response.ErrorInfo.ErrorMessage;
        }
        const site = json[json.length - 1];
        return site.Status === 'Recycled';
    },
    /**
    * Deletes a site from the recycle bin
    * @param url The url of the site
    * @param logger the Logger object
    * @param verbose set if verbose logging should be logged
    * @param wait set to wait until finished
    */
    async deleteSiteFromTheRecycleBin(url, logger, verbose, wait) {
        const spoAdminUrl = await spo.getSpoAdminUrl(logger, verbose);
        const context = await spo.ensureFormDigest(spoAdminUrl, logger, undefined, verbose);
        if (verbose) {
            await logger.logToStderr(`Deleting site ${url} from the recycle bin...`);
        }
        const requestOptions = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': context.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
        };
        const response = await request.post(requestOptions);
        const json = JSON.parse(response);
        const responseContent = json[0];
        if (responseContent.ErrorInfo) {
            throw responseContent.ErrorInfo.ErrorMessage;
        }
        const operation = json[json.length - 1];
        const isComplete = operation.IsComplete;
        if (!wait || isComplete) {
            return;
        }
        await timersUtil.setTimeout(pollingInterval);
        await spo.waitUntilFinished({
            operationId: JSON.stringify(operation._ObjectIdentity_),
            siteUrl: spoAdminUrl,
            logger,
            currentContext: context,
            verbose: verbose,
            debug: verbose
        });
    },
    /**
     * Updates a site with the given properties
     * @param url The url of the site
     * @param logger The logger object
     * @param verbose Set for verbose logging
     * @param title The new title
     * @param classification The classification to be updated
     * @param disableFlows If flows should be disabled or not
     * @param isPublic If site should be public or private
     * @param owners The owners to be updated
     * @param shareByEmailEnabled If share by e-mail should be enabled or not
     * @param siteDesignId The site design to be updated
     * @param sharingCapability The sharing capability to be updated
     */
    async updateSite(url, logger, verbose, title, classification, disableFlows, isPublic, owners, shareByEmailEnabled, siteDesignId, sharingCapability) {
        const tenantId = await spo.getTenantId(logger, verbose);
        const spoAdminUrl = await spo.getSpoAdminUrl(logger, verbose);
        let context = await spo.ensureFormDigest(spoAdminUrl, logger, undefined, verbose);
        if (verbose) {
            await logger.logToStderr('Loading site IDs...');
        }
        const requestOptions = {
            url: `${url}/_api/site?$select=GroupId,Id`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const siteInfo = await request.get(requestOptions);
        const groupId = siteInfo.GroupId;
        const siteId = siteInfo.Id;
        const isGroupConnectedSite = groupId !== '00000000-0000-0000-0000-000000000000';
        if (verbose) {
            await logger.logToStderr(`Retrieved site IDs. siteId: ${siteId}, groupId: ${groupId}`);
        }
        if (isGroupConnectedSite) {
            if (verbose) {
                await logger.logToStderr(`Site attached to group ${groupId}`);
            }
            if (typeof title !== 'undefined' &&
                typeof isPublic !== 'undefined' &&
                typeof owners !== 'undefined') {
                const promises = [];
                if (typeof title !== 'undefined') {
                    const requestOptions = {
                        url: `${spoAdminUrl}/_api/SPOGroup/UpdateGroupPropertiesBySiteId`,
                        headers: {
                            accept: 'application/json;odata=nometadata',
                            'content-type': 'application/json;charset=utf-8',
                            'X-RequestDigest': context.FormDigestValue
                        },
                        data: {
                            groupId: groupId,
                            siteId: siteId,
                            displayName: title
                        },
                        responseType: 'json'
                    };
                    promises.push(request.post(requestOptions));
                }
                if (typeof isPublic !== 'undefined') {
                    promises.push(entraGroup.setGroup(groupId, (isPublic === false), undefined, undefined, logger, verbose));
                }
                if (typeof owners !== 'undefined') {
                    promises.push(spo.setGroupifiedSiteOwners(spoAdminUrl, groupId, owners, logger, verbose));
                }
                await Promise.all(promises);
            }
        }
        else {
            if (verbose) {
                await logger.logToStderr('Site is not group connected');
            }
            if (typeof isPublic !== 'undefined') {
                throw `The isPublic option can't be set on a site that is not groupified`;
            }
            if (owners) {
                await Promise.all(owners.split(',').map(async (o) => {
                    await spo.setSiteAdmin(spoAdminUrl, context, url, o.trim());
                }));
            }
        }
        context = await spo.ensureFormDigest(spoAdminUrl, logger, context, verbose);
        if (verbose) {
            await logger.logToStderr(`Updating site ${url} properties...`);
        }
        let updatedProperties = false;
        if (!isGroupConnectedSite) {
            if (title !== undefined) {
                updatedProperties = true;
            }
        }
        else {
            if (classification !== undefined || disableFlows !== undefined || shareByEmailEnabled !== undefined || sharingCapability !== undefined) {
                updatedProperties = true;
            }
        }
        if (updatedProperties) {
            let propertyId = 27;
            const payload = [];
            if (!isGroupConnectedSite) {
                if (title) {
                    payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="Title"><Parameter Type="String">${formatting.escapeXml(title)}</Parameter></SetProperty>`);
                }
            }
            if (typeof classification === 'string') {
                payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="Classification"><Parameter Type="String">${formatting.escapeXml(classification)}</Parameter></SetProperty>`);
            }
            if (typeof disableFlows !== 'undefined') {
                payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="DisableFlows"><Parameter Type="Boolean">${disableFlows}</Parameter></SetProperty>`);
            }
            if (typeof shareByEmailEnabled !== 'undefined') {
                payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="ShareByEmailEnabled"><Parameter Type="Boolean">${shareByEmailEnabled}</Parameter></SetProperty>`);
            }
            if (sharingCapability) {
                const sharingCapabilityOption = SharingCapabilities[sharingCapability];
                payload.push(`<SetProperty Id="${propertyId}" ObjectPathId="5" Name="SharingCapability"><Parameter Type="Enum">${sharingCapabilityOption}</Parameter></SetProperty>`);
            }
            const pos = tenantId.indexOf('|') + 1;
            const requestOptionsUpdateProperties = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': context.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${payload.join('')}<ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|${tenantId.substring(pos, tenantId.indexOf('&'))}&#xA;SiteProperties&#xA;${formatting.encodeQueryParameter(url)}" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`
            };
            const res = await request.post(requestOptionsUpdateProperties);
            const json = JSON.parse(res);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            else {
                const operation = json[json.length - 1];
                const isComplete = operation.IsComplete;
                if (!isComplete) {
                    await timersUtil.setTimeout(pollingInterval);
                    await spo.waitUntilFinished({
                        operationId: JSON.stringify(operation._ObjectIdentity_),
                        siteUrl: spoAdminUrl,
                        logger,
                        currentContext: context,
                        verbose: verbose,
                        debug: verbose
                    });
                }
            }
        }
        if (siteDesignId) {
            await spo.applySiteDesign(siteDesignId, url, logger, verbose);
        }
    },
    /**
     * Updates the groupified site owners
     * @param spoAdminUrl The SharePoint admin url
     * @param groupId The ID of the group
     * @param owners The owners to be updated
     * @param logger The logger object
     * @param verbose Set for verbose logging
     */
    async setGroupifiedSiteOwners(spoAdminUrl, groupId, owners, logger, verbose) {
        const splittedOwners = owners.split(',').map(o => o.trim());
        if (verbose) {
            await logger.logToStderr('Retrieving user information to set group owners...');
        }
        const requestOptions = {
            url: `https://graph.microsoft.com/v1.0/users?$filter=${splittedOwners.map(o => `userPrincipalName eq '${o}'`).join(' or ')}&$select=id`,
            headers: {
                'content-type': 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const res = await request.get(requestOptions);
        if (res.value.length === 0) {
            return;
        }
        await Promise.all(res.value.map(user => {
            const requestOptions = {
                url: `${spoAdminUrl}/_api/SP.Directory.DirectorySession/Group('${groupId}')/Owners/Add(objectId='${user.id}', principalName='')`,
                headers: {
                    'content-type': 'application/json;odata=verbose'
                }
            };
            return request.post(requestOptions);
        }));
    },
    /**
     * Updates the site admin
     * @param spoAdminUrl The SharePoint admin url
     * @param context The FormDigestInfo
     * @param siteUrl The url of the site where the owners has to be updated
     * @param principal The principal of the admin
     * @param logger The logger object
     * @param verbose Set for verbose logging
     */
    async setSiteAdmin(spoAdminUrl, context, siteUrl, principal, logger, verbose) {
        if (verbose && logger) {
            await logger.logToStderr(`Updating site admin ${principal} on ${siteUrl}...`);
        }
        const requestOptions = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': context.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Method Id="47" ParentId="34" Name="SetSiteAdmin"><Parameters><Parameter Type="String">${formatting.escapeXml(siteUrl)}</Parameter><Parameter Type="String">${formatting.escapeXml(principal)}</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Constructor Id="34" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
        };
        const res = await request.post(requestOptions);
        const json = JSON.parse(res);
        const response = json[0];
        if (response.ErrorInfo) {
            throw response.ErrorInfo.ErrorMessage;
        }
        else {
            return;
        }
    },
    /**
     * Applies a site design
     * @param id The ID of the site design (GUID)
     * @param webUrl The url of the site where the design should be applied
     * @param logger The logger object
     * @param verbose Set for verbose logging
     */
    async applySiteDesign(id, webUrl, logger, verbose) {
        if (verbose) {
            await logger.logToStderr(`Applying site design ${id}...`);
        }
        const spoUrl = await spo.getSpoUrl(logger, verbose);
        const requestOptions = {
            url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.ApplySiteDesign`,
            headers: {
                'content-type': 'application/json;charset=utf-8',
                accept: 'application/json;odata=nometadata'
            },
            data: {
                siteDesignId: id,
                webUrl: webUrl
            },
            responseType: 'json'
        };
        await request.post(requestOptions);
    },
    /**
     * Gets the web properties for a given url
     * @param url The url of the web
     * @param logger The logger object
     * @param verbose Set for verbose logging
     */
    async getWeb(url, logger, verbose) {
        if (verbose && logger) {
            await logger.logToStderr(`Getting the web properties for ${url}`);
        }
        const requestOptions = {
            url: `${url}/_api/web`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const webProperties = await request.get(requestOptions);
        return webProperties;
    },
    /**
     * Applies the retention label to the items in the given list.
     * @param webUrl The url of the web
     * @param name The name of the label
     * @param listAbsoluteUrl The absolute Url to the list
     * @param itemIds The list item Ids to apply the label to. (A maximum 100 is allowed)
     * @param logger The logger object
     * @param verbose Set for verbose logging
     */
    async applyRetentionLabelToListItems(webUrl, name, listAbsoluteUrl, itemIds, logger, verbose) {
        if (verbose && logger) {
            await logger.logToStderr(`Applying retention label '${name}' to item(s) in list '${listAbsoluteUrl}'...`);
        }
        const requestOptions = {
            url: `${webUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            data: {
                listUrl: listAbsoluteUrl,
                complianceTagValue: name,
                itemIds: itemIds
            },
            responseType: 'json'
        };
        await request.post(requestOptions);
    },
    /**
    * Gets a file as list item by url
    * @param absoluteListUrl The absolute url to the list
    * @param url The url of the file
    * @param logger The logger object
    * @param verbose If in verbose mode
    * @returns The list item object
    */
    async getFileAsListItemByUrl(absoluteListUrl, url, logger, verbose) {
        if (verbose && logger) {
            await logger.logToStderr(`Getting the file properties with url ${url}`);
        }
        const serverRelativePath = urlUtil.getServerRelativePath(absoluteListUrl, url);
        const requestUrl = `${absoluteListUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl=@f)?$expand=ListItemAllFields&@f='${formatting.encodeQueryParameter(serverRelativePath)}'`;
        const requestOptions = {
            url: requestUrl,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const file = await request.get(requestOptions);
        return file.ListItemAllFields;
    },
    /**
    * Updates a list item with system update
    * @param absoluteListUrl The absolute base URL without query parameters, pointing to the specific list where the item resides. This URL should represent the list.
    * @param itemId The id of the list item
    * @param properties An object of the properties that should be updated
    * @param contentTypeName The name of the content type to update
    * @param logger The logger object
    * @param verbose If in verbose mode
    * @returns The updated list item object
    */
    async systemUpdateListItem(absoluteListUrl, itemId, logger, verbose, properties, contentTypeName) {
        if (!properties && !contentTypeName) {
            // Neither properties nor contentTypeName provided, no need to proceed
            throw 'Either properties or contentTypeName must be provided for systemUpdateListItem.';
        }
        const parsedUrl = new URL(absoluteListUrl);
        const serverRelativeSiteMatch = absoluteListUrl.match(new RegExp('/sites/[^/]+'));
        const webUrl = `${parsedUrl.protocol}//${parsedUrl.host}${serverRelativeSiteMatch ?? ''}`;
        if (verbose && logger) {
            await logger.logToStderr(`Getting list id...`);
        }
        const listRequestOptions = {
            url: `${absoluteListUrl}?$select=Id`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const list = await request.get(listRequestOptions);
        const listId = list.Id;
        if (verbose && logger) {
            await logger.logToStderr(`Getting request digest for systemUpdate request`);
        }
        const res = await spo.getRequestDigest(webUrl);
        const formDigestValue = res.FormDigestValue;
        const objectIdentity = await spo.requestObjectIdentity(webUrl, logger, verbose);
        let index = 0;
        const requestBodyOptions = properties ? Object.keys(properties).map(key => `
    <Method Name="ParseAndSetFieldValue" Id="${++index}" ObjectPathId="147">
      <Parameters>
        <Parameter Type="String">${key}</Parameter>
        <Parameter Type="String">${properties[key].toString()}</Parameter>
      </Parameters>
    </Method>`) : [];
        const additionalContentType = contentTypeName ? `
    <Method Name="ParseAndSetFieldValue" Id="${++index}" ObjectPathId="147">
      <Parameters>
        <Parameter Type="String">ContentType</Parameter>
        <Parameter Type="String">${contentTypeName}</Parameter>
      </Parameters>
    </Method>` : '';
        const requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
      <Actions>
        ${requestBodyOptions.join('')}${additionalContentType}
        <Method Name="SystemUpdate" Id="${++index}" ObjectPathId="147" />
      </Actions>
      <ObjectPaths>
        <Identity Id="147" Name="${objectIdentity}:list:${listId}:item:${itemId},1" />
      </ObjectPaths>
    </Request>`;
        const requestOptions = {
            url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'Content-Type': 'text/xml',
                'X-RequestDigest': formDigestValue
            },
            data: requestBody
        };
        const response = await request.post(requestOptions);
        if (response.indexOf("ErrorMessage") > -1) {
            throw `Error occurred in systemUpdate operation - ${response}`;
        }
        const requestOptionsItems = {
            url: `${absoluteListUrl}/items(${itemId})`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const itemsResponse = await request.get(requestOptionsItems);
        return (itemsResponse);
    },
    /**
     * Removes the retention label from the items in the given list.
     * @param webUrl The url of the web
     * @param listAbsoluteUrl The absolute Url to the list
     * @param itemIds The list item Ids to clear the label from (A maximum 100 is allowed)
     * @param logger The logger object
     * @param verbose Set for verbose logging
     */
    async removeRetentionLabelFromListItems(webUrl, listAbsoluteUrl, itemIds, logger, verbose) {
        if (verbose && logger) {
            await logger.logToStderr(`Removing retention label from item(s) in list '${listAbsoluteUrl}'...`);
        }
        const requestOptions = {
            url: `${webUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            data: {
                listUrl: listAbsoluteUrl,
                complianceTagValue: '',
                itemIds: itemIds
            },
            responseType: 'json'
        };
        await request.post(requestOptions);
    },
    /**
     * Applies a default retention label to a list or library.
     * @param webUrl The url of the web
     * @param name The name of the label
     * @param listAbsoluteUrl The absolute Url to the list
     * @param syncToItems If the label needs to be synced to existing items/files.
     * @param logger The logger object
     * @param verbose Set for verbose logging
     */
    async applyDefaultRetentionLabelToList(webUrl, name, listAbsoluteUrl, syncToItems, logger, verbose) {
        if (verbose && logger) {
            await logger.logToStderr(`Applying default retention label '${name}' to the list '${listAbsoluteUrl}'...`);
        }
        const requestOptions = {
            url: `${webUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            data: {
                listUrl: listAbsoluteUrl,
                complianceTagValue: name,
                blockDelete: false,
                blockEdit: false,
                syncToItems: syncToItems || false
            },
            responseType: 'json'
        };
        await request.post(requestOptions);
    },
    /**
     * Removes the default retention label from a list or library.
     * @param webUrl The url of the web
     * @param listAbsoluteUrl The absolute Url to the list
     * @param logger The logger object
     * @param verbose Set for verbose logging
     */
    async removeDefaultRetentionLabelFromList(webUrl, listAbsoluteUrl, logger, verbose) {
        if (verbose && logger) {
            await logger.logToStderr(`Removing the default retention label from the list '${listAbsoluteUrl}'...`);
        }
        const requestOptions = {
            url: `${webUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            data: {
                listUrl: listAbsoluteUrl,
                complianceTagValue: '',
                blockDelete: false,
                blockEdit: false,
                syncToItems: false
            },
            responseType: 'json'
        };
        await request.post(requestOptions);
    },
    /**
     * Retrieves the site ID for a given web URL by the MS Graph.
     * @param webUrl The web URL for which to retrieve the site ID.
     * @param logger The logger object.
     * @param verbose Set for verbose logging
     * @returns The site ID as a string.
     */
    async getSiteIdByMSGraph(webUrl, logger, verbose) {
        if (verbose && logger) {
            await logger.logToStderr(`Getting site id for URL: ${webUrl}...`);
        }
        const url = new URL(webUrl);
        const requestOptions = {
            url: `https://graph.microsoft.com/v1.0/sites/${formatting.encodeQueryParameter(url.host)}:${url.pathname}?$select=id`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const site = await request.get(requestOptions);
        return site.id;
    },
    /**
     * Retrieves the SharePoint Online site ID for the specified web URL by the SharePoint REST API.
     * @param webUrl The web URL of the SharePoint Online site to retrieve the site ID for.
     * @param logger The logger object.
     * @param verbose Set for verbose logging
     * @returns The site ID as a string.
     */
    async getSiteIdBySPApi(webUrl, logger, verbose) {
        if (verbose && logger) {
            await logger.logToStderr(`Getting site id for URL: ${webUrl}...`);
        }
        const requestOptions = {
            url: `${webUrl}/_api/site?$select=Id`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const siteResponse = await request.get(requestOptions);
        return siteResponse.Id;
    },
    /**
     * Retrieves the id of a SharePoint web for the specified web URL.
     * @param webUrl The web URL for which to retrieve the web ID.
     * @param logger The logger object for logging messages.
     * @param verbose Set to true for verbose logging.
     * @returns The web ID as a string.
     */
    async getWebId(webUrl, logger, verbose) {
        if (verbose && logger) {
            await logger.logToStderr(`Getting web id for URL: ${webUrl}...`);
        }
        const requestOptions = {
            url: `${webUrl}/_api/web?$select=Id`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const webResponse = await request.get(requestOptions);
        return webResponse.Id;
    },
    /**
     * Retrieves the ID of a SharePoint list by its title or URL.
     * @param webUrl The base URL of the SharePoint site.
     * @param listTitle The title of the list (optional).
     * @param listUrl The server-relative URL of the list (optional).
     * @param logger The logger object for logging messages (optional).
     * @param verbose Set to true for verbose logging (optional).
     * @returns The list ID as a string.
     */
    async getListId(webUrl, listTitle, listUrl, logger, verbose) {
        if (verbose && logger) {
            await logger.logToStderr(`Retrieving list id...`);
        }
        if (!listTitle && !listUrl) {
            throw new Error('Either listTitle or listUrl must be provided.');
        }
        let listId = '';
        if (listTitle) {
            const requestOptions = {
                url: `${webUrl}/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')?$select=Id`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const listResponse = await request.get(requestOptions);
            listId = listResponse.Id;
        }
        else if (listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(webUrl, listUrl);
            const requestOptions = {
                url: `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')?$select=Id`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const listResponse = await request.get(requestOptions);
            listId = listResponse.Id;
        }
        return listId;
    },
    /**
     * Retrieves the server-relative URL of a folder.
     * @param webUrl Web URL
     * @param folderUrl Folder URL
     * @param folderId Folder ID
     * @param logger The logger object
     * @param verbose Set for verbose logging
     * @returns The server-relative URL of the folder
     */
    async getFolderServerRelativeUrl(webUrl, folderUrl, folderId, logger, verbose) {
        if (verbose && logger) {
            await logger.logToStderr(`Retrieving server-relative URL for folder ${folderUrl ? `URL: ${folderUrl}` : `ID: ${folderId}`}`);
        }
        let requestUrl = `${webUrl}/_api/web/`;
        if (folderUrl) {
            const folderServerRelativeUrl = urlUtil.getServerRelativePath(webUrl, folderUrl);
            requestUrl += `GetFolderByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(folderServerRelativeUrl)}')`;
        }
        else {
            requestUrl += `GetFolderById('${folderId}')`;
        }
        requestUrl += '?$select=ServerRelativeUrl';
        const requestOptions = {
            url: requestUrl,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const res = await request.get(requestOptions);
        return res.ServerRelativeUrl;
    },
    /**
     * Retrieves the ObjectIdentity from a SharePoint site
     * @param webUrl web url
     * @param logger The logger object
     * @param verbose If in verbose mode
     * @return The ObjectIdentity as string
     */
    async requestObjectIdentity(webUrl, logger, verbose) {
        const res = await spo.getRequestDigest(webUrl);
        const formDigestValue = res.FormDigestValue;
        const requestOptions = {
            url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': formDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="1" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`
        };
        const response = await request.post(requestOptions);
        if (verbose) {
            await logger.logToStderr('Attempt to get _ObjectIdentity_ key values');
        }
        const json = JSON.parse(response);
        const contents = json.find(x => { return x.ErrorInfo; });
        if (contents && contents.ErrorInfo) {
            throw contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error';
        }
        const identityObject = json.find(x => { return x._ObjectIdentity_; });
        if (identityObject) {
            return identityObject._ObjectIdentity_;
        }
        throw 'Cannot proceed. _ObjectIdentity_ not found'; // this is not supposed to happen
    },
    /**
    * Updates a list item without system update
    * @param absoluteListUrl The absolute base URL without query parameters, pointing to the specific list where the item resides. This URL should represent the list.
    * @param itemId The id of the list item
    * @param properties An object of the properties that should be updated
    * @param contentTypeName The name of the content type to update
    * @returns The updated listitem object
    */
    async updateListItem(absoluteListUrl, itemId, properties, contentTypeName) {
        const requestBodyOptions = [
            ...(properties
                ? Object.keys(properties).map((key) => ({
                    FieldName: key,
                    FieldValue: properties[key].toString()
                }))
                : [])
        ];
        const requestBody = {
            formValues: requestBodyOptions
        };
        if (contentTypeName) {
            requestBody.formValues.push({
                FieldName: 'ContentType',
                FieldValue: contentTypeName
            });
        }
        const requestOptions = {
            url: `${absoluteListUrl}/items(${itemId})/ValidateUpdateListItem()`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            data: requestBody,
            responseType: 'json'
        };
        const response = await request.post(requestOptions);
        // Response is from /ValidateUpdateListItem POST call, perform get on updated item to get all field values
        const fieldValues = response.value;
        if (fieldValues.some(f => f.HasException)) {
            throw `Updating the items has failed with the following errors: ${os.EOL}${fieldValues.filter(f => f.HasException).map(f => { return `- ${f.FieldName} - ${f.ErrorMessage}`; }).join(os.EOL)}`;
        }
        const requestOptionsItems = {
            url: `${absoluteListUrl}/items(${itemId})`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const itemsResponse = await request.get(requestOptionsItems);
        return (itemsResponse);
    },
    /**
    * Retrieves the file by id.
    * Returns a FileProperties object
    * @param webUrl Web url
    * @param id the id of the file
    * @param logger the Logger object
    * @param verbose set for verbose logging
    */
    async getFileById(webUrl, id, logger, verbose) {
        if (verbose && logger) {
            await logger.logToStderr(`Retrieving the file with id ${id}`);
        }
        const requestUrl = `${webUrl}/_api/web/GetFileById('${formatting.encodeQueryParameter(id)}')`;
        const requestOptions = {
            url: requestUrl,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const file = await request.get(requestOptions);
        return file;
    },
    /**
     * Create a SharePoint copy job to copy a file to another location.
     * @param webUrl Absolute web URL where the source file is located.
     * @param sourceUrl Absolute URL of the source file.
     * @param destinationUrl Absolute URL of the destination folder.
     * @param options Options for the copy job.
     * @returns Copy job information. Use {@link spo.getCopyJobResult} to get the result of the copy job.
     */
    async createFileCopyJob(webUrl, sourceUrl, destinationUrl, options) {
        const requestOptions = {
            url: `${webUrl}/_api/Site/CreateCopyJobs`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json',
            data: {
                destinationUri: destinationUrl,
                exportObjectUris: [sourceUrl],
                options: {
                    NameConflictBehavior: options?.nameConflictBehavior ?? CreateFileCopyJobsNameConflictBehavior.Fail,
                    AllowSchemaMismatch: true,
                    BypassSharedLock: !!options?.bypassSharedLock,
                    IgnoreVersionHistory: !!options?.ignoreVersionHistory,
                    IncludeItemPermissions: !!options?.includeItemPermissions,
                    CustomizedItemName: options?.newName ? [options.newName] : undefined,
                    SameWebCopyMoveOptimization: true,
                    IsMoveMode: options?.operation === 'move'
                }
            }
        };
        const response = await request.post(requestOptions);
        return response.value[0];
    },
    /**
     * Create a SharePoint copy job to copy a folder to another location.
     * @param webUrl Absolute web URL where the source folder is located.
     * @param sourceUrl Absolute URL of the source folder.
     * @param destinationUrl Absolute URL of the destination folder.
     * @param options Options for the copy job.
     * @returns Copy job information. Use {@link spo.getCopyJobResult} to get the result of the copy job.
     */
    async createFolderCopyJob(webUrl, sourceUrl, destinationUrl, options) {
        const requestOptions = {
            url: `${webUrl}/_api/Site/CreateCopyJobs`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json',
            data: {
                destinationUri: destinationUrl,
                exportObjectUris: [sourceUrl],
                options: {
                    NameConflictBehavior: options?.nameConflictBehavior ?? CreateFolderCopyJobsNameConflictBehavior.Fail,
                    AllowSchemaMismatch: true,
                    CustomizedItemName: options?.newName ? [options.newName] : undefined,
                    SameWebCopyMoveOptimization: true,
                    IsMoveMode: options?.operation === 'move'
                }
            }
        };
        const response = await request.post(requestOptions);
        return response.value[0];
    },
    /**
     * Poll until the copy job is finished and return the result.
     * @param webUrl Absolute web URL where the copy job was created.
     * @param copyJobInfo Information about the copy job.
     * @throws Error if the copy job has failed.
     * @returns Information about the destination object.
     */
    async getCopyJobResult(webUrl, copyJobInfo) {
        const requestOptions = {
            url: `${webUrl}/_api/Site/GetCopyJobProgress`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json',
            data: {
                copyJobInfo: copyJobInfo
            }
        };
        const logs = [];
        let progress = await request.post(requestOptions);
        const newLogs = progress.Logs?.map(l => JSON.parse(l));
        if (newLogs?.length > 0) {
            logs.push(...newLogs);
        }
        while (progress.JobState !== 0) {
            await timersUtil.setTimeout(pollingInterval);
            progress = await request.post(requestOptions);
            const newLogs = progress.Logs?.map(l => JSON.parse(l));
            if (newLogs?.length > 0) {
                logs.push(...newLogs);
            }
        }
        // Check if the job has failed
        const errorLog = logs.find(l => l.Event === 'JobError');
        if (errorLog) {
            throw new Error(errorLog.Message);
        }
        // Get the destination object information
        let objectInfo = logs.find(l => l.Event === 'JobFinishedObjectInfo');
        // In rare cases, the object info may not be available yet
        if (!objectInfo) {
            // By doing a final poll, we can get the object info
            progress = await request.post(requestOptions);
            const newLogs = progress.Logs?.map(l => JSON.parse(l));
            objectInfo = newLogs.find(l => l.Event === 'JobFinishedObjectInfo');
        }
        return objectInfo;
    },
    /**
     * Gets the primary owner login from a site as admin.
     * @param adminUrl The SharePoint admin URL
     * @param siteId The site ID
     * @param logger The logger object
     * @param verbose If in verbose mode
     * @returns Owner login name
     */
    async getPrimaryAdminLoginNameAsAdmin(adminUrl, siteId, logger, verbose) {
        if (verbose) {
            await logger.logToStderr('Getting the primary admin login name...');
        }
        const requestOptions = {
            url: `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        return response.OwnerLoginName;
    },
    /**
     * Gets the primary owner login from a site.
     * @param siteUrl The site URL
     * @param logger The logger object
     * @param verbose If in verbose mode
     * @returns Owner login name
     */
    async getPrimaryOwnerLoginFromSite(siteUrl, logger, verbose) {
        if (verbose) {
            await logger.logToStderr('Getting the primary admin login name...');
        }
        const requestOptions = {
            url: `${siteUrl}/_api/site/owner`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const responseContent = await request.get(requestOptions);
        return responseContent?.LoginName;
    },
    /**
    * Retrieves the site admin properties for a given site URL.
    * @param adminUrl The SharePoint admin url.
    * @param siteUrl URL of the site for which to retrieve properties.
    * @param includeDetail Set to true to include detailed properties.
    * @param logger The logger object.
    * @param verbose Set for verbose logging.
    * @returns Tenant Site properties.
    */
    async getSiteAdminPropertiesByUrl(siteUrl, includeDetail, logger, verbose) {
        if (verbose) {
            await logger.logToStderr(`Getting site admin properties for URL: ${siteUrl}...`);
        }
        const adminUrl = await spo.getSpoAdminUrl(logger, !!verbose);
        const requestOptions = {
            url: `${adminUrl}/_api/SPO.Tenant/GetSitePropertiesByUrl`,
            headers: {
                accept: 'application/json;odata=nometadata',
                'content-type': 'application/json;charset=utf-8'
            },
            data: {
                url: siteUrl,
                includeDetail: includeDetail
            },
            responseType: 'json'
        };
        return request.post(requestOptions);
    },
    /**
    * Get a role definition by name
    * Returns a RoleDefinition object
    * @param webUrl The web url
    * @param name  the name of the role definition
    * @param logger The logger object
    * @param verbose Set for verbose logging
    */
    async getRoleDefintionByName(webUrl, name, logger, verbose) {
        if (verbose && logger) {
            await logger.logToStderr(`Retrieving the role definition by name ${name}`);
        }
        const response = await odata.getAllItems(`${webUrl}/_api/web/roledefinitions`);
        const roleDefinition = response.find((role) => role.Name === name);
        if (!roleDefinition) {
            throw new Error(`The specified role definition name '${name}' does not exist.`);
        }
        const permissions = new BasePermissions();
        permissions.high = roleDefinition.BasePermissions.High;
        permissions.low = roleDefinition.BasePermissions.Low;
        roleDefinition.BasePermissionsValue = permissions.parse();
        roleDefinition.RoleTypeKindValue = RoleType[roleDefinition.RoleTypeKind];
        return roleDefinition;
    },
    /**
     * Converts SharePoint default column values XML to a structured JSON array.
     * Parses the XML content from client_LocationBasedDefaults.html file and extracts
     * default field values configured for folders in a SharePoint document library.
     *
     * @param xml The XML string containing default column value definitions
     * @returns An array of default column values with field names, values, and associated folder URLs
     */
    convertDefaultColumnXmlToJson(xml) {
        const results = [];
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml, 'application/xml');
        const folderLinks = doc.getElementsByTagName('a');
        for (let i = 0; i < folderLinks.length; i++) {
            const folderUrl = folderLinks[i].getAttribute('href');
            const defaultValues = folderLinks[i].getElementsByTagName('DefaultValue');
            for (let j = 0; j < defaultValues.length; j++) {
                const fieldName = defaultValues[j].getAttribute('FieldName');
                const fieldValue = defaultValues[j].textContent;
                results.push({
                    fieldName: fieldName,
                    fieldValue: fieldValue,
                    folderUrl: decodeURIComponent(folderUrl)
                });
            }
        }
        return results;
    }
};
//# sourceMappingURL=spo.js.map