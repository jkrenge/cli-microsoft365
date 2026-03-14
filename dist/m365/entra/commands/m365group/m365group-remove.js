var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraM365GroupRemoveCommand_instances, _EntraM365GroupRemoveCommand_initTelemetry, _EntraM365GroupRemoveCommand_initOptions, _EntraM365GroupRemoveCommand_initValidators, _EntraM365GroupRemoveCommand_initOptionSets, _EntraM365GroupRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import config from '../../../../config.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { setTimeout } from 'timers/promises';
class EntraM365GroupRemoveCommand extends GraphCommand {
    get name() {
        return commands.M365GROUP_REMOVE;
    }
    get description() {
        return 'Removes a Microsoft 365 Group';
    }
    constructor() {
        super();
        _EntraM365GroupRemoveCommand_instances.add(this);
        this.intervalInMs = 5000;
        __classPrivateFieldGet(this, _EntraM365GroupRemoveCommand_instances, "m", _EntraM365GroupRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupRemoveCommand_instances, "m", _EntraM365GroupRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupRemoveCommand_instances, "m", _EntraM365GroupRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupRemoveCommand_instances, "m", _EntraM365GroupRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupRemoveCommand_instances, "m", _EntraM365GroupRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const removeGroup = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Removing Microsoft 365 Group: ${args.options.id || args.options.displayName}...`);
            }
            try {
                let groupId = args.options.id;
                if (args.options.displayName) {
                    groupId = await entraGroup.getGroupIdByDisplayName(args.options.displayName);
                }
                const isUnifiedGroup = await entraGroup.isUnifiedGroup(groupId);
                if (!isUnifiedGroup) {
                    throw Error(`Specified group with id '${groupId}' is not a Microsoft 365 group.`);
                }
                const siteUrl = await this.getM365GroupSiteUrl(logger, groupId);
                const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
                // Delete the Microsoft 365 group site. This operation will also delete the group.
                await this.deleteM365GroupSite(logger, siteUrl, spoAdminUrl);
                if (args.options.skipRecycleBin) {
                    await this.deleteM365GroupFromRecycleBin(logger, groupId);
                    await this.deleteSiteFromRecycleBin(logger, siteUrl, spoAdminUrl);
                }
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeGroup();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the group ${args.options.id || args.options.displayName}?` });
            if (result) {
                await removeGroup();
            }
        }
    }
    async getM365GroupSiteUrl(logger, id) {
        if (this.verbose) {
            await logger.logToStderr(`Getting the site URL of Microsoft 365 Group: ${id}...`);
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/groups/${id}/drive?$select=webUrl`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const res = await request.get(requestOptions);
        // Extract the base URL by removing everything after the last '/' character in the URL.
        const baseUrl = res.webUrl.substring(0, res.webUrl.lastIndexOf('/'));
        return baseUrl;
    }
    async deleteM365GroupSite(logger, url, spoAdminUrl) {
        if (this.verbose) {
            await logger.logToStderr(`Deleting the group site: '${url}'...`);
        }
        const requestOptions = {
            url: `${spoAdminUrl}/_api/GroupSiteManager/Delete?siteUrl='${formatting.encodeQueryParameter(url)}'`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        await request.post(requestOptions);
    }
    async deleteM365GroupFromRecycleBin(logger, id) {
        for (let retries = 0; retries < EntraM365GroupRemoveCommand.maxRetries; retries++) {
            if (await this.isM365GroupInDeletedItemsList(id)) {
                await this.removeM365GroupPermanently(logger, id);
                return;
            }
            else {
                if (this.verbose) {
                    await logger.logToStderr(`Group has not been deleted yet. Waiting for ${this.intervalInMs / 1000} seconds before the next attempt. ${EntraM365GroupRemoveCommand.maxRetries - retries} attempts remaining...`);
                }
                await setTimeout(this.intervalInMs);
            }
        }
        await logger.logToStderr(`Group has been successfully deleted, but it couldn't be permanently removed from the recycle bin after all retries. It will still appear in the deleted groups list.`);
    }
    async removeM365GroupPermanently(logger, id) {
        if (this.verbose) {
            await logger.logToStderr(`Group has been deleted and is now available in the deleted groups list. Removing permanently...`);
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/directory/deletedItems/${id}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            }
        };
        await request.delete(requestOptions);
    }
    async isM365GroupInDeletedItemsList(id) {
        const requestOptions = {
            url: `${this.resource}/v1.0/directory/deletedItems/${id}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const response = await request.get(requestOptions);
            return Boolean(response && response.id);
        }
        catch (error) {
            if (error.response && error.response.status === 404) {
                return false;
            }
            else {
                throw error;
            }
        }
    }
    async deleteSiteFromRecycleBin(logger, url, spoAdminUrl) {
        if (this.verbose) {
            await logger.logToStderr(`Deleting the M365 group site '${url}' from the recycle bin...`);
        }
        const res = await spo.ensureFormDigest(spoAdminUrl, logger, undefined, this.debug);
        const requestOptions = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': res.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
        };
        const processQuery = await request.post(requestOptions);
        const json = JSON.parse(processQuery);
        const response = json[0];
        if (response.ErrorInfo) {
            throw response.ErrorInfo.ErrorMessage;
        }
    }
}
_EntraM365GroupRemoveCommand_instances = new WeakSet(), _EntraM365GroupRemoveCommand_initTelemetry = function _EntraM365GroupRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: (!(!args.options.force)).toString(),
            skipRecycleBin: args.options.skipRecycleBin
        });
    });
}, _EntraM365GroupRemoveCommand_initOptions = function _EntraM365GroupRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --displayName [displayName]'
    }, {
        option: '-f, --force'
    }, {
        option: '--skipRecycleBin'
    });
}, _EntraM365GroupRemoveCommand_initValidators = function _EntraM365GroupRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
}, _EntraM365GroupRemoveCommand_initOptionSets = function _EntraM365GroupRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'displayName'] });
}, _EntraM365GroupRemoveCommand_initTypes = function _EntraM365GroupRemoveCommand_initTypes() {
    this.types.string.push('id', 'displayName');
};
EntraM365GroupRemoveCommand.maxRetries = 10;
export default new EntraM365GroupRemoveCommand();
//# sourceMappingURL=m365group-remove.js.map