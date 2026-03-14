var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoContentTypeAddCommand_instances, _SpoContentTypeAddCommand_initTelemetry, _SpoContentTypeAddCommand_initOptions, _SpoContentTypeAddCommand_initValidators, _SpoContentTypeAddCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoContentTypeGetCommand from './contenttype-get.js';
class SpoContentTypeAddCommand extends SpoCommand {
    get name() {
        return commands.CONTENTTYPE_ADD;
    }
    get description() {
        return 'Adds a new list or site content type';
    }
    constructor() {
        super();
        _SpoContentTypeAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoContentTypeAddCommand_instances, "m", _SpoContentTypeAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeAddCommand_instances, "m", _SpoContentTypeAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeAddCommand_instances, "m", _SpoContentTypeAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeAddCommand_instances, "m", _SpoContentTypeAddCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            let parentInfo = '';
            if (!args.options.listId && !args.options.listTitle && !args.options.listUrl) {
                parentInfo = '<Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" />';
            }
            else {
                parentInfo = await this.getParentInfo(args.options, logger);
            }
            if (this.verbose) {
                await logger.logToStderr(`Retrieving request digest...`);
            }
            const reqDigest = await spo.getRequestDigest(args.options.webUrl);
            const description = args.options.description ?
                `<Property Name="Description" Type="String">${formatting.escapeXml(args.options.description)}</Property>` :
                '<Property Name="Description" Type="Null" />';
            const group = args.options.group ?
                `<Property Name="Group" Type="String">${formatting.escapeXml(args.options.group)}</Property>` :
                '<Property Name="Group" Type="Null" />';
            const requestOptions = {
                url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="8" ObjectPathId="7" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /></Actions><ObjectPaths><Property Id="7" ParentId="5" Name="ContentTypes" /><Method Id="9" ParentId="7" Name="Add"><Parameters><Parameter TypeId="{168f3091-4554-4f14-8866-b20d48e45b54}">${description}${group}<Property Name="Id" Type="String">${formatting.escapeXml(args.options.id)}</Property><Property Name="Name" Type="String">${formatting.escapeXml(args.options.name)}</Property><Property Name="ParentContentType" Type="Null" /></Parameter></Parameters></Method>${parentInfo}</ObjectPaths></Request>`
            };
            const res = await request.post(requestOptions);
            const json = JSON.parse(res);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            const options = {
                webUrl: args.options.webUrl,
                listTitle: args.options.listTitle,
                listUrl: args.options.listUrl,
                listId: args.options.listId,
                id: args.options.id,
                output: 'json',
                debug: this.debug,
                verbose: this.verbose
            };
            try {
                const output = await cli.executeCommandWithOutput(spoContentTypeGetCommand, { options: { ...options, _: [] } });
                if (this.debug) {
                    await logger.logToStderr(output.stderr);
                }
                await logger.log(JSON.parse(output.stdout));
            }
            catch (cmdError) {
                throw cmdError.error;
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getParentInfo(options, logger) {
        const siteId = await spo.getSiteIdBySPApi(options.webUrl, logger, this.verbose);
        const webId = await spo.getWebId(options.webUrl, logger, this.verbose);
        const listId = options.listId ? options.listId : await spo.getListId(options.webUrl, options.listTitle, options.listUrl, logger, this.verbose);
        return `<Identity Id="5" Name="1a48869e-c092-0000-1f61-81ec89809537|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:list:${listId}" />`;
    }
}
_SpoContentTypeAddCommand_instances = new WeakSet(), _SpoContentTypeAddCommand_initTelemetry = function _SpoContentTypeAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listTitle: typeof args.options.listTitle !== 'undefined',
            description: typeof args.options.description !== 'undefined',
            group: typeof args.options.group !== 'undefined'
        });
    });
}, _SpoContentTypeAddCommand_initOptions = function _SpoContentTypeAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --listTitle [listTitle]'
    }, {
        option: '--listId [listId]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-i, --id <id>'
    }, {
        option: '-n, --name <name>'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '-g, --group [group]'
    });
}, _SpoContentTypeAddCommand_initValidators = function _SpoContentTypeAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.listId) {
            if (!validation.isValidGuid(args.options.listId)) {
                return `${args.options.listId} is not a valid GUID`;
            }
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoContentTypeAddCommand_initTypes = function _SpoContentTypeAddCommand_initTypes() {
    this.types.string.push('id', 'i');
};
export default new SpoContentTypeAddCommand();
//# sourceMappingURL=contenttype-add.js.map