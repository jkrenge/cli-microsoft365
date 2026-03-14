var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFieldSetCommand_instances, _SpoFieldSetCommand_initTelemetry, _SpoFieldSetCommand_initOptions, _SpoFieldSetCommand_initValidators, _SpoFieldSetCommand_initOptionSets;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFieldSetCommand extends SpoCommand {
    get name() {
        return commands.FIELD_SET;
    }
    get description() {
        return 'Updates existing list or site column';
    }
    constructor() {
        super();
        _SpoFieldSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFieldSetCommand_instances, "m", _SpoFieldSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFieldSetCommand_instances, "m", _SpoFieldSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFieldSetCommand_instances, "m", _SpoFieldSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFieldSetCommand_instances, "m", _SpoFieldSetCommand_initOptionSets).call(this);
    }
    allowUnknownOptions() {
        return true;
    }
    async commandAction(logger, args) {
        try {
            const reqDigest = await spo.getRequestDigest(args.options.webUrl);
            const requestDigest = reqDigest.FormDigestValue;
            let fieldsParentIdentity = '<Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" />';
            if (args.options.listId || args.options.listTitle || args.options.listUrl) {
                let requestData = '';
                if (args.options.listId) {
                    requestData = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="664" ObjectPathId="663" /><Query Id="665" ObjectPathId="663"><Query SelectAllProperties="false"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="663" ParentId="7" Name="GetById"><Parameters><Parameter Type="Guid">${formatting.escapeXml(args.options.listId)}</Parameter></Parameters></Method><Property Id="7" ParentId="5" Name="Lists" /><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`;
                }
                else if (args.options.listTitle) {
                    requestData = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="664" ObjectPathId="663" /><Query Id="665" ObjectPathId="663"><Query SelectAllProperties="false"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="663" ParentId="7" Name="GetByTitle"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.listTitle)}</Parameter></Parameters></Method><Property Id="7" ParentId="5" Name="Lists" /><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`;
                }
                else if (args.options.listUrl) {
                    const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
                    requestData = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="5"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticProperty Id="1" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /><Property Id="3" ParentId="1" Name="Web" /><Method Id="5" ParentId="3" Name="GetList"><Parameters><Parameter Type="String">${listServerRelativeUrl}</Parameter></Parameters></Method></ObjectPaths></Request>`;
                }
                const requestOptions = {
                    url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
                    headers: {
                        'X-RequestDigest': requestDigest
                    },
                    data: requestData
                };
                const list = await request.post(requestOptions);
                const json = JSON.parse(list);
                const response = json[0];
                if (response.ErrorInfo) {
                    throw response.ErrorInfo.ErrorMessage;
                }
                const result = json[json.length - 1];
                fieldsParentIdentity = `<Identity Id="5" Name="${result._ObjectIdentity_}" />`;
            }
            // retrieve column CSOM object id
            const fieldQuery = args.options.id ?
                `<Method Id="663" ParentId="7" Name="GetById"><Parameters><Parameter Type="Guid">${formatting.escapeXml(args.options.id)}</Parameter></Parameters></Method>` :
                `<Method Id="663" ParentId="7" Name="GetByInternalNameOrTitle"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.title || args.options.internalName)}</Parameter></Parameters></Method>`;
            let requestOptions = {
                url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': requestDigest
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="664" ObjectPathId="663" /><Query Id="665" ObjectPathId="663"><Query SelectAllProperties="false"><Properties /></Query></Query></Actions><ObjectPaths>${fieldQuery}<Property Id="7" ParentId="5" Name="Fields" />${fieldsParentIdentity}</ObjectPaths></Request>`
            };
            const field = await request.post(requestOptions);
            let json = JSON.parse(field);
            let response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            const result = json[json.length - 1];
            const fieldId = result._ObjectIdentity_;
            requestOptions = {
                url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': requestDigest
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${this.getPayload(args.options)}<Method Name="UpdateAndPushChanges" Id="9000" ObjectPathId="663"><Parameters><Parameter Type="Boolean">${args.options.updateExistingLists ? 'true' : 'false'}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="663" Name="${fieldId}" /></ObjectPaths></Request>`
            };
            const res = await request.post(requestOptions);
            json = JSON.parse(res);
            response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
    getPayload(options) {
        const excludeOptions = [
            'webUrl',
            'listId',
            'listTitle',
            'listUrl',
            'id',
            'title',
            'internalName',
            'updateExistingLists',
            'debug',
            'verbose',
            'output'
        ];
        let i = 667;
        const payload = Object.keys(options).map(key => {
            return excludeOptions.indexOf(key) === -1 ? `<SetProperty Id="${i++}" ObjectPathId="663" Name="${key}"><Parameter Type="String">${formatting.escapeXml(options[key])}</Parameter></SetProperty>` : '';
        }).join('');
        return payload;
    }
}
_SpoFieldSetCommand_instances = new WeakSet(), _SpoFieldSetCommand_initTelemetry = function _SpoFieldSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            internalName: typeof args.options.internalName !== 'undefined',
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            updateExistingLists: !!args.options.updateExistingLists
        });
    });
}, _SpoFieldSetCommand_initOptions = function _SpoFieldSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--listId [listId]'
    }, {
        option: '--listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '--internalName [internalName]'
    }, {
        option: '--updateExistingLists'
    });
}, _SpoFieldSetCommand_initValidators = function _SpoFieldSetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        const listOptions = [args.options.listId, args.options.listTitle, args.options.listUrl];
        if (listOptions.some(item => item !== undefined) && listOptions.filter(item => item !== undefined).length > 1) {
            return `Specify either list id or title or list url, but not multiple`;
        }
        if (args.options.listId &&
            !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} in option listId is not a valid GUID`;
        }
        if (args.options.id &&
            !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} in option id is not a valid GUID`;
        }
        return true;
    });
}, _SpoFieldSetCommand_initOptionSets = function _SpoFieldSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title', 'internalName'] });
};
export default new SpoFieldSetCommand();
//# sourceMappingURL=field-set.js.map