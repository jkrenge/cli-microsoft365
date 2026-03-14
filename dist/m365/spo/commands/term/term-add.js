var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTermAddCommand_instances, _SpoTermAddCommand_initTelemetry, _SpoTermAddCommand_initOptions, _SpoTermAddCommand_initValidators, _SpoTermAddCommand_initOptionSets;
import { v4 } from 'uuid';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTermAddCommand extends SpoCommand {
    get name() {
        return commands.TERM_ADD;
    }
    get description() {
        return 'Adds taxonomy term';
    }
    constructor() {
        super();
        _SpoTermAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTermAddCommand_instances, "m", _SpoTermAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoTermAddCommand_instances, "m", _SpoTermAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTermAddCommand_instances, "m", _SpoTermAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoTermAddCommand_instances, "m", _SpoTermAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let term;
        let formDigest;
        try {
            const spoWebUrl = args.options.webUrl ? args.options.webUrl : await spo.getSpoAdminUrl(logger, this.debug);
            const res = await spo.getRequestDigest(spoWebUrl);
            formDigest = res.FormDigestValue;
            if (this.verbose) {
                await logger.logToStderr(`Adding taxonomy term...`);
            }
            const termGroupQuery = args.options.termGroupId ? `<Method Id="11" ParentId="9" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termGroupId}}</Parameter></Parameters></Method>` : `<Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termGroupName)}</Parameter></Parameters></Method>`;
            const termParentQuery = args.options.parentTermId ?
                // get parent term by ID
                `<Method Id="16" ParentId="6" Name="GetTerm"><Parameters><Parameter Type="Guid">{${args.options.parentTermId}}</Parameter></Parameters></Method>` :
                // no parent term specified, add to term set
                args.options.termSetId ? `<Method Id="16" ParentId="14" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termSetId}}</Parameter></Parameters></Method>` : `<Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termSetName)}</Parameter></Parameters></Method>`;
            const termId = args.options.id || v4();
            const data = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" />${termGroupQuery}<Property Id="14" ParentId="11" Name="TermSets" />${termParentQuery}<Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.name)}</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{${termId}}</Parameter></Parameters></Method></ObjectPaths></Request>`;
            const requestOptionsPost = {
                url: `${spoWebUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': res.FormDigestValue
                },
                data: data
            };
            const processQuery = await request.post(requestOptionsPost);
            const json = JSON.parse(processQuery);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            term = json[json.length - 1];
            let terms = undefined;
            if (!(!args.options.description &&
                !args.options.customProperties &&
                !args.options.localCustomProperties)) {
                if (this.verbose) {
                    await logger.logToStderr(`Setting term properties...`);
                }
                const properties = [];
                let i = 127;
                if (args.options.description) {
                    properties.push(`<Method Name="SetDescription" Id="${i++}" ObjectPathId="117"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.description)}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method>`);
                    term.Description = args.options.description;
                }
                if (args.options.customProperties) {
                    const customProperties = JSON.parse(args.options.customProperties);
                    Object.keys(customProperties).forEach(k => {
                        properties.push(`<Method Name="SetCustomProperty" Id="${i++}" ObjectPathId="117"><Parameters><Parameter Type="String">${formatting.escapeXml(k)}</Parameter><Parameter Type="String">${formatting.escapeXml(customProperties[k])}</Parameter></Parameters></Method>`);
                    });
                    term.CustomProperties = customProperties;
                }
                if (args.options.localCustomProperties) {
                    const localCustomProperties = JSON.parse(args.options.localCustomProperties);
                    Object.keys(localCustomProperties).forEach(k => {
                        properties.push(`<Method Name="SetLocalCustomProperty" Id="${i++}" ObjectPathId="117"><Parameters><Parameter Type="String">${formatting.escapeXml(k)}</Parameter><Parameter Type="String">${formatting.escapeXml(localCustomProperties[k])}</Parameter></Parameters></Method>`);
                    });
                    term.LocalCustomProperties = localCustomProperties;
                }
                let termStoreObjectIdentity = '';
                // get term store object identity
                for (let i = 0; i < json.length; i++) {
                    if (json[i] !== 8) {
                        continue;
                    }
                    termStoreObjectIdentity = json[i + 1]._ObjectIdentity_;
                    break;
                }
                const requestOptions = {
                    url: `${spoWebUrl}/_vti_bin/client.svc/ProcessQuery`,
                    headers: {
                        'X-RequestDigest': formDigest
                    },
                    data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${properties.join('')}<Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="${term._ObjectIdentity_}" /><Identity Id="109" Name="${termStoreObjectIdentity}" /></ObjectPaths></Request>`
                };
                terms = await request.post(requestOptions);
            }
            if (terms) {
                const json = JSON.parse(terms);
                const response = json[0];
                if (response.ErrorInfo) {
                    throw response.ErrorInfo.ErrorMessage;
                }
            }
            delete term._ObjectIdentity_;
            delete term._ObjectType_;
            term.CreatedDate = new Date(Number(term.CreatedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
            term.Id = term.Id.replace('/Guid(', '').replace(')/', '');
            term.LastModifiedDate = new Date(Number(term.LastModifiedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
            await logger.log(term);
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
_SpoTermAddCommand_instances = new WeakSet(), _SpoTermAddCommand_initTelemetry = function _SpoTermAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            webUrl: typeof args.options.webUrl !== 'undefined',
            customProperties: typeof args.options.customProperties !== 'undefined',
            description: typeof args.options.description !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            localCustomProperties: typeof args.options.localCustomProperties !== 'undefined',
            parentTermId: typeof args.options.parentTermId !== 'undefined',
            termGroupId: typeof args.options.termGroupId !== 'undefined',
            termGroupName: typeof args.options.termGroupName !== 'undefined',
            termSetId: typeof args.options.termSetId !== 'undefined',
            termSetName: typeof args.options.termSetName !== 'undefined'
        });
    });
}, _SpoTermAddCommand_initOptions = function _SpoTermAddCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '-u, --webUrl [webUrl]'
    }, {
        option: '--termSetId [termSetId]'
    }, {
        option: '--termSetName [termSetName]'
    }, {
        option: '--termGroupId [termGroupId]'
    }, {
        option: '--termGroupName [termGroupName]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '--parentTermId [parentTermId]'
    }, {
        option: '--customProperties [customProperties]'
    }, {
        option: '--localCustomProperties [localCustomProperties]'
    });
}, _SpoTermAddCommand_initValidators = function _SpoTermAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.webUrl) {
            const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
            if (isValidSharePointUrl !== true) {
                return isValidSharePointUrl;
            }
        }
        if (args.options.id) {
            if (!validation.isValidGuid(args.options.id)) {
                return `${args.options.id} is not a valid GUID`;
            }
        }
        if (args.options.parentTermId) {
            if (!validation.isValidGuid(args.options.parentTermId)) {
                return `${args.options.parentTermId} is not a valid GUID`;
            }
            if (args.options.termSetId || args.options.termSetName) {
                return 'Specify either parentTermId, termSetId or termSetName but not both';
            }
        }
        if (args.options.termGroupId) {
            if (!validation.isValidGuid(args.options.termGroupId)) {
                return `${args.options.termGroupId} is not a valid GUID`;
            }
        }
        if (!args.options.termSetId && !args.options.termSetName && !args.options.parentTermId) {
            return 'Specify termSetId, termSetName or parentTermId';
        }
        if (args.options.termSetId && args.options.termSetName) {
            return 'Specify termSetId or termSetName but not both';
        }
        if (args.options.termSetId) {
            if (!validation.isValidGuid(args.options.termSetId)) {
                return `${args.options.termSetId} is not a valid GUID`;
            }
        }
        if (args.options.customProperties) {
            try {
                JSON.parse(args.options.customProperties);
            }
            catch (e) {
                return `An error has occurred while parsing customProperties: ${e}`;
            }
        }
        if (args.options.localCustomProperties) {
            try {
                JSON.parse(args.options.localCustomProperties);
            }
            catch (e) {
                return `An error has occurred while parsing localCustomProperties: ${e}`;
            }
        }
        return true;
    });
}, _SpoTermAddCommand_initOptionSets = function _SpoTermAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['termGroupId', 'termGroupName'] });
};
export default new SpoTermAddCommand();
//# sourceMappingURL=term-add.js.map