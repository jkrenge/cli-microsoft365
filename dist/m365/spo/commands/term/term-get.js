var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTermGetCommand_instances, _SpoTermGetCommand_initTelemetry, _SpoTermGetCommand_initOptions, _SpoTermGetCommand_initValidators, _SpoTermGetCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTermGetCommand extends SpoCommand {
    get name() {
        return commands.TERM_GET;
    }
    get description() {
        return 'Gets information about the specified taxonomy term';
    }
    constructor() {
        super();
        _SpoTermGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTermGetCommand_instances, "m", _SpoTermGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoTermGetCommand_instances, "m", _SpoTermGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTermGetCommand_instances, "m", _SpoTermGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoTermGetCommand_instances, "m", _SpoTermGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoWebUrl = args.options.webUrl ? args.options.webUrl : await spo.getSpoAdminUrl(logger, this.debug);
            const res = await spo.getRequestDigest(spoWebUrl);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving taxonomy term...`);
            }
            let data = '';
            if (args.options.id) {
                data = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="6" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="7" ParentId="6" Name="GetDefaultSiteCollectionTermStore" /><Method Id="13" ParentId="7" Name="GetTerm"><Parameters><Parameter Type="Guid">{${args.options.id}}</Parameter></Parameters></Method></ObjectPaths></Request>`;
            }
            else {
                data = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectIdentityQuery Id="3" ObjectPathId="1" /><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><ObjectPath Id="8" ObjectPathId="7" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><ObjectPath Id="13" ObjectPathId="12" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectIdentityQuery Id="16" ObjectPathId="14" /><ObjectPath Id="18" ObjectPathId="17" /><SetProperty Id="19" ObjectPathId="17" Name="TrimUnavailable"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="20" ObjectPathId="17" Name="TermLabel"><Parameter Type="String">${formatting.escapeXml(args.options.name)}</Parameter></SetProperty><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="1" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="4" ParentId="1" Name="GetDefaultSiteCollectionTermStore" /><Property Id="7" ParentId="4" Name="Groups" /><Method Id="9" ParentId="7" Name="${args.options.termGroupName === undefined ? "GetById" : "GetByName"}"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termGroupName) || args.options.termGroupId}</Parameter></Parameters></Method><Property Id="12" ParentId="9" Name="TermSets" /><Method Id="14" ParentId="12" Name="${args.options.termSetName === undefined ? "GetById" : "GetByName"}"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termSetName) || args.options.termSetId}</Parameter></Parameters></Method><Constructor Id="17" TypeId="{61a1d689-2744-4ea3-a88b-c95bee9803aa}" /><Method Id="21" ParentId="14" Name="GetTerms"><Parameters><Parameter ObjectPathId="17" /></Parameters></Method></ObjectPaths></Request>`;
            }
            let term;
            const csomResponse = await this.executeCsomCall(data, spoWebUrl, res);
            if (csomResponse === null) {
                throw `Term with id '${args.options.id}' could not be found.`;
            }
            else if (csomResponse._Child_Items_ !== undefined) {
                const terms = csomResponse;
                if (terms._Child_Items_.length === 0) {
                    throw `Term with name '${args.options.name}' could not be found.`;
                }
                if (terms._Child_Items_.length > 1) {
                    const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', terms._Child_Items_);
                    term = await cli.handleMultipleResultsFound(`Multiple terms with the specific term name found.`, resultAsKeyValuePair);
                }
                else {
                    term = terms._Child_Items_[0];
                }
            }
            else {
                term = csomResponse;
            }
            delete term._ObjectIdentity_;
            delete term._ObjectType_;
            term.CreatedDate = this.parseTermDateToIsoString(term.CreatedDate);
            term.Id = this.getTermId(term.Id);
            term.LastModifiedDate = this.parseTermDateToIsoString(term.LastModifiedDate);
            await logger.log(term);
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
    async executeCsomCall(data, spoWebUrl, res) {
        const requestOptions = {
            url: `${spoWebUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': res.FormDigestValue
            },
            data: data
        };
        const processQuery = await request.post(requestOptions);
        const json = JSON.parse(processQuery);
        const response = json[0];
        if (response.ErrorInfo) {
            throw response.ErrorInfo.ErrorMessage;
        }
        const responseObject = json[json.length - 1];
        return responseObject;
    }
    getTermId(termId) {
        return termId.replace('/Guid(', '').replace(')/', '');
    }
    parseTermDateToIsoString(dateAsString) {
        return new Date(Number(dateAsString.replace('/Date(', '').replace(')/', ''))).toISOString();
    }
}
_SpoTermGetCommand_instances = new WeakSet(), _SpoTermGetCommand_initTelemetry = function _SpoTermGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            webUrl: typeof args.options.webUrl !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            termGroupId: typeof args.options.termGroupId !== 'undefined',
            termGroupName: typeof args.options.termGroupName !== 'undefined',
            termSetId: typeof args.options.termSetId !== 'undefined',
            termSetName: typeof args.options.termSetName !== 'undefined'
        });
    });
}, _SpoTermGetCommand_initOptions = function _SpoTermGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl [webUrl]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '--termGroupId [termGroupId]'
    }, {
        option: '--termGroupName [termGroupName]'
    }, {
        option: '--termSetId [termSetId]'
    }, {
        option: '--termSetName [termSetName]'
    });
}, _SpoTermGetCommand_initValidators = function _SpoTermGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.webUrl) {
            const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
            if (isValidSharePointUrl !== true) {
                return isValidSharePointUrl;
            }
        }
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.termGroupId && !validation.isValidGuid(args.options.termGroupId)) {
            return `${args.options.termGroupId} is not a valid GUID`;
        }
        if (args.options.termSetId && !validation.isValidGuid(args.options.termSetId)) {
            return `${args.options.termSetId} is not a valid GUID`;
        }
        return true;
    });
}, _SpoTermGetCommand_initOptionSets = function _SpoTermGetCommand_initOptionSets() {
    this.optionSets.push({
        options: ['id', 'name']
    }, {
        options: ['termGroupId', 'termGroupName'],
        runsWhen: (args) => {
            return args.options.name !== undefined;
        }
    }, {
        options: ['termSetId', 'termSetName'],
        runsWhen: (args) => {
            return args.options.name !== undefined;
        }
    });
};
export default new SpoTermGetCommand();
//# sourceMappingURL=term-get.js.map