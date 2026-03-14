var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTermListCommand_instances, _SpoTermListCommand_initTelemetry, _SpoTermListCommand_initOptions, _SpoTermListCommand_initValidators, _SpoTermListCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTermListCommand extends SpoCommand {
    get name() {
        return commands.TERM_LIST;
    }
    get description() {
        return 'Lists taxonomy terms from the given term set';
    }
    defaultProperties() {
        return ['Id', 'Name', 'ParentTermId'];
    }
    constructor() {
        super();
        _SpoTermListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTermListCommand_instances, "m", _SpoTermListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoTermListCommand_instances, "m", _SpoTermListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTermListCommand_instances, "m", _SpoTermListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoTermListCommand_instances, "m", _SpoTermListCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoWebUrl = args.options.webUrl ? args.options.webUrl : await spo.getSpoAdminUrl(logger, this.debug);
            const res = await spo.getRequestDigest(spoWebUrl);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving taxonomy term sets...`);
            }
            const termGroupQuery = args.options.termGroupId ? `<Method Id="77" ParentId="75" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termGroupId}}</Parameter></Parameters></Method>` : `<Method Id="77" ParentId="75" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termGroupName)}</Parameter></Parameters></Method>`;
            const termSetQuery = args.options.termSetId ? `<Method Id="82" ParentId="80" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termSetId}}</Parameter></Parameters></Method>` : `<Method Id="82" ParentId="80" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termSetName)}</Parameter></Parameters></Method>`;
            const data = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" />${termGroupQuery}<Property Id="80" ParentId="77" Name="TermSets" />${termSetQuery}<Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`;
            const shouldIncludeChildTerms = args.options.withChildTerms;
            const result = await this.executeCsomCall(data, spoWebUrl, res);
            const terms = [];
            if (result._Child_Items_ && result._Child_Items_.length > 0) {
                for (const term of result._Child_Items_) {
                    this.setTermDetails(term);
                    terms.push(term);
                    if (shouldIncludeChildTerms && term.TermsCount > 0) {
                        await this.getChildTerms(spoWebUrl, res, term);
                    }
                }
            }
            if (!args.options.output || !cli.shouldTrimOutput(args.options.output)) {
                await logger.log(terms);
            }
            else if (!shouldIncludeChildTerms) {
                // Converted to text friendly output
                await logger.log(terms.map(i => {
                    return {
                        Id: i.Id,
                        Name: i.Name
                    };
                }));
            }
            else {
                // Converted to text friendly output
                const friendlyOutput = [];
                terms.forEach(term => {
                    term.ParentTermId = '';
                    friendlyOutput.push(term);
                    if (term.Children && term.Children.length > 0) {
                        this.getFriendlyChildTerms(term, friendlyOutput);
                    }
                });
                await logger.log(friendlyOutput.map(i => {
                    return {
                        Id: i.Id,
                        Name: i.Name,
                        ParentTermId: i.ParentTermId
                    };
                }));
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getFriendlyChildTerms(term, friendlyOutput) {
        term.Children.forEach(childTerm => {
            childTerm.ParentTermId = term.Id;
            friendlyOutput.push(childTerm);
            if (childTerm.Children && childTerm.Children.length > 0) {
                this.getFriendlyChildTerms(childTerm, friendlyOutput);
            }
        });
    }
    async getChildTerms(spoWebUrl, res, parentTerm) {
        parentTerm.Children = [];
        const data = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="20" ObjectPathId="19" /><Query Id="21" ObjectPathId="19"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="CustomSortOrder" ScalarProperty="true" /><Property Name="CustomProperties" ScalarProperty="true" /><Property Name="LocalCustomProperties" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><Property Id="19" ParentId="16" Name="Terms" /><Identity Id="16" Name="${parentTerm._ObjectIdentity_}" /></ObjectPaths></Request>`;
        const result = await this.executeCsomCall(data, spoWebUrl, res);
        if (result._Child_Items_ && result._Child_Items_.length > 0) {
            for (const term of result._Child_Items_) {
                this.setTermDetails(term);
                parentTerm.Children.push(term);
                if (term.TermsCount > 0) {
                    await this.getChildTerms(spoWebUrl, res, term);
                }
            }
        }
    }
    setTermDetails(term) {
        term.CreatedDate = this.parseTermDateToIsoString(term.CreatedDate);
        term.Id = term.Id.replace('/Guid(', '').replace(')/', '');
        term.LastModifiedDate = this.parseTermDateToIsoString(term.LastModifiedDate);
    }
    parseTermDateToIsoString(dateAsString) {
        return new Date(Number(dateAsString.replace('/Date(', '').replace(')/', ''))).toISOString();
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
        return json[json.length - 1];
    }
}
_SpoTermListCommand_instances = new WeakSet(), _SpoTermListCommand_initTelemetry = function _SpoTermListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            webUrl: typeof args.options.webUrl !== 'undefined',
            termGroupId: typeof args.options.termGroupId !== 'undefined',
            termGroupName: typeof args.options.termGroupName !== 'undefined',
            termSetId: typeof args.options.termSetId !== 'undefined',
            termSetName: typeof args.options.termSetName !== 'undefined',
            withChildTerms: !!args.options.withChildTerms
        });
    });
}, _SpoTermListCommand_initOptions = function _SpoTermListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl [webUrl]'
    }, {
        option: '--termGroupId [termGroupId]'
    }, {
        option: '--termGroupName [termGroupName]'
    }, {
        option: '--termSetId [termSetId]'
    }, {
        option: '--termSetName [termSetName]'
    }, {
        option: '--withChildTerms'
    });
}, _SpoTermListCommand_initValidators = function _SpoTermListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.webUrl) {
            const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
            if (isValidSharePointUrl !== true) {
                return isValidSharePointUrl;
            }
        }
        if (args.options.termGroupId && !validation.isValidGuid(args.options.termGroupId)) {
            return `${args.options.termGroupId} is not a valid GUID`;
        }
        if (args.options.termSetId && !validation.isValidGuid(args.options.termSetId)) {
            return `${args.options.termSetId} is not a valid GUID`;
        }
        return true;
    });
}, _SpoTermListCommand_initOptionSets = function _SpoTermListCommand_initOptionSets() {
    this.optionSets.push({ options: ['termGroupId', 'termGroupName'] }, { options: ['termSetId', 'termSetName'] });
};
export default new SpoTermListCommand();
//# sourceMappingURL=term-list.js.map