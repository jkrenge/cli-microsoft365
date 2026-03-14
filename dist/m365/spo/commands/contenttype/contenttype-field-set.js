var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoContentTypeFieldSetCommand_instances, _SpoContentTypeFieldSetCommand_initTelemetry, _SpoContentTypeFieldSetCommand_initOptions, _SpoContentTypeFieldSetCommand_initValidators, _SpoContentTypeFieldSetCommand_initTypes;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoContentTypeFieldSetCommand extends SpoCommand {
    get name() {
        return commands.CONTENTTYPE_FIELD_SET;
    }
    get description() {
        return 'Adds or updates a site column reference in a site content type';
    }
    constructor() {
        super();
        _SpoContentTypeFieldSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoContentTypeFieldSetCommand_instances, "m", _SpoContentTypeFieldSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeFieldSetCommand_instances, "m", _SpoContentTypeFieldSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeFieldSetCommand_instances, "m", _SpoContentTypeFieldSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeFieldSetCommand_instances, "m", _SpoContentTypeFieldSetCommand_initTypes).call(this);
        this.requestDigest = '';
        this.siteId = '';
        this.webId = '';
        this.fieldLink = null;
    }
    async commandAction(logger, args) {
        try {
            let schemaXmlWithResourceTokens = '';
            if (this.verbose) {
                await logger.logToStderr(`Retrieving field link for field ${args.options.id}...`);
            }
            let requestOptions = {
                url: `${args.options.webUrl}/_api/web/contenttypes('${formatting.encodeQueryParameter(args.options.contentTypeId)}')/fieldlinks('${args.options.id}')`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const fieldLink = await request.get(requestOptions);
            if (fieldLink["odata.null"] !== true) {
                if (this.verbose) {
                    await logger.logToStderr('Field link found');
                }
                this.fieldLink = fieldLink;
            }
            else {
                if (this.verbose) {
                    await logger.logToStderr('Field link not found. Creating...');
                    await logger.logToStderr(`Retrieving information about site column ${args.options.id}...`);
                }
                requestOptions = {
                    url: `${args.options.webUrl}/_api/web/fields('${args.options.id}')?$select=SchemaXmlWithResourceTokens`,
                    headers: {
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                const field = await request.get(requestOptions);
                schemaXmlWithResourceTokens = field.SchemaXmlWithResourceTokens;
                await this.createFieldLink(logger, args, schemaXmlWithResourceTokens);
            }
            if (!this.fieldLink) {
                if (this.verbose) {
                    await logger.logToStderr(`Retrieving information about field link for field ${args.options.id}...`);
                }
                requestOptions = {
                    url: `${args.options.webUrl}/_api/web/contenttypes('${formatting.encodeQueryParameter(args.options.contentTypeId)}')/fieldlinks('${args.options.id}')`,
                    headers: {
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                const fieldLinkResult = await request.get(requestOptions);
                if (fieldLinkResult && fieldLinkResult["odata.null"] !== true) {
                    this.fieldLink = fieldLinkResult;
                }
            }
            if (!this.fieldLink) {
                throw `Couldn't find field link for field ${args.options.id}`;
            }
            let updateHidden = false;
            let updateRequired = false;
            if (typeof args.options.hidden !== 'undefined' &&
                this.fieldLink.Hidden !== args.options.hidden) {
                updateHidden = true;
            }
            if (typeof args.options.required !== 'undefined' &&
                this.fieldLink.Required !== args.options.required) {
                updateRequired = true;
            }
            if (!updateHidden && !updateRequired) {
                if (this.verbose) {
                    await logger.logToStderr('Field link already up-to-date');
                }
                throw 'DONE';
            }
            if (!this.siteId) {
                this.siteId = await spo.getSiteIdBySPApi(args.options.webUrl, logger, this.verbose);
            }
            if (!this.webId) {
                this.webId = await spo.getWebId(args.options.webUrl, logger, this.verbose);
            }
            if (this.verbose) {
                await logger.logToStderr(`Updating field link...`);
            }
            const requiredProperty = typeof args.options.required !== 'undefined' &&
                this.fieldLink.Required !== args.options.required ? `<SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">${args.options.required}</Parameter></SetProperty>` : '';
            const hiddenProperty = typeof args.options.hidden !== 'undefined' &&
                this.fieldLink.Hidden !== args.options.hidden ? `<SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">${args.options.hidden}</Parameter></SetProperty>` : '';
            requestOptions = {
                url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': this.requestDigest
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${requiredProperty}${hiddenProperty}<Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${this.siteId}:web:${this.webId}:contenttype:${formatting.escapeXml(args.options.contentTypeId)}:fl:${this.fieldLink.Id}" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${this.siteId}:web:${this.webId}:contenttype:${formatting.escapeXml(args.options.contentTypeId)}" /></ObjectPaths></Request>`
            };
            const res = await request.post(requestOptions);
            const json = JSON.parse(res);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
        }
        catch (err) {
            if (err !== 'DONE') {
                this.handleRejectedODataJsonPromise(err);
            }
        }
    }
    async createFieldLink(logger, args, schemaXmlWithResourceTokens) {
        let requiresUpdate = false;
        const match = /(<Field[^>]+>)(.*)/.exec(schemaXmlWithResourceTokens);
        let xField = match[1];
        const allowDeletion = /AllowDeletion="([^"]+)"/.exec(xField);
        if (!allowDeletion) {
            requiresUpdate = true;
            xField = xField.replace('>', ' AllowDeletion="TRUE">') + match[2];
        }
        else {
            if (allowDeletion[1] !== 'TRUE') {
                requiresUpdate = true;
                xField = xField.replace(allowDeletion[0], 'AllowDeletion="TRUE"') + match[2];
            }
        }
        await this.updateField(xField, requiresUpdate, logger, args);
        this.siteId = await spo.getSiteIdBySPApi(args.options.webUrl, logger, this.verbose);
        this.webId = await spo.getWebId(args.options.webUrl, logger, this.verbose);
        await this.ensureRequestDigest(args.options.webUrl, logger);
        const requestOptions = {
            url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': this.requestDigest
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><Method Name="Update" Id="7" ObjectPathId="1"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="2" Name="d6667b9e-50fb-0000-2693-032ae7a0df25|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${this.siteId}:web:${this.webId}:field:${args.options.id}" /><Method Id="4" ParentId="3" Name="Add"><Parameters><Parameter TypeId="{63fb2c92-8f65-4bbb-a658-b6cd294403f4}"><Property Name="Field" ObjectPathId="2" /></Parameter></Parameters></Method><Identity Id="1" Name="d6667b9e-80f4-0000-2693-05528ff416bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${this.siteId}:web:${this.webId}:contenttype:${formatting.escapeXml(args.options.contentTypeId)}" /><Property Id="3" ParentId="1" Name="FieldLinks" /></ObjectPaths></Request>`
        };
        const res = await request.post(requestOptions);
        const json = JSON.parse(res);
        const response = json[0];
        if (response.ErrorInfo) {
            throw response.ErrorInfo.ErrorMessage;
        }
    }
    async updateField(schemaXml, requiresUpdate, logger, args) {
        if (!requiresUpdate) {
            if (this.verbose) {
                await logger.logToStderr(`Schema of field ${args.options.id} is already up-to-date`);
            }
            return;
        }
        await this.ensureRequestDigest(args.options.webUrl, logger);
        if (this.verbose) {
            await logger.logToStderr(`Updating field schema...`);
        }
        const requestOptions = {
            url: `${args.options.webUrl}/_api/web/fields('${args.options.id}')`,
            headers: {
                accept: 'application/json;odata=nometadata',
                'content-type': 'application/json;odata=nometadata',
                'X-HTTP-Method': 'MERGE',
                'x-requestdigest': this.requestDigest
            },
            data: {
                SchemaXml: schemaXml
            },
            responseType: 'json'
        };
        await request.post(requestOptions);
    }
    async ensureRequestDigest(siteUrl, logger) {
        if (this.requestDigest) {
            if (this.debug) {
                await logger.logToStderr('Request digest already present');
            }
            return;
        }
        if (this.debug) {
            await logger.logToStderr('Retrieving request digest...');
        }
        const res = await spo.getRequestDigest(siteUrl);
        this.requestDigest = res.FormDigestValue;
    }
}
_SpoContentTypeFieldSetCommand_instances = new WeakSet(), _SpoContentTypeFieldSetCommand_initTelemetry = function _SpoContentTypeFieldSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            hidden: args.options.hidden,
            required: args.options.required
        });
    });
}, _SpoContentTypeFieldSetCommand_initOptions = function _SpoContentTypeFieldSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--contentTypeId <contentTypeId>'
    }, {
        option: '-i, --id <id>'
    }, {
        option: '-r, --required [required]',
        autocomplete: ['true', 'false']
    }, {
        option: '--hidden [hidden]',
        autocomplete: ['true', 'false']
    });
}, _SpoContentTypeFieldSetCommand_initValidators = function _SpoContentTypeFieldSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoContentTypeFieldSetCommand_initTypes = function _SpoContentTypeFieldSetCommand_initTypes() {
    this.types.string.push('contentTypeId', 'c');
    this.types.boolean.push('required', 'hidden');
};
export default new SpoContentTypeFieldSetCommand();
//# sourceMappingURL=contenttype-field-set.js.map