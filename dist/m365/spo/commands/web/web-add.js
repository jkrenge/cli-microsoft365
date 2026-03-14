var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoWebAddCommand_instances, _SpoWebAddCommand_initTelemetry, _SpoWebAddCommand_initOptions, _SpoWebAddCommand_initValidators;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { BasePermissions, PermissionKind } from '../../base-permissions.js';
import commands from '../../commands.js';
class SpoWebAddCommand extends SpoCommand {
    get name() {
        return commands.WEB_ADD;
    }
    get description() {
        return 'Create new subsite';
    }
    constructor() {
        super();
        _SpoWebAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoWebAddCommand_instances, "m", _SpoWebAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoWebAddCommand_instances, "m", _SpoWebAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoWebAddCommand_instances, "m", _SpoWebAddCommand_initValidators).call(this);
    }
    getExcludedOptionsWithUrls() {
        return ['url'];
    }
    async commandAction(logger, args) {
        try {
            const res = await spo.getRequestDigest(args.options.parentWebUrl);
            const requestOptionsPost = {
                url: `${args.options.parentWebUrl}/_api/web/webinfos/add`,
                headers: {
                    'content-type': 'application/json;odata=nometadata',
                    accept: 'application/json;odata=nometadata',
                    'X-RequestDigest': res.FormDigestValue
                },
                responseType: 'json',
                data: {
                    parameters: {
                        Url: args.options.url,
                        Title: args.options.title,
                        Description: args.options.description,
                        Language: args.options.locale,
                        WebTemplate: args.options.webTemplate,
                        UseUniquePermissions: args.options.breakInheritance
                    }
                }
            };
            if (this.verbose) {
                await logger.logToStderr(`Creating subsite ${args.options.parentWebUrl}/${args.options.webUrl}...`);
            }
            const siteInfo = await request.post(requestOptionsPost);
            if (args.options.inheritNavigation) {
                if (this.verbose) {
                    await logger.logToStderr("Setting inheriting navigation from the parent site...");
                }
                const subsiteFullUrl = `${args.options.parentWebUrl}/${formatting.encodeQueryParameter(args.options.url)}`;
                const requestOptionsPer = {
                    url: `${subsiteFullUrl}/_api/web/effectivebasepermissions`,
                    headers: {
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                const effectivebasepermissions = await request.get(requestOptionsPer);
                const permissions = new BasePermissions();
                permissions.high = effectivebasepermissions.High;
                permissions.low = effectivebasepermissions.Low;
                /// Detects if the site in question has no script enabled or not. 
                /// Detection is done by verifying if the AddAndCustomizePages permission is missing.
                /// 
                /// See https://support.office.com/en-us/article/Turn-scripting-capabilities-on-or-off-1f2c515f-5d7e-448a-9fd7-835da935584f
                /// for the effects of NoScript
                if (permissions.has(PermissionKind.AddAndCustomizePages)) {
                    const digest = await spo.getRequestDigest(subsiteFullUrl);
                    const requestOptionsQuery = {
                        url: `${subsiteFullUrl}/_vti_bin/client.svc/ProcessQuery`,
                        headers: {
                            'X-RequestDigest': digest.FormDigestValue
                        },
                        data: `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="1" ObjectPathId="0" /><ObjectPath Id="3" ObjectPathId="2" /><ObjectPath Id="5" ObjectPathId="4" /><SetProperty Id="6" ObjectPathId="4" Name="UseShared"><Parameter Type="Boolean">true</Parameter></SetProperty></Actions><ObjectPaths><StaticProperty Id="0" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /><Property Id="2" ParentId="0" Name="Web" /><Property Id="4" ParentId="2" Name="Navigation" /></ObjectPaths></Request>`
                    };
                    const query = await request.post(requestOptionsQuery);
                    const json = JSON.parse(query);
                    const response = json[0];
                    if (response.ErrorInfo) {
                        throw response.ErrorInfo.ErrorMessage;
                    }
                }
                else {
                    if (this.verbose) {
                        await logger.logToStderr("No script is enabled. Skipping the InheritParentNavigation settings.");
                    }
                }
            }
            await logger.log(siteInfo);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoWebAddCommand_instances = new WeakSet(), _SpoWebAddCommand_initTelemetry = function _SpoWebAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            description: (!(!args.options.description)).toString(),
            locale: args.options.locale || '1033',
            breakInheritance: args.options.breakInheritance || false,
            inheritNavigation: args.options.inheritNavigation || false
        });
    });
}, _SpoWebAddCommand_initOptions = function _SpoWebAddCommand_initOptions() {
    this.options.unshift({
        option: '-t, --title <title>'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '-u, --url <url>'
    }, {
        option: '-w, --webTemplate <webTemplate>'
    }, {
        option: '-p, --parentWebUrl <parentWebUrl>'
    }, {
        option: '-l, --locale [locale]'
    }, {
        option: '--breakInheritance'
    }, {
        option: '--inheritNavigation'
    });
}, _SpoWebAddCommand_initValidators = function _SpoWebAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.parentWebUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.locale) {
            const locale = parseInt(args.options.locale);
            if (isNaN(locale)) {
                return `${args.options.locale} is not a valid locale number`;
            }
        }
        return true;
    });
};
export default new SpoWebAddCommand();
//# sourceMappingURL=web-add.js.map