var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _FlowExportCommand_instances, _FlowExportCommand_initTelemetry, _FlowExportCommand_initOptions, _FlowExportCommand_initValidators;
import fs from 'fs';
import path from 'path';
import request from '../../../request.js';
import { formatting } from '../../../utils/formatting.js';
import { validation } from '../../../utils/validation.js';
import PowerPlatformCommand from '../../base/PowerPlatformCommand.js';
import commands from '../commands.js';
import PowerAutomateCommand from '../../base/PowerAutomateCommand.js';
class FlowExportCommand extends PowerPlatformCommand {
    get name() {
        return commands.EXPORT;
    }
    get description() {
        return 'Exports the specified Microsoft Flow as a file';
    }
    constructor() {
        super();
        _FlowExportCommand_instances.add(this);
        __classPrivateFieldGet(this, _FlowExportCommand_instances, "m", _FlowExportCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _FlowExportCommand_instances, "m", _FlowExportCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _FlowExportCommand_instances, "m", _FlowExportCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const formatArgument = args.options.format?.toLowerCase() || '';
        if (this.verbose) {
            await logger.logToStderr(`Retrieving package resources for Microsoft Flow ${args.options.name}...`);
        }
        try {
            let res;
            if (formatArgument === 'json') {
                if (this.verbose) {
                    await logger.logToStderr('format = json, skipping listing package resources step.');
                }
            }
            else {
                const requestOptions = {
                    url: `${this.resource}/providers/Microsoft.BusinessAppPlatform/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/listPackageResources?api-version=2016-11-01`,
                    headers: {
                        accept: 'application/json'
                    },
                    data: {
                        "baseResourceIds": [
                            `/providers/Microsoft.Flow/flows/${args.options.name}`
                        ]
                    },
                    responseType: 'json'
                };
                res = await request.post(requestOptions);
            }
            if (typeof res !== 'undefined' && res.errors && res.errors.length && res.errors.length > 0) {
                throw res.errors[0].message;
            }
            if (this.verbose) {
                await logger.logToStderr(`Initiating package export for Microsoft Flow ${args.options.name}...`);
            }
            let requestOptions = {
                url: formatArgument === 'json' ?
                    `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.name)}?api-version=2016-11-01`
                    : `${this.resource}/providers/Microsoft.BusinessAppPlatform/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/exportPackage?api-version=2016-11-01`,
                headers: {
                    accept: 'application/json'
                },
                responseType: 'json'
            };
            if (formatArgument !== 'json') {
                // adds suggestedCreationType property to all resources
                // see https://github.com/pnp/cli-microsoft365/issues/1845
                Object.keys(res.resources).forEach((key) => {
                    if (res.resources[key].type === 'Microsoft.Flow/flows') {
                        res.resources[key].suggestedCreationType = 'Update';
                    }
                    else {
                        res.resources[key].suggestedCreationType = 'Existing';
                    }
                });
                requestOptions.data = {
                    includedResourceIds: [
                        `/providers/Microsoft.Flow/flows/${args.options.name}`
                    ],
                    details: {
                        displayName: args.options.packageDisplayName,
                        description: args.options.packageDescription,
                        creator: args.options.packageCreatedBy,
                        sourceEnvironment: args.options.packageSourceEnvironment
                    },
                    resources: res.resources
                };
            }
            res = formatArgument === 'json' ? await request.get(requestOptions) : await request.post(requestOptions);
            if (this.verbose) {
                await logger.logToStderr(`Getting file for Microsoft Flow ${args.options.name}...`);
            }
            const downloadFileUrl = formatArgument === 'json' ? '' : res.packageLink.value;
            const filenameRegEx = /([^/]+\.zip)/i;
            let filenameFromApi = formatArgument === 'json' ? `${res.properties.displayName}.json` : (filenameRegEx.exec(downloadFileUrl) || ['output.zip'])[0];
            // Replace all illegal characters from the file name
            const illegalCharsRegEx = /[\\/:*?"<>|]/g;
            filenameFromApi = filenameFromApi.replace(illegalCharsRegEx, '_');
            if (this.verbose) {
                await logger.logToStderr(`Filename from PowerApps API: ${filenameFromApi}.`);
                await logger.logToStderr('');
            }
            requestOptions = {
                url: formatArgument === 'json' ?
                    `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.name)}/exportToARMTemplate?api-version=2016-11-01`
                    : downloadFileUrl,
                // Set responseType to arraybuffer, otherwise binary data will be encoded
                // to utf8 and binary data is corrupt
                responseType: 'arraybuffer',
                headers: formatArgument === 'json' ?
                    {
                        accept: 'application/json'
                    } : {
                    'x-anonymous': true
                }
            };
            const file = formatArgument === 'json' ?
                await request.post(requestOptions)
                : await request.get(requestOptions);
            const path = args.options.path ? args.options.path : `./${filenameFromApi}`;
            fs.writeFileSync(path, file, 'binary');
            if (!args.options.path || this.verbose) {
                if (this.verbose) {
                    await logger.logToStderr(`File saved to path '${path}'.`);
                }
                else {
                    await logger.log(path);
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_FlowExportCommand_instances = new WeakSet(), _FlowExportCommand_initTelemetry = function _FlowExportCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            packageDisplayName: typeof args.options.packageDisplayName !== 'undefined',
            packageDescription: typeof args.options.packageDescription !== 'undefined',
            packageCreatedBy: typeof args.options.packageCreatedBy !== 'undefined',
            packageSourceEnvironment: typeof args.options.packageSourceEnvironment !== 'undefined',
            format: args.options.format,
            path: typeof args.options.path !== 'undefined'
        });
    });
}, _FlowExportCommand_initOptions = function _FlowExportCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '-d, --packageDisplayName [packageDisplayName]'
    }, {
        option: '--packageDescription [packageDescription]'
    }, {
        option: '-c, --packageCreatedBy [packageCreatedBy]'
    }, {
        option: '-s, --packageSourceEnvironment [packageSourceEnvironment]'
    }, {
        option: '-f, --format [format]'
    }, {
        option: '-p, --path [path]'
    });
}, _FlowExportCommand_initValidators = function _FlowExportCommand_initValidators() {
    this.validators.push(async (args) => {
        const lowerCaseFormat = args.options.format ? args.options.format.toLowerCase() : '';
        if (!validation.isValidGuid(args.options.name)) {
            return `${args.options.name} is not a valid GUID`;
        }
        if (args.options.format && (lowerCaseFormat !== 'json' && lowerCaseFormat !== 'zip')) {
            return 'Option format must be json or zip. Default is zip';
        }
        if (lowerCaseFormat === 'json') {
            if (args.options.packageCreatedBy) {
                return 'packageCreatedBy cannot be specified with output of json';
            }
            if (args.options.packageDescription) {
                return 'packageDescription cannot be specified with output of json';
            }
            if (args.options.packageDisplayName) {
                return 'packageDisplayName cannot be specified with output of json';
            }
            if (args.options.packageSourceEnvironment) {
                return 'packageSourceEnvironment cannot be specified with output of json';
            }
        }
        if (args.options.path && !fs.existsSync(path.dirname(args.options.path))) {
            return 'Specified path where to save the file does not exist';
        }
        return true;
    });
};
export default new FlowExportCommand();
//# sourceMappingURL=flow-export.js.map