var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PaAppExportCommand_instances, _PaAppExportCommand_initTelemetry, _PaAppExportCommand_initOptions, _PaAppExportCommand_initValidators, _PaAppExportCommand_initTypes;
import fs from 'fs';
import path from 'path';
import { setTimeout } from 'timers/promises';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
class PaAppExportCommand extends PowerPlatformCommand {
    get name() {
        return commands.APP_EXPORT;
    }
    get description() {
        return 'Exports the specified Power App';
    }
    constructor() {
        super();
        _PaAppExportCommand_instances.add(this);
        this.pollingInterval = 5000;
        __classPrivateFieldGet(this, _PaAppExportCommand_instances, "m", _PaAppExportCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PaAppExportCommand_instances, "m", _PaAppExportCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PaAppExportCommand_instances, "m", _PaAppExportCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PaAppExportCommand_instances, "m", _PaAppExportCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const location = await this.exportPackage(args, logger);
            const packageLink = await this.getPackageLink(logger, location);
            let filename = args.options.name;
            if (args.options.packageDisplayName) {
                //Replace all illegal characters from the file name
                const illegalCharsRegEx = /[\\/:*?"<>|]/g;
                filename = args.options.packageDisplayName.replace(illegalCharsRegEx, '_');
            }
            const requestOptions = {
                url: packageLink,
                // Set responseType to arraybuffer, otherwise binary data will be encoded
                // to utf8 and binary data is corrupt
                responseType: 'arraybuffer',
                headers: {
                    'x-anonymous': true
                }
            };
            const file = await request.get(requestOptions);
            let path = args.options.path || './';
            if (!path.endsWith('/')) {
                path += '/';
            }
            path += `${filename}.zip`;
            fs.writeFileSync(path, file, 'binary');
            if (this.verbose) {
                await logger.logToStderr(`File saved to path '${path}'`);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getPackageResources(args, logger) {
        if (this.verbose) {
            await logger.logToStderr('Getting the Microsoft Power App resources...');
        }
        const requestOptions = {
            url: `${this.resource}/providers/Microsoft.BusinessAppPlatform/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/listPackageResources?api-version=2016-11-01`,
            headers: {
                accept: 'application/json'
            },
            data: {
                baseResourceIds: [
                    `/providers/Microsoft.PowerApps/apps/${args.options.name}`
                ]
            },
            responseType: 'json'
        };
        const response = await request.post(requestOptions);
        Object.keys(response.resources).forEach((key) => {
            response.resources[key].suggestedCreationType = 'Update';
        });
        return response.resources;
    }
    async exportPackage(args, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Initiating package export for Microsoft Power App ${args.options.name}...`);
        }
        const resources = await this.getPackageResources(args, logger);
        const requestOptions = {
            url: `${this.resource}/providers/Microsoft.BusinessAppPlatform/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/exportPackage?api-version=2016-11-01`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json',
            data: {
                includedResourceIds: [
                    `/providers/Microsoft.PowerApps/apps/${args.options.name}`
                ],
                details: {
                    creator: args.options.packageCreatedBy,
                    description: args.options.packageDescription,
                    displayName: args.options.packageDisplayName || args.options.name,
                    sourceEnvironment: args.options.packageSourceEnvironment
                },
                resources: resources
            },
            fullResponse: true
        };
        const response = await request.post(requestOptions);
        return response.headers.location;
    }
    async getPackageLink(logger, location) {
        if (this.verbose) {
            await logger.logToStderr('Retrieving the package link and waiting on the exported package.');
        }
        let status;
        let link;
        const requestOptions = {
            url: location,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        do {
            const response = await request.get(requestOptions);
            status = response.properties.status;
            if (status === "Succeeded") {
                link = response.properties.packageLink.value;
            }
            else {
                await setTimeout(this.pollingInterval);
            }
            if (this.verbose) {
                await logger.logToStderr(`Current status of the get package link: ${status}`);
            }
        } while (status === 'Running');
        return link;
    }
}
_PaAppExportCommand_instances = new WeakSet(), _PaAppExportCommand_initTelemetry = function _PaAppExportCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            packageDescription: typeof args.options.packageDescription !== 'undefined',
            packageCreatedBy: typeof args.options.packageCreatedBy !== 'undefined',
            packageSourceEnvironment: typeof args.options.packageSourceEnvironment !== 'undefined',
            path: typeof args.options.path !== 'undefined'
        });
    });
}, _PaAppExportCommand_initOptions = function _PaAppExportCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '--packageDisplayName [packageDisplayName]'
    }, {
        option: '-d, --packageDescription [packageDescription]'
    }, {
        option: '-c, --packageCreatedBy [packageCreatedBy]'
    }, {
        option: '-s, --packageSourceEnvironment [packageSourceEnvironment]'
    }, {
        option: '-p, --path [path]'
    });
}, _PaAppExportCommand_initValidators = function _PaAppExportCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.name)) {
            return `${args.options.name} is not a valid GUID for option name`;
        }
        if (args.options.path && !fs.existsSync(path.dirname(args.options.path))) {
            return 'Specified path where to save the file does not exist';
        }
        return true;
    });
}, _PaAppExportCommand_initTypes = function _PaAppExportCommand_initTypes() {
    this.types.string.push('name', 'environmentName', 'packageDisplayName', 'packageDescription', 'packageCreatedBy', 'packageSourceEnvironment', 'path');
};
export default new PaAppExportCommand();
//# sourceMappingURL=app-export.js.map