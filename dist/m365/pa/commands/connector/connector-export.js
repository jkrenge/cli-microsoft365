var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PaConnectorExportCommand_instances, _PaConnectorExportCommand_initTelemetry, _PaConnectorExportCommand_initOptions, _PaConnectorExportCommand_initValidators;
import fs from 'fs';
import path from 'path';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import flowCommands from '../../../flow/commands.js';
import commands from '../../commands.js';
class PaConnectorExportCommand extends PowerAppsCommand {
    get name() {
        return commands.CONNECTOR_EXPORT;
    }
    get description() {
        return 'Exports the specified power automate or power apps custom connector';
    }
    alias() {
        return [flowCommands.CONNECTOR_EXPORT];
    }
    constructor() {
        super();
        _PaConnectorExportCommand_instances.add(this);
        __classPrivateFieldGet(this, _PaConnectorExportCommand_instances, "m", _PaConnectorExportCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PaConnectorExportCommand_instances, "m", _PaConnectorExportCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PaConnectorExportCommand_instances, "m", _PaConnectorExportCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const outputFolder = path.resolve(args.options.outputFolder || '.', args.options.name);
        const requestOptions = {
            url: `${this.resource}/providers/Microsoft.PowerApps/apis/${formatting.encodeQueryParameter(args.options.name)}?api-version=2016-11-01&$filter=environment%20eq%20%27${formatting.encodeQueryParameter(args.options.environmentName)}%27%20and%20IsCustomApi%20eq%20%27True%27`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        let connector;
        if (this.verbose) {
            await logger.logToStderr('Downloading connector...');
        }
        try {
            connector = await request.get(requestOptions);
            if (!connector.properties) {
                throw 'Properties not present in the api registration information.';
            }
            if (this.verbose) {
                await logger.logToStderr(`Creating output folder ${outputFolder}...`);
            }
            fs.mkdirSync(outputFolder);
            const settings = {
                apiDefinition: "apiDefinition.swagger.json",
                apiProperties: "apiProperties.json",
                connectorId: args.options.name,
                environment: args.options.environmentName,
                icon: "icon.png",
                powerAppsApiVersion: "2016-11-01",
                powerAppsUrl: "https://api.powerapps.com"
            };
            if (this.verbose) {
                await logger.logToStderr('Exporting settings...');
            }
            fs.writeFileSync(path.join(outputFolder, 'settings.json'), JSON.stringify(settings, null, 2), 'utf8');
            const propertiesWhitelist = [
                "connectionParameters",
                "iconBrandColor",
                "capabilities",
                "policyTemplateInstances"
            ];
            const apiProperties = {
                properties: JSON.parse(JSON.stringify(connector.properties))
            };
            Object.keys(apiProperties.properties).forEach(k => {
                if (propertiesWhitelist.indexOf(k) < 0) {
                    delete apiProperties.properties[k];
                }
            });
            if (this.verbose) {
                await logger.logToStderr('Exporting API properties...');
            }
            fs.writeFileSync(path.join(outputFolder, 'apiProperties.json'), JSON.stringify(apiProperties, null, 2), 'utf8');
            let swagger = '';
            if (connector.properties.apiDefinitions &&
                connector.properties.apiDefinitions.originalSwaggerUrl) {
                if (this.verbose) {
                    await logger.logToStderr(`Downloading swagger from ${connector.properties.apiDefinitions.originalSwaggerUrl}...`);
                }
                swagger = await request
                    .get({
                    url: connector.properties.apiDefinitions.originalSwaggerUrl,
                    headers: {
                        'x-anonymous': 'true'
                    }
                });
            }
            else {
                if (this.debug) {
                    await logger.logToStderr('originalSwaggerUrl not set. Skipping');
                }
            }
            if (swagger && swagger.length > 0) {
                if (this.debug) {
                    await logger.logToStderr('Downloaded swagger');
                    await logger.logToStderr(swagger);
                }
                if (this.verbose) {
                    await logger.logToStderr('Exporting swagger...');
                }
                fs.writeFileSync(path.join(outputFolder, 'apiDefinition.swagger.json'), swagger, 'utf8');
            }
            let icon = '';
            if (connector.properties.iconUri) {
                if (this.verbose) {
                    await logger.logToStderr(`Downloading icon from ${connector.properties.iconUri}...`);
                }
                icon = await request
                    .get({
                    url: connector.properties.iconUri,
                    responseType: 'arraybuffer',
                    headers: {
                        'x-anonymous': 'true'
                    }
                });
            }
            else {
                if (this.debug) {
                    await logger.logToStderr('iconUri not set. Skipping');
                }
            }
            if (icon) {
                if (this.verbose) {
                    await logger.logToStderr('Exporting icon...');
                }
                const iconBuffer = Buffer.from(icon, 'utf8');
                fs.writeFileSync(path.join(outputFolder, 'icon.png'), iconBuffer);
            }
            else {
                if (this.debug) {
                    await logger.logToStderr('No icon retrieved');
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PaConnectorExportCommand_instances = new WeakSet(), _PaConnectorExportCommand_initTelemetry = function _PaConnectorExportCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            outputFolder: typeof args.options.outputFolder !== 'undefined'
        });
    });
}, _PaConnectorExportCommand_initOptions = function _PaConnectorExportCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '-n, --name <name>'
    }, {
        option: '--outputFolder [outputFolder]'
    });
}, _PaConnectorExportCommand_initValidators = function _PaConnectorExportCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.outputFolder &&
            !fs.existsSync(path.resolve(args.options.outputFolder))) {
            return `Specified output folder ${args.options.outputFolder} doesn't exist`;
        }
        const outputFolder = path.resolve(args.options.outputFolder || '.', args.options.name);
        if (fs.existsSync(outputFolder)) {
            return `Connector output folder ${outputFolder} already exists`;
        }
        return true;
    });
};
export default new PaConnectorExportCommand();
//# sourceMappingURL=connector-export.js.map