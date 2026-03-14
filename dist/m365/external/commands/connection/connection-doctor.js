var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _ExternalConnectionDoctorCommand_instances, _a, _ExternalConnectionDoctorCommand_initOptions, _ExternalConnectionDoctorCommand_initValidators;
import os from 'os';
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { CheckStatus, formatting } from '../../../../utils/formatting.js';
class ExternalConnectionDoctorCommand extends GraphCommand {
    get name() {
        return commands.CONNECTION_DOCTOR;
    }
    get description() {
        return 'Checks if the external connection is correctly configured for use with the specified Microsoft 365 experience';
    }
    constructor() {
        super();
        _ExternalConnectionDoctorCommand_instances.add(this);
        this.checksStatus = [];
        __classPrivateFieldGet(this, _ExternalConnectionDoctorCommand_instances, "m", _ExternalConnectionDoctorCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _ExternalConnectionDoctorCommand_instances, "m", _ExternalConnectionDoctorCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const ux = args.options.ux ?? 'all';
        const output = args.options.output;
        this.checksStatus = [];
        const show = output === 'text';
        let checks = [
            {
                id: 'loadExternalConnection',
                text: 'Load connection',
                fn: this.loadConnection,
                type: 'required'
            },
            {
                id: 'loadSchema',
                text: 'Load schema',
                fn: this.loadSchema,
                type: 'required'
            }
        ];
        if (ux === 'copilot' || ux === 'all') {
            checks.push({
                id: 'copilotRequiredSemanticLabels',
                text: 'Required semantic labels',
                fn: this.checkCopilotRequiredSemanticLabels,
                type: 'required'
            }, {
                id: 'searchableProperties',
                text: 'Searchable properties',
                fn: this.checkSearchableProperties,
                type: 'required'
            }, {
                id: 'contentIngested',
                text: 'Items have content ingested',
                fn: this.checkContentIngested,
                type: 'required'
            }, {
                id: 'enabledForInlineResults',
                text: 'Connection configured for inline results',
                type: 'required'
            }, {
                id: 'itemsHaveActivities',
                text: 'Items have activities recorded',
                type: 'recommended'
            }, {
                id: 'meaningfulNameAndDescription',
                text: 'Meaningful connection name and description',
                type: 'required'
            });
        }
        if (ux === 'search' || ux === 'all') {
            checks.push({
                id: 'semanticLabels',
                text: 'Semantic labels',
                fn: this.checkSemanticLabels,
                type: 'recommended'
            }, {
                id: 'searchableProperties',
                text: 'Searchable properties',
                fn: this.checkSearchableProperties,
                type: 'recommended'
            }, {
                id: 'resultType',
                text: 'Result type',
                fn: this.checkResultType,
                type: 'recommended'
            }, {
                id: 'contentIngested',
                text: 'Items have content ingested',
                fn: this.checkContentIngested,
                type: 'recommended'
            }, {
                id: 'itemsHaveActivities',
                text: 'Items have activities recorded',
                type: 'recommended'
            });
        }
        checks.push({
            id: 'urlToItemResolver',
            text: 'urlToItemResolver configured',
            fn: this.checkUrlToItemResolverConfigured,
            type: 'recommended'
        });
        // filter out duplicate checks based on their IDs
        checks = checks.filter((check, index, self) => self.findIndex(c => c.id === check.id) === index);
        for (const check of checks) {
            if (this.debug) {
                await logger.logToStderr(`Running check ${check.id}...`);
            }
            // only automated checks have functions
            if (!check.fn) {
                if (show) {
                    await logger.log(formatting.getStatus(CheckStatus.Information, `${check.text} (manual)`));
                }
                this.checksStatus.push({
                    ...check,
                    status: 'manual'
                });
                continue;
            }
            const result = await check.fn.bind(this)(check.id, args);
            this.checksStatus.push({ ...check, ...result });
            if (result.status === 'passed') {
                if (show) {
                    await logger.log(formatting.getStatus(CheckStatus.Success, check.text));
                }
                continue;
            }
            if (result.status === 'failed') {
                if (show) {
                    const message = `${check.text}: ${result.errorMessage}`;
                    if (check.type === 'required') {
                        await logger.log(formatting.getStatus(CheckStatus.Failure, message));
                    }
                    else {
                        await logger.log(formatting.getStatus(CheckStatus.Warning, check.text));
                    }
                }
                if (result.shouldStop) {
                    break;
                }
            }
        }
        if (show || output === 'none') {
            return;
        }
        this.checksStatus.forEach(s => {
            delete s.data;
            delete s.fn;
            delete s.shouldStop;
        });
        if (output === 'json' || output === 'md') {
            await logger.log(this.checksStatus);
            return;
        }
        if (output === 'csv') {
            this.checksStatus.forEach(r => {
                // we need to set errorMessage to empty string so that it's not
                // removed from the CSV output
                r.errorMessage = r.errorMessage ?? '';
            });
            await logger.log(this.checksStatus);
        }
    }
    getMdOutput(logStatement, command, options) {
        const output = [
            `# ${command.getCommandName()} ${Object.keys(options).filter(o => o !== 'output').map(k => `--${k} "${options[k]}"`).join(' ')}`, os.EOL,
            os.EOL,
            `Date: ${(new Date().toLocaleDateString())}`, os.EOL,
            os.EOL
        ];
        if (logStatement && logStatement.length > 0) {
            const properties = ['text', 'type', 'status', 'errorMessage'];
            output.push('Check|Type|Status|Error message', os.EOL);
            output.push(':----|:--:|:----:|:------------', os.EOL);
            logStatement.forEach(r => {
                output.push(properties.map(p => r[p] ?? '').join('|'), os.EOL);
            });
            logStatement.push(os.EOL);
        }
        return output.join('').trimEnd();
    }
    async loadConnection(id, args) {
        const requestOptions = {
            url: `${this.resource}/v1.0/external/connections/${args.options.id}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const externalConnection = await request.get(requestOptions);
            return {
                id,
                data: externalConnection,
                status: 'passed'
            };
        }
        catch (ex) {
            return {
                id,
                error: ex?.response?.data?.error?.innerError?.message,
                errorMessage: 'Connection not found',
                shouldStop: true,
                status: 'failed'
            };
        }
    }
    async loadSchema(id, args) {
        const requestOptions = {
            url: `${this.resource}/v1.0/external/connections/${args.options.id}/schema`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                prefer: 'include-unknown-enum-members'
            },
            responseType: 'json'
        };
        try {
            const schema = await request.get(requestOptions);
            return {
                id,
                data: schema,
                status: 'passed'
            };
        }
        catch (ex) {
            return {
                id,
                errorMessage: 'Schema not found',
                error: ex?.response?.data?.error?.innerError?.message,
                shouldStop: true,
                status: 'failed'
            };
        }
    }
    async checkCopilotRequiredSemanticLabels(id) {
        const schema = this.checksStatus.find(r => r.id === 'loadSchema').data;
        const requiredLabels = ['title', 'url', 'iconUrl'];
        for (const label of requiredLabels) {
            if (!schema.properties?.find(p => p.labels?.find(l => l.toString() === label))) {
                return {
                    id,
                    errorMessage: `Missing label ${label}`,
                    status: 'failed'
                };
            }
        }
        return {
            id,
            status: 'passed'
        };
    }
    async checkSearchableProperties(id) {
        const schema = this.checksStatus.find(r => r.id === 'loadSchema').data;
        if (!schema.properties?.some(p => p.isSearchable)) {
            return {
                id,
                errorMessage: 'Schema does not have any searchable properties',
                status: 'failed'
            };
        }
        return {
            id,
            status: 'passed'
        };
    }
    async checkContentIngested(id, args) {
        try {
            // find items that belong to the connection
            const searchRequestOptions = {
                url: `${this.resource}/v1.0/search/query`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    requests: [
                        {
                            entityTypes: [
                                'externalItem'
                            ],
                            contentSources: [
                                `/external/connections/${args.options.id}`
                            ],
                            query: {
                                queryString: '*'
                            },
                            from: 0,
                            size: 1
                        }
                    ]
                }
            };
            const result = await request.post(searchRequestOptions);
            const hit = result.value?.[0].hitsContainers?.[0]?.hits?.[0];
            if (!hit) {
                return {
                    id,
                    errorMessage: 'No items found that belong to the connection',
                    status: 'failed'
                };
            }
            // something@tenant,itemId
            const itemId = hit.resource?.properties?.substrateContentDomainId?.split(',')?.[1];
            if (!itemId) {
                return {
                    id,
                    errorMessage: 'Item does not have substrateContentDomainId property or the property is invalid',
                    status: 'failed'
                };
            }
            const externalItemRequestOptions = {
                url: `${this.resource}/v1.0/external/connections/${args.options.id}/items/${itemId}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const externalItem = await request.get(externalItemRequestOptions);
            if (!externalItem.content?.value) {
                return {
                    id,
                    data: externalItem,
                    errorMessage: 'Item does not have content or content is empty',
                    status: 'failed'
                };
            }
            return {
                id,
                data: externalItem,
                status: 'passed'
            };
        }
        catch (ex) {
            return {
                id,
                error: ex?.response?.data?.error?.innerError?.message,
                errorMessage: 'Error while checking if content is ingested',
                status: 'failed'
            };
        }
    }
    async checkUrlToItemResolverConfigured(id) {
        const externalConnection = this.checksStatus.find(r => r.id === 'loadExternalConnection').data;
        if (!externalConnection.activitySettings?.urlToItemResolvers?.some(r => r)) {
            return {
                id,
                errorMessage: 'urlToItemResolver is not configured',
                status: 'failed'
            };
        }
        return {
            id,
            status: 'passed'
        };
    }
    async checkSemanticLabels(id) {
        const schema = this.checksStatus.find(r => r.id === 'loadSchema').data;
        const hasLabels = schema.properties?.some(p => p.labels?.some(l => l));
        if (!hasLabels) {
            return {
                id,
                errorMessage: `Schema does not have semantic labels`,
                status: 'failed'
            };
        }
        return {
            id,
            status: 'passed'
        };
    }
    async checkResultType(id) {
        const externalConnection = this.checksStatus.find(r => r.id === 'loadExternalConnection').data;
        if (!externalConnection.searchSettings?.searchResultTemplates?.some(t => t)) {
            return {
                id,
                errorMessage: `Connection has no result types`,
                status: 'failed'
            };
        }
        return {
            id,
            status: 'passed'
        };
    }
}
_a = ExternalConnectionDoctorCommand, _ExternalConnectionDoctorCommand_instances = new WeakSet(), _ExternalConnectionDoctorCommand_initOptions = function _ExternalConnectionDoctorCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '--ux [ux]',
        autocomplete: _a.supportedUx
    });
}, _ExternalConnectionDoctorCommand_initValidators = function _ExternalConnectionDoctorCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.ux) {
            if (!_a.supportedUx.find(u => u === args.options.ux)) {
                return `${args.options.ux} is not a valid UX. Allowed values are ${_a.supportedUx.join(', ')}`;
            }
        }
        return true;
    });
};
ExternalConnectionDoctorCommand.supportedUx = ['copilot', 'search', 'all'];
export default new ExternalConnectionDoctorCommand();
//# sourceMappingURL=connection-doctor.js.map