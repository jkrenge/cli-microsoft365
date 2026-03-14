var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SearchCommand_instances, _SearchCommand_initTelemetry, _SearchCommand_initOptions, _SearchCommand_initValidators;
import request from '../../request.js';
import GraphCommand from '../base/GraphCommand.js';
import commands from './commands.js';
class SearchCommand extends GraphCommand {
    get name() {
        return commands.SEARCH;
    }
    get description() {
        return 'Uses the Microsoft Search to query Microsoft 365 data';
    }
    constructor() {
        super();
        _SearchCommand_instances.add(this);
        this.allowedScopes = ['chatMessage', 'message', 'event', 'drive', 'driveItem', 'list', 'listItem', 'site', 'bookmark', 'acronym', 'person'];
        __classPrivateFieldGet(this, _SearchCommand_instances, "m", _SearchCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SearchCommand_instances, "m", _SearchCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SearchCommand_instances, "m", _SearchCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let result;
        const searchHits = [];
        try {
            let allResults = args.options.allResults ? args.options.allResults : false;
            let startIndex = args.options.startIndex ? args.options.startIndex : 0;
            const pageSize = args.options.pageSize ? args.options.pageSize : 25;
            do {
                const requestOptions = {
                    url: `${this.resource}/v1.0/search/query`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json',
                    data: {
                        requests: [
                            {
                                "entityTypes": args.options.scopes.split(',').map(scope => scope.trim()),
                                "query": {
                                    "queryString": args.options.queryText ?? '*'
                                },
                                "enableTopResults": args.options.enableTopResults,
                                "from": startIndex,
                                "size": pageSize,
                                "fields": this.getProperties(args.options),
                                "sortProperties": this.getSortProperties(args.options),
                                "queryAlterationOptions": {
                                    "enableModification": args.options.enableSpellingModification,
                                    "enableSuggestion": args.options.enableSpellingSuggestion
                                }
                            }
                        ]
                    }
                };
                const response = await request.post(requestOptions);
                result = response.value[0];
                if (allResults && result.hitsContainers) {
                    allResults = result.hitsContainers[0].moreResultsAvailable;
                }
                if (allResults) {
                    startIndex += pageSize;
                }
                if (result.hitsContainers && result.hitsContainers[0].hits) {
                    searchHits.push(...result.hitsContainers[0].hits);
                }
            } while (allResults);
            if (args.options.resultsOnly) {
                await logger.log(searchHits);
            }
            else {
                if (result.hitsContainers && result.hitsContainers[0].hits) {
                    result.hitsContainers[0].hits = searchHits;
                }
                await logger.log(result);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getProperties(options) {
        if (!options.select) {
            return undefined;
        }
        return options.select.split(',').map(prop => prop.trim());
    }
    getSortProperties(options) {
        if (!options.sortBy) {
            return undefined;
        }
        const properties = options.sortBy.split(',').map(prop => prop.trim()).map(property => {
            const sortDefinitions = property.split(':');
            const name = sortDefinitions[0];
            let isDescending = false;
            if (sortDefinitions.length === 2) {
                const order = sortDefinitions[1].trim();
                isDescending = order === 'desc';
            }
            return {
                "name": name,
                "isDescending": isDescending
            };
        });
        return properties;
    }
}
_SearchCommand_instances = new WeakSet(), _SearchCommand_initTelemetry = function _SearchCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            queryText: typeof args.options.queryText !== 'undefined',
            startIndex: typeof args.options.startIndex !== 'undefined',
            pageSize: typeof args.options.pageSize !== 'undefined',
            allResults: !!args.options.allResults,
            resultsOnly: !!args.options.resultsOnly,
            enableTopResults: !!args.options.enableTopResults,
            select: typeof args.options.select !== 'undefined',
            sortBy: typeof args.options.sortBy !== 'undefined',
            enableSpellingSuggestion: !!args.options.enableSpellingSuggestion,
            enableSpellingModification: !!args.options.enableSpellingModification
        });
    });
}, _SearchCommand_initOptions = function _SearchCommand_initOptions() {
    this.options.unshift({
        option: '-q, --queryText [queryText]'
    }, {
        option: '-s, --scopes <scopes>',
        autocomplete: this.allowedScopes
    }, {
        option: '--startIndex [startIndex]'
    }, {
        option: '--pageSize [pageSize]'
    }, {
        option: '--allResults'
    }, {
        option: '--resultsOnly'
    }, {
        option: '--enableTopResults'
    }, {
        option: '--select [select]'
    }, {
        option: '--sortBy [sortBy]'
    }, {
        option: '--enableSpellingSuggestion'
    }, {
        option: '--enableSpellingModification'
    });
}, _SearchCommand_initValidators = function _SearchCommand_initValidators() {
    this.validators.push(async (args) => {
        const scopes = args.options.scopes.split(',').map(x => x.trim());
        if (!scopes.every(scope => this.allowedScopes.indexOf(scope) > -1)) {
            const invalidScope = scopes.find(scope => this.allowedScopes.indexOf(scope) === -1);
            return `'${invalidScope}'' is not a valid scope. Allowed scopes are ${this.allowedScopes.join(', ')}.`;
        }
        if (args.options.startIndex !== undefined && args.options.startIndex < 0) {
            return `'${args.options.startIndex}' is not a valid value for option 'startIndex'. Start index must be greater or equal to 0.`;
        }
        if (args.options.pageSize !== undefined && (args.options.pageSize < 1 || args.options.pageSize > 500)) {
            return `'${args.options.pageSize}' is not a valid value for option 'pageSize'. Page size must be between 1 and 500.`;
        }
        if (args.options.sortBy && scopes.some(scope => scope === 'message' || scope === 'event')) {
            return 'Sorting the results is not supported for messages and events.';
        }
        if (args.options.enableTopResults &&
            ((scopes.length === 1 && scopes.indexOf('message') === -1 && scopes.indexOf('chatMessage') === -1) ||
                (scopes.length === 2) && !(scopes.indexOf('message') > -1 && scopes.indexOf('chatMessage') > -1))) {
            return 'Top results are only supported for messages and chat messages.';
        }
        return true;
    });
};
export default new SearchCommand();
//# sourceMappingURL=search.js.map