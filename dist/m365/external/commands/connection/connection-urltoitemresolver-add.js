var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _ExternalConnectionUrlToItemResolverAddCommand_instances, _ExternalConnectionUrlToItemResolverAddCommand_initOptions;
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class ExternalConnectionUrlToItemResolverAddCommand extends GraphCommand {
    get name() {
        return commands.CONNECTION_URLTOITEMRESOLVER_ADD;
    }
    get description() {
        return 'Adds a URL to item resolver to an external connection';
    }
    constructor() {
        super();
        _ExternalConnectionUrlToItemResolverAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _ExternalConnectionUrlToItemResolverAddCommand_instances, "m", _ExternalConnectionUrlToItemResolverAddCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        const baseUrls = args.options.baseUrls.split(',').map(b => b.trim());
        const itemIdResolver = {
            itemId: args.options.itemId,
            priority: args.options.priority,
            urlMatchInfo: {
                baseUrls: baseUrls,
                urlPattern: args.options.urlPattern
            }
        };
        // not a part of the type definition, but required by the API
        itemIdResolver['@odata.type'] = '#microsoft.graph.externalConnectors.itemIdResolver';
        const requestOptions = {
            url: `${this.resource}/v1.0/external/connections/${args.options.externalConnectionId}`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json'
            },
            responseType: 'json',
            data: {
                activitySettings: {
                    urlToItemResolvers: [itemIdResolver]
                }
            }
        };
        try {
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_ExternalConnectionUrlToItemResolverAddCommand_instances = new WeakSet(), _ExternalConnectionUrlToItemResolverAddCommand_initOptions = function _ExternalConnectionUrlToItemResolverAddCommand_initOptions() {
    this.options.unshift({
        option: '-c, --externalConnectionId <externalConnectionId>'
    }, {
        option: '--baseUrls <baseUrls>'
    }, {
        option: '--urlPattern <urlPattern>'
    }, {
        option: '-i, --itemId <itemId>'
    }, {
        option: '-p, --priority <priority>'
    });
};
export default new ExternalConnectionUrlToItemResolverAddCommand();
//# sourceMappingURL=connection-urltoitemresolver-add.js.map