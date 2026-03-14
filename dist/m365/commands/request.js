var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _RequestCommand_instances, _RequestCommand_initTelemetry, _RequestCommand_initOptions, _RequestCommand_initValidators;
import fs from 'fs';
import path from 'path';
import auth from '../../Auth.js';
import Command from '../../Command.js';
import request from '../../request.js';
import { mailSenderWhitelist } from '../../utils/mailSenderWhitelist.js';
import commands from './commands.js';
class RequestCommand extends Command {
    get name() {
        return commands.REQUEST;
    }
    get description() {
        return 'Executes the specified web request using CLI for Microsoft 365';
    }
    allowUnknownOptions() {
        return true;
    }
    constructor() {
        super();
        _RequestCommand_instances.add(this);
        this.allowedMethods = ['get', 'post', 'put', 'patch', 'delete', 'head', 'options'];
        __classPrivateFieldGet(this, _RequestCommand_instances, "m", _RequestCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _RequestCommand_instances, "m", _RequestCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _RequestCommand_instances, "m", _RequestCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.debug) {
            await logger.logToStderr(`Preparing request...`);
        }
        try {
            const url = this.resolveUrlTokens(args.options.url);
            const method = (args.options.method || 'get').toUpperCase();
            const headers = {};
            this.addUnknownOptionsToPayload(headers, args.options);
            if (!headers.accept) {
                headers.accept = 'application/json';
            }
            if (args.options.resource) {
                headers['x-resource'] = args.options.resource;
            }
            const config = {
                url: url,
                headers,
                method,
                data: args.options.body
            };
            if (headers.accept.toString().startsWith('application/json')) {
                config.responseType = 'json';
            }
            if (args.options.filePath) {
                config.responseType = 'stream';
            }
            if (this.verbose) {
                await logger.logToStderr(`Executing request...`);
            }
            if (args.options.filePath) {
                const file = await request.execute(config);
                const filePath = await new Promise((resolve, reject) => {
                    const writer = fs.createWriteStream(args.options.filePath);
                    file.data.pipe(writer);
                    writer.on('error', err => {
                        reject(err);
                    });
                    writer.on('close', () => {
                        resolve(args.options.filePath);
                    });
                });
                if (this.verbose) {
                    await logger.logToStderr(`File saved to path ${filePath}`);
                }
            }
            else {
                let res = await request.execute(config);
                if (mailSenderWhitelist.shouldFilterInInteractiveMode() &&
                    this.isGraphMailMessageRequest(url) &&
                    typeof res === 'object' &&
                    res !== null) {
                    const filterResult = mailSenderWhitelist.filterResponse(res);
                    res = filterResult.filteredResponse;
                    await mailSenderWhitelist.logFilteredSummary(logger, filterResult.filteredCount);
                }
                await logger.log(res);
            }
        }
        catch (err) {
            this.handleError(err);
        }
    }
    resolveUrlTokens(url) {
        if (url.startsWith('@graphbeta')) {
            return url.replace('@graphbeta', 'https://graph.microsoft.com/beta');
        }
        if (url.startsWith('@graph')) {
            return url.replace('@graph', 'https://graph.microsoft.com/v1.0');
        }
        if (url.startsWith('@spo')) {
            if (auth.connection.spoUrl) {
                return url.replace('@spo', auth.connection.spoUrl);
            }
            throw `SharePoint root site URL is unknown. Please set your SharePoint URL using command 'spo set'.`;
        }
        return url;
    }
    isGraphMailMessageRequest(url) {
        try {
            const parsedUrl = new URL(url);
            if (parsedUrl.hostname.toLowerCase() !== 'graph.microsoft.com') {
                return false;
            }
            const pathName = parsedUrl.pathname.toLowerCase();
            if (!pathName.startsWith('/v1.0/') && !pathName.startsWith('/beta/')) {
                return false;
            }
            const graphPath = pathName.replace(/^\/(v1\.0|beta)\//, '');
            const isSupportedMailboxPath = graphPath.startsWith('me/messages') ||
                graphPath.startsWith('me/mailfolders/') ||
                graphPath.startsWith('users/');
            return isSupportedMailboxPath && graphPath.includes('/messages');
        }
        catch {
            return false;
        }
    }
}
_RequestCommand_instances = new WeakSet(), _RequestCommand_initTelemetry = function _RequestCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            method: args.options.method || 'get',
            resource: typeof args.options.resource !== 'undefined',
            accept: args.options.accept || 'application/json',
            body: typeof args.options.body !== 'undefined',
            filePath: typeof args.options.filePath !== 'undefined'
        });
        this.trackUnknownOptions(this.telemetryProperties, args.options);
    });
}, _RequestCommand_initOptions = function _RequestCommand_initOptions() {
    this.options.unshift({
        option: '-u, --url <url>'
    }, {
        option: '-m, --method [method]',
        autocomplete: this.allowedMethods
    }, {
        option: '-r, --resource [resource]'
    }, {
        option: '-b, --body [body]'
    }, {
        option: '-p, --filePath [filePath]'
    });
}, _RequestCommand_initValidators = function _RequestCommand_initValidators() {
    this.validators.push(async (args) => {
        const currentMethod = args.options.method || 'get';
        if (this.allowedMethods.indexOf(currentMethod) === -1) {
            return `${currentMethod} is not a valid value for method. Allowed values: ${this.allowedMethods.join(', ')}`;
        }
        if (args.options.body && (!args.options.method || args.options.method === 'get')) {
            return 'Specify a different method when using body';
        }
        if (args.options.body && !args.options['content-type']) {
            return 'Specify the content-type when using body';
        }
        if (args.options.filePath && !fs.existsSync(path.dirname(args.options.filePath))) {
            return 'The location specified in the filePath does not exist';
        }
        return true;
    });
};
export default new RequestCommand();
//# sourceMappingURL=request.js.map