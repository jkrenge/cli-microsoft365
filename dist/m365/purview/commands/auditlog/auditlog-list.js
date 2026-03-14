var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PurviewAuditLogListCommand_instances, _PurviewAuditLogListCommand_initTelemetry, _PurviewAuditLogListCommand_initOptions, _PurviewAuditLogListCommand_initValidators;
import Auth from '../../../../Auth.js';
import request from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { validation } from '../../../../utils/validation.js';
import O365MgmtCommand from '../../../base/O365MgmtCommand.js';
import commands from '../../commands.js';
class PurviewAuditLogListCommand extends O365MgmtCommand {
    get name() {
        return commands.AUDITLOG_LIST;
    }
    get description() {
        return 'List audit logs within your tenant';
    }
    defaultProperties() {
        return ['CreationTime', 'UserId', 'Operation', 'ObjectId'];
    }
    constructor() {
        super();
        _PurviewAuditLogListCommand_instances.add(this);
        this.contentTypeOptions = ['AzureActiveDirectory', 'Exchange', 'SharePoint', 'General', 'DLP'];
        __classPrivateFieldGet(this, _PurviewAuditLogListCommand_instances, "m", _PurviewAuditLogListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PurviewAuditLogListCommand_instances, "m", _PurviewAuditLogListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PurviewAuditLogListCommand_instances, "m", _PurviewAuditLogListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        // If we don't create a now object, start and end date can be an few extra ms apart due to execution time between code lines
        const now = new Date();
        try {
            let startTime;
            if (args.options.startTime) {
                startTime = new Date(args.options.startTime);
            }
            else {
                startTime = new Date(now);
                startTime.setDate(startTime.getDate() - 1);
            }
            const endTime = args.options.endTime ? new Date(args.options.endTime) : new Date(now);
            if (this.verbose) {
                await logger.logToStderr(`Getting audit logs for content type '${args.options.contentType}' within a time frame from '${startTime.toISOString()}' to '${endTime.toISOString()}'.`);
            }
            const tenantId = accessToken.getTenantIdFromAccessToken(Auth.connection.accessTokens[Auth.defaultResource].accessToken);
            const contentTypeValue = args.options.contentType === 'DLP' ? 'DLP.All' : 'Audit.' + args.options.contentType;
            await this.ensureSubscription(tenantId, contentTypeValue);
            if (this.verbose) {
                await logger.logToStderr(`'${args.options.contentType}' subscription is active.`);
            }
            const contentUris = [];
            for (const time = startTime; time < endTime; time.setDate(time.getDate() + 1)) {
                const differenceInMs = endTime.getTime() - time.getTime();
                const endTimeBatch = new Date(time.getTime() + Math.min(differenceInMs, 1000 * 60 * 60 * 24)); // ms difference cannot be greater than 1 day
                if (this.verbose) {
                    await logger.logToStderr(`Get content URIs for date range from '${time.toISOString()}' to '${endTimeBatch.toISOString()}'.`);
                }
                const contentUrisBatch = await this.getContentUris(tenantId, contentTypeValue, time, endTimeBatch);
                contentUris.push(...contentUrisBatch);
            }
            if (this.verbose) {
                await logger.logToStderr(`Get content from ${contentUris.length} content URIs.`);
            }
            const logs = await this.getContent(logger, contentUris);
            const sortedLogs = logs.sort(this.auditLogsCompare);
            await logger.log(sortedLogs);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async ensureSubscription(tenantId, contentType) {
        const requestOptions = {
            url: `${this.resource}/api/v1.0/${tenantId}/activity/feed/subscriptions/list`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        const subscriptions = await request.get(requestOptions);
        if (subscriptions.some(s => s.contentType === contentType && s.status === 'enabled')) {
            return;
        }
        requestOptions.url = `${this.resource}/api/v1.0/${tenantId}/activity/feed/subscriptions/start?contentType=${contentType}`;
        const subscription = await request.post(requestOptions);
        if (subscription.status !== 'enabled') {
            throw `Unable to start subscription '${contentType}'`;
        }
    }
    async getContentUris(tenantId, contentType, startTime, endTime) {
        const contentUris = [];
        const requestOptions = {
            url: `${this.resource}/api/v1.0/${tenantId}/activity/feed/subscriptions/content?contentType=${contentType}&startTime=${startTime.toISOString()}&endTime=${endTime.toISOString()}`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json',
            fullResponse: true
        };
        do {
            const response = await request.get(requestOptions);
            const uris = response.data.map(d => d.contentUri);
            contentUris.push(...uris);
            requestOptions.url = response.headers.nextpageuri;
        } while (requestOptions.url);
        return contentUris;
    }
    async getContent(logger, contentUris) {
        const logs = [];
        const batchSize = 30;
        for (let i = 0; i < contentUris.length; i += batchSize) {
            const contentUrisBatch = contentUris.slice(i, i + batchSize);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving content from next ${contentUrisBatch.length} content URIs. Progress: ${Math.round(i / contentUris.length * 100)}%`);
            }
            const batchResult = await Promise.all(contentUrisBatch.map(uri => request.get({
                url: uri,
                headers: {
                    accept: 'application/json'
                },
                responseType: 'json'
            })));
            batchResult.forEach(res => logs.push(...res));
        }
        return logs;
    }
    auditLogsCompare(a, b) {
        if (a.CreationTime < b.CreationTime) {
            return -1;
        }
        if (a.CreationTime > b.CreationTime) {
            return 1;
        }
        return 0;
    }
}
_PurviewAuditLogListCommand_instances = new WeakSet(), _PurviewAuditLogListCommand_initTelemetry = function _PurviewAuditLogListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            startTime: typeof args.options.startTime !== 'undefined',
            endTime: typeof args.options.endTime !== 'undefined'
        });
    });
}, _PurviewAuditLogListCommand_initOptions = function _PurviewAuditLogListCommand_initOptions() {
    this.options.unshift({
        option: '--contentType <contentType>',
        autocomplete: this.contentTypeOptions
    }, {
        option: '--startTime [startTime]'
    }, {
        option: '--endTime [endTime]'
    });
}, _PurviewAuditLogListCommand_initValidators = function _PurviewAuditLogListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (this.contentTypeOptions.indexOf(args.options.contentType) === -1) {
            return `'${args.options.contentType}' is not a valid contentType value. Allowed values: ${this.contentTypeOptions}.`;
        }
        if (args.options.startTime) {
            if (!validation.isValidISODateTime(args.options.startTime)) {
                return `'${args.options.startTime}' is not a valid ISO date time string.`;
            }
            const lowerDateLimit = new Date();
            lowerDateLimit.setDate(lowerDateLimit.getDate() - 7);
            lowerDateLimit.setHours(lowerDateLimit.getHours() - 1); // Min date is 7 days ago, however there seems to be an 1h margin
            if (new Date(args.options.startTime) < lowerDateLimit) {
                return 'startTime value cannot be more than 7 days in the past.';
            }
        }
        if (args.options.endTime) {
            if (!validation.isValidISODateTime(args.options.endTime)) {
                return `'${args.options.endTime}' is not a valid ISO date time string.`;
            }
            if (new Date(args.options.endTime) > new Date()) {
                return 'endTime value cannot be in the future.';
            }
        }
        if (args.options.startTime && args.options.endTime && new Date(args.options.startTime) >= new Date(args.options.endTime)) {
            return 'startTime value must be before endTime.';
        }
        return true;
    });
};
export default new PurviewAuditLogListCommand();
//# sourceMappingURL=auditlog-list.js.map