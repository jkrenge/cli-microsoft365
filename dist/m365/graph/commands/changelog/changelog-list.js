var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _GraphChangelogListCommand_instances, _GraphChangelogListCommand_initTelemetry, _GraphChangelogListCommand_initOptions, _GraphChangelogListCommand_initValidators;
import { DOMParser } from '@xmldom/xmldom';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { md } from '../../../../utils/md.js';
import { validation } from '../../../../utils/validation.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';
class GraphChangelogListCommand extends AnonymousCommand {
    get name() {
        return commands.CHANGELOG_LIST;
    }
    get description() {
        return 'Gets an overview of specific API-level changes in Microsoft Graph v1.0 and beta';
    }
    defaultProperties() {
        return ['category', 'title', 'description'];
    }
    constructor() {
        super();
        _GraphChangelogListCommand_instances.add(this);
        this.allowedVersions = ['beta', 'v1.0'];
        this.allowedChangeTypes = ['Addition', 'Change', 'Deletion', 'Deprecation'];
        this.allowedServices = [
            'Applications', 'Calendar', 'Change notifications', 'Cloud communications',
            'Compliance', 'Cross-device experiences', 'Customer booking', 'Device and app management',
            'Education', 'Files', 'Financials', 'Groups',
            'Identity and access', 'Mail', 'Notes', 'Notifications',
            'People and workplace intelligence', 'Personal contacts', 'Reports', 'Search',
            'Security', 'Sites and lists', 'Tasks and plans', 'Teamwork',
            'To-do tasks', 'Users', 'Workbooks and charts'
        ];
        __classPrivateFieldGet(this, _GraphChangelogListCommand_instances, "m", _GraphChangelogListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _GraphChangelogListCommand_instances, "m", _GraphChangelogListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _GraphChangelogListCommand_instances, "m", _GraphChangelogListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const allowedChangeType = args.options.changeType && this.allowedChangeTypes.find(x => x.toLocaleLowerCase() === args.options.changeType.toLocaleLowerCase());
            const searchParam = args.options.changeType ? `/?filterBy=${allowedChangeType}` : '';
            const requestOptions = {
                url: `https://developer.microsoft.com/en-us/graph/changelog/rss${searchParam}`,
                headers: {
                    'accept': 'text/xml',
                    'x-anonymous': 'true'
                }
            };
            const output = await request.get(requestOptions);
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(output.toString(), "text/xml");
            const changelog = this.filterThroughOptions(args.options, this.mapChangelog(xmlDoc, args));
            await logger.log(changelog.items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    filterThroughOptions(options, changelog) {
        let items = changelog.items;
        if (options.services) {
            const allowedServices = this.allowedServices
                .filter(allowedService => options.services.toLocaleLowerCase().split(',').includes(allowedService.toLocaleLowerCase()));
            items = changelog.items.filter(item => allowedServices.includes(item.title));
        }
        if (options.versions) {
            const allowedVersions = this.allowedVersions
                .filter(allowedVersion => options.versions.toLocaleLowerCase().split(',').includes(allowedVersion.toLocaleLowerCase()));
            items = items.filter(item => allowedVersions.includes(item.category));
        }
        if (options.startDate) {
            const startDate = new Date(options.startDate);
            items = items.filter(item => item.pubDate >= startDate);
        }
        if (options.endDate) {
            const endDate = new Date(options.endDate);
            items = items.filter(item => item.pubDate <= endDate);
        }
        // Make sure everything is unique based on the item guid
        items = [...new Map(items.map((item) => [item.guid, item])).values()];
        changelog.items = items.sort((itemA, itemB) => Number(itemB.pubDate) - Number(itemA.pubDate));
        return changelog;
    }
    mapChangelog(xmlDoc, args) {
        const channel = xmlDoc.getElementsByTagName('channel').item(0);
        const changelog = {
            title: channel.getElementsByTagName('title').item(0).textContent,
            description: channel.getElementsByTagName('description').item(0).textContent,
            url: channel.getElementsByTagName('link').item(0).textContent,
            items: []
        };
        Array.from(xmlDoc.getElementsByTagName('item')).forEach((item) => {
            const description = cli.shouldTrimOutput(args.options.output) ?
                md.md2plain(item.getElementsByTagName('description').item(0).textContent, '') :
                item.getElementsByTagName('description').item(0).textContent;
            changelog.items.push({
                guid: item.getElementsByTagName('guid').item(0).textContent,
                category: item.getElementsByTagName('category').item(1).textContent,
                title: item.getElementsByTagName('title').item(0).textContent,
                description: cli.shouldTrimOutput(args.options.output) ?
                    description.length > 50 ? `${description.substring(0, 47)}...` : description :
                    description,
                pubDate: new Date(item.getElementsByTagName('pubDate').item(0).textContent)
            });
        });
        return changelog;
    }
}
_GraphChangelogListCommand_instances = new WeakSet(), _GraphChangelogListCommand_initTelemetry = function _GraphChangelogListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            versions: typeof args.options.versions !== 'undefined',
            changeType: typeof args.options.changeType !== 'undefined',
            services: typeof args.options.services !== 'undefined',
            startDate: typeof args.options.startDate !== 'undefined',
            endDate: typeof args.options.endDate !== 'undefined'
        });
    });
}, _GraphChangelogListCommand_initOptions = function _GraphChangelogListCommand_initOptions() {
    this.options.unshift({ option: '-v, --versions [versions]', autocomplete: this.allowedVersions }, { option: "-c, --changeType [changeType]", autocomplete: this.allowedChangeTypes }, { option: "-s, --services [services]", autocomplete: this.allowedServices }, { option: "--startDate [startDate]" }, { option: "--endDate [endDate]" });
}, _GraphChangelogListCommand_initValidators = function _GraphChangelogListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.versions &&
            args.options.versions.toLocaleLowerCase().split(',').some(x => !this.allowedVersions.map(y => y.toLocaleLowerCase()).includes(x))) {
            return `The verions contains an invalid value. Specify either ${this.allowedVersions.join(', ')} as properties`;
        }
        if (args.options.changeType &&
            !this.allowedChangeTypes.map(x => x.toLocaleLowerCase()).includes(args.options.changeType.toLocaleLowerCase())) {
            return `The change type contain an invalid value. Specify either ${this.allowedChangeTypes.join(', ')} as properties`;
        }
        if (args.options.services &&
            args.options.services.toLocaleLowerCase().split(',').some(x => !this.allowedServices.map(y => y.toLocaleLowerCase()).includes(x))) {
            return `The services contains invalid value. Specify either ${this.allowedServices.join(', ')} as properties`;
        }
        if (args.options.startDate && !validation.isValidISODate(args.options.startDate)) {
            return 'The startDate is not a valid ISO date string';
        }
        if (args.options.endDate && !validation.isValidISODate(args.options.endDate)) {
            return 'The endDate is not a valid ISO date string';
        }
        if (args.options.endDate && args.options.startDate && new Date(args.options.endDate) < new Date(args.options.startDate)) {
            return 'The endDate should be later than startDate';
        }
        return true;
    });
};
export default new GraphChangelogListCommand();
//# sourceMappingURL=changelog-list.js.map