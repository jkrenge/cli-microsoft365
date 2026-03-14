const allowedExactHosts = new Set([
    '127.0.0.1',
    '::1',
    '169.254.169.254',
    'analysis.windows.net',
    'api.bap.microsoft.com',
    'api.flow.microsoft.com',
    'api.powerapps.com',
    'api.powerbi.com',
    'api.powerplatform.com',
    'api.yammer.com',
    'dod-graph.microsoft.us',
    'graph.microsoft.com',
    'graph.microsoft.us',
    'localhost',
    'login.chinacloudapi.cn',
    'login.microsoftonline.com',
    'login.microsoftonline.us',
    'login.windows.net',
    'manage.office.com',
    'management.azure.com',
    'management.chinacloudapi.cn',
    'management.usgovcloudapi.net',
    'microsoft.com',
    'microsoftgraph.chinacloudapi.cn',
    'service.powerapps.com',
    'tasks.office.com',
    'www.microsoft.com',
    'www.yammer.com'
]);
const allowedHostSuffixes = [
    '.api.bap.microsoft.com',
    '.appsplatform.us',
    '.dynamics.com',
    '.microsoftdynamics.us',
    '.powerapps.com',
    '.powerplatform.com',
    '.sharepoint.cn',
    '.sharepoint.com',
    '.sharepoint-df.com',
    '.sharepoint-mil.us',
    '.sharepoint.us'
];
function parseBoolean(value) {
    if (!value) {
        return undefined;
    }
    const normalizedValue = value.trim().toLowerCase();
    if (['1', 'true', 'yes', 'on'].includes(normalizedValue)) {
        return true;
    }
    if (['0', 'false', 'no', 'off'].includes(normalizedValue)) {
        return false;
    }
    return undefined;
}
function getAdditionalAllowedHosts() {
    return (process.env.CLIMICROSOFT365_ALLOWED_HOSTS ?? '')
        .split(',')
        .map(host => host.trim().toLowerCase())
        .filter(host => host.length > 0);
}
export const networkAccess = {
    isRestricted() {
        if (parseBoolean(process.env.CLIMICROSOFT365_ALLOW_EXTERNAL) === true) {
            return false;
        }
        const configuredValue = parseBoolean(process.env.CLIMICROSOFT365_RESTRICT_EXTERNAL);
        if (typeof configuredValue !== 'undefined') {
            return configuredValue;
        }
        return true;
    },
    isAllowedHost(host) {
        const normalizedHost = host.toLowerCase();
        if (allowedExactHosts.has(normalizedHost)) {
            return true;
        }
        if (getAdditionalAllowedHosts().includes(normalizedHost)) {
            return true;
        }
        return allowedHostSuffixes.some(suffix => normalizedHost.endsWith(suffix));
    },
    assertAllowed(url) {
        if (!this.isRestricted()) {
            return;
        }
        const host = new URL(url).hostname;
        if (this.isAllowedHost(host)) {
            return;
        }
        throw new Error(`External network access to '${host}' is blocked by default. ` +
            `Set CLIMICROSOFT365_ALLOW_EXTERNAL=1 to disable the restriction ` +
            `or add the host to CLIMICROSOFT365_ALLOWED_HOSTS.`);
    }
};
//# sourceMappingURL=networkAccess.js.map