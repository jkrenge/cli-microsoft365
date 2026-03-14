import request from '../request.js';
export class MsalNetworkClient {
    sendGetRequestAsync(url, options) {
        return this.sendRequestAsync(url, 'GET', options);
    }
    sendPostRequestAsync(url, options) {
        return this.sendRequestAsync(url, 'POST', options);
    }
    async sendRequestAsync(url, method, options = {}) {
        const requestOptions = {
            url: url,
            method: method,
            headers: {
                'x-anonymous': true,
                ...options.headers
            },
            data: options.body,
            fullResponse: true
        };
        try {
            const res = await request.execute(requestOptions);
            const headersObj = {};
            for (const [key, value] of Object.entries(res.headers)) {
                headersObj[key] = typeof value === 'string' ? value : String(value);
            }
            return {
                headers: headersObj,
                body: JSON.parse(res.data),
                status: res.status
            };
        }
        catch (ex) {
            const error = ex;
            const headersObj = {};
            if (error.response?.headers) {
                for (const [key, value] of Object.entries(error.response.headers)) {
                    headersObj[key] = typeof value === 'string' ? value : String(value);
                }
            }
            return {
                headers: headersObj,
                body: JSON.parse(error.response?.data ?? '{}'),
                status: error.response?.status ?? 400
            };
        }
    }
}
//# sourceMappingURL=MsalNetworkClient.js.map