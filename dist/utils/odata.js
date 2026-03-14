import request from "../request.js";
/* eslint-enable no-redeclare */
// eslint-disable-next-line no-redeclare
async function getAllItems(param1, metadata) {
    const requestOptions = typeof param1 !== 'string' ? param1 : {
        url: param1,
        headers: {
            accept: `application/json;odata.metadata=${metadata ?? 'none'}`,
            'odata-version': '4.0'
        },
        responseType: 'json'
    };
    const res = await request.get(requestOptions);
    let items = res.value;
    const nextLink = res['@odata.nextLink'] ?? res.nextLink;
    if (nextLink) {
        const nextPageItems = await odata.getAllItems({ ...requestOptions, url: nextLink });
        items = items.concat(nextPageItems);
    }
    return items;
}
export const odata = {
    getAllItems
};
//# sourceMappingURL=odata.js.map