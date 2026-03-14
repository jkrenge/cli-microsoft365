/**
 * Specifies a set of built-in permissions.
 * See: https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.basepermissions.aspx
 */
export class BasePermissions {
    constructor() {
        this._high = 0;
        this._low = 0;
    }
    get high() {
        return this._high;
    }
    set high(value) {
        this._high = value;
    }
    get low() {
        return this._low;
    }
    set low(value) {
        this._low = value;
    }
    parse() {
        const result = [];
        for (const permissionKind in PermissionKind) {
            if (this.has(PermissionKind[permissionKind])) {
                result.push(permissionKind);
            }
        }
        return result;
    }
    has(perm) {
        let hasPermission = false;
        if (perm === PermissionKind.FullMask) {
            hasPermission = (this.high & 32767) === 32767 && this.low === 65535;
        }
        let a = perm;
        a = a - 1;
        let b = 1;
        if (a >= 0 && a < 32) {
            b = b << a;
            hasPermission = 0 !== (this.low & b);
        }
        else if (a >= 32 && a < 64) {
            b = b << a - 32;
            hasPermission = (0 !== (this.high & b));
        }
        return hasPermission;
    }
    set(perm) {
        if (perm === PermissionKind.FullMask) {
            this._low = 65535;
            this._high = 32767;
        }
        else if (perm === PermissionKind.EmptyMask) {
            this._low = 0;
            this._high = 0;
        }
        else {
            const num1 = (perm - 1);
            const num2 = 1;
            if (num1 >= 0 && num1 < 32) {
                this._low = this._low | num2 << num1;
            }
            else {
                if (num1 < 32 || num1 >= 64) {
                    return;
                }
                this._high = this._high | num2 << num1 - 32;
            }
        }
    }
}
/**
 * Specifies permissions that are used to define user roles.
 * See: https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.permissionkind.aspx
 */
export var PermissionKind;
(function (PermissionKind) {
    PermissionKind[PermissionKind["EmptyMask"] = 0] = "EmptyMask";
    PermissionKind[PermissionKind["ViewListItems"] = 1] = "ViewListItems";
    PermissionKind[PermissionKind["AddListItems"] = 2] = "AddListItems";
    PermissionKind[PermissionKind["EditListItems"] = 3] = "EditListItems";
    PermissionKind[PermissionKind["DeleteListItems"] = 4] = "DeleteListItems";
    PermissionKind[PermissionKind["ApproveItems"] = 5] = "ApproveItems";
    PermissionKind[PermissionKind["OpenItems"] = 6] = "OpenItems";
    PermissionKind[PermissionKind["ViewVersions"] = 7] = "ViewVersions";
    PermissionKind[PermissionKind["DeleteVersions"] = 8] = "DeleteVersions";
    PermissionKind[PermissionKind["CancelCheckout"] = 9] = "CancelCheckout";
    PermissionKind[PermissionKind["ManagePersonalViews"] = 10] = "ManagePersonalViews";
    PermissionKind[PermissionKind["ManageLists"] = 12] = "ManageLists";
    PermissionKind[PermissionKind["ViewFormPages"] = 13] = "ViewFormPages";
    PermissionKind[PermissionKind["AnonymousSearchAccessList"] = 14] = "AnonymousSearchAccessList";
    PermissionKind[PermissionKind["Open"] = 17] = "Open";
    PermissionKind[PermissionKind["ViewPages"] = 18] = "ViewPages";
    PermissionKind[PermissionKind["AddAndCustomizePages"] = 19] = "AddAndCustomizePages";
    PermissionKind[PermissionKind["ApplyThemeAndBorder"] = 20] = "ApplyThemeAndBorder";
    PermissionKind[PermissionKind["ApplyStyleSheets"] = 21] = "ApplyStyleSheets";
    PermissionKind[PermissionKind["ViewUsageData"] = 22] = "ViewUsageData";
    PermissionKind[PermissionKind["CreateSSCSite"] = 23] = "CreateSSCSite";
    PermissionKind[PermissionKind["ManageSubwebs"] = 24] = "ManageSubwebs";
    PermissionKind[PermissionKind["CreateGroups"] = 25] = "CreateGroups";
    PermissionKind[PermissionKind["ManagePermissions"] = 26] = "ManagePermissions";
    PermissionKind[PermissionKind["BrowseDirectories"] = 27] = "BrowseDirectories";
    PermissionKind[PermissionKind["BrowseUserInfo"] = 28] = "BrowseUserInfo";
    PermissionKind[PermissionKind["AddDelPrivateWebParts"] = 29] = "AddDelPrivateWebParts";
    PermissionKind[PermissionKind["UpdatePersonalWebParts"] = 30] = "UpdatePersonalWebParts";
    PermissionKind[PermissionKind["ManageWeb"] = 31] = "ManageWeb";
    PermissionKind[PermissionKind["AnonymousSearchAccessWebLists"] = 32] = "AnonymousSearchAccessWebLists";
    PermissionKind[PermissionKind["UseClientIntegration"] = 37] = "UseClientIntegration";
    PermissionKind[PermissionKind["UseRemoteAPIs"] = 38] = "UseRemoteAPIs";
    PermissionKind[PermissionKind["ManageAlerts"] = 39] = "ManageAlerts";
    PermissionKind[PermissionKind["CreateAlerts"] = 40] = "CreateAlerts";
    PermissionKind[PermissionKind["EditMyUserInfo"] = 41] = "EditMyUserInfo";
    PermissionKind[PermissionKind["EnumeratePermissions"] = 63] = "EnumeratePermissions";
    PermissionKind[PermissionKind["FullMask"] = 65] = "FullMask";
})(PermissionKind || (PermissionKind = {}));
//# sourceMappingURL=base-permissions.js.map