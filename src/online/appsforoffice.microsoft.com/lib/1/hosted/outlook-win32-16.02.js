/* Outlook Win32 specific API library */
/* osfweb version: 16.0.19009.20000 */
/* office-js-api version: 20250612.2 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/
/*
    Your use of this file is governed by the license terms for the Microsoft Office JavaScript (Office.js) API library: https://github.com/OfficeDev/office-js/blob/release/LICENSE.md

    This file also contains the following Promise implementation (with a few small modifications):
        * @overview es6-promise - a tiny implementation of Promises/A+.
        * @copyright Copyright (c) 2014 Yehuda Katz, Tom Dale, Stefan Penner and contributors (Conversion to ES6 API by Jake Archibald)
        * @license   Licensed under MIT license
        *            See https://raw.githubusercontent.com/jakearchibald/es6-promise/master/LICENSE
        * @version   2.3.0
*/
"undefined" !== typeof OSFPerformance && (OSFPerformance.hostInitializationStart = OSFPerformance.now());
/* Outlook rich client specific API library */
/* Version: 16.0.19009.20000 */
var __extends = this && this.__extends || function() {
    var a = function(c, b) {
        a = Object.setPrototypeOf || {
            __proto__: []
        }instanceof Array && function(b, a) {
            b.__proto__ = a
        }
        || function(c, a) {
            for (var b in a)
                if (a.hasOwnProperty(b))
                    c[b] = a[b]
        }
        ;
        return a(c, b)
    };
    return function(c, b) {
        a(c, b);
        function d() {
            this.constructor = c
        }
        c.prototype = b === null ? Object.create(b) : (d.prototype = b.prototype,
        new d)
    }
}(), __assign = this && this.__assign || function() {
    __assign = Object.assign || function(d) {
        for (var a, b = 1, e = arguments.length; b < e; b++) {
            a = arguments[b];
            for (var c in a)
                if (Object.prototype.hasOwnProperty.call(a, c))
                    d[c] = a[c]
        }
        return d
    }
    ;
    return __assign.apply(this, arguments)
}
, OfficeExt;
(function(b) {
    var a = function() {
        var a = true;
        function b() {}
        b.prototype.isMsAjaxLoaded = function() {
            var b = "function"
              , c = "undefined";
            if (typeof Sys !== c && typeof Type !== c && Sys.StringBuilder && typeof Sys.StringBuilder === b && Type.registerNamespace && typeof Type.registerNamespace === b && Type.registerClass && typeof Type.registerClass === b && typeof Function._validateParams === b && Sys.Serialization && Sys.Serialization.JavaScriptSerializer && typeof Sys.Serialization.JavaScriptSerializer.serialize === b)
                return a;
            else
                return false
        }
        ;
        b.prototype.loadMsAjaxFull = function(b) {
            var a = (window.location.protocol.toLowerCase() === "https:" ? "https:" : "http:") + "//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js";
            OSF.OUtil.loadScript(a, b)
        }
        ;
        Object.defineProperty(b.prototype, "msAjaxError", {
            "get": function() {
                var a = this;
                if (a._msAjaxError == null && a.isMsAjaxLoaded())
                    a._msAjaxError = Error;
                return a._msAjaxError
            },
            "set": function(a) {
                this._msAjaxError = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(b.prototype, "msAjaxString", {
            "get": function() {
                var a = this;
                if (a._msAjaxString == null && a.isMsAjaxLoaded())
                    a._msAjaxString = String;
                return a._msAjaxString
            },
            "set": function(a) {
                this._msAjaxString = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(b.prototype, "msAjaxDebug", {
            "get": function() {
                var a = this;
                if (a._msAjaxDebug == null && a.isMsAjaxLoaded())
                    a._msAjaxDebug = Sys.Debug;
                return a._msAjaxDebug
            },
            "set": function(a) {
                this._msAjaxDebug = a
            },
            enumerable: a,
            configurable: a
        });
        return b
    }();
    b.MicrosoftAjaxFactory = a
}
)(OfficeExt || (OfficeExt = {}));
var OsfMsAjaxFactory = new OfficeExt.MicrosoftAjaxFactory
  , OSF = OSF || {};
(function(b) {
    var a = function() {
        function a(a) {
            this._internalStorage = a
        }
        a.prototype.getItem = function(a) {
            try {
                return this._internalStorage && this._internalStorage.getItem(a)
            } catch (b) {
                return null
            }
        }
        ;
        a.prototype.setItem = function(b, a) {
            try {
                this._internalStorage && this._internalStorage.setItem(b, a)
            } catch (c) {}
        }
        ;
        a.prototype.clear = function() {
            try {
                this._internalStorage && this._internalStorage.clear()
            } catch (a) {}
        }
        ;
        a.prototype.removeItem = function(a) {
            try {
                this._internalStorage && this._internalStorage.removeItem(a)
            } catch (b) {}
        }
        ;
        a.prototype.getKeysWithPrefix = function(d) {
            var b = [];
            try {
                for (var e = this._internalStorage && this._internalStorage.length || 0, a = 0; a < e; a++) {
                    var c = this._internalStorage.key(a);
                    c.indexOf(d) === 0 && b.push(c)
                }
            } catch (f) {}
            return b
        }
        ;
        a.prototype.isLocalStorageAvailable = function() {
            return this._internalStorage != null
        }
        ;
        return a
    }();
    b.SafeStorage = a
}
)(OfficeExt || (OfficeExt = {}));
OSF.XdmFieldName = {
    ConversationUrl: "ConversationUrl",
    AppId: "AppId"
};
OSF.TestFlightStart = 1e3;
OSF.TestFlightEnd = 1009;
OSF.FlightNames = {
    UseOriginNotUrl: 0,
    AddinEnforceHttps: 2,
    FirstPartyAnonymousProxyReadyCheckTimeout: 6,
    AddinRibbonIdAllowUnknown: 9,
    ManifestParserDevConsoleLog: 15,
    AddinActionDefinitionHybridMode: 18,
    UseActionIdForUILessCommand: 20,
    RequirementSetRibbonApiOnePointTwo: 21,
    SetFocusToTaskpaneIsEnabled: 22,
    ShortcutInfoArrayInUserPreferenceData: 23,
    OSFTestFlight1000: OSF.TestFlightStart,
    OSFTestFlight1001: OSF.TestFlightStart + 1,
    OSFTestFlight1002: OSF.TestFlightStart + 2,
    OSFTestFlight1003: OSF.TestFlightStart + 3,
    OSFTestFlight1004: OSF.TestFlightStart + 4,
    OSFTestFlight1005: OSF.TestFlightStart + 5,
    OSFTestFlight1006: OSF.TestFlightStart + 6,
    OSFTestFlight1007: OSF.TestFlightStart + 7,
    OSFTestFlight1008: OSF.TestFlightStart + 8,
    OSFTestFlight1009: OSF.TestFlightEnd
};
OSF.TrustUXFlightValues = {
    TrustUXControlA: 0,
    TrustUXExperimentB: 1,
    TrustUXExperimentC: 2
};
OSF.FlightTreatmentNames = {
    AddinTrustUXImprovement: "Microsoft.Office.SharedOnline.AddinTrustUXImprovement",
    BlockAutoOpenAddInIfStoreDisabled: "Microsoft.Office.SharedOnline.BlockAutoOpenAddInIfStoreDisabled",
    Bug7083046KillSwitch: "Microsoft.Office.SharedOnline.Bug7083046KillSwitch",
    CheckProxyIsReadyRetry: "Microsoft.Office.SharedOnline.OEP.CheckProxyIsReadyRetry",
    InsertionDialogFixesEnabled: "Microsoft.Office.SharedOnline.InsertionDialogFixesEnabled",
    WopiPreinstalledAddInsEnabled: "Microsoft.Office.SharedOnline.WopiPreinstalledAddInsEnabled",
    RemoveGetTrustNoPrompt: "Microsoft.Office.SharedOnline.removeGetTrustNoPrompt",
    HostTrustDialog: "Microsoft.Office.SharedOnline.HostTrustDialog",
    GetAddinFlyoutEnabled: "Microsoft.Office.SharedOnline.GetAddinFlyoutEnabled",
    BackstageEnabled: "Microsoft.Office.SharedOnline.NewBackstageEnabled",
    EnablingWindowOpenUsageLogging: "Microsoft.Office.SharedOnline.EnablingWindowOpenUsageLogging",
    MosForWXPEnabled: "Microsoft.Office.SharedOnline.MosForWXPEnabled",
    EnableMsal3SsoApi: "Microsoft.Office.SharedOnline.EnableMsal3SsoApi"
};
OSF.Flights = [];
OSF.DisabledChangeGates = [];
OSF.IntFlights = {};
OSF.Settings = {};
OSF.WindowNameItemKeys = {
    BaseFrameName: "baseFrameName",
    HostInfo: "hostInfo",
    XdmInfo: "xdmInfo",
    AppContext: "appContext",
    Flights: "flights",
    DisabledChangeGates: "disabledChangeGates"
};
OSF.OUtil = function() {
    var l = "focus"
      , k = "https:"
      , j = "on"
      , q = "configurable"
      , p = "writable"
      , i = "enumerable"
      , e = ""
      , f = "undefined"
      , c = false
      , b = true
      , h = "string"
      , m = 2147483647
      , a = null
      , g = "#"
      , d = -1
      , w = d
      , B = "&_xdm_Info="
      , A = "&_flights="
      , F = "&_disabledChangeGates="
      , z = "_xdm_"
      , G = "_flights="
      , D = "_disabledChangeGates="
      , s = g
      , y = "&"
      , n = "class"
      , v = {}
      , E = 3e4
      , r = a
      , u = a
      , o = (new Date).getTime();
    function C() {
        var a = m * Math.random();
        a ^= o ^ (new Date).getMilliseconds() << Math.floor(Math.random() * (31 - 10));
        return a.toString(16)
    }
    function t() {
        if (!r) {
            try {
                var b = window.sessionStorage
            } catch (c) {
                b = a
            }
            r = new OfficeExt.SafeStorage(b)
        }
        return r
    }
    function x(e) {
        for (var c = [], b = [], f = e.length, a, d = 0; d < f; d++) {
            a = e[d];
            if (a.tabIndex)
                if (a.tabIndex > 0)
                    b.push(a);
                else
                    a.tabIndex === 0 && c.push(a);
            else
                c.push(a)
        }
        b = b.sort(function(d, c) {
            var a = d.tabIndex - c.tabIndex;
            if (a === 0)
                a = b.indexOf(d) - b.indexOf(c);
            return a
        });
        return [].concat(b, c)
    }
    return {
        set_entropy: function(a) {
            if (typeof a == h)
                for (var b = 0; b < a.length; b += 4) {
                    for (var d = 0, c = 0; c < 4 && b + c < a.length; c++)
                        d = (d << 8) + a.charCodeAt(b + c);
                    o ^= d
                }
            else if (typeof a == "number")
                o ^= a;
            else
                o ^= m * Math.random();
            o &= m
        },
        extend: function(b, a) {
            var c = function() {};
            c.prototype = a.prototype;
            b.prototype = new c;
            b.prototype.constructor = b;
            b.uber = a.prototype;
            if (a.prototype.constructor === Object.prototype.constructor)
                a.prototype.constructor = a
        },
        setNamespace: function(b, a) {
            if (a && b && !a[b])
                a[b] = {}
        },
        unsetNamespace: function(b, a) {
            if (a && b && a[b])
                delete a[b]
        },
        serializeSettings: function(b) {
            var d = {};
            for (var c in b) {
                var a = b[c];
                try {
                    if (JSON)
                        a = JSON.stringify(a, function(a, b) {
                            return OSF.OUtil.isDate(this[a]) ? OSF.DDA.SettingsManager.DateJSONPrefix + this[a].getTime() + OSF.DDA.SettingsManager.DataJSONSuffix : b
                        });
                    else
                        a = Sys.Serialization.JavaScriptSerializer.serialize(a);
                    d[c] = a
                } catch (e) {}
            }
            return d
        },
        deserializeSettings: function(c) {
            var f = {};
            c = c || {};
            for (var e in c) {
                var a = c[e];
                try {
                    if (JSON)
                        a = JSON.parse(a, function(c, a) {
                            var b;
                            if (typeof a === h && a && a.length > 6 && a.slice(0, 5) === OSF.DDA.SettingsManager.DateJSONPrefix && a.slice(d) === OSF.DDA.SettingsManager.DataJSONSuffix) {
                                b = new Date(parseInt(a.slice(5, d)));
                                if (b)
                                    return b
                            }
                            return a
                        });
                    else
                        a = Sys.Serialization.JavaScriptSerializer.deserialize(a, b);
                    f[e] = a
                } catch (g) {}
            }
            return f
        },
        loadScript: function(f, g, i, j) {
            if (f && g) {
                var l = window.document
                  , d = v[f];
                if (!d) {
                    var e = l.createElement("script");
                    e.type = "text/javascript";
                    d = {
                        loaded: c,
                        pendingCallbacks: [g],
                        timer: a
                    };
                    v[f] = d;
                    var k = function() {
                        if (d.timer != a) {
                            clearTimeout(d.timer);
                            delete d.timer
                        }
                        d.loaded = b;
                        for (var e = d.pendingCallbacks.length, c = 0; c < e; c++) {
                            var f = d.pendingCallbacks.shift();
                            f()
                        }
                    }
                      , m = function() {
                        if (window.navigator.userAgent.indexOf("Trident") > 0)
                            h(a);
                        else
                            h(new Event("Script load timed out"))
                    }
                      , h = function(g) {
                        delete v[f];
                        if (d.timer != a) {
                            clearTimeout(d.timer);
                            delete d.timer
                        }
                        for (var c = d.pendingCallbacks.length, b = 0; b < c; b++) {
                            var e = d.pendingCallbacks.shift();
                            e(g)
                        }
                    };
                    if (e.readyState)
                        e.onreadystatechange = function() {
                            if (e.readyState == "loaded" || e.readyState == "complete") {
                                e.onreadystatechange = a;
                                k()
                            }
                        }
                        ;
                    else
                        e.onload = k;
                    e.onerror = h;
                    i = i || E;
                    d.timer = setTimeout(m, i);
                    e.setAttribute("crossOrigin", "anonymous");
                    e.src = j ? j.createScriptURL(f) : f;
                    l.getElementsByTagName("head")[0].appendChild(e)
                } else if (d.loaded)
                    g();
                else
                    d.pendingCallbacks.push(g)
            }
        },
        loadCSS: function(c) {
            if (c) {
                var b = window.document
                  , a = b.createElement("link");
                a.type = "text/css";
                a.rel = "stylesheet";
                a.href = c;
                b.getElementsByTagName("head")[0].appendChild(a)
            }
        },
        parseEnum: function(b, c) {
            var a = c[b.trim()];
            if (typeof a == f) {
                OsfMsAjaxFactory.msAjaxDebug.trace("invalid enumeration string:" + b);
                throw OsfMsAjaxFactory.msAjaxError.argument("str")
            }
            return a
        },
        delayExecutionAndCache: function() {
            var a = {
                calc: arguments[0]
            };
            return function() {
                if (a.calc) {
                    a.val = a.calc.apply(this, arguments);
                    delete a.calc
                }
                return a.val
            }
        },
        getUniqueId: function() {
            w = w + 1;
            return w.toString()
        },
        formatString: function() {
            var a = arguments
              , b = a[0];
            return b.replace(/{(\d+)}/gm, function(d, b) {
                var c = parseInt(b, 10) + 1;
                return a[c] === undefined ? "{" + b + "}" : a[c]
            })
        },
        generateConversationId: function() {
            return [C(), C(), (new Date).getTime().toString()].join("_")
        },
        getFrameName: function(a) {
            return z + a + this.generateConversationId()
        },
        addXdmInfoAsHash: function(b, a) {
            return OSF.OUtil.addInfoAsHash(b, B, a, c)
        },
        addFlightsAsHash: function(c, a) {
            return OSF.OUtil.addInfoAsHash(c, A, a, b)
        },
        addInfoAsHash: function(b, g, c, i) {
            b = b.trim() || e;
            var f = b.split(s), h = f.shift(), d = f.join(s), a;
            if (i)
                a = [g, encodeURIComponent(c), d].join(e);
            else
                a = [d, g, c].join(e);
            return [h, s, a].join(e)
        },
        parseHostInfoFromWindowName: function(a, b) {
            return OSF.OUtil.parseInfoFromWindowName(a, b, OSF.WindowNameItemKeys.HostInfo)
        },
        parseXdmInfo: function(b) {
            var a = OSF.OUtil.parseXdmInfoWithGivenFragment(b, window.location.hash);
            if (!a)
                a = OSF.OUtil.parseXdmInfoFromWindowName(b, window.name);
            return a
        },
        parseXdmInfoFromWindowName: function(a, b) {
            return OSF.OUtil.parseInfoFromWindowName(a, b, OSF.WindowNameItemKeys.XdmInfo)
        },
        parseXdmInfoWithGivenFragment: function(a, b) {
            return OSF.OUtil.parseInfoWithGivenFragment(B, z, c, a, b)
        },
        parseFlights: function(b) {
            var a = OSF.OUtil.parseFlightsWithGivenFragment(b, window.location.hash);
            if (a.length == 0)
                a = OSF.OUtil.parseFlightsFromWindowName(b, window.name);
            return a
        },
        parseDisabledChangeGates: function(b) {
            var a = OSF.OUtil.parseDisabledChangeGatesWithGivenFragment(b, window.location.hash);
            if (a.length == 0)
                a = OSF.OUtil.parseDisabledChangeGatesFromWindowName(b, window.name);
            return a
        },
        checkFlight: function(a) {
            return OSF.Flights && OSF.Flights.indexOf(a) >= 0
        },
        isChangeGateEnabled: function(a) {
            return !OSF.DisabledChangeGates || OSF.DisabledChangeGates.indexOf(a) === d
        },
        pushFlight: function(a) {
            if (OSF.Flights.indexOf(a) < 0) {
                OSF.Flights.push(a);
                return b
            }
            return c
        },
        getBooleanSetting: function(a) {
            return OSF.OUtil.getBooleanFromDictionary(OSF.Settings, a)
        },
        getBooleanFromDictionary: function(b, a) {
            var d = b && a && b[a] !== undefined && b[a] && (typeof b[a] === h && b[a].toUpperCase() === "TRUE" || typeof b[a] === "boolean" && b[a]);
            return d !== undefined ? d : c
        },
        getIntFromDictionary: function(b, a) {
            if (b && a && b[a] !== undefined && typeof b[a] === h)
                return parseInt(b[a]);
            else
                return NaN
        },
        pushIntFlight: function(a, d) {
            if (!(a in OSF.IntFlights)) {
                OSF.IntFlights[a] = d;
                return b
            }
            return c
        },
        getIntFlight: function(a) {
            if (OSF.IntFlights && a in OSF.IntFlights)
                return OSF.IntFlights[a];
            else
                return NaN
        },
        parseFlightsFromWindowName: function(a, b) {
            return OSF.OUtil.parseArrayWithDefault(OSF.OUtil.parseInfoFromWindowName(a, b, OSF.WindowNameItemKeys.Flights))
        },
        parseDisabledChangeGatesFromWindowName: function(a, b) {
            return OSF.OUtil.parseArrayWithDefault(OSF.OUtil.parseInfoFromWindowName(a, b, OSF.WindowNameItemKeys.DisabledChangeGates))
        },
        parseFlightsWithGivenFragment: function(a, c) {
            return OSF.OUtil.parseArrayWithDefault(OSF.OUtil.parseInfoWithGivenFragment(A, G, b, a, c))
        },
        parseDisabledChangeGatesWithGivenFragment: function(a, c) {
            return OSF.OUtil.parseArrayWithDefault(OSF.OUtil.parseInfoWithGivenFragment(F, D, b, a, c))
        },
        parseArrayWithDefault: function(b) {
            var a = [];
            try {
                a = JSON.parse(b)
            } catch (c) {}
            if (!Array.isArray(a))
                a = [];
            return a
        },
        parseInfoFromWindowName: function(g, h, f) {
            try {
                var b = JSON.parse(h)
                  , c = b != a ? b[f] : a
                  , d = t();
                if (!g && d && b != a) {
                    var e = b[OSF.WindowNameItemKeys.BaseFrameName] + f;
                    if (c)
                        d.setItem(e, c);
                    else
                        c = d.getItem(e)
                }
                return c
            } catch (i) {
                return a
            }
        },
        parseInfoWithGivenFragment: function(m, j, k, i, l) {
            var f = l.split(m)
              , b = f.length > 1 ? f[f.length - 1] : a;
            if (k && b != a) {
                if (b.indexOf(y) >= 0)
                    b = b.split(y)[0];
                b = decodeURIComponent(b)
            }
            var c = t();
            if (!i && c) {
                var e = window.name.indexOf(j);
                if (e > d) {
                    var g = window.name.indexOf(";", e);
                    if (g == d)
                        g = window.name.length;
                    var h = window.name.substring(e, g);
                    if (b)
                        c.setItem(h, b);
                    else
                        b = c.getItem(h)
                }
            }
            return b
        },
        getConversationId: function() {
            var c = window.location.search
              , b = a;
            if (c) {
                var d = c.indexOf("&");
                b = d > 0 ? c.substring(1, d) : c.substr(1);
                if (b && b.charAt(b.length - 1) === "=") {
                    b = b.substring(0, b.length - 1);
                    if (b)
                        b = decodeURIComponent(b)
                }
            }
            return b
        },
        getInfoItems: function(b) {
            var a = b.split("$");
            if (typeof a[1] == f)
                a = b.split("|");
            if (typeof a[1] == f)
                a = b.split("%7C");
            return a
        },
        getXdmFieldValue: function(f, d) {
            var b = e
              , c = OSF.OUtil.parseXdmInfo(d);
            if (c) {
                var a = OSF.OUtil.getInfoItems(c);
                if (a != undefined && a.length >= 3)
                    switch (f) {
                    case OSF.XdmFieldName.ConversationUrl:
                        b = a[2];
                        break;
                    case OSF.XdmFieldName.AppId:
                        b = a[1]
                    }
            }
            return b
        },
        validateParamObject: function(f, e) {
            var a = Function._validateParams(arguments, [{
                name: "params",
                type: Object,
                mayBeNull: c
            }, {
                name: "expectedProperties",
                type: Object,
                mayBeNull: c
            }, {
                name: "callback",
                type: Function,
                mayBeNull: b
            }]);
            if (a)
                throw a;
            for (var d in e) {
                a = Function._validateParameter(f[d], e[d], d);
                if (a)
                    throw a
            }
        },
        writeProfilerMark: function(a) {
            if (window.msWriteProfilerMark) {
                window.msWriteProfilerMark(a);
                OsfMsAjaxFactory.msAjaxDebug.trace(a)
            }
        },
        outputDebug: function(a) {
            typeof OsfMsAjaxFactory !== f && OsfMsAjaxFactory.msAjaxDebug && OsfMsAjaxFactory.msAjaxDebug.trace && OsfMsAjaxFactory.msAjaxDebug.trace(a)
        },
        defineNondefaultProperty: function(e, f, a, c) {
            a = a || {};
            for (var g in c) {
                var d = c[g];
                if (a[d] == undefined)
                    a[d] = b
            }
            Object.defineProperty(e, f, a);
            return e
        },
        defineNondefaultProperties: function(c, a, d) {
            a = a || {};
            for (var b in a)
                OSF.OUtil.defineNondefaultProperty(c, b, a[b], d);
            return c
        },
        defineEnumerableProperty: function(c, b, a) {
            return OSF.OUtil.defineNondefaultProperty(c, b, a, [i])
        },
        defineEnumerableProperties: function(b, a) {
            return OSF.OUtil.defineNondefaultProperties(b, a, [i])
        },
        defineMutableProperty: function(c, b, a) {
            return OSF.OUtil.defineNondefaultProperty(c, b, a, [p, i, q])
        },
        defineMutableProperties: function(b, a) {
            return OSF.OUtil.defineNondefaultProperties(b, a, [p, i, q])
        },
        finalizeProperties: function(e, d) {
            d = d || {};
            for (var g = Object.getOwnPropertyNames(e), i = g.length, f = 0; f < i; f++) {
                var h = g[f]
                  , a = Object.getOwnPropertyDescriptor(e, h);
                if (!a.get && !a.set)
                    a.writable = d.writable || c;
                a.configurable = d.configurable || c;
                a.enumerable = d.enumerable || b;
                Object.defineProperty(e, h, a)
            }
            return e
        },
        mapList: function(a, c) {
            var b = [];
            if (a)
                for (var d in a)
                    b.push(c(a[d]));
            return b
        },
        listContainsKey: function(d, e) {
            for (var a in d)
                if (e == a)
                    return b;
            return c
        },
        listContainsValue: function(a, d) {
            for (var e in a)
                if (d == a[e])
                    return b;
            return c
        },
        augmentList: function(a, b) {
            var d = a.push ? function(c, b) {
                a.push(b)
            }
            : function(c, b) {
                a[c] = b
            }
            ;
            for (var c in b)
                d(c, b[c])
        },
        redefineList: function(a, b) {
            for (var d in a)
                delete a[d];
            for (var c in b)
                a[c] = b[c]
        },
        isArray: function(a) {
            return Object.prototype.toString.apply(a) === "[object Array]"
        },
        isFunction: function(a) {
            return Object.prototype.toString.apply(a) === "[object Function]"
        },
        isDate: function(a) {
            return Object.prototype.toString.apply(a) === "[object Date]"
        },
        addEventListener: function(a, b, d) {
            if (a.addEventListener)
                a.addEventListener(b, d, c);
            else if (Sys.Browser.agent === Sys.Browser.InternetExplorer && a.attachEvent)
                a.attachEvent(j + b, d);
            else
                a[j + b] = d
        },
        removeEventListener: function(b, d, e) {
            if (b.removeEventListener)
                b.removeEventListener(d, e, c);
            else if (Sys.Browser.agent === Sys.Browser.InternetExplorer && b.detachEvent)
                b.detachEvent(j + d, e);
            else
                b[j + d] = a
        },
        xhrGet: function(f, e, c) {
            var a;
            try {
                a = new XMLHttpRequest;
                a.onreadystatechange = function() {
                    if (a.readyState == 4)
                        if (a.status == 200)
                            e(a.responseText);
                        else
                            c(a.status)
                }
                ;
                a.open("GET", f, b);
                a.send()
            } catch (d) {
                c(d)
            }
        },
        encodeBase64: function(c) {
            if (!c)
                return c;
            var o = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=", m = [], b = [], i = 0, k, h, j, d, f, g, a, n = c.length;
            do {
                k = c.charCodeAt(i++);
                h = c.charCodeAt(i++);
                j = c.charCodeAt(i++);
                a = 0;
                d = k & 255;
                f = k >> 8;
                g = h & 255;
                b[a++] = d >> 2;
                b[a++] = (d & 3) << 4 | f >> 4;
                b[a++] = (f & 15) << 2 | g >> 6;
                b[a++] = g & 63;
                if (!isNaN(h)) {
                    d = h >> 8;
                    f = j & 255;
                    g = j >> 8;
                    b[a++] = d >> 2;
                    b[a++] = (d & 3) << 4 | f >> 4;
                    b[a++] = (f & 15) << 2 | g >> 6;
                    b[a++] = g & 63
                }
                if (isNaN(h))
                    b[a - 1] = 64;
                else if (isNaN(j)) {
                    b[a - 2] = 64;
                    b[a - 1] = 64
                }
                for (var l = 0; l < a; l++)
                    m.push(o.charAt(b[l]))
            } while (i < n);
            return m.join(e)
        },
        getSessionStorage: function() {
            return t()
        },
        getLocalStorage: function() {
            if (!u) {
                try {
                    var b = window.localStorage
                } catch (c) {
                    b = a
                }
                u = new OfficeExt.SafeStorage(b)
            }
            return u
        },
        convertIntToCssHexColor: function(b) {
            var a = g + (Number(b) + 16777216).toString(16).slice(-6);
            return a
        },
        attachClickHandler: function(a, b) {
            a.onclick = function() {
                b()
            }
            ;
            a.ontouchend = function(a) {
                b();
                a.preventDefault()
            }
        },
        getQueryStringParamValue: function(a, d) {
            var f = Function._validateParams(arguments, [{
                name: "queryString",
                type: String,
                mayBeNull: c
            }, {
                name: "paramName",
                type: String,
                mayBeNull: c
            }]);
            if (f) {
                OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: Parameters cannot be null.");
                return e
            }
            var b = new RegExp("[\\?&]" + d + "=([^&#]*)","i");
            if (!b.test(a)) {
                OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: The parameter is not found.");
                return e
            }
            return b.exec(a)[1]
        },
        getHostnamePortionForLogging: function(d) {
            var f = Function._validateParams(arguments, [{
                name: "hostname",
                type: String,
                mayBeNull: c
            }]);
            if (f)
                return e;
            var a = d.split(".")
              , b = a.length;
            if (b >= 2)
                return a[b - 2] + "." + a[b - 1];
            else if (b == 1)
                return a[0]
        },
        isiOS: function() {
            return window.navigator.userAgent.match(/(iPad|iPhone|iPod)/g) ? b : c
        },
        isChrome: function() {
            return window.navigator.userAgent.indexOf("Chrome") > 0 && !OSF.OUtil.isEdge()
        },
        isEdge: function() {
            return window.navigator.userAgent.indexOf("Edge") > 0
        },
        isIE: function() {
            return window.navigator.userAgent.indexOf("Trident") > 0
        },
        isFirefox: function() {
            return window.navigator.userAgent.indexOf("Firefox") > 0
        },
        startsWith: function(b, a, c) {
            if (c)
                return b.substr(0, a.length) === a;
            else
                return b.startsWith(a)
        },
        containsPort: function(d, e, c, a) {
            return this.startsWith(d, e + "//" + c + ":" + a, b) || this.startsWith(d, c + ":" + a, b)
        },
        getRedundandPortString: function(b, a) {
            if (!b || !a)
                return e;
            if (a.protocol == k && this.containsPort(b, k, a.hostname, "443"))
                return ":443";
            else if (a.protocol == "http:" && this.containsPort(b, "http:", a.hostname, "80"))
                return ":80";
            return e
        },
        removeChar: function(a, b) {
            if (b < a.length - 1)
                return a.substring(0, b) + a.substring(b + 1);
            else if (b == a.length - 1)
                return a.substring(0, a.length - 1);
            else
                return a
        },
        cleanUrlOfChar: function(a, c) {
            for (var b = 0; b < a.length; b++)
                if (a.charAt(b) === c)
                    if (b + 1 >= a.length)
                        return this.removeChar(a, b);
                    else if (c === "/") {
                        if (a.charAt(b + 1) === "?" || a.charAt(b + 1) === g)
                            return this.removeChar(a, b)
                    } else if (c === "?")
                        if (a.charAt(b + 1) === g)
                            return this.removeChar(a, b);
            return a
        },
        cleanUrl: function(a) {
            a = this.cleanUrlOfChar(a, "/");
            a = this.cleanUrlOfChar(a, "?");
            a = this.cleanUrlOfChar(a, g);
            if (a.substr(0, 8) == "https://") {
                var b = a.indexOf(":443");
                if (b != d)
                    if (b == a.length - 4 || a.charAt(b + 4) == "/" || a.charAt(b + 4) == "?" || a.charAt(b + 4) == g)
                        a = a.substring(0, b) + a.substring(b + 4)
            } else if (a.substr(0, 7) == "http://") {
                var b = a.indexOf(":80");
                if (b != d)
                    if (b == a.length - 3 || a.charAt(b + 3) == "/" || a.charAt(b + 3) == "?" || a.charAt(b + 3) == g)
                        a = a.substring(0, b) + a.substring(b + 3)
            }
            return a
        },
        parseUrl: function(g, i) {
            var h = this;
            if (i === void 0)
                i = c;
            if (typeof g === f || !g)
                return undefined;
            var j = "NotHttps"
              , o = "InvalidUrl"
              , n = h.isIE()
              , b = {
                protocol: undefined,
                hostname: undefined,
                host: undefined,
                port: undefined,
                pathname: undefined,
                search: undefined,
                hash: undefined,
                isPortPartOfUrl: undefined
            };
            try {
                if (n) {
                    var a = document.createElement("a");
                    a.href = g;
                    if (!a || !a.protocol || !a.host || !a.hostname || !a.href || h.cleanUrl(a.href).toLowerCase() !== h.cleanUrl(g).toLowerCase())
                        throw o;
                    if (OSF.OUtil.checkFlight(OSF.FlightNames.AddinEnforceHttps))
                        if (i && a.protocol != k)
                            throw new Error(j);
                    var m = h.getRedundandPortString(g, a);
                    b.protocol = a.protocol;
                    b.hostname = a.hostname;
                    b.port = m == e ? a.port : e;
                    b.host = m != e ? a.hostname : a.host;
                    b.pathname = (n ? "/" : e) + a.pathname;
                    b.search = a.search;
                    b.hash = a.hash;
                    b.isPortPartOfUrl = h.containsPort(g, a.protocol, a.hostname, a.port)
                } else {
                    var d = new URL(g);
                    if (d && d.protocol && d.host && d.hostname) {
                        if (OSF.OUtil.checkFlight(OSF.FlightNames.AddinEnforceHttps))
                            if (i && d.protocol != k)
                                throw new Error(j);
                        b.protocol = d.protocol;
                        b.hostname = d.hostname;
                        b.port = d.port;
                        b.host = d.host;
                        b.pathname = d.pathname;
                        b.search = d.search;
                        b.hash = d.hash;
                        b.isPortPartOfUrl = d.host.lastIndexOf(":" + d.port) == d.host.length - d.port.length - 1
                    }
                }
            } catch (l) {
                if (l.message === j)
                    throw l
            }
            return b
        },
        shallowCopy: function(b) {
            if (b == a)
                return a;
            else if (!(b instanceof Object))
                return b;
            else if (Array.isArray(b)) {
                for (var e = [], d = 0; d < b.length; d++)
                    e.push(b[d]);
                return e
            } else {
                var f = b.constructor();
                for (var c in b)
                    if (b.hasOwnProperty(c))
                        f[c] = b[c];
                return f
            }
        },
        createObject: function(b) {
            var d = a;
            if (b) {
                d = {};
                for (var e = b.length, c = 0; c < e; c++)
                    d[b[c].name] = b[c].value
            }
            return d
        },
        addClass: function(a, b) {
            if (!OSF.OUtil.hasClass(a, b)) {
                var c = a.getAttribute(n);
                if (c)
                    a.setAttribute(n, c + " " + b);
                else
                    a.setAttribute(n, b)
            }
        },
        removeClass: function(b, c) {
            if (OSF.OUtil.hasClass(b, c)) {
                var a = b.getAttribute(n)
                  , d = new RegExp("(\\s|^)" + c + "(\\s|$)");
                a = a.replace(d, e);
                b.setAttribute(n, a)
            }
        },
        hasClass: function(c, b) {
            var a = c.getAttribute(n);
            return a && a.match(new RegExp("(\\s|^)" + b + "(\\s|$)"))
        },
        focusToFirstTabbable: function(e, i) {
            var g, h = c, f, j = function() {
                h = b
            }, k = function(c, a, b) {
                if (a < 0 || a > c)
                    return d;
                else if (a === 0 && b)
                    return d;
                else if (a === c - 1 && !b)
                    return d;
                if (b)
                    return a - 1;
                else
                    return a + 1
            };
            e = x(e);
            g = i ? e.length - 1 : 0;
            if (e.length === 0)
                return a;
            while (!h && g >= 0 && g < e.length) {
                f = e[g];
                window.focus();
                f.addEventListener(l, j);
                f.focus();
                f.removeEventListener(l, j);
                g = k(e.length, g, i);
                if (!h && f === document.activeElement)
                    h = b
            }
            if (h)
                return f;
            else
                return a
        },
        focusToNextTabbable: function(f, o, m) {
            var j, e, h = c, g, k = function() {
                h = b
            }, n = function(b, c) {
                for (var a = 0; a < b.length; a++)
                    if (b[a] === c)
                        return a;
                return d
            }, i = function(c, a, b) {
                if (a < 0 || a > c)
                    return d;
                else if (a === 0 && b)
                    return d;
                else if (a === c - 1 && !b)
                    return d;
                if (b)
                    return a - 1;
                else
                    return a + 1
            };
            f = x(f);
            j = n(f, o);
            e = i(f.length, j, m);
            if (e < 0)
                return a;
            while (!h && e >= 0 && e < f.length) {
                g = f[e];
                g.addEventListener(l, k);
                g.focus();
                g.removeEventListener(l, k);
                e = i(f.length, e, m);
                if (!h && g === document.activeElement)
                    h = b
            }
            if (h)
                return g;
            else
                return a
        },
        isNullOrUndefined: function(d) {
            if (typeof d === f)
                return b;
            if (d === a)
                return b;
            return c
        },
        stringEndsWith: function(d, a) {
            if (!OSF.OUtil.isNullOrUndefined(d) && !OSF.OUtil.isNullOrUndefined(a)) {
                if (a.length > d.length)
                    return c;
                if (d.substr(d.length - a.length) === a)
                    return b
            }
            return c
        },
        hashCode: function(b) {
            var a = 0;
            if (!OSF.OUtil.isNullOrUndefined(b)) {
                var c = 0
                  , d = b.length;
                while (c < d)
                    a = (a << 5) - a + b.charCodeAt(c++) | 0
            }
            return a
        },
        getValue: function(a, b) {
            if (OSF.OUtil.isNullOrUndefined(a))
                return b;
            return a
        },
        externalNativeFunctionExists: function(a) {
            return a === "unknown" || a !== f
        },
        isMosForWXPEnabled: function() {
            return this.getBooleanSetting("MosForWXPEnabled")
        }
    }
}();
OSF.OUtil.Guid = function() {
    var a = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"];
    return {
        generateNewGuid: function() {
            for (var c = "", d = (new Date).getTime(), b = 0; b < 32 && d > 0; b++) {
                if (b == 8 || b == 12 || b == 16 || b == 20)
                    c += "-";
                c += a[d % 16];
                d = Math.floor(d / 16)
            }
            for (; b < 32; b++) {
                if (b == 8 || b == 12 || b == 16 || b == 20)
                    c += "-";
                c += a[Math.floor(Math.random() * 16)]
            }
            return c
        }
    }
}();
try {
    (function() {
        OSF.Flights = OSF.OUtil.parseFlights(true);
        OSF.DisabledChangeGates = OSF.OUtil.parseDisabledChangeGates(true)
    }
    )()
} catch (ex) {}
window.OSF = OSF;
OSF.OUtil.setNamespace("OSF", window);
OSF.MessageIDs = {
    FetchBundleUrl: 0,
    LoadReactBundle: 1,
    LoadBundleSuccess: 2,
    LoadBundleError: 3
};
OSF.AppName = {
    Unsupported: 0,
    Excel: 1,
    Word: 2,
    PowerPoint: 4,
    Outlook: 8,
    ExcelWebApp: 16,
    WordWebApp: 32,
    OutlookWebApp: 64,
    Project: 128,
    AccessWebApp: 256,
    PowerpointWebApp: 512,
    ExcelIOS: 1024,
    Sway: 2048,
    WordIOS: 4096,
    PowerPointIOS: 8192,
    Access: 16384,
    Lync: 32768,
    OutlookIOS: 65536,
    OneNoteWebApp: 131072,
    OneNote: 262144,
    ExcelWinRT: 524288,
    WordWinRT: 1048576,
    PowerpointWinRT: 2097152,
    OutlookAndroid: 4194304,
    OneNoteWinRT: 8388608,
    ExcelAndroid: 8388609,
    VisioWebApp: 8388610,
    OneNoteIOS: 8388611,
    WordAndroid: 8388613,
    PowerpointAndroid: 8388614,
    Visio: 8388615,
    OneNoteAndroid: 4194305
};
OSF.InternalPerfMarker = {
    DataCoercionBegin: "Agave.HostCall.CoerceDataStart",
    DataCoercionEnd: "Agave.HostCall.CoerceDataEnd"
};
OSF.HostCallPerfMarker = {
    IssueCall: "Agave.HostCall.IssueCall",
    ReceiveResponse: "Agave.HostCall.ReceiveResponse",
    RuntimeExceptionRaised: "Agave.HostCall.RuntimeExecptionRaised"
};
OSF.AgaveHostAction = {
    Select: 0,
    UnSelect: 1,
    CancelDialog: 2,
    InsertAgave: 3,
    CtrlF6In: 4,
    CtrlF6Exit: 5,
    CtrlF6ExitShift: 6,
    SelectWithError: 7,
    NotifyHostError: 8,
    RefreshAddinCommands: 9,
    PageIsReady: 10,
    TabIn: 11,
    TabInShift: 12,
    TabExit: 13,
    TabExitShift: 14,
    EscExit: 15,
    F2Exit: 16,
    ExitNoFocusable: 17,
    ExitNoFocusableShift: 18,
    MouseEnter: 19,
    MouseLeave: 20,
    UpdateTargetUrl: 21,
    InstallCustomFunctions: 22,
    SendTelemetryEvent: 23,
    UninstallCustomFunctions: 24,
    SendMessage: 25,
    LaunchExtensionComponent: 26,
    StopExtensionComponent: 27,
    RestartExtensionComponent: 28,
    EnableTaskPaneHeaderButton: 29,
    DisableTaskPaneHeaderButton: 30,
    TaskPaneHeaderButtonClicked: 31,
    RemoveAppCommandsAddin: 32,
    RefreshRibbonGallery: 33,
    GetOriginalControlId: 34,
    OfficeJsReady: 35,
    InsertDevManifest: 36,
    InsertDevManifestError: 37,
    SendCustomerContent: 38,
    KeyboardShortcuts: 39,
    ReportAddinSkillResult: 47
};
OSF.SharedConstants = {
    NotificationConversationIdSuffix: "_ntf"
};
OSF.DialogMessageType = {
    DialogMessageReceived: 0,
    DialogParentMessageReceived: 1,
    DialogClosed: 12006
};
OSF.OfficeAppContext = function(C, y, t, q, v, z, u, x, B, l, A, n, m, p, j, i, h, g, k, d, f, w, s, b, o, r, e, c) {
    var a = this;
    a._id = C;
    a._appName = y;
    a._appVersion = t;
    a._appUILocale = q;
    a._dataLocale = v;
    a._docUrl = z;
    a._clientMode = u;
    a._settings = x;
    a._reason = B;
    a._osfControlType = l;
    a._eToken = A;
    a._correlationId = n;
    a._appInstanceId = m;
    a._touchEnabled = p;
    a._commerceAllowed = j;
    a._appMinorVersion = i;
    a._requirementMatrix = h;
    a._hostCustomMessage = g;
    a._hostFullVersion = k;
    a._isDialog = false;
    a._clientWindowHeight = d;
    a._clientWindowWidth = f;
    a._addinName = w;
    a._appDomains = s;
    a._dialogRequirementMatrix = b;
    a._featureGates = o;
    a._officeTheme = r;
    a._initialDisplayMode = e;
    a._nestedAppAuthBridgeType = c;
    a.get_id = function() {
        return this._id
    }
    ;
    a.get_appName = function() {
        return this._appName
    }
    ;
    a.get_appVersion = function() {
        return this._appVersion
    }
    ;
    a.get_appUILocale = function() {
        return this._appUILocale
    }
    ;
    a.get_dataLocale = function() {
        return this._dataLocale
    }
    ;
    a.get_docUrl = function() {
        return this._docUrl
    }
    ;
    a.get_clientMode = function() {
        return this._clientMode
    }
    ;
    a.get_bindings = function() {
        return this._bindings
    }
    ;
    a.get_settings = function() {
        return this._settings
    }
    ;
    a.get_reason = function() {
        return this._reason
    }
    ;
    a.get_osfControlType = function() {
        return this._osfControlType
    }
    ;
    a.get_eToken = function() {
        return this._eToken
    }
    ;
    a.get_correlationId = function() {
        return this._correlationId
    }
    ;
    a.get_appInstanceId = function() {
        return this._appInstanceId
    }
    ;
    a.get_touchEnabled = function() {
        return this._touchEnabled
    }
    ;
    a.get_commerceAllowed = function() {
        return this._commerceAllowed
    }
    ;
    a.get_appMinorVersion = function() {
        return this._appMinorVersion
    }
    ;
    a.get_requirementMatrix = function() {
        return this._requirementMatrix
    }
    ;
    a.get_dialogRequirementMatrix = function() {
        return this._dialogRequirementMatrix
    }
    ;
    a.get_hostCustomMessage = function() {
        return this._hostCustomMessage
    }
    ;
    a.get_hostFullVersion = function() {
        return this._hostFullVersion
    }
    ;
    a.get_isDialog = function() {
        return this._isDialog
    }
    ;
    a.get_clientWindowHeight = function() {
        return this._clientWindowHeight
    }
    ;
    a.get_clientWindowWidth = function() {
        return this._clientWindowWidth
    }
    ;
    a.get_addinName = function() {
        return this._addinName
    }
    ;
    a.get_appDomains = function() {
        return this._appDomains
    }
    ;
    a.get_featureGates = function() {
        return this._featureGates
    }
    ;
    a.get_officeTheme = function() {
        return this._officeTheme
    }
    ;
    a.get_initialDisplayMode = function() {
        return this._initialDisplayMode ? this._initialDisplayMode : 0
    }
    ;
    a.get_nestedAppAuthBridgeType = function() {
        return this._nestedAppAuthBridgeType
    }
}
;
OSF.OsfControlType = {
    DocumentLevel: 0,
    ContainerLevel: 1
};
OSF.ClientMode = {
    ReadOnly: 0,
    ReadWrite: 1
};
OSF.OUtil.setNamespace("Microsoft", window);
OSF.OUtil.setNamespace("Office", Microsoft);
OSF.OUtil.setNamespace("Client", Microsoft.Office);
OSF.OUtil.setNamespace("WebExtension", Microsoft.Office);
Microsoft.Office.WebExtension.InitializationReason = {
    Inserted: "inserted",
    DocumentOpened: "documentOpened",
    ControlActivation: "controlActivation"
};
Microsoft.Office.WebExtension.ValueFormat = {
    Unformatted: "unformatted",
    Formatted: "formatted"
};
Microsoft.Office.WebExtension.FilterType = {
    All: "all"
};
Microsoft.Office.WebExtension.Parameters = {
    BindingType: "bindingType",
    CoercionType: "coercionType",
    ValueFormat: "valueFormat",
    FilterType: "filterType",
    Columns: "columns",
    SampleData: "sampleData",
    GoToType: "goToType",
    SelectionMode: "selectionMode",
    Id: "id",
    PromptText: "promptText",
    ItemName: "itemName",
    FailOnCollision: "failOnCollision",
    StartRow: "startRow",
    StartColumn: "startColumn",
    RowCount: "rowCount",
    ColumnCount: "columnCount",
    Callback: "callback",
    AsyncContext: "asyncContext",
    Data: "data",
    JsonData: "jsonData",
    Rows: "rows",
    OverwriteIfStale: "overwriteIfStale",
    FileType: "fileType",
    EventType: "eventType",
    Handler: "handler",
    SliceSize: "sliceSize",
    SliceIndex: "sliceIndex",
    ActiveView: "activeView",
    Status: "status",
    PlatformType: "platformType",
    HostType: "hostType",
    ForceConsent: "forceConsent",
    ForceAddAccount: "forceAddAccount",
    AuthChallenge: "authChallenge",
    AllowConsentPrompt: "allowConsentPrompt",
    ForMSGraphAccess: "forMSGraphAccess",
    AllowSignInPrompt: "allowSignInPrompt",
    JsonPayload: "jsonPayload",
    EnableNewHosts: "enableNewHosts",
    AccountTypeFilter: "accountTypeFilter",
    AddinTrustId: "addinTrustId",
    Reserved: "reserved",
    Tcid: "tcid",
    Xml: "xml",
    Namespace: "namespace",
    Prefix: "prefix",
    XPath: "xPath",
    Text: "text",
    ImageLeft: "imageLeft",
    ImageTop: "imageTop",
    ImageWidth: "imageWidth",
    ImageHeight: "imageHeight",
    TaskId: "taskId",
    FieldId: "fieldId",
    FieldValue: "fieldValue",
    ServerUrl: "serverUrl",
    ListName: "listName",
    ResourceId: "resourceId",
    ViewType: "viewType",
    ViewName: "viewName",
    GetRawValue: "getRawValue",
    CellFormat: "cellFormat",
    TableOptions: "tableOptions",
    TaskIndex: "taskIndex",
    ResourceIndex: "resourceIndex",
    CustomFieldId: "customFieldId",
    Url: "url",
    MessageHandler: "messageHandler",
    Width: "width",
    Height: "height",
    RequireHTTPs: "requireHTTPS",
    MessageToParent: "messageToParent",
    DisplayInIframe: "displayInIframe",
    MessageContent: "messageContent",
    HideTitle: "hideTitle",
    UseDeviceIndependentPixels: "useDeviceIndependentPixels",
    PromptBeforeOpen: "promptBeforeOpen",
    EnforceAppDomain: "enforceAppDomain",
    UrlNoHostInfo: "urlNoHostInfo",
    TargetOrigin: "targetOrigin",
    AppCommandInvocationCompletedData: "appCommandInvocationCompletedData",
    Base64: "base64",
    FormId: "formId"
};
OSF.OUtil.setNamespace("DDA", OSF);
OSF.DDA.DocumentMode = {
    ReadOnly: 1,
    ReadWrite: 0
};
OSF.DDA.PropertyDescriptors = {
    AsyncResultStatus: "AsyncResultStatus"
};
OSF.DDA.EventDescriptors = {};
OSF.DDA.ListDescriptors = {};
OSF.DDA.UI = {};
OSF.DDA.getXdmEventName = function(b, a) {
    if (a == Microsoft.Office.WebExtension.EventType.BindingSelectionChanged || a == Microsoft.Office.WebExtension.EventType.BindingDataChanged || a == Microsoft.Office.WebExtension.EventType.DataNodeDeleted || a == Microsoft.Office.WebExtension.EventType.DataNodeInserted || a == Microsoft.Office.WebExtension.EventType.DataNodeReplaced)
        return b + "_" + a;
    else
        return a
}
;
OSF.DDA.MethodDispId = {
    dispidMethodMin: 64,
    dispidGetSelectedDataMethod: 64,
    dispidSetSelectedDataMethod: 65,
    dispidAddBindingFromSelectionMethod: 66,
    dispidAddBindingFromPromptMethod: 67,
    dispidGetBindingMethod: 68,
    dispidReleaseBindingMethod: 69,
    dispidGetBindingDataMethod: 70,
    dispidSetBindingDataMethod: 71,
    dispidAddRowsMethod: 72,
    dispidClearAllRowsMethod: 73,
    dispidGetAllBindingsMethod: 74,
    dispidLoadSettingsMethod: 75,
    dispidSaveSettingsMethod: 76,
    dispidGetDocumentCopyMethod: 77,
    dispidAddBindingFromNamedItemMethod: 78,
    dispidAddColumnsMethod: 79,
    dispidGetDocumentCopyChunkMethod: 80,
    dispidReleaseDocumentCopyMethod: 81,
    dispidNavigateToMethod: 82,
    dispidGetActiveViewMethod: 83,
    dispidGetDocumentThemeMethod: 84,
    dispidGetOfficeThemeMethod: 85,
    dispidGetFilePropertiesMethod: 86,
    dispidClearFormatsMethod: 87,
    dispidSetTableOptionsMethod: 88,
    dispidSetFormatsMethod: 89,
    dispidExecuteRichApiRequestMethod: 93,
    dispidAppCommandInvocationCompletedMethod: 94,
    dispidCloseContainerMethod: 97,
    dispidGetAccessTokenMethod: 98,
    dispidGetAuthContextMethod: 99,
    dispidOpenBrowserWindow: 102,
    dispidCreateDocumentMethod: 105,
    dispidInsertFormMethod: 106,
    dispidDisplayRibbonCalloutAsyncMethod: 109,
    dispidGetSelectedTaskMethod: 110,
    dispidGetSelectedResourceMethod: 111,
    dispidGetTaskMethod: 112,
    dispidGetResourceFieldMethod: 113,
    dispidGetWSSUrlMethod: 114,
    dispidGetTaskFieldMethod: 115,
    dispidGetProjectFieldMethod: 116,
    dispidGetSelectedViewMethod: 117,
    dispidGetTaskByIndexMethod: 118,
    dispidGetResourceByIndexMethod: 119,
    dispidSetTaskFieldMethod: 120,
    dispidSetResourceFieldMethod: 121,
    dispidGetMaxTaskIndexMethod: 122,
    dispidGetMaxResourceIndexMethod: 123,
    dispidCreateTaskMethod: 124,
    dispidAddDataPartMethod: 128,
    dispidGetDataPartByIdMethod: 129,
    dispidGetDataPartsByNamespaceMethod: 130,
    dispidGetDataPartXmlMethod: 131,
    dispidGetDataPartNodesMethod: 132,
    dispidDeleteDataPartMethod: 133,
    dispidGetDataNodeValueMethod: 134,
    dispidGetDataNodeXmlMethod: 135,
    dispidGetDataNodesMethod: 136,
    dispidSetDataNodeValueMethod: 137,
    dispidSetDataNodeXmlMethod: 138,
    dispidAddDataNamespaceMethod: 139,
    dispidGetDataUriByPrefixMethod: 140,
    dispidGetDataPrefixByUriMethod: 141,
    dispidGetDataNodeTextMethod: 142,
    dispidSetDataNodeTextMethod: 143,
    dispidMessageParentMethod: 144,
    dispidSendMessageMethod: 145,
    dispidExecuteFeature: 146,
    dispidQueryFeature: 147,
    dispidGetNestedAppAuthContextMethod: 205,
    dispidSdxSendMessage: 208,
    dispidNestedAppAuthRequestMethod: 209,
    dispidAddinSkillActionReply: 259
};
OSF.DDA.EventDispId = {
    dispidEventMin: 0,
    dispidInitializeEvent: 0,
    dispidSettingsChangedEvent: 1,
    dispidDocumentSelectionChangedEvent: 2,
    dispidBindingSelectionChangedEvent: 3,
    dispidBindingDataChangedEvent: 4,
    dispidDocumentOpenEvent: 5,
    dispidDocumentCloseEvent: 6,
    dispidActiveViewChangedEvent: 7,
    dispidDocumentThemeChangedEvent: 8,
    dispidOfficeThemeChangedEvent: 9,
    dispidDialogMessageReceivedEvent: 10,
    dispidDialogNotificationShownInAddinEvent: 11,
    dispidDialogParentMessageReceivedEvent: 12,
    dispidObjectDeletedEvent: 13,
    dispidObjectSelectionChangedEvent: 14,
    dispidObjectDataChangedEvent: 15,
    dispidContentControlAddedEvent: 16,
    dispidSuspend: 19,
    dispidResume: 20,
    dispidActivationStatusChangedEvent: 32,
    dispidRichApiMessageEvent: 33,
    dispidAppCommandInvokedEvent: 39,
    dispidOnSdxSendMessageEvent: 40,
    dispidOlkItemSelectedChangedEvent: 46,
    dispidOlkRecipientsChangedEvent: 47,
    dispidOlkAppointmentTimeChangedEvent: 48,
    dispidOlkRecurrenceChangedEvent: 49,
    dispidOlkAttachmentsChangedEvent: 50,
    dispidOlkEnhancedLocationsChangedEvent: 51,
    dispidOlkInfobarClickedEvent: 52,
    dispidOlkSelectedItemsChangedEvent: 53,
    dispidOlkSensitivityLabelChangedEvent: 54,
    dispidOlkInitializationContextChangedEvent: 55,
    dispidTaskSelectionChangedEvent: 56,
    dispidResourceSelectionChangedEvent: 57,
    dispidViewSelectionChangedEvent: 58,
    dispidDataNodeAddedEvent: 60,
    dispidDataNodeReplacedEvent: 61,
    dispidDataNodeDeletedEvent: 62,
    dispidEventMax: 63,
    dispidOlkDragAndDropEvent: 10001
};
OSF.DDA.ErrorCodeManager = function() {
    var a = {};
    return {
        getErrorArgs: function(c) {
            var b = a[c];
            if (!b)
                b = a[this.errorCodes.ooeInternalError];
            else {
                if (!b.name)
                    b.name = a[this.errorCodes.ooeInternalError].name;
                if (!b.message)
                    b.message = a[this.errorCodes.ooeInternalError].message
            }
            return b
        },
        addErrorMessage: function(c, b) {
            a[c] = b
        },
        errorCodes: {
            ooeSuccess: 0,
            ooeChunkResult: 1,
            ooeCoercionTypeNotSupported: 1e3,
            ooeGetSelectionNotMatchDataType: 1001,
            ooeCoercionTypeNotMatchBinding: 1002,
            ooeInvalidGetRowColumnCounts: 1003,
            ooeSelectionNotSupportCoercionType: 1004,
            ooeInvalidGetStartRowColumn: 1005,
            ooeNonUniformPartialGetNotSupported: 1006,
            ooeGetDataIsTooLarge: 1008,
            ooeFileTypeNotSupported: 1009,
            ooeGetDataParametersConflict: 1010,
            ooeInvalidGetColumns: 1011,
            ooeInvalidGetRows: 1012,
            ooeInvalidReadForBlankRow: 1013,
            ooeUnsupportedDataObject: 2e3,
            ooeCannotWriteToSelection: 2001,
            ooeDataNotMatchSelection: 2002,
            ooeOverwriteWorksheetData: 2003,
            ooeDataNotMatchBindingSize: 2004,
            ooeInvalidSetStartRowColumn: 2005,
            ooeInvalidDataFormat: 2006,
            ooeDataNotMatchCoercionType: 2007,
            ooeDataNotMatchBindingType: 2008,
            ooeSetDataIsTooLarge: 2009,
            ooeNonUniformPartialSetNotSupported: 2010,
            ooeInvalidSetColumns: 2011,
            ooeInvalidSetRows: 2012,
            ooeSetDataParametersConflict: 2013,
            ooeCellDataAmountBeyondLimits: 2014,
            ooeSelectionCannotBound: 3e3,
            ooeBindingNotExist: 3002,
            ooeBindingToMultipleSelection: 3003,
            ooeInvalidSelectionForBindingType: 3004,
            ooeOperationNotSupportedOnThisBindingType: 3005,
            ooeNamedItemNotFound: 3006,
            ooeMultipleNamedItemFound: 3007,
            ooeInvalidNamedItemForBindingType: 3008,
            ooeUnknownBindingType: 3009,
            ooeOperationNotSupportedOnMatrixData: 3010,
            ooeInvalidColumnsForBinding: 3011,
            ooeSettingNameNotExist: 4e3,
            ooeSettingsCannotSave: 4001,
            ooeSettingsAreStale: 4002,
            ooeOperationNotSupported: 5e3,
            ooeInternalError: 5001,
            ooeDocumentReadOnly: 5002,
            ooeEventHandlerNotExist: 5003,
            ooeInvalidApiCallInContext: 5004,
            ooeShuttingDown: 5005,
            ooeUnsupportedEnumeration: 5007,
            ooeIndexOutOfRange: 5008,
            ooeBrowserAPINotSupported: 5009,
            ooeInvalidParam: 5010,
            ooeRequestTimeout: 5011,
            ooeInvalidOrTimedOutSession: 5012,
            ooeInvalidApiArguments: 5013,
            ooeOperationCancelled: 5014,
            ooeWorkbookHidden: 5015,
            ooeWriteNotSupportedWhenModalDialogOpen: 5016,
            ooeUndoNotSupported: 5017,
            ooeTooManyIncompleteRequests: 5100,
            ooeRequestTokenUnavailable: 5101,
            ooeActivityLimitReached: 5102,
            ooeRequestPayloadSizeLimitExceeded: 5103,
            ooeResponsePayloadSizeLimitExceeded: 5104,
            ooeCustomXmlNodeNotFound: 6e3,
            ooeCustomXmlError: 6100,
            ooeCustomXmlExceedQuota: 6101,
            ooeCustomXmlOutOfDate: 6102,
            ooeNoCapability: 7e3,
            ooeCannotNavTo: 7001,
            ooeSpecifiedIdNotExist: 7002,
            ooeNavOutOfBound: 7004,
            ooeElementMissing: 8e3,
            ooeProtectedError: 8001,
            ooeInvalidCellsValue: 8010,
            ooeInvalidTableOptionValue: 8011,
            ooeInvalidFormatValue: 8012,
            ooeRowIndexOutOfRange: 8020,
            ooeColIndexOutOfRange: 8021,
            ooeFormatValueOutOfRange: 8022,
            ooeCellFormatAmountBeyondLimits: 8023,
            ooeMemoryFileLimit: 11000,
            ooeNetworkProblemRetrieveFile: 11001,
            ooeInvalidSliceSize: 11002,
            ooeInvalidCallback: 11101,
            ooeInvalidWidth: 12000,
            ooeInvalidHeight: 12001,
            ooeNavigationError: 12002,
            ooeInvalidScheme: 12003,
            ooeAppDomains: 12004,
            ooeRequireHTTPS: 12005,
            ooeWebDialogClosed: 12006,
            ooeDialogAlreadyOpened: 12007,
            ooeEndUserAllow: 12008,
            ooeEndUserIgnore: 12009,
            ooeNotUILessDialog: 12010,
            ooeCrossZone: 12011,
            ooeModalDialogOpen: 12012,
            ooeDocumentIsInactive: 12013,
            ooeDialogParentIsMinimized: 12014,
            ooeNotSSOAgave: 13000,
            ooeSSOUserNotSignedIn: 13001,
            ooeSSOUserAborted: 13002,
            ooeSSOUnsupportedUserIdentity: 13003,
            ooeSSOInvalidResourceUrl: 13004,
            ooeSSOInvalidGrant: 13005,
            ooeSSOClientError: 13006,
            ooeSSOServerError: 13007,
            ooeAddinIsAlreadyRequestingToken: 13008,
            ooeSSOUserConsentNotSupportedByCurrentAddinCategory: 13009,
            ooeSSOConnectionLost: 13010,
            ooeResourceNotAllowed: 13011,
            ooeSSOUnsupportedPlatform: 13012,
            ooeSSOCallThrottled: 13013,
            ooeAccessDenied: 13990,
            ooeGeneralException: 13991
        },
        initializeErrorMessages: function(b) {
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotSupported] = {
                name: b.L_InvalidCoercion,
                message: b.L_CoercionTypeNotSupported
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetSelectionNotMatchDataType] = {
                name: b.L_DataReadError,
                message: b.L_GetSelectionNotSupported
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding] = {
                name: b.L_InvalidCoercion,
                message: b.L_CoercionTypeNotMatchBinding
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRowColumnCounts] = {
                name: b.L_DataReadError,
                message: b.L_InvalidGetRowColumnCounts
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionNotSupportCoercionType] = {
                name: b.L_DataReadError,
                message: b.L_SelectionNotSupportCoercionType
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetStartRowColumn] = {
                name: b.L_DataReadError,
                message: b.L_InvalidGetStartRowColumn
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialGetNotSupported] = {
                name: b.L_DataReadError,
                message: b.L_NonUniformPartialGetNotSupported
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataIsTooLarge] = {
                name: b.L_DataReadError,
                message: b.L_GetDataIsTooLarge
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeFileTypeNotSupported] = {
                name: b.L_DataReadError,
                message: b.L_FileTypeNotSupported
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataParametersConflict] = {
                name: b.L_DataReadError,
                message: b.L_GetDataParametersConflict
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetColumns] = {
                name: b.L_DataReadError,
                message: b.L_InvalidGetColumns
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRows] = {
                name: b.L_DataReadError,
                message: b.L_InvalidGetRows
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidReadForBlankRow] = {
                name: b.L_DataReadError,
                message: b.L_InvalidReadForBlankRow
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject] = {
                name: b.L_DataWriteError,
                message: b.L_UnsupportedDataObject
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotWriteToSelection] = {
                name: b.L_DataWriteError,
                message: b.L_CannotWriteToSelection
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchSelection] = {
                name: b.L_DataWriteError,
                message: b.L_DataNotMatchSelection
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeOverwriteWorksheetData] = {
                name: b.L_DataWriteError,
                message: b.L_OverwriteWorksheetData
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingSize] = {
                name: b.L_DataWriteError,
                message: b.L_DataNotMatchBindingSize
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetStartRowColumn] = {
                name: b.L_DataWriteError,
                message: b.L_InvalidSetStartRowColumn
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidDataFormat] = {
                name: b.L_InvalidFormat,
                message: b.L_InvalidDataFormat
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchCoercionType] = {
                name: b.L_InvalidDataObject,
                message: b.L_DataNotMatchCoercionType
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingType] = {
                name: b.L_InvalidDataObject,
                message: b.L_DataNotMatchBindingType
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataIsTooLarge] = {
                name: b.L_DataWriteError,
                message: b.L_SetDataIsTooLarge
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialSetNotSupported] = {
                name: b.L_DataWriteError,
                message: b.L_NonUniformPartialSetNotSupported
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetColumns] = {
                name: b.L_DataWriteError,
                message: b.L_InvalidSetColumns
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetRows] = {
                name: b.L_DataWriteError,
                message: b.L_InvalidSetRows
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataParametersConflict] = {
                name: b.L_DataWriteError,
                message: b.L_SetDataParametersConflict
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionCannotBound] = {
                name: b.L_BindingCreationError,
                message: b.L_SelectionCannotBound
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingNotExist] = {
                name: b.L_InvalidBindingError,
                message: b.L_BindingNotExist
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingToMultipleSelection] = {
                name: b.L_BindingCreationError,
                message: b.L_BindingToMultipleSelection
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSelectionForBindingType] = {
                name: b.L_BindingCreationError,
                message: b.L_InvalidSelectionForBindingType
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnThisBindingType] = {
                name: b.L_InvalidBindingOperation,
                message: b.L_OperationNotSupportedOnThisBindingType
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNamedItemNotFound] = {
                name: b.L_BindingCreationError,
                message: b.L_NamedItemNotFound
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeMultipleNamedItemFound] = {
                name: b.L_BindingCreationError,
                message: b.L_MultipleNamedItemFound
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidNamedItemForBindingType] = {
                name: b.L_BindingCreationError,
                message: b.L_InvalidNamedItemForBindingType
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnknownBindingType] = {
                name: b.L_InvalidBinding,
                message: b.L_UnknownBindingType
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnMatrixData] = {
                name: b.L_InvalidBindingOperation,
                message: b.L_OperationNotSupportedOnMatrixData
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidColumnsForBinding] = {
                name: b.L_InvalidBinding,
                message: b.L_InvalidColumnsForBinding
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingNameNotExist] = {
                name: b.L_ReadSettingsError,
                message: b.L_SettingNameNotExist
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsCannotSave] = {
                name: b.L_SaveSettingsError,
                message: b.L_SettingsCannotSave
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsAreStale] = {
                name: b.L_SettingsStaleError,
                message: b.L_SettingsAreStale
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupported] = {
                name: b.L_HostError,
                message: b.L_OperationNotSupported
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError] = {
                name: b.L_InternalError,
                message: b.L_InternalErrorDescription
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDocumentReadOnly] = {
                name: b.L_PermissionDenied,
                message: b.L_DocumentReadOnly
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist] = {
                name: b.L_EventRegistrationError,
                message: b.L_EventHandlerNotExist
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext] = {
                name: b.L_InvalidAPICall,
                message: b.L_InvalidApiCallInContext
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeShuttingDown] = {
                name: b.L_ShuttingDown,
                message: b.L_ShuttingDown
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration] = {
                name: b.L_UnsupportedEnumeration,
                message: b.L_UnsupportedEnumerationMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange] = {
                name: b.L_IndexOutOfRange,
                message: b.L_IndexOutOfRange
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeBrowserAPINotSupported] = {
                name: b.L_APINotSupported,
                message: b.L_BrowserAPINotSupported
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTimeout] = {
                name: b.L_APICallFailed,
                message: b.L_RequestTimeout
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidOrTimedOutSession] = {
                name: b.L_InvalidOrTimedOutSession,
                message: b.L_InvalidOrTimedOutSessionMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiArguments] = {
                name: b.L_APICallFailed,
                message: b.L_InvalidApiArgumentsMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeWorkbookHidden] = {
                name: b.L_APICallFailed,
                message: b.L_WorkbookHiddenMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeWriteNotSupportedWhenModalDialogOpen] = {
                name: b.L_APICallFailed,
                message: b.L_WriteNotSupportedWhenModalDialogOpen
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeUndoNotSupported] = {
                name: b.L_APICallFailed,
                message: b.L_UndoNotSupportedMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeTooManyIncompleteRequests] = {
                name: b.L_APICallFailed,
                message: b.L_TooManyIncompleteRequests
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTokenUnavailable] = {
                name: b.L_APICallFailed,
                message: b.L_RequestTokenUnavailable
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeActivityLimitReached] = {
                name: b.L_APICallFailed,
                message: b.L_ActivityLimitReached
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestPayloadSizeLimitExceeded] = {
                name: b.L_APICallFailed,
                message: b.L_RequestPayloadSizeLimitExceededMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeResponsePayloadSizeLimitExceeded] = {
                name: b.L_APICallFailed,
                message: b.L_ResponsePayloadSizeLimitExceededMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlNodeNotFound] = {
                name: b.L_InvalidNode,
                message: b.L_CustomXmlNodeNotFound
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlError] = {
                name: b.L_CustomXmlError,
                message: b.L_CustomXmlError
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlExceedQuota] = {
                name: b.L_CustomXmlExceedQuotaName,
                message: b.L_CustomXmlExceedQuotaMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlOutOfDate] = {
                name: b.L_CustomXmlOutOfDateName,
                message: b.L_CustomXmlOutOfDateMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability] = {
                name: b.L_PermissionDenied,
                message: b.L_NoCapability
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotNavTo] = {
                name: b.L_CannotNavigateTo,
                message: b.L_CannotNavigateTo
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSpecifiedIdNotExist] = {
                name: b.L_SpecifiedIdNotExist,
                message: b.L_SpecifiedIdNotExist
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavOutOfBound] = {
                name: b.L_NavOutOfBound,
                message: b.L_NavOutOfBound
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellDataAmountBeyondLimits] = {
                name: b.L_DataWriteReminder,
                message: b.L_CellDataAmountBeyondLimits
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeElementMissing] = {
                name: b.L_MissingParameter,
                message: b.L_ElementMissing
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeProtectedError] = {
                name: b.L_PermissionDenied,
                message: b.L_NoCapability
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCellsValue] = {
                name: b.L_InvalidValue,
                message: b.L_InvalidCellsValue
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidTableOptionValue] = {
                name: b.L_InvalidValue,
                message: b.L_InvalidTableOptionValue
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidFormatValue] = {
                name: b.L_InvalidValue,
                message: b.L_InvalidFormatValue
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeRowIndexOutOfRange] = {
                name: b.L_OutOfRange,
                message: b.L_RowIndexOutOfRange
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeColIndexOutOfRange] = {
                name: b.L_OutOfRange,
                message: b.L_ColIndexOutOfRange
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeFormatValueOutOfRange] = {
                name: b.L_OutOfRange,
                message: b.L_FormatValueOutOfRange
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellFormatAmountBeyondLimits] = {
                name: b.L_FormattingReminder,
                message: b.L_CellFormatAmountBeyondLimits
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeMemoryFileLimit] = {
                name: b.L_MemoryLimit,
                message: b.L_CloseFileBeforeRetrieve
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNetworkProblemRetrieveFile] = {
                name: b.L_NetworkProblem,
                message: b.L_NetworkProblemRetrieveFile
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize] = {
                name: b.L_InvalidValue,
                message: b.L_SliceSizeNotSupported
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened] = {
                name: b.L_DisplayDialogError,
                message: b.L_DialogAlreadyOpened
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidWidth] = {
                name: b.L_IndexOutOfRange,
                message: b.L_IndexOutOfRange
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidHeight] = {
                name: b.L_IndexOutOfRange,
                message: b.L_IndexOutOfRange
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavigationError] = {
                name: b.L_DisplayDialogError,
                message: b.L_NetworkProblem
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidScheme] = {
                name: b.L_DialogNavigateError,
                message: b.L_DialogInvalidScheme
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeAppDomains] = {
                name: b.L_DisplayDialogError,
                message: b.L_DialogAddressNotTrusted
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequireHTTPS] = {
                name: b.L_DisplayDialogError,
                message: b.L_DialogRequireHTTPS
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeEndUserIgnore] = {
                name: b.L_DisplayDialogError,
                message: b.L_UserClickIgnore
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeCrossZone] = {
                name: b.L_DisplayDialogError,
                message: b.L_NewWindowCrossZoneErrorString
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeModalDialogOpen] = {
                name: b.L_DisplayDialogError,
                message: b.L_ModalDialogOpen
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDocumentIsInactive] = {
                name: b.L_DisplayDialogError,
                message: b.L_DocumentIsInactive
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogParentIsMinimized] = {
                name: b.L_DisplayDialogError,
                message: b.L_DialogParentIsMinimized
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeNotSSOAgave] = {
                name: b.L_APINotSupported,
                message: b.L_InvalidSSOAddinMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserNotSignedIn] = {
                name: b.L_UserNotSignedIn,
                message: b.L_UserNotSignedIn
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserAborted] = {
                name: b.L_UserAborted,
                message: b.L_UserAbortedMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUnsupportedUserIdentity] = {
                name: b.L_UnsupportedUserIdentity,
                message: b.L_UnsupportedUserIdentityMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidResourceUrl] = {
                name: b.L_InvalidResourceUrl,
                message: b.L_InvalidResourceUrlMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidGrant] = {
                name: b.L_InvalidGrant,
                message: b.L_InvalidGrantMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOClientError] = {
                name: b.L_SSOClientError,
                message: b.L_SSOClientErrorMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOServerError] = {
                name: b.L_SSOServerError,
                message: b.L_SSOServerErrorMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeAddinIsAlreadyRequestingToken] = {
                name: b.L_AddinIsAlreadyRequestingToken,
                message: b.L_AddinIsAlreadyRequestingTokenMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserConsentNotSupportedByCurrentAddinCategory] = {
                name: b.L_SSOUserConsentNotSupportedByCurrentAddinCategory,
                message: b.L_SSOUserConsentNotSupportedByCurrentAddinCategoryMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOConnectionLost] = {
                name: b.L_SSOConnectionLostError,
                message: b.L_SSOConnectionLostErrorMessage
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUnsupportedPlatform] = {
                name: b.L_APINotSupported,
                message: b.L_SSOUnsupportedPlatform
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOCallThrottled] = {
                name: b.L_APICallFailed,
                message: b.L_RequestTokenUnavailable
            };
            a[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationCancelled] = {
                name: b.L_OperationCancelledError,
                message: b.L_OperationCancelledErrorMessage
            }
        }
    }
}();
(function(a) {
    var b;
    (function(b) {
        var a = 1.1
          , A = function() {
            function a() {}
            return a
        }();
        b.RequirementVersion = A;
        var d = function() {
            function a(b) {
                var a = this;
                a.isSetSupported = function(d, b) {
                    if (d == undefined)
                        return false;
                    if (b == undefined)
                        b = 0;
                    var f = this._setMap
                      , e = f._sets;
                    if (e.hasOwnProperty(d.toLowerCase())) {
                        var g = e[d.toLowerCase()];
                        try {
                            var a = this._getVersion(g);
                            b = b + "";
                            var c = this._getVersion(b);
                            if (a.major > 0 && a.major > c.major)
                                return true;
                            if (a.major > 0 && a.minor >= 0 && a.major == c.major && a.minor >= c.minor)
                                return true
                        } catch (h) {
                            return false
                        }
                    }
                    return false
                }
                ;
                a._getVersion = function(b) {
                    var a = "version format incorrect";
                    b = b + "";
                    var c = b.split(".")
                      , d = 0
                      , e = 0;
                    if (c.length < 2 && isNaN(Number(b)))
                        throw a;
                    else {
                        d = Number(c[0]);
                        if (c.length >= 2)
                            e = Number(c[1]);
                        if (isNaN(d) || isNaN(e))
                            throw a
                    }
                    var f = {
                        minor: e,
                        major: d
                    };
                    return f
                }
                ;
                a._setMap = b;
                a.isSetSupported = a.isSetSupported.bind(a)
            }
            return a
        }();
        b.RequirementMatrix = d;
        var c = function() {
            function a(a) {
                this._addSetMap = function(a) {
                    for (var b in a)
                        this._sets[b] = a[b]
                }
                ;
                this._sets = a
            }
            return a
        }();
        b.DefaultSetRequirement = c;
        var l = function(c) {
            __extends(b, c);
            function b() {
                return c.call(this, {
                    dialogapi: a
                }) || this
            }
            return b
        }(c);
        b.DefaultRequiredDialogSetRequirement = l;
        var k = function(c) {
            __extends(b, c);
            function b() {
                return c.call(this, {
                    dialogorigin: a
                }) || this
            }
            return b
        }(c);
        b.DefaultOptionalDialogSetRequirement = k;
        var f = function(c) {
            __extends(b, c);
            function b() {
                return c.call(this, {
                    bindingevents: a,
                    documentevents: a,
                    excelapi: a,
                    matrixbindings: a,
                    matrixcoercion: a,
                    selection: a,
                    settings: a,
                    tablebindings: a,
                    tablecoercion: a,
                    textbindings: a,
                    textcoercion: a
                }) || this
            }
            return b
        }(c);
        b.ExcelClientDefaultSetRequirement = f;
        var m = function(c) {
            __extends(b, c);
            function b() {
                var b = c.call(this) || this;
                b._addSetMap({
                    imagecoercion: a
                });
                return b
            }
            return b
        }(f);
        b.ExcelClientV1DefaultSetRequirement = m;
        var n = function(b) {
            __extends(a, b);
            function a() {
                return b.call(this, {
                    mailbox: 1.3
                }) || this
            }
            return a
        }(c);
        b.OutlookClientDefaultSetRequirement = n;
        var h = function(c) {
            __extends(b, c);
            function b() {
                return c.call(this, {
                    bindingevents: a,
                    compressedfile: a,
                    customxmlparts: a,
                    documentevents: a,
                    file: a,
                    htmlcoercion: a,
                    matrixbindings: a,
                    matrixcoercion: a,
                    ooxmlcoercion: a,
                    pdffile: a,
                    selection: a,
                    settings: a,
                    tablebindings: a,
                    tablecoercion: a,
                    textbindings: a,
                    textcoercion: a,
                    textfile: a,
                    wordapi: a
                }) || this
            }
            return b
        }(c);
        b.WordClientDefaultSetRequirement = h;
        var r = function(c) {
            __extends(b, c);
            function b() {
                var b = c.call(this) || this;
                b._addSetMap({
                    customxmlparts: 1.2,
                    wordapi: 1.2,
                    imagecoercion: a
                });
                return b
            }
            return b
        }(h);
        b.WordClientV1DefaultSetRequirement = r;
        var e = function(c) {
            __extends(b, c);
            function b() {
                return c.call(this, {
                    activeview: a,
                    compressedfile: a,
                    documentevents: a,
                    file: a,
                    pdffile: a,
                    selection: a,
                    settings: a,
                    textcoercion: a
                }) || this
            }
            return b
        }(c);
        b.PowerpointClientDefaultSetRequirement = e;
        var j = function(c) {
            __extends(b, c);
            function b() {
                var b = c.call(this) || this;
                b._addSetMap({
                    imagecoercion: a
                });
                return b
            }
            return b
        }(e);
        b.PowerpointClientV1DefaultSetRequirement = j;
        var q = function(c) {
            __extends(b, c);
            function b() {
                return c.call(this, {
                    selection: a,
                    textcoercion: a
                }) || this
            }
            return b
        }(c);
        b.ProjectClientDefaultSetRequirement = q;
        var w = function(c) {
            __extends(b, c);
            function b() {
                return c.call(this, {
                    bindingevents: a,
                    documentevents: a,
                    matrixbindings: a,
                    matrixcoercion: a,
                    selection: a,
                    settings: a,
                    tablebindings: a,
                    tablecoercion: a,
                    textbindings: a,
                    textcoercion: a,
                    file: a
                }) || this
            }
            return b
        }(c);
        b.ExcelWebDefaultSetRequirement = w;
        var y = function(c) {
            __extends(b, c);
            function b() {
                return c.call(this, {
                    compressedfile: a,
                    documentevents: a,
                    file: a,
                    imagecoercion: a,
                    matrixcoercion: a,
                    ooxmlcoercion: a,
                    pdffile: a,
                    selection: a,
                    settings: a,
                    tablecoercion: a,
                    textcoercion: a,
                    textfile: a
                }) || this
            }
            return b
        }(c);
        b.WordWebDefaultSetRequirement = y;
        var p = function(c) {
            __extends(b, c);
            function b() {
                return c.call(this, {
                    activeview: a,
                    settings: a
                }) || this
            }
            return b
        }(c);
        b.PowerpointWebDefaultSetRequirement = p;
        var g = function(b) {
            __extends(a, b);
            function a() {
                return b.call(this, {
                    mailbox: 1.3
                }) || this
            }
            return a
        }(c);
        b.OutlookWebDefaultSetRequirement = g;
        var x = function(c) {
            __extends(b, c);
            function b() {
                return c.call(this, {
                    activeview: a,
                    documentevents: a,
                    selection: a,
                    settings: a,
                    textcoercion: a
                }) || this
            }
            return b
        }(c);
        b.SwayWebDefaultSetRequirement = x;
        var t = function(c) {
            __extends(b, c);
            function b() {
                return c.call(this, {
                    bindingevents: a,
                    partialtablebindings: a,
                    settings: a,
                    tablebindings: a,
                    tablecoercion: a
                }) || this
            }
            return b
        }(c);
        b.AccessWebDefaultSetRequirement = t;
        var v = function(c) {
            __extends(b, c);
            function b() {
                return c.call(this, {
                    bindingevents: a,
                    documentevents: a,
                    matrixbindings: a,
                    matrixcoercion: a,
                    selection: a,
                    settings: a,
                    tablebindings: a,
                    tablecoercion: a,
                    textbindings: a,
                    textcoercion: a
                }) || this
            }
            return b
        }(c);
        b.ExcelIOSDefaultSetRequirement = v;
        var i = function(c) {
            __extends(b, c);
            function b() {
                return c.call(this, {
                    bindingevents: a,
                    compressedfile: a,
                    customxmlparts: a,
                    documentevents: a,
                    file: a,
                    htmlcoercion: a,
                    matrixbindings: a,
                    matrixcoercion: a,
                    ooxmlcoercion: a,
                    pdffile: a,
                    selection: a,
                    settings: a,
                    tablebindings: a,
                    tablecoercion: a,
                    textbindings: a,
                    textcoercion: a,
                    textfile: a
                }) || this
            }
            return b
        }(c);
        b.WordIOSDefaultSetRequirement = i;
        var u = function(b) {
            __extends(a, b);
            function a() {
                var a = b.call(this) || this;
                a._addSetMap({
                    customxmlparts: 1.2,
                    wordapi: 1.2
                });
                return a
            }
            return a
        }(i);
        b.WordIOSV1DefaultSetRequirement = u;
        var o = function(c) {
            __extends(b, c);
            function b() {
                return c.call(this, {
                    activeview: a,
                    compressedfile: a,
                    documentevents: a,
                    file: a,
                    pdffile: a,
                    selection: a,
                    settings: a,
                    textcoercion: a
                }) || this
            }
            return b
        }(c);
        b.PowerpointIOSDefaultSetRequirement = o;
        var s = function(c) {
            __extends(b, c);
            function b() {
                return c.call(this, {
                    mailbox: a
                }) || this
            }
            return b
        }(c);
        b.OutlookIOSDefaultSetRequirement = s;
        var z = function() {
            var b = "undefined";
            function a() {}
            a.initializeOsfDda = function() {
                OSF.OUtil.setNamespace("Requirement", OSF.DDA)
            }
            ;
            a.getDefaultRequirementMatrix = function(f) {
                this.initializeDefaultSetMatrix();
                var e = undefined
                  , g = f.get_requirementMatrix();
                if (g != undefined && g.length > 0 && typeof JSON !== b) {
                    var i = JSON.parse(f.get_requirementMatrix().toLowerCase());
                    e = new d(new c(i))
                } else {
                    var h = a.getClientFullVersionString(f);
                    if (a.DefaultSetArrayMatrix != undefined && a.DefaultSetArrayMatrix[h] != undefined)
                        e = new d(a.DefaultSetArrayMatrix[h]);
                    else
                        e = new d(new c({}))
                }
                return e
            }
            ;
            a.getDefaultDialogRequirementMatrix = function(h) {
                var a = undefined
                  , i = h.get_dialogRequirementMatrix();
                if (i != undefined && i.length > 0 && typeof JSON !== b) {
                    var f = JSON.parse(h.get_requirementMatrix().toLowerCase());
                    a = new c(f)
                } else {
                    a = new l;
                    var g = h.get_requirementMatrix();
                    if (g != undefined && g.length > 0 && typeof JSON !== b) {
                        var f = JSON.parse(g.toLowerCase());
                        for (var e in a._sets)
                            if (f.hasOwnProperty(e))
                                a._sets[e] = f[e];
                        var j = new k;
                        for (var e in j._sets)
                            if (f.hasOwnProperty(e))
                                a._sets[e] = f[e]
                    }
                }
                return new d(a)
            }
            ;
            a.getClientFullVersionString = function(a) {
                var d = a.get_appMinorVersion()
                  , e = ""
                  , b = ""
                  , c = a.get_appName()
                  , f = c == 1024 || c == 4096 || c == 8192 || c == 65536;
                if (f && a.get_appVersion() == 1)
                    if (c == 4096 && d >= 15)
                        b = "16.00.01";
                    else
                        b = "16.00";
                else if (a.get_appName() == 64)
                    b = a.get_appVersion();
                else {
                    if (d < 10)
                        e = "0" + d;
                    else
                        e = "" + d;
                    b = a.get_appVersion() + "." + e
                }
                return a.get_appName() + "-" + b
            }
            ;
            a.initializeDefaultSetMatrix = function() {
                a.DefaultSetArrayMatrix[a.Excel_RCLIENT_1600] = new f;
                a.DefaultSetArrayMatrix[a.Word_RCLIENT_1600] = new h;
                a.DefaultSetArrayMatrix[a.PowerPoint_RCLIENT_1600] = new e;
                a.DefaultSetArrayMatrix[a.Excel_RCLIENT_1601] = new m;
                a.DefaultSetArrayMatrix[a.Word_RCLIENT_1601] = new r;
                a.DefaultSetArrayMatrix[a.PowerPoint_RCLIENT_1601] = new j;
                a.DefaultSetArrayMatrix[a.Outlook_RCLIENT_1600] = new n;
                a.DefaultSetArrayMatrix[a.Excel_WAC_1600] = new w;
                a.DefaultSetArrayMatrix[a.Word_WAC_1600] = new y;
                a.DefaultSetArrayMatrix[a.Outlook_WAC_1600] = new g;
                a.DefaultSetArrayMatrix[a.Outlook_WAC_1601] = new g;
                a.DefaultSetArrayMatrix[a.Project_RCLIENT_1600] = new q;
                a.DefaultSetArrayMatrix[a.Access_WAC_1600] = new t;
                a.DefaultSetArrayMatrix[a.PowerPoint_WAC_1600] = new p;
                a.DefaultSetArrayMatrix[a.Excel_IOS_1600] = new v;
                a.DefaultSetArrayMatrix[a.SWAY_WAC_1600] = new x;
                a.DefaultSetArrayMatrix[a.Word_IOS_1600] = new i;
                a.DefaultSetArrayMatrix[a.Word_IOS_16001] = new u;
                a.DefaultSetArrayMatrix[a.PowerPoint_IOS_1600] = new o;
                a.DefaultSetArrayMatrix[a.Outlook_IOS_1600] = new s
            }
            ;
            a.Excel_RCLIENT_1600 = "1-16.00";
            a.Excel_RCLIENT_1601 = "1-16.01";
            a.Word_RCLIENT_1600 = "2-16.00";
            a.Word_RCLIENT_1601 = "2-16.01";
            a.PowerPoint_RCLIENT_1600 = "4-16.00";
            a.PowerPoint_RCLIENT_1601 = "4-16.01";
            a.Outlook_RCLIENT_1600 = "8-16.00";
            a.Excel_WAC_1600 = "16-16.00";
            a.Word_WAC_1600 = "32-16.00";
            a.Outlook_WAC_1600 = "64-16.00";
            a.Outlook_WAC_1601 = "64-16.01";
            a.Project_RCLIENT_1600 = "128-16.00";
            a.Access_WAC_1600 = "256-16.00";
            a.PowerPoint_WAC_1600 = "512-16.00";
            a.Excel_IOS_1600 = "1024-16.00";
            a.SWAY_WAC_1600 = "2048-16.00";
            a.Word_IOS_1600 = "4096-16.00";
            a.Word_IOS_16001 = "4096-16.00.01";
            a.PowerPoint_IOS_1600 = "8192-16.00";
            a.Outlook_IOS_1600 = "65536-16.00";
            a.DefaultSetArrayMatrix = {};
            return a
        }();
        b.RequirementsMatrixFactory = z
    }
    )(b = a.Requirement || (a.Requirement = {}))
}
)(OfficeExt || (OfficeExt = {}));
OfficeExt.Requirement.RequirementsMatrixFactory.initializeOsfDda();
Microsoft.Office.WebExtension.ApplicationMode = {
    WebEditor: "webEditor",
    WebViewer: "webViewer",
    Client: "client"
};
Microsoft.Office.WebExtension.DocumentMode = {
    ReadOnly: "readOnly",
    ReadWrite: "readWrite"
};
OSF.NamespaceManager = function() {
    var b, a = false;
    return {
        enableShortcut: function() {
            if (!a) {
                if (window.Office)
                    b = window.Office;
                else
                    OSF.OUtil.setNamespace("Office", window);
                window.Office = Microsoft.Office.WebExtension;
                a = true
            }
        },
        disableShortcut: function() {
            if (a) {
                if (b)
                    window.Office = b;
                else
                    OSF.OUtil.unsetNamespace("Office", window);
                a = false
            }
        }
    }
}();
OSF.NamespaceManager.enableShortcut();
Microsoft.Office.WebExtension.useShortNamespace = function(a) {
    if (a)
        OSF.NamespaceManager.enableShortcut();
    else
        OSF.NamespaceManager.disableShortcut()
}
;
Microsoft.Office.WebExtension.select = function(a, b) {
    var c;
    if (a && typeof a == "string") {
        var d = a.indexOf("#");
        if (d != -1) {
            var h = a.substring(0, d)
              , g = a.substring(d + 1);
            switch (h) {
            case "binding":
            case "bindings":
                if (g)
                    c = new OSF.DDA.BindingPromise(g)
            }
        }
    }
    if (!c) {
        if (b) {
            var e = typeof b;
            if (e == "function") {
                var f = {};
                f[Microsoft.Office.WebExtension.Parameters.Callback] = b;
                OSF.DDA.issueAsyncResult(f, OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext, OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext))
            } else
                throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction, e)
        }
    } else {
        c.onFail = b;
        return c
    }
}
;
OSF.DDA.Context = function(a, j, k, c, d) {
    var i = "officeTheme"
      , h = "requirements"
      , b = this;
    OSF.OUtil.defineEnumerableProperties(b, {
        contentLanguage: {
            value: a.get_dataLocale()
        },
        displayLanguage: {
            value: a.get_appUILocale()
        },
        touchEnabled: {
            value: a.get_touchEnabled()
        },
        commerceAllowed: {
            value: a.get_commerceAllowed()
        },
        host: {
            value: OfficeExt.HostName.Host.getInstance().getHost()
        },
        platform: {
            value: OfficeExt.HostName.Host.getInstance().getPlatform()
        },
        isDialog: {
            value: OSF._OfficeAppFactory.getHostInfo().isDialog
        },
        diagnostics: {
            value: OfficeExt.HostName.Host.getInstance().getDiagnostics(a.get_hostFullVersion())
        }
    });
    k && OSF.OUtil.defineEnumerableProperty(b, "license", {
        value: k
    });
    a.ui && OSF.OUtil.defineEnumerableProperty(b, "ui", {
        value: a.ui
    });
    a.auth && OSF.OUtil.defineEnumerableProperty(b, "auth", {
        value: a.auth
    });
    a.webAuth && OSF.OUtil.defineEnumerableProperty(b, "webAuth", {
        value: a.webAuth
    });
    a.partitionKey && OSF.OUtil.defineEnumerableProperty(b, "partitionKey", {
        value: a.partitionKey
    });
    a.application && OSF.OUtil.defineEnumerableProperty(b, "application", {
        value: a.application
    });
    a.extensionLifeCycle && OSF.OUtil.defineEnumerableProperty(b, "extensionLifeCycle", {
        value: a.extensionLifeCycle
    });
    a.messaging && OSF.OUtil.defineEnumerableProperty(b, "messaging", {
        value: a.messaging
    });
    a.ui && a.ui.taskPaneAction && OSF.OUtil.defineEnumerableProperty(b, "taskPaneAction", {
        value: a.ui.taskPaneAction
    });
    a.ui && a.ui.ribbonGallery && OSF.OUtil.defineEnumerableProperty(b, "ribbonGallery", {
        value: a.ui.ribbonGallery
    });
    if (a.get_isDialog()) {
        var f = OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultDialogRequirementMatrix(a);
        OSF.OUtil.defineEnumerableProperty(b, h, {
            value: f
        })
    } else {
        j && OSF.OUtil.defineEnumerableProperty(b, "document", {
            value: j
        });
        if (c) {
            var l = c.displayName || "appOM";
            delete c.displayName;
            OSF.OUtil.defineEnumerableProperty(b, l, {
                value: c
            })
        }
        if (a.get_officeTheme())
            OSF.OUtil.defineEnumerableProperty(b, i, {
                "get": function() {
                    return a.get_officeTheme()
                }
            });
        else
            d && OSF.OUtil.defineEnumerableProperty(b, i, {
                "get": function() {
                    return d()
                }
            });
        var f = OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultRequirementMatrix(a);
        OSF.OUtil.defineEnumerableProperty(b, h, {
            value: f
        });
        if (a.get_featureGates) {
            var e = a.get_featureGates();
            if (e && e["EnablePublicThemeManager"]) {
                var g = new OSF.DDA.Theming.InternalThemeHandler;
                g.InitializeThemeManager();
                OSF.OUtil.defineEnumerableProperty(b, "themeManager", {
                    value: g
                })
            }
        }
    }
}
;
OSF.DDA.OutlookContext = function(a, f, h, i, g) {
    var e = "roamingSettings"
      , b = this;
    OSF.DDA.OutlookContext.uber.constructor.call(b, a, null, h, i, g);
    if (a && a.appOM && a.appOM.initialData && a.appOM.initialData.roamingSettings) {
        var d;
        d = OSF.InitializationHelper.prototype.deserializeSettings({
            SettingsKey: a.appOM.initialData.roamingSettings
        }, false);
        OSF.OUtil.defineEnumerableProperty(b, e, {
            value: d
        })
    } else
        f && OSF.OUtil.defineEnumerableProperty(b, e, {
            value: f
        });
    a.sensitivityLabelsCatalog && OSF.OUtil.defineEnumerableProperty(b, "sensitivityLabelsCatalog", {
        value: a.sensitivityLabelsCatalog()
    });
    a.devicePermission && a.get_appName() == 64 && OSF.OUtil.defineEnumerableProperty(window.Office, "devicePermission", {
        value: a.devicePermission()
    });
    if (a.urls) {
        var c = {};
        try {
            c = JSON.parse(a.urls)
        } catch (j) {}
        c && OSF.OUtil.defineEnumerableProperty(b, "urls", {
            value: c
        })
    }
}
;
OSF.OUtil.extend(OSF.DDA.OutlookContext, OSF.DDA.Context);
OSF.DDA.OutlookAppOm = function() {}
;
OSF.DDA.Application = function() {}
;
OSF.DDA.Document = function(b, c) {
    var a;
    switch (b.get_clientMode()) {
    case OSF.ClientMode.ReadOnly:
        a = Microsoft.Office.WebExtension.DocumentMode.ReadOnly;
        break;
    case OSF.ClientMode.ReadWrite:
        a = Microsoft.Office.WebExtension.DocumentMode.ReadWrite
    }
    c && OSF.OUtil.defineEnumerableProperty(this, "settings", {
        value: c
    });
    OSF.OUtil.defineMutableProperties(this, {
        mode: {
            value: a
        },
        url: {
            value: b.get_docUrl()
        }
    })
}
;
OSF.DDA.JsomDocument = function(d, b, e) {
    var a = this;
    OSF.DDA.JsomDocument.uber.constructor.call(a, d, e);
    b && OSF.OUtil.defineEnumerableProperty(a, "bindings", {
        "get": function() {
            return b
        }
    });
    var c = OSF.DDA.AsyncMethodNames;
    OSF.DDA.DispIdHost.addAsyncMethods(a, [c.GetSelectedDataAsync, c.SetSelectedDataAsync]);
    OSF.DDA.DispIdHost.addEventSupport(a, new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged]))
}
;
OSF.OUtil.extend(OSF.DDA.JsomDocument, OSF.DDA.Document);
OSF.OUtil.defineEnumerableProperty(Microsoft.Office.WebExtension, "context", {
    "get": function() {
        var a;
        if (OSF && OSF._OfficeAppFactory)
            a = OSF._OfficeAppFactory.getContext();
        return a
    }
});
OSF.DDA.License = function(a) {
    OSF.OUtil.defineEnumerableProperty(this, "value", {
        value: a
    })
}
;
OSF.DDA.ApiMethodCall = function(d, f, c, g, h) {
    var a = this
      , e = d.length
      , b = OSF.OUtil.delayExecutionAndCache(function() {
        return OSF.OUtil.formatString(Strings.OfficeOM.L_InvalidParameters, h)
    });
    a.verifyArguments = function(d, f) {
        for (var e in d) {
            var a = d[e]
              , c = f[e];
            if (a["enum"])
                switch (typeof c) {
                case "string":
                    if (OSF.OUtil.listContainsValue(a["enum"], c))
                        break;
                case "undefined":
                    throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration;
                default:
                    throw b()
                }
            if (a["types"])
                if (!OSF.OUtil.listContainsValue(a["types"], typeof c))
                    throw b()
        }
    }
    ;
    a.extractRequiredArguments = function(g, l, j) {
        if (g.length < e)
            throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_MissingRequiredArguments);
        for (var c = [], a = 0; a < e; a++)
            c.push(g[a]);
        this.verifyArguments(d, c);
        var i = {};
        for (a = 0; a < e; a++) {
            var f = d[a]
              , h = c[a];
            if (f.verify) {
                var k = f.verify(h, l, j);
                if (!k)
                    throw b()
            }
            i[f.name] = h
        }
        return i
    }
    ,
    a.fillOptions = function(a, e, h, g) {
        a = a || {};
        for (var d in f)
            if (!OSF.OUtil.listContainsKey(a, d)) {
                var c = undefined
                  , b = f[d];
                if (b.calculate && e)
                    c = b.calculate(e, h, g);
                if (!c && b.defaultValue !== undefined)
                    c = b.defaultValue;
                a[d] = c
            }
        return a
    }
    ;
    a.constructCallArgs = function(e, f, h, d) {
        var a = {};
        for (var j in e)
            a[j] = e[j];
        for (var i in f)
            a[i] = f[i];
        for (var b in c)
            if (c.hasOwnProperty(b))
                a[b] = c[b](h, d);
        if (g)
            a = g(a, h, d);
        return a
    }
}
;
OSF.OUtil.setNamespace("AsyncResultEnum", OSF.DDA);
OSF.DDA.AsyncResultEnum.Properties = {
    Context: "Context",
    Value: "Value",
    Status: "Status",
    Error: "Error"
};
Microsoft.Office.WebExtension.AsyncResultStatus = {
    Succeeded: "succeeded",
    Failed: "failed"
};
OSF.DDA.AsyncResultEnum.ErrorCode = {
    Success: 0,
    Failed: 1
};
OSF.DDA.AsyncResultEnum.ErrorProperties = {
    Name: "Name",
    Message: "Message",
    Code: "Code"
};
OSF.DDA.AsyncMethodNames = {};
OSF.DDA.AsyncMethodNames.addNames = function(b) {
    for (var a in b) {
        var c = {};
        OSF.OUtil.defineEnumerableProperties(c, {
            id: {
                value: a
            },
            displayName: {
                value: b[a]
            }
        });
        OSF.DDA.AsyncMethodNames[a] = c
    }
}
;
OSF.DDA.AsyncMethodCall = function(d, e, i, f, g, j, k) {
    var a = "function"
      , c = d.length
      , b = new OSF.DDA.ApiMethodCall(d,e,i,j,k);
    function h(h, j, l, k) {
        if (h.length > c + 2)
            throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);
        for (var d, f, i = h.length - 1; i >= c; i--) {
            var g = h[i];
            switch (typeof g) {
            case "object":
                if (d)
                    throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);
                else
                    d = g;
                break;
            case a:
                if (f)
                    throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalFunction);
                else
                    f = g;
                break;
            default:
                throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument)
            }
        }
        d = b.fillOptions(d, j, l, k);
        if (f)
            if (d[Microsoft.Office.WebExtension.Parameters.Callback])
                throw Strings.OfficeOM.L_RedundantCallbackSpecification;
            else
                d[Microsoft.Office.WebExtension.Parameters.Callback] = f;
        b.verifyArguments(e, d);
        return d
    }
    this.verifyAndExtractCall = function(e, c, a) {
        var d = b.extractRequiredArguments(e, c, a)
          , g = h(e, d, c, a)
          , f = b.constructCallArgs(d, g, c, a);
        return f
    }
    ;
    this.processResponse = function(c, b, e, d) {
        var a;
        if (c == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
            if (f)
                a = f(b, e, d);
            else
                a = b;
        else if (g)
            a = g(c, b);
        else
            a = OSF.DDA.ErrorCodeManager.getErrorArgs(c);
        return a
    }
    ;
    this.getCallArgs = function(g) {
        for (var b, d, f = g.length - 1; f >= c; f--) {
            var e = g[f];
            switch (typeof e) {
            case "object":
                b = e;
                break;
            case a:
                d = e
            }
        }
        b = b || {};
        if (d)
            b[Microsoft.Office.WebExtension.Parameters.Callback] = d;
        return b
    }
}
;
OSF.DDA.AsyncMethodCallFactory = function() {
    return {
        manufacture: function(a) {
            var c = a.supportedOptions ? OSF.OUtil.createObject(a.supportedOptions) : []
              , b = a.privateStateCallbacks ? OSF.OUtil.createObject(a.privateStateCallbacks) : [];
            return new OSF.DDA.AsyncMethodCall(a.requiredArguments || [],c,b,a.onSucceeded,a.onFailed,a.checkCallArgs,a.method.displayName)
        }
    }
}();
OSF.DDA.AsyncMethodCalls = {};
OSF.DDA.AsyncMethodCalls.define = function(a) {
    OSF.DDA.AsyncMethodCalls[a.method.id] = OSF.DDA.AsyncMethodCallFactory.manufacture(a)
}
;
OSF.DDA.Error = function(c, a, b) {
    OSF.OUtil.defineEnumerableProperties(this, {
        name: {
            value: c
        },
        message: {
            value: a
        },
        code: {
            value: b
        }
    })
}
;
OSF.DDA.AsyncResult = function(b, a) {
    OSF.OUtil.defineEnumerableProperties(this, {
        value: {
            value: b[OSF.DDA.AsyncResultEnum.Properties.Value]
        },
        status: {
            value: a ? Microsoft.Office.WebExtension.AsyncResultStatus.Failed : Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded
        }
    });
    b[OSF.DDA.AsyncResultEnum.Properties.Context] && OSF.OUtil.defineEnumerableProperty(this, "asyncContext", {
        value: b[OSF.DDA.AsyncResultEnum.Properties.Context]
    });
    a && OSF.OUtil.defineEnumerableProperty(this, "error", {
        value: new OSF.DDA.Error(a[OSF.DDA.AsyncResultEnum.ErrorProperties.Name],a[OSF.DDA.AsyncResultEnum.ErrorProperties.Message],a[OSF.DDA.AsyncResultEnum.ErrorProperties.Code])
    })
}
;
OSF.DDA.issueAsyncResult = function(d, f, a) {
    var e = d[Microsoft.Office.WebExtension.Parameters.Callback];
    if (e) {
        var c = {};
        c[OSF.DDA.AsyncResultEnum.Properties.Context] = d[Microsoft.Office.WebExtension.Parameters.AsyncContext];
        var b;
        if (f == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
            c[OSF.DDA.AsyncResultEnum.Properties.Value] = a;
        else {
            b = {};
            a = a || OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
            b[OSF.DDA.AsyncResultEnum.ErrorProperties.Code] = f || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
            b[OSF.DDA.AsyncResultEnum.ErrorProperties.Name] = a.name || a;
            b[OSF.DDA.AsyncResultEnum.ErrorProperties.Message] = a.message || a
        }
        e(new OSF.DDA.AsyncResult(c,b))
    }
}
;
OSF.DDA.SyncMethodNames = {};
OSF.DDA.SyncMethodNames.addNames = function(b) {
    for (var a in b) {
        var c = {};
        OSF.OUtil.defineEnumerableProperties(c, {
            id: {
                value: a
            },
            displayName: {
                value: b[a]
            }
        });
        OSF.DDA.SyncMethodNames[a] = c
    }
}
;
OSF.DDA.SyncMethodCall = function(b, c, f, g, h) {
    var d = b.length
      , a = new OSF.DDA.ApiMethodCall(b,c,f,g,h);
    function e(e, h, j, i) {
        if (e.length > d + 1)
            throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);
        for (var b, k, f = e.length - 1; f >= d; f--) {
            var g = e[f];
            switch (typeof g) {
            case "object":
                if (b)
                    throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);
                else
                    b = g;
                break;
            default:
                throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument)
            }
        }
        b = a.fillOptions(b, h, j, i);
        a.verifyArguments(c, b);
        return b
    }
    this.verifyAndExtractCall = function(f, c, b) {
        var d = a.extractRequiredArguments(f, c, b)
          , h = e(f, d, c, b)
          , g = a.constructCallArgs(d, h, c, b);
        return g
    }
}
;
OSF.DDA.SyncMethodCallFactory = function() {
    return {
        manufacture: function(a) {
            var b = a.supportedOptions ? OSF.OUtil.createObject(a.supportedOptions) : [];
            return new OSF.DDA.SyncMethodCall(a.requiredArguments || [],b,a.privateStateCallbacks,a.checkCallArgs,a.method.displayName)
        }
    }
}();
OSF.DDA.SyncMethodCalls = {};
OSF.DDA.SyncMethodCalls.define = function(a) {
    OSF.DDA.SyncMethodCalls[a.method.id] = OSF.DDA.SyncMethodCallFactory.manufacture(a)
}
;
OSF.DDA.ListType = function() {
    var a = {};
    return {
        setListType: function(c, b) {
            a[c] = b
        },
        isListType: function(b) {
            return OSF.OUtil.listContainsKey(a, b)
        },
        getDescriptor: function(b) {
            return a[b]
        }
    }
}();
OSF.DDA.HostParameterMap = function(b, c) {
    var j = "fromHost"
      , a = this
      , i = "toHost"
      , e = j
      , l = "sourceData"
      , g = "self"
      , d = {};
    d[Microsoft.Office.WebExtension.Parameters.Data] = {
        toHost: function(a) {
            if (a != null && a.rows !== undefined) {
                var b = {};
                b[OSF.DDA.TableDataProperties.TableRows] = a.rows;
                b[OSF.DDA.TableDataProperties.TableHeaders] = a.headers;
                a = b
            }
            return a
        },
        fromHost: function(a) {
            return a
        }
    };
    d[Microsoft.Office.WebExtension.Parameters.JsonData] = {
        toHost: function(a) {
            return a
        },
        fromHost: function(a) {
            return typeof a === "string" ? JSON.parse(a) : a
        }
    };
    d[Microsoft.Office.WebExtension.Parameters.SampleData] = d[Microsoft.Office.WebExtension.Parameters.Data];
    function f(j, i) {
        var m = j ? {} : undefined;
        for (var h in j) {
            var g = j[h], a;
            if (OSF.DDA.ListType.isListType(h)) {
                a = [];
                for (var n in g)
                    a.push(f(g[n], i))
            } else if (OSF.OUtil.listContainsKey(d, h))
                a = d[h][i](g);
            else if (i == e && b.preserveNesting(h))
                a = f(g, i);
            else {
                var k = c[h];
                if (k) {
                    var l = k[i];
                    if (l) {
                        a = l[g];
                        if (a === undefined)
                            a = g
                    }
                } else
                    a = g
            }
            m[h] = a
        }
        return m
    }
    function k(j, h) {
        var e;
        for (var a in h) {
            var d;
            if (b.isComplexType(a))
                d = k(j, c[a][i]);
            else
                d = j[a];
            if (d != undefined) {
                if (!e)
                    e = {};
                var f = h[a];
                if (f == g)
                    f = a;
                e[f] = b.pack(a, d)
            }
        }
        return e
    }
    function h(j, n, f) {
        if (!f)
            f = {};
        for (var a in n) {
            var k = n[a], d;
            if (k == g)
                d = j;
            else if (k == l) {
                f[a] = j.toArray();
                continue
            } else
                d = j[k];
            if (d === null || d === undefined)
                f[a] = undefined;
            else {
                d = b.unpack(a, d);
                var i;
                if (b.isComplexType(a)) {
                    i = c[a][e];
                    if (b.preserveNesting(a))
                        f[a] = h(d, i);
                    else
                        h(d, i, f)
                } else if (OSF.DDA.ListType.isListType(a)) {
                    i = {};
                    var p = OSF.DDA.ListType.getDescriptor(a);
                    i[p] = g;
                    var m = new Array(d.length);
                    for (var o in d)
                        m[o] = h(d[o], i);
                    f[a] = m
                } else
                    f[a] = d
            }
        }
        return f
    }
    function m(l, e, a) {
        var d = c[l][a], b;
        if (a == "toHost") {
            var i = f(e, a);
            b = k(i, d)
        } else if (a == j) {
            var g = h(e, d);
            b = f(g, a)
        }
        return b
    }
    if (!c)
        c = {};
    a.addMapping = function(l, h) {
        var a, d;
        if (h.map) {
            a = h.map;
            d = {};
            for (var j in a) {
                var k = a[j];
                if (k == g)
                    k = j;
                d[k] = j
            }
        } else {
            a = h.toHost;
            d = h.fromHost
        }
        var b = c[l];
        if (b) {
            var f = b[i];
            for (var n in f)
                a[n] = f[n];
            f = b[e];
            for (var m in f)
                d[m] = f[m]
        } else
            b = c[l] = {};
        b[i] = a;
        b[e] = d
    }
    ;
    a.toHost = function(b, a) {
        return m(b, a, i)
    }
    ;
    a.fromHost = function(a, b) {
        return m(a, b, e)
    }
    ;
    a.self = g;
    a.sourceData = l;
    a.addComplexType = function(a) {
        b.addComplexType(a)
    }
    ;
    a.getDynamicType = function(a) {
        return b.getDynamicType(a)
    }
    ;
    a.setDynamicType = function(c, a) {
        b.setDynamicType(c, a)
    }
    ;
    a.dynamicTypes = d;
    a.doMapValues = function(a, b) {
        return f(a, b)
    }
}
;
OSF.DDA.SpecialProcessor = function(c, b) {
    var a = this;
    a.addComplexType = function(a) {
        c.push(a)
    }
    ;
    a.getDynamicType = function(a) {
        return b[a]
    }
    ;
    a.setDynamicType = function(c, a) {
        b[c] = a
    }
    ;
    a.isComplexType = function(a) {
        return OSF.OUtil.listContainsValue(c, a)
    }
    ;
    a.isDynamicType = function(a) {
        return OSF.OUtil.listContainsKey(b, a)
    }
    ;
    a.preserveNesting = function(b) {
        var a = [];
        OSF.DDA.PropertyDescriptors && a.push(OSF.DDA.PropertyDescriptors.Subset);
        if (OSF.DDA.DataNodeEventProperties)
            a = a.concat([OSF.DDA.DataNodeEventProperties.OldNode, OSF.DDA.DataNodeEventProperties.NewNode, OSF.DDA.DataNodeEventProperties.NextSiblingNode]);
        return OSF.OUtil.listContainsValue(a, b)
    }
    ;
    a.pack = function(c, d) {
        var a;
        if (this.isDynamicType(c))
            a = b[c].toHost(d);
        else
            a = d;
        return a
    }
    ;
    a.unpack = function(c, d) {
        var a;
        if (this.isDynamicType(c))
            a = b[c].fromHost(d);
        else
            a = d;
        return a
    }
}
;
OSF.DDA.getDecoratedParameterMap = function(d, c) {
    var a = new OSF.DDA.HostParameterMap(d)
      , f = a.self;
    function b(a) {
        var c = null;
        if (a) {
            c = {};
            for (var d = a.length, b = 0; b < d; b++)
                c[a[b].name] = a[b].value
        }
        return c
    }
    a.define = function(c) {
        var d = {}
          , e = b(c.toHost);
        if (c.invertible)
            d.map = e;
        else if (c.canonical)
            d.toHost = d.fromHost = e;
        else {
            d.toHost = e;
            d.fromHost = b(c.fromHost)
        }
        a.addMapping(c.type, d);
        c.isComplexType && a.addComplexType(c.type)
    }
    ;
    for (var e in c)
        a.define(c[e]);
    return a
}
;
OSF.OUtil.setNamespace("DispIdHost", OSF.DDA);
OSF.DDA.DispIdHost.Methods = {
    InvokeMethod: "invokeMethod",
    AddEventHandler: "addEventHandler",
    RemoveEventHandler: "removeEventHandler",
    OpenDialog: "openDialog",
    CloseDialog: "closeDialog",
    MessageParent: "messageParent",
    SendMessage: "sendMessage"
};
OSF.DDA.DispIdHost.Delegates = {
    ExecuteAsync: "executeAsync",
    RegisterEventAsync: "registerEventAsync",
    UnregisterEventAsync: "unregisterEventAsync",
    ParameterMap: "parameterMap",
    OpenDialog: "openDialog",
    CloseDialog: "closeDialog",
    MessageParent: "messageParent",
    SendMessage: "sendMessage"
};
OSF.DDA.DispIdHost.Facade = function(f, h) {
    var c = false
      , b = null
      , g = this
      , d = {}
      , e = OSF.DDA.AsyncMethodNames
      , a = OSF.DDA.MethodDispId
      , n = {
        GoToByIdAsync: a.dispidNavigateToMethod,
        GetSelectedDataAsync: a.dispidGetSelectedDataMethod,
        SetSelectedDataAsync: a.dispidSetSelectedDataMethod,
        GetDocumentCopyChunkAsync: a.dispidGetDocumentCopyChunkMethod,
        ReleaseDocumentCopyAsync: a.dispidReleaseDocumentCopyMethod,
        GetDocumentCopyAsync: a.dispidGetDocumentCopyMethod,
        AddFromSelectionAsync: a.dispidAddBindingFromSelectionMethod,
        AddFromPromptAsync: a.dispidAddBindingFromPromptMethod,
        AddFromNamedItemAsync: a.dispidAddBindingFromNamedItemMethod,
        GetAllAsync: a.dispidGetAllBindingsMethod,
        GetByIdAsync: a.dispidGetBindingMethod,
        ReleaseByIdAsync: a.dispidReleaseBindingMethod,
        GetDataAsync: a.dispidGetBindingDataMethod,
        SetDataAsync: a.dispidSetBindingDataMethod,
        AddRowsAsync: a.dispidAddRowsMethod,
        AddColumnsAsync: a.dispidAddColumnsMethod,
        DeleteAllDataValuesAsync: a.dispidClearAllRowsMethod,
        RefreshAsync: a.dispidLoadSettingsMethod,
        SaveAsync: a.dispidSaveSettingsMethod,
        GetActiveViewAsync: a.dispidGetActiveViewMethod,
        GetFilePropertiesAsync: a.dispidGetFilePropertiesMethod,
        GetOfficeThemeAsync: a.dispidGetOfficeThemeMethod,
        GetDocumentThemeAsync: a.dispidGetDocumentThemeMethod,
        ClearFormatsAsync: a.dispidClearFormatsMethod,
        SetTableOptionsAsync: a.dispidSetTableOptionsMethod,
        SetFormatsAsync: a.dispidSetFormatsMethod,
        GetAccessTokenAsync: a.dispidGetAccessTokenMethod,
        GetAuthContextAsync: a.dispidGetAuthContextMethod,
        GetNestedAppAuthContextAsync: a.dispidGetNestedAppAuthContextMethod,
        NestedAppAuthRequestAsync: a.dispidNestedAppAuthRequestMethod,
        ExecuteRichApiRequestAsync: a.dispidExecuteRichApiRequestMethod,
        AppCommandInvocationCompletedAsync: a.dispidAppCommandInvocationCompletedMethod,
        CloseContainerAsync: a.dispidCloseContainerMethod,
        OpenBrowserWindow: a.dispidOpenBrowserWindow,
        CreateDocumentAsync: a.dispidCreateDocumentMethod,
        InsertFormAsync: a.dispidInsertFormMethod,
        ExecuteFeature: a.dispidExecuteFeature,
        QueryFeature: a.dispidQueryFeature,
        AddDataPartAsync: a.dispidAddDataPartMethod,
        GetDataPartByIdAsync: a.dispidGetDataPartByIdMethod,
        GetDataPartsByNameSpaceAsync: a.dispidGetDataPartsByNamespaceMethod,
        GetPartXmlAsync: a.dispidGetDataPartXmlMethod,
        GetPartNodesAsync: a.dispidGetDataPartNodesMethod,
        DeleteDataPartAsync: a.dispidDeleteDataPartMethod,
        GetNodeValueAsync: a.dispidGetDataNodeValueMethod,
        GetNodeXmlAsync: a.dispidGetDataNodeXmlMethod,
        GetRelativeNodesAsync: a.dispidGetDataNodesMethod,
        SetNodeValueAsync: a.dispidSetDataNodeValueMethod,
        SetNodeXmlAsync: a.dispidSetDataNodeXmlMethod,
        AddDataPartNamespaceAsync: a.dispidAddDataNamespaceMethod,
        GetDataPartNamespaceAsync: a.dispidGetDataUriByPrefixMethod,
        GetDataPartPrefixAsync: a.dispidGetDataPrefixByUriMethod,
        GetNodeTextAsync: a.dispidGetDataNodeTextMethod,
        SetNodeTextAsync: a.dispidSetDataNodeTextMethod,
        GetSelectedTask: a.dispidGetSelectedTaskMethod,
        GetTask: a.dispidGetTaskMethod,
        GetWSSUrl: a.dispidGetWSSUrlMethod,
        GetTaskField: a.dispidGetTaskFieldMethod,
        GetSelectedResource: a.dispidGetSelectedResourceMethod,
        GetResourceField: a.dispidGetResourceFieldMethod,
        GetProjectField: a.dispidGetProjectFieldMethod,
        GetSelectedView: a.dispidGetSelectedViewMethod,
        GetTaskByIndex: a.dispidGetTaskByIndexMethod,
        GetResourceByIndex: a.dispidGetResourceByIndexMethod,
        SetTaskField: a.dispidSetTaskFieldMethod,
        SetResourceField: a.dispidSetResourceFieldMethod,
        GetMaxTaskIndex: a.dispidGetMaxTaskIndexMethod,
        GetMaxResourceIndex: a.dispidGetMaxResourceIndexMethod,
        CreateTask: a.dispidCreateTaskMethod
    };
    for (var i in n)
        if (e[i])
            d[e[i].id] = n[i];
    e = OSF.DDA.SyncMethodNames;
    a = OSF.DDA.MethodDispId;
    var m = {
        MessageParent: a.dispidMessageParentMethod,
        SendMessage: a.dispidSendMessageMethod
    };
    for (var i in m)
        if (e[i])
            d[e[i].id] = m[i];
    e = Microsoft.Office.WebExtension.EventType;
    a = OSF.DDA.EventDispId;
    var o = {
        SettingsChanged: a.dispidSettingsChangedEvent,
        DocumentSelectionChanged: a.dispidDocumentSelectionChangedEvent,
        BindingSelectionChanged: a.dispidBindingSelectionChangedEvent,
        BindingDataChanged: a.dispidBindingDataChangedEvent,
        ActiveViewChanged: a.dispidActiveViewChangedEvent,
        OfficeThemeChanged: a.dispidOfficeThemeChangedEvent,
        DocumentThemeChanged: a.dispidDocumentThemeChangedEvent,
        AppCommandInvoked: a.dispidAppCommandInvokedEvent,
        DialogMessageReceived: a.dispidDialogMessageReceivedEvent,
        DialogParentMessageReceived: a.dispidDialogParentMessageReceivedEvent,
        ObjectDeleted: a.dispidObjectDeletedEvent,
        ObjectSelectionChanged: a.dispidObjectSelectionChangedEvent,
        ObjectDataChanged: a.dispidObjectDataChangedEvent,
        ContentControlAdded: a.dispidContentControlAddedEvent,
        Suspend: a.dispidSuspend,
        Resume: a.dispidResume,
        RichApiMessage: a.dispidRichApiMessageEvent,
        ItemChanged: a.dispidOlkItemSelectedChangedEvent,
        RecipientsChanged: a.dispidOlkRecipientsChangedEvent,
        AppointmentTimeChanged: a.dispidOlkAppointmentTimeChangedEvent,
        RecurrenceChanged: a.dispidOlkRecurrenceChangedEvent,
        AttachmentsChanged: a.dispidOlkAttachmentsChangedEvent,
        EnhancedLocationsChanged: a.dispidOlkEnhancedLocationsChangedEvent,
        InfobarClicked: a.dispidOlkInfobarClickedEvent,
        SelectedItemsChanged: a.dispidOlkSelectedItemsChangedEvent,
        SensitivityLabelChanged: a.dispidOlkSensitivityLabelChangedEvent,
        InitializationContextChanged: a.dispidOlkInitializationContextChangedEvent,
        DragAndDropEvent: a.dispidOlkDragAndDropEvent,
        TaskSelectionChanged: a.dispidTaskSelectionChangedEvent,
        ResourceSelectionChanged: a.dispidResourceSelectionChangedEvent,
        ViewSelectionChanged: a.dispidViewSelectionChangedEvent,
        DataNodeInserted: a.dispidDataNodeAddedEvent,
        DataNodeReplaced: a.dispidDataNodeReplacedEvent,
        DataNodeDeleted: a.dispidDataNodeDeletedEvent
    };
    for (var k in o)
        if (e[k])
            d[e[k]] = o[k];
    function l(a) {
        return a == OSF.DDA.EventDispId.dispidObjectDeletedEvent || a == OSF.DDA.EventDispId.dispidObjectSelectionChangedEvent || a == OSF.DDA.EventDispId.dispidObjectDataChangedEvent || a == OSF.DDA.EventDispId.dispidContentControlAddedEvent
    }
    function j(a, c, d, b) {
        if (typeof a == "number") {
            if (!b)
                b = c.getCallArgs(d);
            OSF.DDA.issueAsyncResult(b, a, OSF.DDA.ErrorCodeManager.getErrorArgs(a))
        } else
            throw a
    }
    g[OSF.DDA.DispIdHost.Methods.InvokeMethod] = function(t, m, n, q) {
        var a;
        try {
            var i = t.id
              , l = OSF.DDA.AsyncMethodCalls[i];
            a = l.verifyAndExtractCall(m, n, q);
            var k = d[i]
              , s = f(i)
              , c = b;
            if (window.Excel && window.Office.context.requirements.isSetSupported("RedirectV1Api"))
                window.Excel._RedirectV1APIs = true;
            if (window.Excel && window.Excel._RedirectV1APIs && (c = window.Excel._V1APIMap[i])) {
                var e = OSF.OUtil.shallowCopy(a);
                delete e[Microsoft.Office.WebExtension.Parameters.AsyncContext];
                if (c.preprocess)
                    e = c.preprocess(e);
                var o = new window.Excel.RequestContext
                  , u = c.call(o, e);
                o.sync().then(function() {
                    var b = u.value
                      , d = b.status;
                    delete b["status"];
                    delete b["@odata.type"];
                    if (c.postprocess)
                        b = c.postprocess(b, e);
                    if (d != 0)
                        b = OSF.DDA.ErrorCodeManager.getErrorArgs(d);
                    OSF.DDA.issueAsyncResult(a, d, b)
                })["catch"](function() {
                    OSF.DDA.issueAsyncResult(a, OSF.DDA.ErrorCodeManager.errorCodes.ooeFailure, b)
                })
            } else {
                var g;
                if (h.toHost)
                    g = h.toHost(k, a);
                else
                    g = a;
                var r = (new Date).getTime();
                s[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]({
                    dispId: k,
                    hostCallArgs: g,
                    onCalling: function() {},
                    onReceiving: function() {},
                    onComplete: function(c, d) {
                        var b;
                        if (c == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
                            if (h.fromHost)
                                b = h.fromHost(k, d);
                            else
                                b = d;
                        else
                            b = d;
                        var e = l.processResponse(c, b, n, a);
                        OSF.DDA.issueAsyncResult(a, c, e);
                        OSF.AppTelemetry && !(OSF.ConstantNames && OSF.ConstantNames.IsCustomFunctionsRuntime) && OSF.AppTelemetry.onMethodDone(k, g, Math.abs((new Date).getTime() - r), c)
                    }
                })
            }
        } catch (p) {
            j(p, l, m, a)
        }
    }
    ;
    g[OSF.DDA.DispIdHost.Methods.AddEventHandler] = function(q, e, p, t) {
        var g, a, o, i = c;
        function n(b) {
            if (b == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                var d = !i ? e.addEventHandler(a, o) : e.addObjectEventHandler(a, g[Microsoft.Office.WebExtension.Parameters.Id], o);
                if (!d)
                    b = OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerAdditionFailed
            }
            var c;
            if (b != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
                c = OSF.DDA.ErrorCodeManager.getErrorArgs(b);
            OSF.DDA.issueAsyncResult(g, b, c)
        }
        try {
            var r = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.AddHandlerAsync.id];
            g = r.verifyAndExtractCall(q, p, e);
            a = g[Microsoft.Office.WebExtension.Parameters.EventType];
            o = g[Microsoft.Office.WebExtension.Parameters.Handler];
            if (t) {
                n(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
                return
            }
            var m = d[a];
            i = l(m);
            var k = i ? g[Microsoft.Office.WebExtension.Parameters.Id] : p.id || ""
              , v = i ? e.getObjectEventHandlerCount(a, k) : e.getEventHandlerCount(a);
            if (v == 0) {
                var u = f(a)[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];
                u({
                    eventType: a,
                    dispId: m,
                    targetId: k,
                    onCalling: function() {
                        OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)
                    },
                    onReceiving: function() {
                        OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)
                    },
                    onComplete: n,
                    onEvent: function(c) {
                        var d = b
                          , j = OSF._OfficeAppFactory.getHostInfo();
                        if (j && j.hostPlatform.toLowerCase() == "web" && m == OSF.DDA.EventDispId.dispidOfficeThemeChangedEvent) {
                            d = c;
                            if (c.KeepHexColors)
                                try {
                                    var g = JSON.parse(c.OfficeThemeData[0]);
                                    for (var f in OSF._OfficeAppFactory.getContext().officeTheme)
                                        if (g[f])
                                            OSF._OfficeAppFactory.getContext().officeTheme[f] = g[f]
                                } catch (l) {}
                        } else
                            d = h.fromHost(m, c);
                        if (!i)
                            e.fireEvent(OSF.DDA.OMFactory.manufactureEventArgs(a, p, d));
                        else
                            e.fireObjectEvent(k, OSF.DDA.OMFactory.manufactureEventArgs(a, k, d))
                    }
                })
            } else
                n(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
        } catch (s) {
            j(s, r, q, g)
        }
    }
    ;
    g[OSF.DDA.DispIdHost.Methods.RemoveEventHandler] = function(p, e, r) {
        var g, a, m, h = c;
        function o(a) {
            var b;
            if (a != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
                b = OSF.DDA.ErrorCodeManager.getErrorArgs(a);
            OSF.DDA.issueAsyncResult(g, a, b)
        }
        try {
            var q = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.id];
            g = q.verifyAndExtractCall(p, r, e);
            a = g[Microsoft.Office.WebExtension.Parameters.EventType];
            m = g[Microsoft.Office.WebExtension.Parameters.Handler];
            var s = d[a];
            h = l(s);
            var k = h ? g[Microsoft.Office.WebExtension.Parameters.Id] : r.id || "", n, i;
            if (m === b) {
                i = h ? e.clearObjectEventHandlers(a, k) : e.clearEventHandlers(a);
                n = OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess
            } else {
                i = h ? e.removeObjectEventHandler(a, k, m) : e.removeEventHandler(a, m);
                n = i ? OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess : OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist
            }
            var v = h ? e.getObjectEventHandlerCount(a, k) : e.getEventHandlerCount(a);
            if (i && v == 0) {
                var u = f(a)[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];
                u({
                    eventType: a,
                    dispId: s,
                    targetId: k,
                    onCalling: function() {
                        OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)
                    },
                    onReceiving: function() {
                        OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)
                    },
                    onComplete: o
                })
            } else
                o(n)
        } catch (t) {
            j(t, q, p, g)
        }
    }
    ;
    g[OSF.DDA.DispIdHost.Methods.OpenDialog] = function(p, a, o, q) {
        var g, n, k = b, e = Microsoft.Office.WebExtension.EventType.DialogMessageReceived, i = Microsoft.Office.WebExtension.EventType.DialogEventReceived;
        function l(b) {
            var d;
            if (b != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
                d = OSF.DDA.ErrorCodeManager.getErrorArgs(b);
            else {
                var c = {};
                c[Microsoft.Office.WebExtension.Parameters.Id] = n;
                c[Microsoft.Office.WebExtension.Parameters.Data] = a;
                var d = k.processResponse(b, c, o, g);
                OSF.DialogShownStatus.hasDialogShown = true;
                a.clearEventHandlers(e);
                a.clearEventHandlers(i)
            }
            OSF.DDA.issueAsyncResult(g, b, d)
        }
        try {
            (e == undefined || i == undefined) && l(OSF.DDA.ErrorCodeManager.ooeOperationNotSupported);
            if (!q) {
                if (OSF.DDA.AsyncMethodNames.DisplayDialogAsync == b) {
                    l(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
                    return
                }
                k = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.DisplayDialogAsync.id]
            } else {
                if (OSF.DDA.AsyncMethodNames.DisplayModalDialogAsync == b) {
                    l(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
                    return
                }
                k = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.DisplayModalDialogAsync.id]
            }
            g = k.verifyAndExtractCall(p, o, a);
            var r = d[e]
              , m = f(e)
              , t = m[OSF.DDA.DispIdHost.Delegates.OpenDialog] != undefined ? m[OSF.DDA.DispIdHost.Delegates.OpenDialog] : m[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];
            g["isModal"] = q;
            n = JSON.stringify(g);
            if (!OSF.DialogShownStatus.hasDialogShown) {
                a.clearQueuedEvent(e);
                a.clearQueuedEvent(i);
                a.clearQueuedEvent(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived)
            }
            t({
                eventType: e,
                dispId: r,
                targetId: n,
                onCalling: function() {
                    OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)
                },
                onReceiving: function() {
                    OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)
                },
                onComplete: l,
                onEvent: function(j) {
                    var g = h.fromHost(r, j)
                      , f = OSF.DDA.OMFactory.manufactureEventArgs(e, o, g);
                    if (f.type == i) {
                        var d = OSF.DDA.ErrorCodeManager.getErrorArgs(f.error)
                          , b = {};
                        b[OSF.DDA.AsyncResultEnum.ErrorProperties.Code] = status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                        b[OSF.DDA.AsyncResultEnum.ErrorProperties.Name] = d.name || d;
                        b[OSF.DDA.AsyncResultEnum.ErrorProperties.Message] = d.message || d;
                        f.error = new OSF.DDA.Error(b[OSF.DDA.AsyncResultEnum.ErrorProperties.Name],b[OSF.DDA.AsyncResultEnum.ErrorProperties.Message],b[OSF.DDA.AsyncResultEnum.ErrorProperties.Code])
                    }
                    a.fireOrQueueEvent(f);
                    if (g[OSF.DDA.PropertyDescriptors.MessageType] == OSF.DialogMessageType.DialogClosed) {
                        a.clearEventHandlers(e);
                        a.clearEventHandlers(i);
                        a.clearEventHandlers(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived);
                        OSF.DialogShownStatus.hasDialogShown = c
                    }
                }
            })
        } catch (s) {
            j(s, k, p, g)
        }
    }
    ;
    g[OSF.DDA.DispIdHost.Methods.CloseDialog] = function(h, o, e, q) {
        var l, a, i, g = OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
        function n(a) {
            g = a;
            OSF.DialogShownStatus.hasDialogShown = c
        }
        try {
            var k = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.CloseAsync.id];
            l = k.verifyAndExtractCall(h, q, e);
            a = Microsoft.Office.WebExtension.EventType.DialogMessageReceived;
            i = Microsoft.Office.WebExtension.EventType.DialogEventReceived;
            e.clearEventHandlers(a);
            e.clearEventHandlers(i);
            var r = d[a]
              , b = f(a)
              , p = b[OSF.DDA.DispIdHost.Delegates.CloseDialog] != undefined ? b[OSF.DDA.DispIdHost.Delegates.CloseDialog] : b[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];
            p({
                eventType: a,
                dispId: r,
                targetId: o,
                onCalling: function() {
                    OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)
                },
                onReceiving: function() {
                    OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)
                },
                onComplete: n
            })
        } catch (m) {
            j(m, k, h, l)
        }
        if (g != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
            throw OSF.OUtil.formatString(Strings.OfficeOM.L_FunctionCallFailed, OSF.DDA.AsyncMethodNames.CloseAsync.displayName, g)
    }
    ;
    g[OSF.DDA.DispIdHost.Methods.MessageParent] = function(a, i) {
        var c = {}
          , b = OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.MessageParent.id]
          , e = b.verifyAndExtractCall(a, i, c)
          , g = f(OSF.DDA.SyncMethodNames.MessageParent.id)
          , h = g[OSF.DDA.DispIdHost.Delegates.MessageParent]
          , j = d[OSF.DDA.SyncMethodNames.MessageParent.id];
        return h({
            dispId: j,
            hostCallArgs: e,
            onCalling: function() {
                OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)
            },
            onReceiving: function() {
                OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)
            }
        })
    }
    ;
    g[OSF.DDA.DispIdHost.Methods.SendMessage] = function(a, k, i) {
        var c = {}
          , b = OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.SendMessage.id]
          , e = b.verifyAndExtractCall(a, i, c)
          , g = f(OSF.DDA.SyncMethodNames.SendMessage.id)
          , h = g[OSF.DDA.DispIdHost.Delegates.SendMessage]
          , j = d[OSF.DDA.SyncMethodNames.SendMessage.id];
        return h({
            dispId: j,
            hostCallArgs: e,
            onCalling: function() {
                OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall)
            },
            onReceiving: function() {
                OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse)
            }
        })
    }
}
;
OSF.DDA.DispIdHost.addAsyncMethods = function(a, b, e) {
    for (var f in b) {
        var c = b[f]
          , d = c.displayName;
        !a[d] && OSF.OUtil.defineEnumerableProperty(a, d, {
            value: function(b) {
                return function() {
                    var c = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.InvokeMethod];
                    c(b, arguments, a, e)
                }
            }(c)
        })
    }
}
;
OSF.DDA.DispIdHost.addEventSupport = function(a, b, e) {
    var d = OSF.DDA.AsyncMethodNames.AddHandlerAsync.displayName
      , c = OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.displayName;
    !a[d] && OSF.OUtil.defineEnumerableProperty(a, d, {
        value: function() {
            var c = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.AddEventHandler];
            c(arguments, b, a, e)
        }
    });
    !a[c] && OSF.OUtil.defineEnumerableProperty(a, c, {
        value: function() {
            var c = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.RemoveEventHandler];
            c(arguments, b, a)
        }
    })
}
;
(function(e) {
    var c = "\n"
      , d = true
      , a = null
      , b = "undefined"
      , g = typeof osfLoadMsAjax === b || osfLoadMsAjax !== false;
    e.MsAjaxTypeHelper = g ? undefined : function() {
        function c() {}
        c.isInstanceOfType = function(f, e) {
            if (typeof e === b || e === a)
                return false;
            if (e instanceof f)
                return d;
            var c = e.constructor;
            if (!c || typeof c !== "function" || !c.__typeName || c.__typeName === "Object")
                c = Object;
            return !!(c === f) || c.__typeName && f.__typeName && c.__typeName === f.__typeName
        }
        ;
        return c
    }();
    e.MsAjaxError = g ? undefined : function() {
        var f = "Parameter name: {0}";
        function d() {}
        d.create = function(c, b) {
            var a = new Error(c);
            a.message = c;
            if (b)
                for (var d in b)
                    a[d] = b[d];
            a.popStackFrame();
            return a
        }
        ;
        d.parameterCount = function(a) {
            var c = "Sys.ParameterCountException: " + (a ? a : "Parameter count mismatch.")
              , b = d.create(c, {
                name: "Sys.ParameterCountException"
            });
            b.popStackFrame();
            return b
        }
        ;
        d.argument = function(a, g) {
            var b = "Sys.ArgumentException: " + (g ? g : "Value does not fall within the expected range.");
            if (a)
                b += c + e.MsAjaxString.format(f, a);
            var h = d.create(b, {
                name: "Sys.ArgumentException",
                paramName: a
            });
            h.popStackFrame();
            return h
        }
        ;
        d.argumentNull = function(a, g) {
            var b = "Sys.ArgumentNullException: " + (g ? g : "Value cannot be null.");
            if (a)
                b += c + e.MsAjaxString.format(f, a);
            var h = d.create(b, {
                name: "Sys.ArgumentNullException",
                paramName: a
            });
            h.popStackFrame();
            return h
        }
        ;
        d.argumentOutOfRange = function(i, g, j) {
            var h = "Sys.ArgumentOutOfRangeException: " + (j ? j : "Specified argument was out of the range of valid values.");
            if (i)
                h += c + e.MsAjaxString.format(f, i);
            if (typeof g !== b && g !== a)
                h += c + e.MsAjaxString.format("Actual value was {0}.", g);
            var k = d.create(h, {
                name: "Sys.ArgumentOutOfRangeException",
                paramName: i,
                actualValue: g
            });
            k.popStackFrame();
            return k
        }
        ;
        d.argumentType = function(h, g, b, i) {
            var a = "Sys.ArgumentTypeException: ";
            if (i)
                a += i;
            else if (g && b)
                a += e.MsAjaxString.format("Object of type '{0}' cannot be converted to type '{1}'.", g.getName ? g.getName() : g, b.getName ? b.getName() : b);
            else
                a += "Object cannot be converted to the required type.";
            if (h)
                a += c + e.MsAjaxString.format(f, h);
            var j = d.create(a, {
                name: "Sys.ArgumentTypeException",
                paramName: h,
                actualType: g,
                expectedType: b
            });
            j.popStackFrame();
            return j
        }
        ;
        d.argumentUndefined = function(a, g) {
            var b = "Sys.ArgumentUndefinedException: " + (g ? g : "Value cannot be undefined.");
            if (a)
                b += c + e.MsAjaxString.format(f, a);
            var h = d.create(b, {
                name: "Sys.ArgumentUndefinedException",
                paramName: a
            });
            h.popStackFrame();
            return h
        }
        ;
        d.invalidOperation = function(a) {
            var c = "Sys.InvalidOperationException: " + (a ? a : "Operation is not valid due to the current state of the object.")
              , b = d.create(c, {
                name: "Sys.InvalidOperationException"
            });
            b.popStackFrame();
            return b
        }
        ;
        return d
    }();
    e.MsAjaxString = g ? undefined : function() {
        function a() {}
        a.format = function(c) {
            for (var b = [], a = 1; a < arguments.length; a++)
                b[a - 1] = arguments[a];
            var d = c;
            return d.replace(/{(\d+)}/gm, function(d, a) {
                var c = parseInt(a, 10);
                return b[c] === undefined ? "{" + a + "}" : b[c]
            })
        }
        ;
        a.startsWith = function(b, a) {
            return b.substr(0, a.length) === a
        }
        ;
        return a
    }();
    e.MsAjaxDebug = g ? undefined : function() {
        function a() {}
        a.trace = function() {}
        ;
        return a
    }();
    if (!g && !OsfMsAjaxFactory.isMsAjaxLoaded()) {
        var f = function(b, d, c) {
            if (b.__typeName === undefined || b.__typeName === a)
                b.__typeName = d;
            if (b.__class === undefined || b.__class === a)
                b.__class = c
        };
        f(Function, "Function", d);
        f(Error, "Error", d);
        f(Object, "Object", d);
        f(String, "String", d);
        f(Boolean, "Boolean", d);
        f(Date, "Date", d);
        f(Number, "Number", d);
        f(RegExp, "RegExp", d);
        f(Array, "Array", d);
        if (!Function.createCallback)
            Function.createCallback = function(b, a) {
                var c = Function._validateParams(arguments, [{
                    name: "method",
                    type: Function
                }, {
                    name: "context",
                    mayBeNull: d
                }]);
                if (c)
                    throw c;
                return function() {
                    var e = arguments.length;
                    if (e > 0) {
                        for (var d = [], c = 0; c < e; c++)
                            d[c] = arguments[c];
                        d[e] = a;
                        return b.apply(this, d)
                    }
                    return b.call(this, a)
                }
            }
            ;
        if (!Function.createDelegate)
            Function.createDelegate = function(b, c) {
                var a = Function._validateParams(arguments, [{
                    name: "instance",
                    mayBeNull: d
                }, {
                    name: "method",
                    type: Function
                }]);
                if (a)
                    throw a;
                return function() {
                    return c.apply(b, arguments)
                }
            }
            ;
        if (!Function._validateParams)
            Function._validateParams = function(i, g, e) {
                var c, f = g.length;
                e = e || typeof e === b;
                c = Function._validateParameterCount(i, g, e);
                if (c) {
                    c.popStackFrame();
                    return c
                }
                for (var d = 0, k = i.length; d < k; d++) {
                    var h = g[Math.min(d, f - 1)]
                      , j = h.name;
                    if (h.parameterArray)
                        j += "[" + (d - f + 1) + "]";
                    else if (!e && d >= f)
                        break;
                    c = Function._validateParameter(i[d], h, j);
                    if (c) {
                        c.popStackFrame();
                        return c
                    }
                }
                return a
            }
            ;
        if (!Function._validateParameterCount)
            Function._validateParameterCount = function(m, g, l) {
                var b, f, c = g.length, h = m.length;
                if (h < c) {
                    var i = c;
                    for (b = 0; b < c; b++) {
                        var j = g[b];
                        if (j.optional || j.parameterArray)
                            i--
                    }
                    if (h < i)
                        f = d
                } else if (l && h > c) {
                    f = d;
                    for (b = 0; b < c; b++)
                        if (g[b].parameterArray) {
                            f = false;
                            break
                        }
                }
                if (f) {
                    var k = e.MsAjaxError.parameterCount();
                    k.popStackFrame();
                    return k
                }
                return a
            }
            ;
        if (!Function._validateParameter)
            Function._validateParameter = function(e, c, j) {
                var d, i = c.type, n = !!c.integer, m = !!c.domElement, o = !!c.mayBeNull;
                d = Function._validateParameterType(e, i, n, m, o, j);
                if (d) {
                    d.popStackFrame();
                    return d
                }
                var g = c.elementType
                  , h = !!c.elementMayBeNull;
                if (i === Array && typeof e !== b && e !== a && (g || !h))
                    for (var l = !!c.elementInteger, k = !!c.elementDomElement, f = 0; f < e.length; f++) {
                        var p = e[f];
                        d = Function._validateParameterType(p, g, l, k, h, j + "[" + f + "]");
                        if (d) {
                            d.popStackFrame();
                            return d
                        }
                    }
                return a
            }
            ;
        if (!Function._validateParameterType)
            Function._validateParameterType = function(d, f, j, i, h, g) {
                var c, k;
                if (typeof d === b)
                    if (h)
                        return a;
                    else {
                        c = e.MsAjaxError.argumentUndefined(g);
                        c.popStackFrame();
                        return c
                    }
                if (d === a)
                    if (h)
                        return a;
                    else {
                        c = e.MsAjaxError.argumentNull(g);
                        c.popStackFrame();
                        return c
                    }
                if (f && !e.MsAjaxTypeHelper.isInstanceOfType(f, d)) {
                    c = e.MsAjaxError.argumentType(g, typeof d, f);
                    c.popStackFrame();
                    return c
                }
                return a
            }
            ;
        if (!window.Type)
            window.Type = Function;
        if (!Type.registerNamespace)
            Type.registerNamespace = function(d) {
                for (var c = d.split("."), b = window, a = 0; a < c.length; a++) {
                    b[c[a]] = b[c[a]] || {};
                    b = b[c[a]]
                }
            }
            ;
        if (!Type.prototype.registerClass)
            Type.prototype.registerClass = function(a) {
                a = {}
            }
            ;
        typeof Sys === b && Type.registerNamespace("Sys");
        if (!Error.prototype.popStackFrame)
            Error.prototype.popStackFrame = function() {
                var d = this;
                if (arguments.length !== 0)
                    throw e.MsAjaxError.parameterCount();
                if (typeof d.stack === b || d.stack === a || typeof d.fileName === b || d.fileName === a || typeof d.lineNumber === b || d.lineNumber === a)
                    return;
                var f = d.stack.split(c)
                  , h = f[0]
                  , j = d.fileName + ":" + d.lineNumber;
                while (typeof h !== b && h !== a && h.indexOf(j) === -1) {
                    f.shift();
                    h = f[0]
                }
                var i = f[1];
                if (typeof i === b || i === a)
                    return;
                var g = i.match(/@(.*):(\d+)$/);
                if (typeof g === b || g === a)
                    return;
                d.fileName = g[1];
                d.lineNumber = parseInt(g[2]);
                f.shift();
                d.stack = f.join(c)
            }
            ;
        OsfMsAjaxFactory.msAjaxError = e.MsAjaxError;
        OsfMsAjaxFactory.msAjaxString = e.MsAjaxString;
        OsfMsAjaxFactory.msAjaxDebug = e.MsAjaxDebug
    }
}
)(OfficeExt || (OfficeExt = {}));
OSF.OUtil.setNamespace("SafeArray", OSF.DDA);
OSF.DDA.SafeArray.Response = {
    Status: 0,
    Payload: 1
};
OSF.DDA.SafeArray.UniqueArguments = {
    Offset: "offset",
    Run: "run",
    BindingSpecificData: "bindingSpecificData",
    MergedCellGuid: "{66e7831f-81b2-42e2-823c-89e872d541b3}"
};
OSF.OUtil.setNamespace("Delegate", OSF.DDA.SafeArray);
OSF.DDA.SafeArray.Delegate._onException = function(d, b) {
    var a, c = d.number;
    if (c)
        switch (c) {
        case -2146828218:
            a = OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
            break;
        case -2147467259:
            if (b.dispId == OSF.DDA.EventDispId.dispidDialogMessageReceivedEvent)
                a = OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened;
            else
                a = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
            break;
        case -2146828283:
            a = OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidParam;
            break;
        case -2147209089:
            a = OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidParam;
            break;
        case -2147208704:
            a = OSF.DDA.ErrorCodeManager.errorCodes.ooeTooManyIncompleteRequests;
            break;
        case -2146827850:
        default:
            a = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError
        }
    b.onComplete && b.onComplete(a || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError)
}
;
OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod = function(c) {
    var a, b = c.number;
    if (b)
        switch (b) {
        case -2146828218:
            a = OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
            break;
        case -2146827850:
        default:
            a = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError
        }
    return a || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError
}
;
OSF.DDA.SafeArray.Delegate.SpecialProcessor = function() {
    function a(a) {
        var b;
        try {
            var h = a.ubound(1)
              , d = a.ubound(2);
            a = a.toArray();
            if (h == 1 && d == 1)
                b = [a];
            else {
                b = [];
                for (var f = 0; f < h; f++) {
                    for (var c = [], e = 0; e < d; e++) {
                        var g = a[f * d + e];
                        g != OSF.DDA.SafeArray.UniqueArguments.MergedCellGuid && c.push(g)
                    }
                    c.length > 0 && b.push(c)
                }
            }
        } catch (i) {}
        return b
    }
    var c = []
      , b = {};
    b[Microsoft.Office.WebExtension.Parameters.Data] = function() {
        var c = 0
          , b = 1;
        return {
            toHost: function(a) {
                if (OSF.DDA.TableDataProperties && typeof a != "string" && a[OSF.DDA.TableDataProperties.TableRows] !== undefined) {
                    var d = [];
                    d[c] = a[OSF.DDA.TableDataProperties.TableRows];
                    d[b] = a[OSF.DDA.TableDataProperties.TableHeaders];
                    a = d
                }
                return a
            },
            fromHost: function(f) {
                var e;
                if (f.toArray) {
                    var g = f.dimensions();
                    if (g === 2)
                        e = a(f);
                    else {
                        var d = f.toArray();
                        if (d.length === 2 && (d[0] != null && d[0].toArray || d[1] != null && d[1].toArray)) {
                            e = {};
                            e[OSF.DDA.TableDataProperties.TableRows] = a(d[c]);
                            e[OSF.DDA.TableDataProperties.TableHeaders] = a(d[b])
                        } else
                            e = d
                    }
                } else
                    e = f;
                return e
            }
        }
    }();
    OSF.DDA.SafeArray.Delegate.SpecialProcessor.uber.constructor.call(this, c, b);
    this.unpack = function(c, a) {
        var d;
        if (this.isComplexType(c) || OSF.DDA.ListType.isListType(c)) {
            var e = a !== undefined && a.toArray !== undefined;
            d = e ? a.toArray() : a || {}
        } else if (this.isDynamicType(c))
            d = b[c].fromHost(a);
        else
            d = a;
        return d
    }
}
;
OSF.OUtil.extend(OSF.DDA.SafeArray.Delegate.SpecialProcessor, OSF.DDA.SpecialProcessor);
OSF.DDA.SafeArray.Delegate.ParameterMap = OSF.DDA.getDecoratedParameterMap(new OSF.DDA.SafeArray.Delegate.SpecialProcessor, [{
    type: Microsoft.Office.WebExtension.Parameters.ValueFormat,
    toHost: [{
        name: Microsoft.Office.WebExtension.ValueFormat.Unformatted,
        value: 0
    }, {
        name: Microsoft.Office.WebExtension.ValueFormat.Formatted,
        value: 1
    }]
}, {
    type: Microsoft.Office.WebExtension.Parameters.FilterType,
    toHost: [{
        name: Microsoft.Office.WebExtension.FilterType.All,
        value: 0
    }]
}]);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.PropertyDescriptors.AsyncResultStatus,
    fromHost: [{
        name: Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded,
        value: 0
    }, {
        name: Microsoft.Office.WebExtension.AsyncResultStatus.Failed,
        value: 1
    }]
});
OSF.DDA.SafeArray.Delegate.executeAsync = function(a) {
    function c(a) {
        var b = a;
        if (OSF.OUtil.isArray(a))
            for (var f = b.length, d = 0; d < f; d++)
                b[d] = c(b[d]);
        else if (OSF.OUtil.isDate(a))
            b = a.getVarDate();
        else if (typeof a === "object" && !OSF.OUtil.isArray(a)) {
            b = [];
            for (var e in a)
                if (!OSF.OUtil.isFunction(a[e]))
                    b[e] = c(a[e])
        }
        return b
    }
    function b(a) {
        var e = a;
        if (a != null && a.toArray) {
            var d = a.toArray();
            e = new Array(d.length);
            for (var c = 0; c < d.length; c++)
                e[c] = b(d[c])
        }
        return e
    }
    try {
        a.onCalling && a.onCalling();
        OSF.ClientHostController.execute(a.dispId, c(a.hostCallArgs), function(g) {
            var d, e;
            if (typeof g === "number") {
                d = [];
                e = g
            } else {
                d = g.toArray();
                e = d[OSF.DDA.SafeArray.Response.Status]
            }
            if (e == OSF.DDA.ErrorCodeManager.errorCodes.ooeChunkResult) {
                var c = d[OSF.DDA.SafeArray.Response.Payload];
                c = b(c);
                if (c != null) {
                    if (!a._chunkResultData)
                        a._chunkResultData = [];
                    a._chunkResultData[c[0]] = c[1]
                }
                return false
            }
            a.onReceiving && a.onReceiving();
            if (a.onComplete) {
                var c;
                if (e == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                    if (d.length > 2) {
                        c = [];
                        for (var f = 1; f < d.length; f++)
                            c[f - 1] = d[f]
                    } else
                        c = d[OSF.DDA.SafeArray.Response.Payload];
                    if (a._chunkResultData) {
                        c = b(c);
                        if (c != null) {
                            var h = c[c.length - 1];
                            if (a._chunkResultData.length == h)
                                c[c.length - 1] = a._chunkResultData;
                            else
                                e = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError
                        }
                    }
                } else
                    c = d[OSF.DDA.SafeArray.Response.Payload];
                a.onComplete(e, c)
            }
            return true
        })
    } catch (d) {
        OSF.DDA.SafeArray.Delegate._onException(d, a)
    }
}
;
OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent = function(c, a) {
    var b = (new Date).getTime();
    return function(d) {
        a.onReceiving && a.onReceiving();
        var e = d.toArray ? d.toArray()[OSF.DDA.SafeArray.Response.Status] : d;
        a.onComplete && a.onComplete(e);
        OSF.AppTelemetry && OSF.AppTelemetry.onRegisterDone(c, a.dispId, Math.abs((new Date).getTime() - b), e)
    }
}
;
OSF.DDA.SafeArray.Delegate.registerEventAsync = function(a) {
    a.onCalling && a.onCalling();
    var c = OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(true, a);
    try {
        OSF.ClientHostController.registerEvent(a.dispId, a.targetId, function(c, b) {
            a.onEvent && a.onEvent(b);
            OSF.AppTelemetry && OSF.AppTelemetry.onEventDone(a.dispId)
        }, c)
    } catch (b) {
        OSF.DDA.SafeArray.Delegate._onException(b, a)
    }
}
;
OSF.DDA.SafeArray.Delegate.unregisterEventAsync = function(a) {
    a.onCalling && a.onCalling();
    var c = OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(false, a);
    try {
        OSF.ClientHostController.unregisterEvent(a.dispId, a.targetId, c)
    } catch (b) {
        OSF.DDA.SafeArray.Delegate._onException(b, a)
    }
}
;
OSF.ClientMode = {
    ReadWrite: 0,
    ReadOnly: 1
};
OSF.DDA.RichInitializationReason = {
    1: Microsoft.Office.WebExtension.InitializationReason.Inserted,
    2: Microsoft.Office.WebExtension.InitializationReason.DocumentOpened
};
OSF.InitializationHelper = function(d, b, f, e, c) {
    var a = this;
    a._hostInfo = d;
    a._webAppState = b;
    a._context = f;
    a._settings = e;
    a._hostFacade = c;
    a._initializeSettings = a.initializeSettings
}
;
OSF.InitializationHelper.prototype.deserializeSettings = function(c, d) {
    var a, b = OSF.DDA.SettingsManager.deserializeSettings(c);
    if (d)
        a = new OSF.DDA.RefreshableSettings(b);
    else
        a = new OSF.DDA.Settings(b);
    return a
}
;
OSF.InitializationHelper.prototype.saveAndSetDialogInfo = function() {}
;
OSF.InitializationHelper.prototype.setAgaveHostCommunication = function() {}
;
OSF.InitializationHelper.prototype.prepareRightBeforeWebExtensionInitialize = function(a) {
    this.prepareApiSurface(a);
    Microsoft.Office.WebExtension.initialize(this.getInitializationReason(a))
}
;
OSF.InitializationHelper.prototype.prepareApiSurface = function(a) {
    var h = new OSF.DDA.License(a.get_eToken())
      , g = OSF.DDA.OfficeTheme && OSF.DDA.OfficeTheme.getOfficeTheme ? OSF.DDA.OfficeTheme.getOfficeTheme : null;
    if (a.get_isDialog()) {
        if (OSF.DDA.UI.ChildUI)
            a.ui = new OSF.DDA.UI.ChildUI
    } else if (OSF.DDA.UI.ParentUI) {
        a.ui = new OSF.DDA.UI.ParentUI;
        OfficeExt.Container && OSF.DDA.DispIdHost.addAsyncMethods(a.ui, [OSF.DDA.AsyncMethodNames.CloseContainerAsync])
    }
    OSF.DDA.OpenBrowser && OSF.DDA.DispIdHost.addAsyncMethods(a.ui, [OSF.DDA.AsyncMethodNames.OpenBrowserWindow]);
    OSF.DDA.ExecuteFeature && OSF.DDA.DispIdHost.addAsyncMethods(a.ui, [OSF.DDA.AsyncMethodNames.ExecuteFeature]);
    OSF.DDA.QueryFeature && OSF.DDA.DispIdHost.addAsyncMethods(a.ui, [OSF.DDA.AsyncMethodNames.QueryFeature]);
    if (OSF.DDA.Auth) {
        a.auth = new OSF.DDA.Auth;
        var b = []
          , c = OSF.DDA.AsyncMethodNames.GetAccessTokenAsync;
        c && b.push(c);
        var e = OSF.DDA.AsyncMethodNames.GetNestedAppAuthContextAsync;
        e && b.push(e);
        OSF.DDA.DispIdHost.addAsyncMethods(a.auth, b)
    }
    OSF._OfficeAppFactory.setContext(new OSF.DDA.Context(a,a.doc,h,null,g));
    var d, f;
    d = OSF.DDA.DispIdHost.getClientDelegateMethods;
    f = OSF.DDA.SafeArray.Delegate.ParameterMap;
    OSF._OfficeAppFactory.setHostFacade(new OSF.DDA.DispIdHost.Facade(d,f))
}
;
OSF.InitializationHelper.prototype.getInitializationReason = function(a) {
    return OSF.DDA.RichInitializationReason[a.get_reason()]
}
;
OSF.DDA.DispIdHost.getClientDelegateMethods = function(b) {
    var a = {};
    a[OSF.DDA.DispIdHost.Delegates.ExecuteAsync] = OSF.DDA.SafeArray.Delegate.executeAsync;
    a[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync] = OSF.DDA.SafeArray.Delegate.registerEventAsync;
    a[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync] = OSF.DDA.SafeArray.Delegate.unregisterEventAsync;
    a[OSF.DDA.DispIdHost.Delegates.OpenDialog] = OSF.DDA.SafeArray.Delegate.openDialog;
    a[OSF.DDA.DispIdHost.Delegates.CloseDialog] = OSF.DDA.SafeArray.Delegate.closeDialog;
    a[OSF.DDA.DispIdHost.Delegates.MessageParent] = OSF.DDA.SafeArray.Delegate.messageParent;
    a[OSF.DDA.DispIdHost.Delegates.SendMessage] = OSF.DDA.SafeArray.Delegate.sendMessage;
    if (OSF.DDA.AsyncMethodNames.RefreshAsync && b == OSF.DDA.AsyncMethodNames.RefreshAsync.id) {
        var d = function(c, b, a) {
            if (typeof OSF.DDA.ClientSettingsManager.refresh === "function")
                return OSF.DDA.ClientSettingsManager.refresh(b, a);
            else
                return OSF.DDA.ClientSettingsManager.read(b, a)
        };
        a[OSF.DDA.DispIdHost.Delegates.ExecuteAsync] = OSF.DDA.ClientSettingsManager.getSettingsExecuteMethod(d)
    }
    if (OSF.DDA.AsyncMethodNames.SaveAsync && b == OSF.DDA.AsyncMethodNames.SaveAsync.id) {
        var c = function(a, c, b) {
            return OSF.DDA.ClientSettingsManager.write(a[OSF.DDA.SettingsManager.SerializedSettings], a[Microsoft.Office.WebExtension.Parameters.OverwriteIfStale], c, b)
        };
        a[OSF.DDA.DispIdHost.Delegates.ExecuteAsync] = OSF.DDA.ClientSettingsManager.getSettingsExecuteMethod(c)
    }
    return a
}
;
(function(b) {
    var a = function() {
        var a = "undefined";
        function b() {}
        b.prototype.execute = function(d, c, b) {
            if (typeof OsfOMToken != a && OsfOMToken)
                window.external.Execute(d, c, b, OsfOMToken);
            else
                window.external.Execute(d, c, b)
        }
        ;
        b.prototype.registerEvent = function(e, c, d, b) {
            if (typeof OsfOMToken != a && OsfOMToken)
                window.external.RegisterEvent(e, c, d, b, OsfOMToken);
            else
                window.external.RegisterEvent(e, c, d, b)
        }
        ;
        b.prototype.unregisterEvent = function(d, c, b) {
            if (typeof OsfOMToken != a && OsfOMToken)
                window.external.UnregisterEvent(d, c, b, OsfOMToken);
            else
                window.external.UnregisterEvent(d, c, b)
        }
        ;
        b.prototype.closeSdxDialog = function(a) {
            OSF.OUtil.externalNativeFunctionExists(typeof window.external.closeSdxDialog) && window.external.closeSdxDialog(a)
        }
        ;
        b.prototype.resizeSdxDialog = function(b, a) {
            OSF.OUtil.externalNativeFunctionExists(typeof window.external.resizeSdxDialog) && window.external.resizeSdxDialog(b, a)
        }
        ;
        return b
    }();
    b.RichClientHostController = a
}
)(OfficeExt || (OfficeExt = {}));
(function(a) {
    var b = function(c) {
        var b = "undefined";
        __extends(a, c);
        function a() {
            return c !== null && c.apply(this, arguments) || this
        }
        a.prototype.messageParent = function(a) {
            if (OSF.OUtil.externalNativeFunctionExists(typeof window.external.MessageParent2)) {
                if (a) {
                    var c = a[Microsoft.Office.WebExtension.Parameters.MessageToParent];
                    if (typeof c === "boolean")
                        if (c === true)
                            a[Microsoft.Office.WebExtension.Parameters.MessageToParent] = "true";
                        else if (c === false)
                            a[Microsoft.Office.WebExtension.Parameters.MessageToParent] = ""
                }
                if (typeof OsfOMToken != b && OsfOMToken)
                    window.external.MessageParent2(JSON.stringify(a), OsfOMToken);
                else
                    window.external.MessageParent2(JSON.stringify(a))
            } else {
                var d = a[Microsoft.Office.WebExtension.Parameters.MessageToParent];
                window.external.MessageParent(d)
            }
        }
        ;
        a.prototype.openDialog = function(d, b, c, a) {
            this.registerEvent(d, b, c, a)
        }
        ;
        a.prototype.closeDialog = function(c, b, a) {
            this.unregisterEvent(c, b, a)
        }
        ;
        a.prototype.sendMessage = function(a) {
            if (OSF.OUtil.externalNativeFunctionExists(typeof window.external.MessageChild2))
                if (typeof OsfOMToken != b && OsfOMToken)
                    window.external.MessageChild2(JSON.stringify(a), OsfOMToken);
                else
                    window.external.MessageChild2(JSON.stringify(a));
            else {
                var c = a[Microsoft.Office.WebExtension.Parameters.MessageContent];
                window.external.MessageChild(c)
            }
        }
        ;
        return a
    }(a.RichClientHostController);
    a.Win32RichClientHostController = b
}
)(OfficeExt || (OfficeExt = {}));
(function(a) {
    var b;
    (function(c) {
        var b = function() {
            var a = null;
            function b() {
                this._osfOfficeTheme = a;
                this._osfOfficeThemeTimeStamp = a
            }
            b.prototype.getOfficeTheme = function() {
                var c = "GetOfficeThemeInfo"
                  , a = this;
                if (OSF.DDA._OsfControlContext) {
                    if (a._osfOfficeTheme && a._osfOfficeThemeTimeStamp && (new Date).getTime() - a._osfOfficeThemeTimeStamp < b._osfOfficeThemeCacheValidPeriod)
                        OSF.AppTelemetry && OSF.AppTelemetry.onPropertyDone(c, 0);
                    else {
                        var g = (new Date).getTime()
                          , f = OSF.DDA._OsfControlContext.GetOfficeThemeInfo()
                          , d = (new Date).getTime();
                        OSF.AppTelemetry && OSF.AppTelemetry.onPropertyDone(c, Math.abs(d - g));
                        a._osfOfficeTheme = JSON.parse(f);
                        if (OSF.DDA.Theming && typeof OSF.DDA.Theming.ConvertToOfficeTheme === "function")
                            a._osfOfficeTheme = OSF.DDA.Theming.ConvertToOfficeTheme(a._osfOfficeTheme);
                        else
                            for (var e in a._osfOfficeTheme)
                                a._osfOfficeTheme[e] = OSF.OUtil.convertIntToCssHexColor(a._osfOfficeTheme[e]);
                        a._osfOfficeThemeTimeStamp = d
                    }
                    return a._osfOfficeTheme
                }
            }
            ;
            b.instance = function() {
                if (b._instance == a)
                    b._instance = new b;
                return b._instance
            }
            ;
            b._osfOfficeThemeCacheValidPeriod = 5e3;
            b._instance = a;
            return b
        }();
        c.OfficeThemeManager = b;
        OSF.OUtil.setNamespace("OfficeTheme", OSF.DDA);
        OSF.DDA.OfficeTheme.getOfficeTheme = a.OfficeTheme.OfficeThemeManager.instance().getOfficeTheme
    }
    )(b = a.OfficeTheme || (a.OfficeTheme = {}))
}
)(OfficeExt || (OfficeExt = {}));
OSF.initializeRichCommon = function() {
    var a = "undefined";
    OSF.DDA.ClientSettingsManager = {
        getSettingsExecuteMethod: function(a) {
            return function(b) {
                var e = function(c, a) {
                    b.onReceiving && b.onReceiving();
                    b.onComplete && b.onComplete(c, a)
                }, c;
                try {
                    c = a(b.hostCallArgs, b.onCalling, e)
                } catch (d) {
                    var f = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                    c = {
                        name: Strings.OfficeOM.L_InternalError,
                        message: d
                    };
                    b.onComplete && b.onComplete(f, c)
                }
            }
        },
        read: function(g, f) {
            var c = []
              , e = [];
            g && g();
            if (typeof OsfOMToken != a && OsfOMToken)
                OSF.DDA._OsfControlContext.GetSettings(OsfOMToken).Read(c, e);
            else
                OSF.DDA._OsfControlContext.GetSettings().Read(c, e);
            for (var d = {}, b = 0; b < c.length; b++)
                d[c[b]] = e[b];
            f && f(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess, d);
            return d
        },
        write: function(f, i, g, c) {
            var e = []
              , d = [];
            for (var h in f) {
                e.push(h);
                d.push(f[h])
            }
            g && g();
            var b;
            if (typeof OsfOMToken != a && OsfOMToken)
                b = OSF.DDA._OsfControlContext.GetSettings(OsfOMToken);
            else
                b = OSF.DDA._OsfControlContext.GetSettings();
            if (typeof b.WriteAsync != a)
                b.WriteAsync(e, d, c);
            else {
                b.Write(e, d);
                c && c(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess)
            }
        },
        refresh: function(f, e) {
            var c = []
              , g = [];
            f && f();
            var b;
            if (typeof OsfOMToken != a && OsfOMToken)
                b = OSF.DDA._OsfControlContext.GetSettings(OsfOMToken);
            else
                b = OSF.DDA._OsfControlContext.GetSettings();
            var d = function() {
                b.Read(c, g);
                for (var d = {}, a = 0; a < c.length; a++)
                    d[c[a]] = g[a];
                e && e(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess, d)
            };
            if (b.RefreshAsync)
                b.RefreshAsync(function() {
                    d()
                });
            else
                d()
        }
    };
    OSF.InitializationHelper.prototype.initializeSettings = function(b) {
        var a = OSF.DDA.ClientSettingsManager.read()
          , c = this.deserializeSettings(a, b);
        return c
    }
    ;
    OSF.InitializationHelper.prototype.getAppContext = function(I, x) {
        var l = "Warning: Office.js is loaded outside of Office client", b;
        try {
            if (window.external && OSF.OUtil.externalNativeFunctionExists(typeof window.external.GetContext))
                b = OSF.DDA._OsfControlContext = window.external.GetContext();
            else {
                console.error('[ERROR] - There is no window.external.');
                OsfMsAjaxFactory.msAjaxDebug.trace(l);
                return
            }
        } catch (d) {
            OsfMsAjaxFactory.msAjaxDebug.trace(l + " Ignoring...Error: " + d);
            return
        }
        var D = b.GetAppType()
          , H = b.GetSolutionRef()
          , E = b.GetAppVersionMajor()
          , y = b.GetAppVersionMinor()
          , C = b.GetAppUILocale()
          , B = b.GetAppDataLocale()
          , F = b.GetDocUrl()
          , A = b.GetAppCapabilities()
          , G = b.GetActivationMode()
          , u = b.GetControlIntegrationLevel()
          , m = {};
        try {
            var f = []
              , g = [];
            if (typeof OsfOMToken != a && OsfOMToken)
                b.GetSettings(OsfOMToken).Read(f, g);
            else
                b.GetSettings().Read(f, g);
            for (var e = 0; e < f.length; e++)
                m[f[e]] = g[e]
        } catch (d) {
            OsfMsAjaxFactory.msAjaxDebug.trace("Ignoring...Error reading settings in getAppContext: " + d)
        }
        var n = "";
        try {
            var k = b.GetSolutionToken();
            n = k ? k.toString() : ""
        } catch (d) {
            OsfMsAjaxFactory.msAjaxDebug.trace("Ignoring...Error getting solution token: " + d)
        }
        var c = function(a) {
            if (OSF.OUtil.externalNativeFunctionExists(typeof b[a]))
                return b[a]();
            return undefined
        }
          , w = c("GetCorrelationId")
          , v = c("GetInstanceId")
          , z = c("GetTouchEnabled")
          , s = c("GetCommerceAllowed")
          , r = c("GetSupportedMatrix")
          , q = c("GetHostCustomMessage")
          , t = c("GetHostFullVersion")
          , o = c("GetDialogRequirementMatrix")
          , p = c("GetInitialDisplayMode") || 0
          , h = c("GetFeaturesForSolution")
          , j = {};
        try {
            if (h)
                j = JSON.parse(h)
        } catch (d) {
            OsfMsAjaxFactory.msAjaxDebug.trace("Exception while creating the SDX FeatureGates object. Details: " + d)
        }
        var i = new OSF.OfficeAppContext(H,D,E,C,B,F,A,m,G,u,n,w,v,z,s,y,r,q,t,undefined,undefined,undefined,undefined,o,j,undefined,p,undefined);
        OSF.AppTelemetry && OSF.AppTelemetry.initialize(i);
        x(i)
    }
}
;
OSF.OUtil.setNamespace("Messaging", OSF);
OSF.Messaging.sendMessage = function(a) {
    var b = JSON.stringify(a);
    OSF.ClientHostController.execute(OSF.DDA.MethodDispId.dispidSdxSendMessage, [b], null)
}
;
OSF.Messaging.registerOnMessageEventHandler = function(a) {
    var b = function(d, c) {
        var b = c.toArray();
        a(JSON.parse(b[0]))
    };
    OSF.ClientHostController.registerEvent(OSF.DDA.EventDispId.dispidOnSdxSendMessageEvent, "", b, null)
}
;
OSF.ClientHostController = new OfficeExt.Win32RichClientHostController;
OSF.initializeRichCommon();
(function() {
    var b = "undefined"
      , a = function() {
        var h = function(a) {
            a && OSF.OUtil.loadScript(a, function() {
                OsfMsAjaxFactory.msAjaxDebug.trace("loaded customized script:" + a)
            })
        }, e, g, a, d = null, f = OSF.OUtil.parseXdmInfo();
        if (f) {
            a = OSF.OUtil.getInfoItems(f);
            if (a && a.length >= 3) {
                e = a[0];
                g = a[2];
                d = Microsoft.Office.Common.XdmCommunicationManager.connect(e, window.parent, g)
            }
        }
        var c = null;
        if (!d) {
            try {
                if (window.external && typeof window.external.getCustomizedScriptPath !== b)
                    c = window.external.getCustomizedScriptPath()
            } catch (i) {
                OsfMsAjaxFactory.msAjaxDebug.trace("no script override through window.external.")
            }
            h(c)
        }
    }
      , c = typeof osfLoadMsAjax === b || osfLoadMsAjax !== false;
    if (c && !OsfMsAjaxFactory.isMsAjaxLoaded())
        if (!(OSF._OfficeAppFactory && OSF._OfficeAppFactory && OSF._OfficeAppFactory.getLoadScriptHelper && OSF._OfficeAppFactory.getLoadScriptHelper().isScriptLoading(OSF.ConstantNames.MicrosoftAjaxId)))
            OsfMsAjaxFactory.loadMsAjaxFull(function() {
                if (OsfMsAjaxFactory.isMsAjaxLoaded())
                    a();
                else
                    throw "Not able to load MicrosoftAjax.js."
            });
        else
            OSF._OfficeAppFactory.getLoadScriptHelper().waitForScripts([OSF.ConstantNames.MicrosoftAjaxId], a);
    else
        a()
}
)();
Microsoft.Office.WebExtension.EventType = {};
OSF.EventDispatch = function(c) {
    var b = this;
    b._eventHandlers = {};
    b._objectEventHandlers = {};
    b._queuedEventsArgs = {};
    if (c != null)
        for (var d = 0; d < c.length; d++) {
            var a = c[d]
              , e = a == "objectDeleted" || a == "objectSelectionChanged" || a == "objectDataChanged" || a == "contentControlAdded";
            if (!e)
                b._eventHandlers[a] = [];
            else
                b._objectEventHandlers[a] = {};
            b._queuedEventsArgs[a] = []
        }
}
;
OSF.EventDispatch.prototype = {
    getSupportedEvents: function() {
        var a = [];
        for (var b in this._eventHandlers)
            a.push(b);
        for (var b in this._objectEventHandlers)
            a.push(b);
        return a
    },
    supportsEvent: function(b) {
        for (var a in this._eventHandlers)
            if (b == a)
                return true;
        for (var a in this._objectEventHandlers)
            if (b == a)
                return true;
        return false
    },
    hasEventHandler: function(c, d) {
        var a = this._eventHandlers[c];
        if (a && a.length > 0)
            for (var b = 0; b < a.length; b++)
                if (a[b] === d)
                    return true;
        return false
    },
    hasObjectEventHandler: function(d, e, f) {
        var c = this._objectEventHandlers[d];
        if (c != null)
            for (var a = c[e], b = 0; a != null && b < a.length; b++)
                if (a[b] === f)
                    return true;
        return false
    },
    addEventHandler: function(b, a) {
        if (typeof a != "function")
            return false;
        var c = this._eventHandlers[b];
        if (c && !this.hasEventHandler(b, a)) {
            c.push(a);
            return true
        } else
            return false
    },
    addObjectEventHandler: function(d, b, c) {
        if (typeof c != "function")
            return false;
        var a = this._objectEventHandlers[d];
        if (a && !this.hasObjectEventHandler(d, b, c)) {
            if (a[b] == null)
                a[b] = [];
            a[b].push(c);
            return true
        }
        return false
    },
    addEventHandlerAndFireQueuedEvent: function(a, e) {
        var d = this._eventHandlers[a]
          , c = d.length == 0
          , b = this.addEventHandler(a, e);
        c && b && this.fireQueuedEvent(a);
        return b
    },
    removeEventHandler: function(c, d) {
        var a = this._eventHandlers[c];
        if (a && a.length > 0)
            for (var b = 0; b < a.length; b++)
                if (a[b] === d) {
                    a.splice(b, 1);
                    return true
                }
        return false
    },
    removeObjectEventHandler: function(d, e, f) {
        var c = this._objectEventHandlers[d];
        if (c != null)
            for (var a = c[e], b = 0; a != null && b < a.length; b++)
                if (a[b] === f) {
                    a.splice(b, 1);
                    return true
                }
        return false
    },
    clearEventHandlers: function(a) {
        if (typeof this._eventHandlers[a] != "undefined" && this._eventHandlers[a].length > 0) {
            this._eventHandlers[a] = [];
            return true
        }
        return false
    },
    clearObjectEventHandlers: function(a, b) {
        if (this._objectEventHandlers[a] != null && this._objectEventHandlers[a][b] != null) {
            this._objectEventHandlers[a][b] = [];
            return true
        }
        return false
    },
    getEventHandlerCount: function(a) {
        return this._eventHandlers[a] != undefined ? this._eventHandlers[a].length : -1
    },
    getObjectEventHandlerCount: function(a, b) {
        if (this._objectEventHandlers[a] == null || this._objectEventHandlers[a][b] == null)
            return 0;
        return this._objectEventHandlers[a][b].length
    },
    fireEvent: function(a) {
        if (a.type == undefined)
            return false;
        var b = a.type;
        if (b && this._eventHandlers[b]) {
            for (var d = this._eventHandlers[b], c = 0; c < d.length; c++)
                d[c](a);
            return true
        } else
            return false
    },
    fireObjectEvent: function(f, a) {
        if (a.type == undefined)
            return false;
        var b = a.type;
        if (b && this._objectEventHandlers[b]) {
            var e = this._objectEventHandlers[b]
              , c = e[f];
            if (c != null) {
                for (var d = 0; d < c.length; d++)
                    c[d](a);
                return true
            }
        }
        return false
    },
    fireOrQueueEvent: function(c) {
        var b = this
          , a = c.type;
        if (a && b._eventHandlers[a]) {
            var d = b._eventHandlers[a]
              , e = b._queuedEventsArgs[a];
            if (d.length == 0)
                e.push(c);
            else
                b.fireEvent(c);
            return true
        } else
            return false
    },
    fireQueuedEvent: function(a) {
        if (a && this._eventHandlers[a]) {
            var b = this._eventHandlers[a]
              , c = this._queuedEventsArgs[a];
            if (b.length > 0) {
                var d = b[0];
                while (c.length > 0) {
                    var e = c.shift();
                    d(e)
                }
                return true
            }
        }
        return false
    },
    clearQueuedEvent: function(a) {
        if (a && this._eventHandlers[a]) {
            var b = this._queuedEventsArgs[a];
            if (b)
                this._queuedEventsArgs[a] = []
        }
    }
};
OSF.DDA.OMFactory = OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureEventArgs = function(c, f, b) {
    var h = "hostPlatform", e = "outlook", d = "hostType", g = this, a;
    switch (c) {
    case Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged:
        a = new OSF.DDA.DocumentSelectionChangedEventArgs(f);
        break;
    case Microsoft.Office.WebExtension.EventType.BindingSelectionChanged:
        a = new OSF.DDA.BindingSelectionChangedEventArgs(g.manufactureBinding(b, f.document),b[OSF.DDA.PropertyDescriptors.Subset]);
        break;
    case Microsoft.Office.WebExtension.EventType.BindingDataChanged:
        a = new OSF.DDA.BindingDataChangedEventArgs(g.manufactureBinding(b, f.document));
        break;
    case Microsoft.Office.WebExtension.EventType.SettingsChanged:
        a = new OSF.DDA.SettingsChangedEventArgs(f);
        break;
    case Microsoft.Office.WebExtension.EventType.ActiveViewChanged:
        a = new OSF.DDA.ActiveViewChangedEventArgs(b);
        break;
    case Microsoft.Office.WebExtension.EventType.OfficeThemeChanged:
        a = new OSF.DDA.Theming.OfficeThemeChangedEventArgs(b);
        break;
    case Microsoft.Office.WebExtension.EventType.DocumentThemeChanged:
        a = new OSF.DDA.Theming.DocumentThemeChangedEventArgs(b);
        break;
    case Microsoft.Office.WebExtension.EventType.Suspend:
        a = new OSF.DDA.SuspendEventArgs(b);
        break;
    case Microsoft.Office.WebExtension.EventType.Resume:
        a = new OSF.DDA.ResumeEventArgs(b);
        break;
    case Microsoft.Office.WebExtension.EventType.AppCommandInvoked:
        a = OSF.DDA.AppCommand.AppCommandInvokedEventArgs.create(b);
        OSF._OfficeAppFactory.getHostInfo()[d] == e && OSF._OfficeAppFactory.getHostInfo()[h] == "mac" && OSF.DDA.convertOlkAppointmentTimeToDateFormat(a);
        break;
    case Microsoft.Office.WebExtension.EventType.ObjectDeleted:
    case Microsoft.Office.WebExtension.EventType.ObjectSelectionChanged:
    case Microsoft.Office.WebExtension.EventType.ObjectDataChanged:
    case Microsoft.Office.WebExtension.EventType.ContentControlAdded:
        a = new OSF.DDA.ObjectEventArgs(c,b[Microsoft.Office.WebExtension.Parameters.Id]);
        break;
    case Microsoft.Office.WebExtension.EventType.RichApiMessage:
        a = new OSF.DDA.RichApiMessageEventArgs(c,b);
        break;
    case Microsoft.Office.WebExtension.EventType.DataNodeInserted:
        a = new OSF.DDA.NodeInsertedEventArgs(g.manufactureDataNode(b[OSF.DDA.DataNodeEventProperties.NewNode]),b[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
        break;
    case Microsoft.Office.WebExtension.EventType.DataNodeReplaced:
        a = new OSF.DDA.NodeReplacedEventArgs(g.manufactureDataNode(b[OSF.DDA.DataNodeEventProperties.OldNode]),g.manufactureDataNode(b[OSF.DDA.DataNodeEventProperties.NewNode]),b[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
        break;
    case Microsoft.Office.WebExtension.EventType.DataNodeDeleted:
        a = new OSF.DDA.NodeDeletedEventArgs(g.manufactureDataNode(b[OSF.DDA.DataNodeEventProperties.OldNode]),g.manufactureDataNode(b[OSF.DDA.DataNodeEventProperties.NextSiblingNode]),b[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
        break;
    case Microsoft.Office.WebExtension.EventType.TaskSelectionChanged:
        a = new OSF.DDA.TaskSelectionChangedEventArgs(f);
        break;
    case Microsoft.Office.WebExtension.EventType.ResourceSelectionChanged:
        a = new OSF.DDA.ResourceSelectionChangedEventArgs(f);
        break;
    case Microsoft.Office.WebExtension.EventType.ViewSelectionChanged:
        a = new OSF.DDA.ViewSelectionChangedEventArgs(f);
        break;
    case Microsoft.Office.WebExtension.EventType.DialogMessageReceived:
        a = new OSF.DDA.DialogEventArgs(b);
        break;
    case Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived:
        a = new OSF.DDA.DialogParentEventArgs(b);
        break;
    case Microsoft.Office.WebExtension.EventType.ItemChanged:
        if (OSF._OfficeAppFactory.getHostInfo()[d] == e) {
            a = new OSF.DDA.OlkItemSelectedChangedEventArgs(b);
            f.initialize(a["initialData"]);
            (OSF._OfficeAppFactory.getHostInfo()[h] == "win32" || OSF._OfficeAppFactory.getHostInfo()[h] == "mac") && f.setCurrentItemNumber(a["itemNumber"].itemNumber)
        } else
            throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, c));
        break;
    case Microsoft.Office.WebExtension.EventType.RecipientsChanged:
        if (OSF._OfficeAppFactory.getHostInfo()[d] == e)
            a = new OSF.DDA.OlkRecipientsChangedEventArgs(b);
        else
            throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, c));
        break;
    case Microsoft.Office.WebExtension.EventType.AppointmentTimeChanged:
        if (OSF._OfficeAppFactory.getHostInfo()[d] == e)
            a = new OSF.DDA.OlkAppointmentTimeChangedEventArgs(b);
        else
            throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, c));
        break;
    case Microsoft.Office.WebExtension.EventType.RecurrenceChanged:
        if (OSF._OfficeAppFactory.getHostInfo()[d] == e)
            a = new OSF.DDA.OlkRecurrenceChangedEventArgs(b);
        else
            throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, c));
        break;
    case Microsoft.Office.WebExtension.EventType.AttachmentsChanged:
        if (OSF._OfficeAppFactory.getHostInfo()[d] == e)
            a = new OSF.DDA.OlkAttachmentsChangedEventArgs(b);
        else
            throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, c));
        break;
    case Microsoft.Office.WebExtension.EventType.EnhancedLocationsChanged:
        if (OSF._OfficeAppFactory.getHostInfo()[d] == e)
            a = new OSF.DDA.OlkEnhancedLocationsChangedEventArgs(b);
        else
            throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, c));
        break;
    case Microsoft.Office.WebExtension.EventType.InfobarClicked:
        if (OSF._OfficeAppFactory.getHostInfo()[d] == e)
            a = new OSF.DDA.OlkInfobarClickedEventArgs(b);
        else
            throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, c));
        break;
    case Microsoft.Office.WebExtension.EventType.SelectedItemsChanged:
        if (OSF._OfficeAppFactory.getHostInfo()[d] == e)
            a = new OSF.DDA.OlkSelectedItemsChangedEventArgs(b);
        else
            throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, c));
        break;
    case Microsoft.Office.WebExtension.EventType.SensitivityLabelChanged:
        if (OSF._OfficeAppFactory.getHostInfo()[d] == e)
            a = new OSF.DDA.OlkSensitivityLabelChangedEventArgs(b);
        else
            throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, c));
        break;
    case Microsoft.Office.WebExtension.EventType.InitializationContextChanged:
        if (OSF._OfficeAppFactory.getHostInfo()[d] == e)
            a = new OSF.DDA.OlkInitializationContextChangedEventArgs(b);
        else
            throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, c));
        break;
    case Microsoft.Office.WebExtension.EventType.DragAndDropEvent:
        if (OSF._OfficeAppFactory.getHostInfo()[d] == e && OSF._OfficeAppFactory.getHostInfo()[h] == "web")
            a = new OSF.DDA.OlkDragAndDropEventArgs(b);
        else
            throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, c));
        break;
    default:
        throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, c))
    }
    return a
}
;
OSF.DDA.AsyncMethodNames.addNames({
    AddHandlerAsync: "addHandlerAsync",
    RemoveHandlerAsync: "removeHandlerAsync"
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.AddHandlerAsync,
    requiredArguments: [{
        name: Microsoft.Office.WebExtension.Parameters.EventType,
        "enum": Microsoft.Office.WebExtension.EventType,
        verify: function(b, c, a) {
            return a.supportsEvent(b)
        }
    }, {
        name: Microsoft.Office.WebExtension.Parameters.Handler,
        types: ["function"]
    }],
    supportedOptions: [],
    privateStateCallbacks: []
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.RemoveHandlerAsync,
    requiredArguments: [{
        name: Microsoft.Office.WebExtension.Parameters.EventType,
        "enum": Microsoft.Office.WebExtension.EventType,
        verify: function(b, c, a) {
            return a.supportsEvent(b)
        }
    }],
    supportedOptions: [{
        name: Microsoft.Office.WebExtension.Parameters.Handler,
        value: {
            types: ["function", "object"],
            defaultValue: null
        }
    }],
    privateStateCallbacks: []
});
OSF.DialogShownStatus = {
    hasDialogShown: false,
    isWindowDialog: false
};
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
    DialogMessageReceivedEvent: "DialogMessageReceivedEvent"
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    DialogMessageReceived: "dialogMessageReceived",
    DialogEventReceived: "dialogEventReceived"
});
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
    MessageType: "messageType",
    MessageContent: "messageContent",
    MessageOrigin: "messageOrigin"
});
OSF.DDA.DialogEventType = {};
OSF.OUtil.augmentList(OSF.DDA.DialogEventType, {
    DialogClosed: "dialogClosed",
    NavigationFailed: "naviationFailed"
});
OSF.DDA.AsyncMethodNames.addNames({
    DisplayDialogAsync: "displayDialogAsync",
    DisplayModalDialogAsync: "displayModalDialogAsync",
    CloseAsync: "close"
});
OSF.DDA.SyncMethodNames.addNames({
    MessageParent: "messageParent",
    MessageChild: "messageChild",
    SendMessage: "sendMessage",
    AddMessageHandler: "addEventHandler"
});
OSF.DDA.UI.ParentUI = function() {
    var a;
    if (Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived != null)
        a = new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DialogMessageReceived, Microsoft.Office.WebExtension.EventType.DialogEventReceived, Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived]);
    else
        a = new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DialogMessageReceived, Microsoft.Office.WebExtension.EventType.DialogEventReceived]);
    var b = this
      , c = function(c, d) {
        !b[c] && OSF.OUtil.defineEnumerableProperty(b, c, {
            value: function() {
                var c = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.OpenDialog];
                c(arguments, a, b, d)
            }
        })
    };
    c(OSF.DDA.AsyncMethodNames.DisplayDialogAsync.displayName, false);
    Microsoft.Office.WebExtension.FeatureGates["ModalWebDialogAPI"] && c(OSF.DDA.AsyncMethodNames.DisplayModalDialogAsync.displayName, true);
    OSF.OUtil.finalizeProperties(this)
}
;
OSF.DDA.UI.ChildUI = function(d) {
    var b = OSF.DDA.SyncMethodNames.MessageParent.displayName
      , a = this;
    !a[b] && OSF.OUtil.defineEnumerableProperty(a, b, {
        value: function() {
            var b = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.MessageParent];
            return b(arguments, a)
        }
    });
    var c = OSF.DDA.SyncMethodNames.AddMessageHandler.displayName;
    !a[c] && typeof OSF.DialogParentMessageEventDispatch != "undefined" && OSF.DDA.DispIdHost.addEventSupport(a, OSF.DialogParentMessageEventDispatch, d);
    OSF.OUtil.finalizeProperties(this)
}
;
OSF.DialogHandler = function() {}
;
OSF.DDA.DialogEventArgs = function(a) {
    if (a[OSF.DDA.PropertyDescriptors.MessageType] == OSF.DialogMessageType.DialogMessageReceived)
        OSF.OUtil.defineEnumerableProperties(this, {
            type: {
                value: Microsoft.Office.WebExtension.EventType.DialogMessageReceived
            },
            message: {
                value: a[OSF.DDA.PropertyDescriptors.MessageContent]
            },
            origin: {
                value: a[OSF.DDA.PropertyDescriptors.MessageOrigin]
            }
        });
    else
        OSF.OUtil.defineEnumerableProperties(this, {
            type: {
                value: Microsoft.Office.WebExtension.EventType.DialogEventReceived
            },
            error: {
                value: a[OSF.DDA.PropertyDescriptors.MessageType]
            }
        })
}
;
OSF.DDA.DialogParentEventArgs = function(a) {
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived
        },
        message: {
            value: a[OSF.DDA.PropertyDescriptors.MessageContent]
        },
        origin: {
            value: a[OSF.DDA.PropertyDescriptors.MessageOrigin]
        }
    })
}
;
var DialogApiManager = function() {
    var d = false
      , c = "boolean"
      , a = true;
    function b() {}
    b.defineApi = function(d, c) {
        var b = OSF.DDA.AsyncMethodCalls;
        b.define({
            method: d,
            requiredArguments: [{
                name: Microsoft.Office.WebExtension.Parameters.Url,
                types: ["string"]
            }],
            supportedOptions: c,
            privateStateCallbacks: [],
            onSucceeded: function(d) {
                var i = d[Microsoft.Office.WebExtension.Parameters.Id]
                  , c = d[Microsoft.Office.WebExtension.Parameters.Data]
                  , b = new OSF.DialogHandler
                  , f = OSF.DDA.AsyncMethodNames.CloseAsync.displayName;
                OSF.OUtil.defineEnumerableProperty(b, f, {
                    value: function() {
                        var a = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.CloseDialog];
                        a(arguments, i, c, b)
                    }
                });
                var h = OSF.DDA.SyncMethodNames.AddMessageHandler.displayName;
                OSF.OUtil.defineEnumerableProperty(b, h, {
                    value: function() {
                        var d = OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.AddMessageHandler.id]
                          , a = d.verifyAndExtractCall(arguments, b, c)
                          , e = a[Microsoft.Office.WebExtension.Parameters.EventType]
                          , f = a[Microsoft.Office.WebExtension.Parameters.Handler];
                        return c.addEventHandlerAndFireQueuedEvent(e, f)
                    }
                });
                if (OSF.DDA.UI.EnableSendMessageDialogAPI === a) {
                    var g = OSF.DDA.SyncMethodNames.SendMessage.displayName;
                    OSF.OUtil.defineEnumerableProperty(b, g, {
                        value: function() {
                            var a = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.SendMessage];
                            return a(arguments, c, b)
                        }
                    })
                }
                if (OSF.DDA.UI.EnableMessageChildDialogAPI === a) {
                    var e = OSF.DDA.SyncMethodNames.MessageChild.displayName;
                    OSF.OUtil.defineEnumerableProperty(b, e, {
                        value: function() {
                            var a = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.SendMessage];
                            return a(arguments, c, b)
                        }
                    })
                }
                return b
            },
            checkCallArgs: function(b) {
                if (b[Microsoft.Office.WebExtension.Parameters.Width] <= 0)
                    b[Microsoft.Office.WebExtension.Parameters.Width] = 1;
                if (!b[Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels] && b[Microsoft.Office.WebExtension.Parameters.Width] > 100)
                    b[Microsoft.Office.WebExtension.Parameters.Width] = 99;
                if (b[Microsoft.Office.WebExtension.Parameters.Height] <= 0)
                    b[Microsoft.Office.WebExtension.Parameters.Height] = 1;
                if (!b[Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels] && b[Microsoft.Office.WebExtension.Parameters.Height] > 100)
                    b[Microsoft.Office.WebExtension.Parameters.Height] = 99;
                if (!b[Microsoft.Office.WebExtension.Parameters.RequireHTTPs])
                    b[Microsoft.Office.WebExtension.Parameters.RequireHTTPs] = a;
                return b
            }
        })
    }
    ;
    b.messageChildRichApiBridge = function() {
        if (OSF.DDA.UI.EnableMessageChildDialogAPI === a) {
            var b = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.SendMessage];
            return b(arguments, null, null)
        }
    }
    ;
    b.initOnce = function() {
        b.defineApi(OSF.DDA.AsyncMethodNames.DisplayDialogAsync, b.displayDialogAsyncApiSupportedOptions);
        b.defineApi(OSF.DDA.AsyncMethodNames.DisplayModalDialogAsync, b.displayModalDialogAsyncApiSupportedOptions)
    }
    ;
    b.displayDialogAsyncApiSupportedOptions = [{
        name: Microsoft.Office.WebExtension.Parameters.Width,
        value: {
            types: ["number"],
            defaultValue: 99
        }
    }, {
        name: Microsoft.Office.WebExtension.Parameters.Height,
        value: {
            types: ["number"],
            defaultValue: 99
        }
    }, {
        name: Microsoft.Office.WebExtension.Parameters.RequireHTTPs,
        value: {
            types: [c],
            defaultValue: a
        }
    }, {
        name: Microsoft.Office.WebExtension.Parameters.DisplayInIframe,
        value: {
            types: [c],
            defaultValue: d
        }
    }, {
        name: Microsoft.Office.WebExtension.Parameters.HideTitle,
        value: {
            types: [c],
            defaultValue: d
        }
    }, {
        name: Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels,
        value: {
            types: [c],
            defaultValue: d
        }
    }, {
        name: Microsoft.Office.WebExtension.Parameters.PromptBeforeOpen,
        value: {
            types: [c],
            defaultValue: a
        }
    }, {
        name: Microsoft.Office.WebExtension.Parameters.EnforceAppDomain,
        value: {
            types: [c],
            defaultValue: a
        }
    }, {
        name: Microsoft.Office.WebExtension.Parameters.UrlNoHostInfo,
        value: {
            types: [c],
            defaultValue: d
        }
    }];
    b.displayModalDialogAsyncApiSupportedOptions = b.displayDialogAsyncApiSupportedOptions.concat([{
        name: "abortWhenParentIsMinimized",
        value: {
            types: [c],
            defaultValue: d
        }
    }, {
        name: "abortWhenDocIsInactive",
        value: {
            types: [c],
            defaultValue: d
        }
    }]);
    return b
}();
DialogApiManager.initOnce();
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.CloseAsync,
    requiredArguments: [],
    supportedOptions: [],
    privateStateCallbacks: []
});
OSF.DDA.SyncMethodCalls.define({
    method: OSF.DDA.SyncMethodNames.MessageParent,
    requiredArguments: [{
        name: Microsoft.Office.WebExtension.Parameters.MessageToParent,
        types: ["string", "number", "boolean"]
    }],
    supportedOptions: [{
        name: Microsoft.Office.WebExtension.Parameters.TargetOrigin,
        value: {
            types: ["string"],
            defaultValue: ""
        }
    }]
});
OSF.DDA.SyncMethodCalls.define({
    method: OSF.DDA.SyncMethodNames.AddMessageHandler,
    requiredArguments: [{
        name: Microsoft.Office.WebExtension.Parameters.EventType,
        "enum": Microsoft.Office.WebExtension.EventType,
        verify: function(b, c, a) {
            return a.supportsEvent(b)
        }
    }, {
        name: Microsoft.Office.WebExtension.Parameters.Handler,
        types: ["function"]
    }],
    supportedOptions: []
});
OSF.DDA.SyncMethodCalls.define({
    method: OSF.DDA.SyncMethodNames.SendMessage,
    requiredArguments: [{
        name: Microsoft.Office.WebExtension.Parameters.MessageContent,
        types: ["string"]
    }],
    supportedOptions: [{
        name: Microsoft.Office.WebExtension.Parameters.TargetOrigin,
        value: {
            types: ["string"],
            defaultValue: ""
        }
    }],
    privateStateCallbacks: []
});
OSF.DDA.SafeArray.Delegate.openDialog = function(a) {
    try {
        a.onCalling && a.onCalling();
        var c = OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(true, a);
        OSF.ClientHostController.openDialog(a.dispId, a.targetId, function(c, b) {
            a.onEvent && a.onEvent(b);
            OSF.AppTelemetry && OSF.AppTelemetry.onEventDone(a.dispId)
        }, c)
    } catch (b) {
        OSF.DDA.SafeArray.Delegate._onException(b, a)
    }
}
;
OSF.DDA.SafeArray.Delegate.closeDialog = function(a) {
    a.onCalling && a.onCalling();
    var c = OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(false, a);
    try {
        OSF.ClientHostController.closeDialog(a.dispId, a.targetId, c)
    } catch (b) {
        OSF.DDA.SafeArray.Delegate._onException(b, a)
    }
}
;
OSF.DDA.SafeArray.Delegate.messageParent = function(a) {
    try {
        a.onCalling && a.onCalling();
        var d = (new Date).getTime()
          , b = OSF.ClientHostController.messageParent(a.hostCallArgs);
        a.onReceiving && a.onReceiving();
        OSF.AppTelemetry && OSF.AppTelemetry.onMethodDone(a.dispId, a.hostCallArgs, Math.abs((new Date).getTime() - d), b);
        return b
    } catch (c) {
        return OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod(c)
    }
}
;
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidDialogMessageReceivedEvent,
    fromHost: [{
        name: OSF.DDA.EventDescriptors.DialogMessageReceivedEvent,
        value: OSF.DDA.SafeArray.Delegate.ParameterMap.self
    }],
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDescriptors.DialogMessageReceivedEvent,
    fromHost: [{
        name: OSF.DDA.PropertyDescriptors.MessageType,
        value: 0
    }, {
        name: OSF.DDA.PropertyDescriptors.MessageContent,
        value: 1
    }, {
        name: OSF.DDA.PropertyDescriptors.MessageOrigin,
        value: 2
    }],
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.sendMessage = function(a) {
    try {
        a.onCalling && a.onCalling();
        var d = (new Date).getTime()
          , c = OSF.ClientHostController.sendMessage(a.hostCallArgs);
        a.onReceiving && a.onReceiving();
        return c
    } catch (b) {
        return OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod(b)
    }
}
;
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    ItemChanged: "olkItemSelectedChanged"
});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
    OlkItemSelectedData: "OlkItemSelectedData"
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    RecipientsChanged: "olkRecipientsChanged"
});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
    OlkRecipientsData: "OlkRecipientsData"
});
OSF.DDA.OlkRecipientsChangedEventArgs = function(b) {
    var a = b[OSF.DDA.EventDescriptors.OlkRecipientsData][0];
    if (a === "")
        a = null;
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.RecipientsChanged
        },
        changedRecipientFields: {
            value: JSON.parse(a)
        }
    })
}
;
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    AppointmentTimeChanged: "olkAppointmentTimeChanged"
});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
    OlkAppointmentTimeChangedData: "OlkAppointmentTimeChangedData"
});
OSF.DDA.OlkAppointmentTimeChangedEventArgs = function(e) {
    var d = e[OSF.DDA.EventDescriptors.OlkAppointmentTimeChangedData][0], a, b;
    try {
        var c = JSON.parse(d);
        a = (new Date(c.start)).toISOString();
        b = (new Date(c.end)).toISOString()
    } catch (f) {
        a = null;
        b = null
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.AppointmentTimeChanged
        },
        start: {
            value: a
        },
        end: {
            value: b
        }
    })
}
;
OSF.DDA.convertOlkAppointmentTimeToDateFormat = function(e) {
    var b = null
      , a = JSON.parse(e.eventObjStr);
    if (a != b && a.type != b && a.type == "olkAppointmentTimeChanged") {
        var c, d;
        try {
            c = (new Date(a.start)).toISOString();
            d = (new Date(a.end)).toISOString()
        } catch (f) {
            c = b;
            d = b
        }
        a.start = c;
        a.end = d
    }
    e.eventObjStr = JSON.stringify(a)
}
;
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    RecurrenceChanged: "olkRecurrenceChanged"
});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
    OlkRecurrenceData: "OlkRecurrenceData"
});
OSF.DDA.OlkRecurrenceChangedEventArgs = function(c) {
    var a = null;
    try {
        var b = JSON.parse(c[OSF.DDA.EventDescriptors.OlkRecurrenceChangedData][0]);
        if (b.recurrence != null) {
            a = JSON.parse(b.recurrence);
            a = Microsoft.Office.WebExtension.OutlookBase.SeriesTimeJsonConverter(a)
        }
    } catch (d) {
        a = null
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.RecurrenceChanged
        },
        recurrence: {
            value: a
        }
    })
}
;
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    OfficeThemeChanged: "officeThemeChanged"
});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
    OfficeThemeData: "OfficeThemeData"
});
OSF.OUtil.setNamespace("Theming", OSF.DDA);
OSF.DDA.Theming.OfficeThemeChangedEventArgs = function(b) {
    var a = {};
    if (b.KeepHexColors)
        try {
            a = JSON.parse(b.OfficeThemeData[0])
        } catch (e) {}
    else {
        var c = JSON.parse(b.OfficeThemeData[0]);
        for (var d in c)
            a[d] = OSF.OUtil.convertIntToCssHexColor(c[d])
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.OfficeThemeChanged
        },
        officeTheme: {
            value: a
        }
    })
}
;
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    AttachmentsChanged: "olkAttachmentsChanged"
});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
    OlkAttachmentsChangedData: "OlkAttachmentsChangedData"
});
OSF.DDA.OlkAttachmentsChangedEventArgs = function(d) {
    var b, a;
    try {
        var c = JSON.parse(d[OSF.DDA.EventDescriptors.OlkAttachmentsChangedData][0]);
        b = c.attachmentStatus;
        a = Microsoft.Office.WebExtension.OutlookBase.CreateAttachmentDetails(c.attachmentDetails)
    } catch (e) {
        b = null;
        a = null
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.AttachmentsChanged
        },
        attachmentStatus: {
            value: b
        },
        attachmentDetails: {
            value: a
        }
    })
}
;
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    EnhancedLocationsChanged: "olkEnhancedLocationsChanged"
});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
    OlkEnhancedLocationsChangedData: "OlkEnhancedLocationsChangedData"
});
OSF.DDA.OlkEnhancedLocationsChangedEventArgs = function(c) {
    var a;
    try {
        var b = JSON.parse(c[OSF.DDA.EventDescriptors.OlkEnhancedLocationsChangedData][0]);
        a = b.enhancedLocations
    } catch (d) {
        a = null
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.EnhancedLocationsChanged
        },
        enhancedLocations: {
            value: a
        }
    })
}
;
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    InfobarClicked: "olkInfobarClicked"
});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
    OlkInfobarClickedData: "OlkInfobarClickedData"
});
OSF.DDA.OlkInfobarClickedEventArgs = function(b) {
    var a;
    try {
        a = b[OSF.DDA.EventDescriptors.OlkInfobarClickedData][0]
    } catch (c) {
        a = null
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.InfobarClicked
        },
        infobarDetails: {
            value: a
        }
    })
}
;
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    SelectedItemsChanged: "olkSelectedItemsChanged"
});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
    OlkSelectionData: "OlkSelectionData"
});
OSF.DDA.OlkSelectedItemsChangedEventArgs = function() {
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.SelectedItemsChanged
        }
    })
}
;
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    SensitivityLabelChanged: "olkSensitivityLabelChanged"
});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
    OlkSensitivityLabelChangedData: "OlkSensitivityLabelChangedData"
});
OSF.DDA.OlkSensitivityLabelChangedEventArgs = function() {
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.SensitivityLabelChanged
        }
    })
}
;
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    InitializationContextChanged: "olkInitializationContextChanged"
});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
    OlkInitializationContextChangedData: "OlkInitializationContextChangedData"
});
OSF.DDA.OlkInitializationContextChangedEventArgs = function(b) {
    var a;
    try {
        a = b[OSF.DDA.EventDescriptors.OlkInitializationContextChangedData][0]
    } catch (c) {
        a = null
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.InitializationContextChanged
        },
        initializationContextData: {
            value: a
        }
    })
}
;
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    DragAndDropEvent: "olkDragAndDropEvent"
});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
    OlkDragAndDropEventData: "OlkDragAndDropEventData"
});
OSF.DDA.OlkDragAndDropEventArgs = function(c) {
    var b;
    try {
        var a = c[OSF.DDA.EventDescriptors.OlkDragAndDropEventData];
        if (a.type === "dragover")
            b = a;
        else if (a.type === "drop")
            b = __assign({}, a, {
                dataTransfer: __assign({}, a.dataTransfer, {
                    files: a.dataTransfer.files.map(function(a) {
                        for (var c = atob(a.fileContent), d = [], b = 0; b < c.length; b++)
                            d.push(c.charCodeAt(b));
                        var e = new Uint8Array(d);
                        return __assign({}, a, {
                            fileContent: new Blob([e],{
                                type: a.type
                            })
                        })
                    })
                })
            })
    } catch (d) {
        b = null
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.DragAndDropEvent
        },
        dragAndDropEventData: {
            value: b
        }
    })
}
;
OSF.DDA.OlkItemSelectedChangedEventArgs = function(b) {
    var a = b[OSF.DDA.EventDescriptors.OlkItemSelectedData][0];
    if (a === "")
        a = null;
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.ItemChanged
        },
        initialData: {
            value: JSON.parse(a)
        },
        itemNumber: {
            value: JSON.parse(b[OSF.DDA.EventDescriptors.OlkItemSelectedData][1])
        }
    })
}
;
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkItemSelectedChangedEvent,
    fromHost: [{
        name: OSF.DDA.EventDescriptors.OlkItemSelectedData,
        value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData
    }],
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkRecipientsChangedEvent,
    fromHost: [{
        name: OSF.DDA.EventDescriptors.OlkRecipientsData,
        value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData
    }],
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkAppointmentTimeChangedEvent,
    fromHost: [{
        name: OSF.DDA.EventDescriptors.OlkAppointmentTimeChangedData,
        value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData
    }],
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkRecurrenceChangedEvent,
    fromHost: [{
        name: OSF.DDA.EventDescriptors.OlkRecurrenceChangedData,
        value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData
    }],
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOfficeThemeChangedEvent,
    fromHost: [{
        name: OSF.DDA.EventDescriptors.OfficeThemeData,
        value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData
    }],
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkAttachmentsChangedEvent,
    fromHost: [{
        name: OSF.DDA.EventDescriptors.OlkAttachmentsChangedData,
        value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData
    }],
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkEnhancedLocationsChangedEvent,
    fromHost: [{
        name: OSF.DDA.EventDescriptors.OlkEnhancedLocationsChangedData,
        value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData
    }],
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkInfobarClickedEvent,
    fromHost: [{
        name: OSF.DDA.EventDescriptors.OlkInfobarClickedData,
        value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData
    }],
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkSelectedItemsChangedEvent,
    fromHost: [{
        name: OSF.DDA.EventDescriptors.OlkSelectionData,
        value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData
    }],
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkSensitivityLabelChangedEvent,
    fromHost: [{
        name: OSF.DDA.EventDescriptors.OlkSensitivityLabelChangedData,
        value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData
    }],
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidOlkInitializationContextChangedEvent,
    fromHost: [{
        name: OSF.DDA.EventDescriptors.OlkInitializationContextChangedData,
        value: OSF.DDA.SafeArray.Delegate.ParameterMap.self
    }],
    isComplexType: true
});
var OSFLog;
(function(g) {
    var e = "ResponseTime"
      , d = "Message"
      , c = "SessionId"
      , b = "CorrelationId"
      , a = true
      , f = function() {
        function b(a) {
            this._table = a;
            this._fields = {}
        }
        Object.defineProperty(b.prototype, "Fields", {
            "get": function() {
                return this._fields
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(b.prototype, "Table", {
            "get": function() {
                return this._table
            },
            enumerable: a,
            configurable: a
        });
        b.prototype.SerializeFields = function() {}
        ;
        b.prototype.SetSerializedField = function(b, a) {
            if (typeof a !== "undefined" && a !== null)
                this._serializedFields[b] = a.toString()
        }
        ;
        b.prototype.SerializeRow = function() {
            var a = this;
            a._serializedFields = {};
            a.SetSerializedField("Table", a._table);
            a.SerializeFields();
            return JSON.stringify(a._serializedFields)
        }
        ;
        return b
    }();
    g.BaseUsageData = f;
    var i = function(y) {
        var x = "LaunchReason"
          , w = "LaunchSource"
          , v = "IsMOS"
          , u = "IsFromWacAutomation"
          , t = "WacHostEnvironment"
          , s = "HostJSVersion"
          , r = "OfficeJSVersion"
          , q = "DocUrl"
          , p = "AppSizeHeight"
          , o = "AppSizeWidth"
          , n = "ClientId"
          , m = "HostVersion"
          , l = "Host"
          , k = "UserId"
          , j = "Browser"
          , i = "AssetId"
          , h = "AppURL"
          , g = "AppInstanceId"
          , f = "AppId";
        __extends(e, y);
        function e() {
            return y.call(this, "AppActivated") || this
        }
        Object.defineProperty(e.prototype, b, {
            "get": function() {
                return this.Fields[b]
            },
            "set": function(a) {
                this.Fields[b] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, c, {
            "get": function() {
                return this.Fields[c]
            },
            "set": function(a) {
                this.Fields[c] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, f, {
            "get": function() {
                return this.Fields[f]
            },
            "set": function(a) {
                this.Fields[f] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, g, {
            "get": function() {
                return this.Fields[g]
            },
            "set": function(a) {
                this.Fields[g] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, h, {
            "get": function() {
                return this.Fields[h]
            },
            "set": function(a) {
                this.Fields[h] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, i, {
            "get": function() {
                return this.Fields[i]
            },
            "set": function(a) {
                this.Fields[i] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, j, {
            "get": function() {
                return this.Fields[j]
            },
            "set": function(a) {
                this.Fields[j] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, k, {
            "get": function() {
                return this.Fields[k]
            },
            "set": function(a) {
                this.Fields[k] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, l, {
            "get": function() {
                return this.Fields[l]
            },
            "set": function(a) {
                this.Fields[l] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, m, {
            "get": function() {
                return this.Fields[m]
            },
            "set": function(a) {
                this.Fields[m] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, n, {
            "get": function() {
                return this.Fields[n]
            },
            "set": function(a) {
                this.Fields[n] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, o, {
            "get": function() {
                return this.Fields[o]
            },
            "set": function(a) {
                this.Fields[o] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, p, {
            "get": function() {
                return this.Fields[p]
            },
            "set": function(a) {
                this.Fields[p] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, d, {
            "get": function() {
                return this.Fields[d]
            },
            "set": function(a) {
                this.Fields[d] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, q, {
            "get": function() {
                return this.Fields[q]
            },
            "set": function(a) {
                this.Fields[q] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, r, {
            "get": function() {
                return this.Fields[r]
            },
            "set": function(a) {
                this.Fields[r] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, s, {
            "get": function() {
                return this.Fields[s]
            },
            "set": function(a) {
                this.Fields[s] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, t, {
            "get": function() {
                return this.Fields[t]
            },
            "set": function(a) {
                this.Fields[t] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, u, {
            "get": function() {
                return this.Fields[u]
            },
            "set": function(a) {
                this.Fields[u] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, v, {
            "get": function() {
                return this.Fields[v]
            },
            "set": function(a) {
                this.Fields[v] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, w, {
            "get": function() {
                return this.Fields[w]
            },
            "set": function(a) {
                this.Fields[w] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(e.prototype, x, {
            "get": function() {
                return this.Fields[x]
            },
            "set": function(a) {
                this.Fields[x] = a
            },
            enumerable: a,
            configurable: a
        });
        e.prototype.SerializeFields = function() {
            var a = this;
            a.SetSerializedField(b, a.CorrelationId);
            a.SetSerializedField(c, a.SessionId);
            a.SetSerializedField(f, a.AppId);
            a.SetSerializedField(g, a.AppInstanceId);
            a.SetSerializedField(h, a.AppURL);
            a.SetSerializedField(i, a.AssetId);
            a.SetSerializedField(j, a.Browser);
            a.SetSerializedField(k, a.UserId);
            a.SetSerializedField(l, a.Host);
            a.SetSerializedField(m, a.HostVersion);
            a.SetSerializedField(n, a.ClientId);
            a.SetSerializedField(o, a.AppSizeWidth);
            a.SetSerializedField(p, a.AppSizeHeight);
            a.SetSerializedField(d, a.Message);
            a.SetSerializedField(q, a.DocUrl);
            a.SetSerializedField(r, a.OfficeJSVersion);
            a.SetSerializedField(s, a.HostJSVersion);
            a.SetSerializedField(t, a.WacHostEnvironment);
            a.SetSerializedField(u, a.IsFromWacAutomation);
            a.SetSerializedField(v, a.IsMOS);
            a.SetSerializedField(w, a.LaunchSource);
            a.SetSerializedField(x, a.LaunchReason)
        }
        ;
        return e
    }(f);
    g.AppActivatedUsageData = i;
    var k = function(h) {
        var f = "StartTime"
          , d = "ScriptId";
        __extends(g, h);
        function g() {
            return h.call(this, "ScriptLoad") || this
        }
        Object.defineProperty(g.prototype, b, {
            "get": function() {
                return this.Fields[b]
            },
            "set": function(a) {
                this.Fields[b] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(g.prototype, c, {
            "get": function() {
                return this.Fields[c]
            },
            "set": function(a) {
                this.Fields[c] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(g.prototype, d, {
            "get": function() {
                return this.Fields[d]
            },
            "set": function(a) {
                this.Fields[d] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(g.prototype, f, {
            "get": function() {
                return this.Fields[f]
            },
            "set": function(a) {
                this.Fields[f] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(g.prototype, e, {
            "get": function() {
                return this.Fields[e]
            },
            "set": function(a) {
                this.Fields[e] = a
            },
            enumerable: a,
            configurable: a
        });
        g.prototype.SerializeFields = function() {
            var a = this;
            a.SetSerializedField(b, a.CorrelationId);
            a.SetSerializedField(c, a.SessionId);
            a.SetSerializedField(d, a.ScriptId);
            a.SetSerializedField(f, a.StartTime);
            a.SetSerializedField(e, a.ResponseTime)
        }
        ;
        return g
    }(f);
    g.ScriptLoadUsageData = k;
    var l = function(j) {
        var h = "CloseMethod"
          , g = "OpenTime"
          , f = "AppSizeFinalHeight"
          , e = "AppSizeFinalWidth"
          , d = "FocusTime";
        __extends(i, j);
        function i() {
            return j.call(this, "AppClosed") || this
        }
        Object.defineProperty(i.prototype, b, {
            "get": function() {
                return this.Fields[b]
            },
            "set": function(a) {
                this.Fields[b] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(i.prototype, c, {
            "get": function() {
                return this.Fields[c]
            },
            "set": function(a) {
                this.Fields[c] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(i.prototype, d, {
            "get": function() {
                return this.Fields[d]
            },
            "set": function(a) {
                this.Fields[d] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(i.prototype, e, {
            "get": function() {
                return this.Fields[e]
            },
            "set": function(a) {
                this.Fields[e] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(i.prototype, f, {
            "get": function() {
                return this.Fields[f]
            },
            "set": function(a) {
                this.Fields[f] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(i.prototype, g, {
            "get": function() {
                return this.Fields[g]
            },
            "set": function(a) {
                this.Fields[g] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(i.prototype, h, {
            "get": function() {
                return this.Fields[h]
            },
            "set": function(a) {
                this.Fields[h] = a
            },
            enumerable: a,
            configurable: a
        });
        i.prototype.SerializeFields = function() {
            var a = this;
            a.SetSerializedField(b, a.CorrelationId);
            a.SetSerializedField(c, a.SessionId);
            a.SetSerializedField(d, a.FocusTime);
            a.SetSerializedField(e, a.AppSizeFinalWidth);
            a.SetSerializedField(f, a.AppSizeFinalHeight);
            a.SetSerializedField(g, a.OpenTime);
            a.SetSerializedField(h, a.CloseMethod)
        }
        ;
        return i
    }(f);
    g.AppClosedUsageData = l;
    var m = function(j) {
        var h = "ErrorType"
          , g = "Parameters"
          , f = "APIID"
          , d = "APIType";
        __extends(i, j);
        function i() {
            return j.call(this, "APIUsage") || this
        }
        Object.defineProperty(i.prototype, b, {
            "get": function() {
                return this.Fields[b]
            },
            "set": function(a) {
                this.Fields[b] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(i.prototype, c, {
            "get": function() {
                return this.Fields[c]
            },
            "set": function(a) {
                this.Fields[c] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(i.prototype, d, {
            "get": function() {
                return this.Fields[d]
            },
            "set": function(a) {
                this.Fields[d] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(i.prototype, f, {
            "get": function() {
                return this.Fields[f]
            },
            "set": function(a) {
                this.Fields[f] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(i.prototype, g, {
            "get": function() {
                return this.Fields[g]
            },
            "set": function(a) {
                this.Fields[g] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(i.prototype, e, {
            "get": function() {
                return this.Fields[e]
            },
            "set": function(a) {
                this.Fields[e] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(i.prototype, h, {
            "get": function() {
                return this.Fields[h]
            },
            "set": function(a) {
                this.Fields[h] = a
            },
            enumerable: a,
            configurable: a
        });
        i.prototype.SerializeFields = function() {
            var a = this;
            a.SetSerializedField(b, a.CorrelationId);
            a.SetSerializedField(c, a.SessionId);
            a.SetSerializedField(d, a.APIType);
            a.SetSerializedField(f, a.APIID);
            a.SetSerializedField(g, a.Parameters);
            a.SetSerializedField(e, a.ResponseTime);
            a.SetSerializedField(h, a.ErrorType)
        }
        ;
        return i
    }(f);
    g.APIUsageUsageData = m;
    var h = function(g) {
        var e = "SuccessCode";
        __extends(f, g);
        function f() {
            return g.call(this, "AppInitialization") || this
        }
        Object.defineProperty(f.prototype, b, {
            "get": function() {
                return this.Fields[b]
            },
            "set": function(a) {
                this.Fields[b] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(f.prototype, c, {
            "get": function() {
                return this.Fields[c]
            },
            "set": function(a) {
                this.Fields[c] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(f.prototype, e, {
            "get": function() {
                return this.Fields[e]
            },
            "set": function(a) {
                this.Fields[e] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(f.prototype, d, {
            "get": function() {
                return this.Fields[d]
            },
            "set": function(a) {
                this.Fields[d] = a
            },
            enumerable: a,
            configurable: a
        });
        f.prototype.SerializeFields = function() {
            var a = this;
            a.SetSerializedField(b, a.CorrelationId);
            a.SetSerializedField(c, a.SessionId);
            a.SetSerializedField(e, a.SuccessCode);
            a.SetSerializedField(d, a.Message)
        }
        ;
        return f
    }(f);
    g.AppInitializationUsageData = h;
    var j = function(i) {
        var g = "knownHostIndex"
          , f = "isLocalStorageAvailable"
          , e = "hostPlatform"
          , d = "hostType"
          , c = "instanceId"
          , b = "isWacKnownHost";
        __extends(h, i);
        function h() {
            return i.call(this, "CheckWACHost") || this
        }
        Object.defineProperty(h.prototype, b, {
            "get": function() {
                return this.Fields[b]
            },
            "set": function(a) {
                this.Fields[b] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(h.prototype, c, {
            "get": function() {
                return this.Fields[c]
            },
            "set": function(a) {
                this.Fields[c] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(h.prototype, d, {
            "get": function() {
                return this.Fields[d]
            },
            "set": function(a) {
                this.Fields[d] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(h.prototype, e, {
            "get": function() {
                return this.Fields[e]
            },
            "set": function(a) {
                this.Fields[e] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(h.prototype, f, {
            "get": function() {
                return this.Fields[f]
            },
            "set": function(a) {
                this.Fields[f] = a
            },
            enumerable: a,
            configurable: a
        });
        Object.defineProperty(h.prototype, g, {
            "get": function() {
                return this.Fields[g]
            },
            "set": function(a) {
                this.Fields[g] = a
            },
            enumerable: a,
            configurable: a
        });
        h.prototype.SerializeFields = function() {
            var a = this;
            a.SetSerializedField(b, a.isWacKnownHost);
            a.SetSerializedField(c, a.instanceId);
            a.SetSerializedField(d, a.hostType);
            a.SetSerializedField(e, a.hostPlatform);
            a.SetSerializedField(f, a.isLocalStorageAvailable);
            a.SetSerializedField(g, a.knownHostIndex)
        }
        ;
        return h
    }(f);
    g.CheckWACHostUsageData = j
}
)(OSFLog || (OSFLog = {}));
var Logger;
(function(a) {
    "use strict";
    var e;
    (function(a) {
        a[a["info"] = 0] = "info";
        a[a["warning"] = 1] = "warning";
        a[a["error"] = 2] = "error"
    }
    )(e = a.TraceLevel || (a.TraceLevel = {}));
    var f;
    (function(a) {
        a[a["none"] = 0] = "none";
        a[a["flush"] = 1] = "flush"
    }
    )(f = a.SendFlag || (a.SendFlag = {}));
    function b() {}
    a.allowUploadingData = b;
    function g() {}
    a.sendLog = g;
    function c() {
        try {
            return new d
        } catch (a) {
            return null
        }
    }
    var d = function() {
        function a() {}
        a.prototype.writeLog = function() {}
        ;
        a.prototype.loadProxyFrame = function() {}
        ;
        return a
    }();
    if (!OSF.Logger)
        OSF.Logger = a;
    a.ulsEndpoint = c()
}
)(Logger || (Logger = {}));
var OSFAriaLogger;
(function(w) {
    var e = "undefined"
      , f = "hostPlatform"
      , h = "hostType"
      , j = "ResponseTime"
      , g = "double"
      , a = false
      , d = "int64"
      , b = "string"
      , c = true
      , l = {
        name: "AppActivated",
        enabled: c,
        critical: c,
        points: [{
            name: "Browser",
            type: b
        }, {
            name: "Message",
            type: b
        }, {
            name: "Host",
            type: b
        }, {
            name: "AppSizeWidth",
            type: d
        }, {
            name: "AppSizeHeight",
            type: d
        }, {
            name: "IsFromWacAutomation",
            type: b
        }, {
            name: "IsMOS",
            type: d
        }, {
            name: "LaunchSource",
            type: b
        }, {
            name: "LaunchReason",
            type: b
        }]
    }
      , n = {
        name: "ScriptLoad",
        enabled: c,
        critical: a,
        points: [{
            name: "ScriptId",
            type: b
        }, {
            name: "StartTime",
            type: g
        }, {
            name: j,
            type: g
        }]
    }
      , v = o()
      , s = {
        name: "APIUsage",
        enabled: v,
        critical: a,
        points: [{
            name: "APIType",
            type: b
        }, {
            name: "APIID",
            type: d
        }, {
            name: "Parameters",
            type: b
        }, {
            name: j,
            type: d
        }, {
            name: "ErrorType",
            type: d
        }]
    }
      , k = {
        name: "AppInitialization",
        enabled: c,
        critical: a,
        points: [{
            name: "SuccessCode",
            type: d
        }, {
            name: "Message",
            type: b
        }]
    }
      , p = {
        name: "AppClosed",
        enabled: c,
        critical: a,
        points: [{
            name: "FocusTime",
            type: d
        }, {
            name: "AppSizeFinalWidth",
            type: d
        }, {
            name: "AppSizeFinalHeight",
            type: d
        }, {
            name: "OpenTime",
            type: d
        }]
    }
      , m = {
        name: "CheckWACHost",
        enabled: c,
        critical: a,
        points: [{
            name: "isWacKnownHost",
            type: d
        }, {
            name: "knownHostIndex",
            type: d
        }, {
            name: "solutionId",
            type: b
        }, {
            name: h,
            type: b
        }, {
            name: f,
            type: b
        }, {
            name: "correlationId",
            type: b
        }, {
            name: "isLocalStorageAvailable",
            type: "boolean"
        }]
    }
      , u = [l, n, s, k, p, m];
    function t(a, e) {
        var f = e.rename === undefined ? e.name : e.rename
          , h = e.type
          , c = undefined;
        switch (h) {
        case b:
            c = oteljs.makeStringDataField(f, a);
            break;
        case g:
            if (typeof a === b)
                a = parseFloat(a);
            c = oteljs.makeDoubleDataField(f, a);
            break;
        case d:
            if (typeof a === b)
                a = parseInt(a);
            c = oteljs.makeInt64DataField(f, a);
            break;
        case "boolean":
            if (typeof a === b)
                a = a === "true";
            c = oteljs.makeBooleanDataField(f, a)
        }
        return c
    }
    function i(d) {
        for (var a = 0, b = u; a < b.length; a++) {
            var c = b[a];
            if (c.name === d)
                return c
        }
        return undefined
    }
    function x(c) {
        var b = i(c);
        if (b === undefined)
            return a;
        return b.enabled
    }
    function o() {
        if (!OSF._OfficeAppFactory || !OSF._OfficeAppFactory.getHostInfo)
            return a;
        var b = OSF._OfficeAppFactory.getHostInfo();
        if (!b)
            return a;
        switch (b[h]) {
        case "outlook":
            switch (b[f]) {
            case "mac":
            case "web":
            case "ios":
            case "android":
                return c;
            default:
                return a
            }
        default:
            return a
        }
    }
    function q(e, l) {
        var a = i(e);
        if (a === undefined)
            return undefined;
        for (var d = [], c = 0, j = a.points; c < j.length; c++) {
            var g = j[c]
              , n = g.name
              , h = l[n];
            if (h === undefined)
                continue;
            var f = t(h, g);
            f !== undefined && d.push(f)
        }
        var b = {
            dataCategories: oteljs.DataCategories.ProductServiceUsage
        };
        if (a.critical)
            b.samplingPolicy = oteljs.SamplingPolicy.CriticalBusinessImpact;
        b.diagnosticLevel = oteljs.DiagnosticLevel.NecessaryServiceDataEvent;
        var k = "Office.Extensibility.OfficeJs." + e + "X"
          , m = {
            eventName: k,
            dataFields: d,
            eventFlags: b
        };
        return m
    }
    function r(a, b) {
        if (x(a))
            typeof OTel !== e && OTel.OTelLogger.onTelemetryLoaded(function() {
                var c = q(a, b);
                if (c === undefined)
                    return;
                Microsoft.Office.WebExtension.sendTelemetryEvent(c)
            })
    }
    var y = function() {
        function b() {}
        b.prototype.getAriaCDNLocation = function() {
            return OSF._OfficeAppFactory.getLoadScriptHelper().getOfficeJsBasePath() + "ariatelemetry/aria-web-telemetry.js"
        }
        ;
        b.getInstance = function() {
            if (b.AriaLoggerObj === undefined)
                b.AriaLoggerObj = new b;
            return b.AriaLoggerObj
        }
        ;
        b.prototype.isIUsageData = function(a) {
            return a["Fields"] !== undefined
        }
        ;
        b.prototype.shouldSendDirectToAria = function(f, h) {
            var k = 10, i = [16, 0, 11601], j = [16, 28], d;
            if (!f)
                return a;
            else if (f.toLowerCase() === "win32")
                d = i;
            else if (f.toLowerCase() === "mac")
                d = j;
            else
                return c;
            if (!h)
                return a;
            for (var g = h.split("."), b = 0; b < d.length && b < g.length; b++) {
                var e = parseInt(g[b], k);
                if (isNaN(e))
                    return a;
                if (e < d[b])
                    return c;
                if (e > d[b])
                    return a
            }
            return a
        }
        ;
        b.prototype.isDirectToAriaEnabled = function() {
            var a = this;
            if (a.EnableDirectToAria === undefined || a.EnableDirectToAria === null) {
                var c = void 0
                  , b = void 0;
                if (OSF._OfficeAppFactory && OSF._OfficeAppFactory.getHostInfo)
                    c = OSF._OfficeAppFactory.getHostInfo()[f];
                if (window.external && typeof window.external.GetContext !== e && typeof window.external.GetContext().GetHostFullVersion !== e)
                    b = window.external.GetContext().GetHostFullVersion();
                a.EnableDirectToAria = a.shouldSendDirectToAria(c, b)
            }
            return a.EnableDirectToAria
        }
        ;
        b.prototype.sendTelemetry = function(c, a) {
            var e = 1e3
              , d = b.EnableSendingTelemetryWithLegacyAria && this.isDirectToAriaEnabled();
            d && OSF.OUtil.loadScript(this.getAriaCDNLocation(), function() {
                try {
                    if (!this.ALogger) {
                        var e = "db334b301e7b474db5e0f02f07c51a47-a1b5bc36-1bbe-482f-a64a-c2d9cb606706-7439";
                        this.ALogger = AWTLogManager.initialize(e)
                    }
                    var b = new AWTEventProperties;
                    b.setName("Office.Extensibility.OfficeJS." + c);
                    for (var d in a)
                        d.toLowerCase() !== "table" && b.setProperty(d, a[d]);
                    var f = new Date;
                    b.setProperty("Date", f.toISOString());
                    this.ALogger.logEvent(b)
                } catch (g) {}
            }, e, OSF.TrustedTypesPolicy);
            b.EnableSendingTelemetryWithOTel && r(c, a)
        }
        ;
        b.prototype.logData = function(a) {
            if (this.isIUsageData(a))
                this.sendTelemetry(a["Table"], a["Fields"]);
            else
                this.sendTelemetry(a["Table"], a)
        }
        ;
        b.EnableSendingTelemetryWithOTel = c;
        b.EnableSendingTelemetryWithLegacyAria = a;
        return b
    }();
    w.AriaLogger = y
}
)(OSFAriaLogger || (OSFAriaLogger = {}));
var OSFAppTelemetry;
(function(d) {
    var e = false
      , l = "Microsoft.Office.SharedOnline.ChangeGate.OfficeVSO_10045620_CopilotAgentTelSettings"
      , b = null
      , f = true
      , c = "";
    "use strict";
    var a, h = OSF.OUtil.Guid.generateNewGuid(), k = c, A = new RegExp("^https?://store\\.office(ppe|-int)?\\.com/","i"), s = "PRIVATE";
    d.enableTelemetry = f;
    var x = {
        outlook: {
            mac: [-10, 4, 10, 12, 37, 38],
            web: [-10, 4, 10, 12, 37, 38],
            ios: [1e4],
            android: [1e4]
        }
    }
      , t = function() {
        function a() {}
        return a
    }();
    d.AppInfo = t;
    var j = function() {
        function a(b, a) {
            this.name = b;
            this.handler = a
        }
        return a
    }()
      , o = function() {
        function a() {
            this.clientIDKey = "Office API client";
            this.logIdSetKey = "Office App Log Id Set"
        }
        a.prototype.getClientId = function() {
            var b = this
              , a = b.getValue(b.clientIDKey);
            if (!a || a.length <= 0 || a.length > 40) {
                a = OSF.OUtil.Guid.generateNewGuid();
                b.setValue(b.clientIDKey, a)
            }
            return a
        }
        ;
        a.prototype.saveLog = function(d, e) {
            var b = this
              , a = b.getValue(b.logIdSetKey);
            a = (a && a.length > 0 ? a + ";" : c) + d;
            b.setValue(b.logIdSetKey, a);
            b.setValue(d, e)
        }
        ;
        a.prototype.enumerateLog = function(c, e) {
            var a = this
              , d = a.getValue(a.logIdSetKey);
            if (d) {
                var f = d.split(";");
                for (var h in f) {
                    var b = f[h]
                      , g = a.getValue(b);
                    if (g) {
                        c && c(b, g);
                        e && a.remove(b)
                    }
                }
                e && a.remove(a.logIdSetKey)
            }
        }
        ;
        a.prototype.getValue = function(d) {
            var a = OSF.OUtil.getLocalStorage()
              , b = c;
            if (a)
                b = a.getItem(d);
            return b
        }
        ;
        a.prototype.setValue = function(c, b) {
            var a = OSF.OUtil.getLocalStorage();
            a && a.setItem(c, b)
        }
        ;
        a.prototype.remove = function(b) {
            var a = OSF.OUtil.getLocalStorage();
            if (a)
                try {
                    a.removeItem(b)
                } catch (c) {}
        }
        ;
        return a
    }()
      , i = function() {
        function a() {}
        a.prototype.LogData = function(a) {
            if (!d.enableTelemetry)
                return;
            try {
                OSFAriaLogger.AriaLogger.getInstance().logData(a)
            } catch (b) {}
        }
        ;
        a.prototype.LogRawData = function(a) {
            if (!d.enableTelemetry)
                return;
            try {
                OSFAriaLogger.AriaLogger.getInstance().logData(JSON.parse(a))
            } catch (b) {}
        }
        ;
        return a
    }();
    function g(a) {
        if (a)
            a = a.replace(/[{}]/g, c).toLowerCase();
        return a || c
    }
    function r(a) {
        try {
            return JSON.parse(a)
        } catch (b) {
            return a
        }
    }
    function J(i) {
        var q = "LaunchReason"
          , m = "LaunchSource";
        if (!d.enableTelemetry)
            return;
        if (a)
            return;
        a = new t;
        if (i.get_hostFullVersion())
            a.hostVersion = i.get_hostFullVersion();
        else
            a.hostVersion = i.get_appVersion();
        a.appId = n() ? i.get_id() : s;
        a.marketplaceType = i._marketplaceType;
        a.browser = window.navigator.userAgent;
        a.correlationId = g(i.get_correlationId());
        a.clientId = (new o).getClientId();
        a.appInstanceId = i.get_appInstanceId();
        if (a.appInstanceId) {
            a.appInstanceId = g(a.appInstanceId);
            a.appInstanceId = p(i.get_id(), a.appInstanceId)
        }
        a.message = i.get_hostCustomMessage();
        a.officeJSVersion = OSF.ConstantNames.FileVersion;
        a.hostJSVersion = "16.0.19009.20000";
        if (i._wacHostEnvironment)
            a.wacHostEnvironment = i._wacHostEnvironment;
        if (i._isFromWacAutomation !== undefined && i._isFromWacAutomation !== b)
            a.isFromWacAutomation = i._isFromWacAutomation.toString().toLowerCase();
        var w = i.get_docUrl();
        a.docUrl = A.test(w) ? w : c;
        var v = location.href;
        if (v)
            v = v.split("?")[0].split("#")[0];
        a.isMos = u();
        if (OSF.OUtil.isChangeGateEnabled(l)) {
            var k = i.get_settings();
            if (k && k[m] && k[q]) {
                a.launchSource = r(k[m]);
                a.launchReason = r(k[q])
            }
        }
        a.appURL = c;
        (function(j, d) {
            var a, h, e;
            d.assetId = c;
            d.userId = c;
            try {
                a = decodeURIComponent(j);
                h = new DOMParser;
                if (OSF.TrustedTypesPolicy && window.trustedTypes && window.trustedTypes.createPolicy) {
                    var i = window.trustedTypes.createPolicy("officejs-domparser", {
                        createHTML: function(a) {
                            return a
                        }
                    });
                    a = i.createHTML(a)
                }
                e = h.parseFromString(a, "text/xml");
                var f = e.getElementsByTagName("t")[0].attributes.getNamedItem("cid")
                  , g = e.getElementsByTagName("t")[0].attributes.getNamedItem("oid");
                if (f && f.nodeValue)
                    d.userId = f.nodeValue;
                else if (g && g.nodeValue)
                    d.userId = g.nodeValue;
                d.assetId = e.getElementsByTagName("t")[0].attributes.getNamedItem("aid").nodeValue
            } catch (k) {} finally {
                a = b;
                e = b;
                h = b
            }
        }
        )(i.get_eToken(), a);
        a.sessionId = h;
        typeof OTel !== "undefined" && OTel.OTelLogger.initialize(a);
        (function() {
            var m = new Date
              , c = b
              , i = 0
              , l = e
              , g = function() {
                if (document.hasFocus()) {
                    if (c == b)
                        c = new Date
                } else if (c) {
                    i += Math.abs((new Date).getTime() - c.getTime());
                    c = b
                }
            }
              , a = [];
            a.push(new j("focus",g));
            a.push(new j("blur",g));
            a.push(new j("focusout",g));
            a.push(new j("focusin",g));
            var k = function() {
                for (var e = 0; e < a.length; e++)
                    OSF.OUtil.removeEventListener(window, a[e].name, a[e].handler);
                a.length = 0;
                if (!l) {
                    if (document.hasFocus() && c) {
                        i += Math.abs((new Date).getTime() - c.getTime());
                        c = b
                    }
                    d.onAppClosed(Math.abs((new Date).getTime() - m.getTime()), i);
                    l = f
                }
            };
            a.push(new j("beforeunload",k));
            a.push(new j("unload",k));
            for (var h = 0; h < a.length; h++)
                OSF.OUtil.addEventListener(window, a[h].name, a[h].handler);
            g()
        }
        )();
        d.onAppActivated()
    }
    d.initialize = J;
    function B() {
        if (!a)
            return;
        (new o).enumerateLog(function(b, a) {
            return (new i).LogRawData(a)
        }, f);
        var d = new OSFLog.AppActivatedUsageData;
        d.SessionId = h;
        d.AppId = a.appId;
        d.AssetId = a.assetId;
        d.AppURL = c;
        d.UserId = c;
        d.ClientId = a.clientId;
        d.Browser = a.browser;
        d.HostVersion = a.hostVersion;
        d.CorrelationId = g(a.correlationId);
        d.AppSizeWidth = window.innerWidth;
        d.AppSizeHeight = window.innerHeight;
        d.AppInstanceId = a.appInstanceId;
        d.Message = a.message;
        d.DocUrl = a.docUrl;
        d.OfficeJSVersion = a.officeJSVersion;
        d.HostJSVersion = a.hostJSVersion;
        if (a.wacHostEnvironment)
            d.WacHostEnvironment = a.wacHostEnvironment;
        if (a.isFromWacAutomation !== undefined && a.isFromWacAutomation !== b)
            d.IsFromWacAutomation = a.isFromWacAutomation;
        d.IsMOS = a.isMos ? 1 : 0;
        if (OSF.OUtil.isChangeGateEnabled(l))
            if (a.launchSource && a.launchReason) {
                d.LaunchSource = a.launchSource;
                d.LaunchReason = a.launchReason
            }
        (new i).LogData(d)
    }
    d.onAppActivated = B;
    function G(e, d, c, b) {
        var a = new OSFLog.ScriptLoadUsageData;
        a.CorrelationId = g(b);
        a.SessionId = h;
        a.ScriptId = e;
        a.StartTime = d;
        a.ResponseTime = c;
        (new i).LogData(a)
    }
    d.onScriptDone = G;
    function K(c, d, f, e, j) {
        if (!a)
            return;
        if (!w(d, c))
            return;
        var b = new OSFLog.APIUsageUsageData;
        b.CorrelationId = g(k);
        b.SessionId = h;
        b.APIType = c;
        b.APIID = d;
        b.Parameters = f;
        b.ResponseTime = e;
        b.ErrorType = j;
        (new i).LogData(b)
    }
    d.onCallDone = K;
    function F(h, d, f, g) {
        var a = b;
        if (d)
            if (typeof d == "number")
                a = String(d);
            else if (typeof d === "object")
                for (var e in d) {
                    if (a !== b)
                        a += ",";
                    else
                        a = c;
                    if (typeof d[e] == "number")
                        a += String(d[e])
                }
            else
                a = c;
        OSF.AppTelemetry.onCallDone("method", h, a, f, g)
    }
    d.onMethodDone = F;
    function D(b, a) {
        OSF.AppTelemetry.onCallDone("property", -1, b, a)
    }
    d.onPropertyDone = D;
    function C(c, d, f, g, e, b) {
        var a = new OSFLog.CheckWACHostUsageData;
        a.isWacKnownHost = c;
        a.knownHostIndex = d;
        a.instanceId = f;
        a.hostType = g;
        a.hostPlatform = e;
        a.isLocalStorageAvailable = b;
        (new i).LogData(a)
    }
    d.onCheckWACHost = C;
    function I(c, a) {
        OSF.AppTelemetry.onCallDone("event", c, b, 0, a)
    }
    d.onEventDone = I;
    function E(d, e, a, c) {
        OSF.AppTelemetry.onCallDone(d ? "registerevent" : "unregisterevent", e, b, a, c)
    }
    d.onRegisterDone = E;
    function H(d, c) {
        if (!a)
            return;
        var b = new OSFLog.AppClosedUsageData;
        b.CorrelationId = g(k);
        b.SessionId = h;
        b.FocusTime = c;
        b.OpenTime = d;
        b.AppSizeFinalWidth = window.innerWidth;
        b.AppSizeFinalHeight = window.innerHeight;
        (new o).saveLog(h, b.SerializeRow())
    }
    d.onAppClosed = H;
    function v(a) {
        k = g(a)
    }
    d.setOsfControlAppCorrelationId = v;
    function m(b, c) {
        var a = new OSFLog.AppInitializationUsageData;
        a.CorrelationId = g(k);
        a.SessionId = h;
        a.SuccessCode = b ? 1 : 0;
        a.Message = c;
        (new i).LogData(a)
    }
    d.doAppInitializationLogging = m;
    function y(a) {
        m(e, a)
    }
    d.logAppCommonMessage = y;
    function z(a) {
        m(f, a)
    }
    d.logAppException = z;
    function w(f, d) {
        if (!OSF._OfficeAppFactory || !OSF._OfficeAppFactory.getHostInfo)
            return e;
        var a = OSF._OfficeAppFactory.getHostInfo();
        if (!a)
            return e;
        if (d === "method") {
            var c = x[a["hostType"]];
            if (c) {
                var b = c[a["hostPlatform"]];
                return b && b.indexOf(f) !== -1
            }
        }
        return e
    }
    function n() {
        var b = (OSF._OfficeAppFactory.getHostInfo().flags & OSF.HostInfoFlags.PublicAddin) != 0;
        if (b)
            return b;
        if (!a)
            return e;
        var c = OSF._OfficeAppFactory.getHostInfo().hostPlatform
          , d = a.hostVersion;
        return q(c, d)
    }
    d.canSendAddinId = n;
    function p(b, a) {
        if (!n() && a === b)
            return s;
        return a
    }
    d.getCompliantAppInstanceId = p;
    function q(d, j) {
        var c = e
          , i = /^(\d+)\.(\d+)\.(\d+)\.(\d+)$/
          , a = i.exec(j);
        if (a) {
            var b = parseInt(a[1])
              , h = parseInt(a[2])
              , g = parseInt(a[3]);
            if (d == "win32") {
                if (b < 16 || b == 16 && g < 14225)
                    c = f
            } else if (d == "mac")
                if (b < 16 || b == 16 && (h < 52 || h == 52 && g < 808))
                    c = f
        }
        return c
    }
    d._isComplianceExceptedHost = q;
    function u() {
        return (OSF._OfficeAppFactory.getHostInfo().flags & OSF.HostInfoFlags.IsMos) != 0
    }
    d.isMos = u;
    OSF.AppTelemetry = d
}
)(OSFAppTelemetry || (OSFAppTelemetry = {}));
var OSFPerfUtil;
(function(c) {
    function a(b, a) {
        a = a + "_Resource";
        if (oteljs !== undefined)
            return [oteljs.makeDoubleDataField(a + "_responseEnd", b.responseEnd), oteljs.makeDoubleDataField(a + "_responseStart", b.responseStart), oteljs.makeDoubleDataField(a + "_startTime", b.startTime), oteljs.makeDoubleDataField(a + "_transferSize", b.transferSize)]
    }
    function b() {
        var b = "undefined";
        if (typeof OTel !== b && OSF.AppTelemetry.enableTelemetry && typeof OSFPerformance !== b && typeof performance != b && performance.getEntriesByType) {
            var d, c, e = performance.getEntriesByType("resource");
            e.forEach(function(b) {
                var a = b.name.toLowerCase();
                if (OSF.OUtil.stringEndsWith(a, OSFPerformance.hostSpecificFileName))
                    d = b;
                else if (OSF.OUtil.stringEndsWith(a, OSF.ConstantNames.OfficeDebugJS) || OSF.OUtil.stringEndsWith(a, OSF.ConstantNames.OfficeJS))
                    c = b
            });
            OTel.OTelLogger.onTelemetryLoaded(function() {
                var b = a(d, "HostJs");
                b = b.concat(a(c, "OfficeJs"));
                b = b.concat([oteljs.makeDoubleDataField("officeExecuteStartDate", OSFPerformance.officeExecuteStartDate), oteljs.makeDoubleDataField("officeExecuteStart", OSFPerformance.officeExecuteStart), oteljs.makeDoubleDataField("officeExecuteEnd", OSFPerformance.officeExecuteEnd), oteljs.makeDoubleDataField("hostInitializationStart", OSFPerformance.hostInitializationStart), oteljs.makeDoubleDataField("hostInitializationEnd", OSFPerformance.hostInitializationEnd), oteljs.makeDoubleDataField("totalJSHeapSize", OSFPerformance.totalJSHeapSize), oteljs.makeDoubleDataField("usedJSHeapSize", OSFPerformance.usedJSHeapSize), oteljs.makeDoubleDataField("jsHeapSizeLimit", OSFPerformance.jsHeapSizeLimit), oteljs.makeDoubleDataField("getAppContextStart", OSFPerformance.getAppContextStart), oteljs.makeDoubleDataField("getAppContextEnd", OSFPerformance.getAppContextEnd), oteljs.makeDoubleDataField("createOMEnd", OSFPerformance.createOMEnd), oteljs.makeDoubleDataField("officeOnReady", OSFPerformance.officeOnReady), oteljs.makeBooleanDataField("isSharedRuntime", (OSF._OfficeAppFactory.getHostInfo().flags & OSF.HostInfoFlags.SharedApp) !== 0)]);
                Microsoft.Office.WebExtension.sendTelemetryEvent({
                    eventName: "Office.Extensibility.OfficeJs.JSPerformanceTelemetryV06",
                    dataFields: b,
                    eventFlags: {
                        dataCategories: oteljs.DataCategories.ProductServiceUsage,
                        diagnosticLevel: oteljs.DiagnosticLevel.NecessaryServiceDataEvent
                    }
                })
            })
        }
    }
    c.sendPerformanceTelemetry = b
}
)(OSFPerfUtil || (OSFPerfUtil = {}));
(function(a) {
    var b;
    (function(b) {
        var e = function() {
            var i = "object"
              , e = true
              , g = false
              , j = "function"
              , h = "string"
              , d = null;
            function f() {
                var a = this
                  , b = a;
                a._pseudoDocument = d;
                a._eventDispatch = d;
                a._useAssociatedActionsOnly = d;
                a._processAppCommandInvocation = function(a) {
                    var c = b._verifyManifestCallback(a.callbackName);
                    if (c.errorCode != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                        b._invokeAppCommandCompletedMethod(a.appCommandId, c.errorCode, "");
                        return
                    }
                    var d = b._constructEventObjectForCallback(a);
                    if (d)
                        window.setTimeout(function() {
                            c.callback(d)
                        }, 0);
                    else
                        b._invokeAppCommandCompletedMethod(a.appCommandId, OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError, "")
                }
            }
            f.initializeOsfDda = function() {
                OSF.DDA.AsyncMethodNames.addNames({
                    AppCommandInvocationCompletedAsync: "appCommandInvocationCompletedAsync"
                });
                OSF.DDA.AsyncMethodCalls.define({
                    method: OSF.DDA.AsyncMethodNames.AppCommandInvocationCompletedAsync,
                    requiredArguments: [{
                        name: Microsoft.Office.WebExtension.Parameters.Id,
                        types: [h]
                    }, {
                        name: Microsoft.Office.WebExtension.Parameters.Status,
                        types: ["number"]
                    }, {
                        name: Microsoft.Office.WebExtension.Parameters.AppCommandInvocationCompletedData,
                        types: [h]
                    }]
                });
                OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
                    AppCommandInvokedEvent: "AppCommandInvokedEvent"
                });
                OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
                    AppCommandInvoked: "appCommandInvoked"
                });
                OSF.OUtil.setNamespace("AppCommand", OSF.DDA);
                OSF.DDA.AppCommand.AppCommandInvokedEventArgs = a.AppCommand.AppCommandInvokedEventArgs
            }
            ;
            f.prototype.initializeAndChangeOnce = function(c) {
                var a = this;
                b.registerDdaFacade();
                a._pseudoDocument = {};
                OSF.DDA.DispIdHost.addAsyncMethods(a._pseudoDocument, [OSF.DDA.AsyncMethodNames.AppCommandInvocationCompletedAsync]);
                a._eventDispatch = new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.AppCommandInvoked]);
                var d = function(a) {
                    if (c)
                        if (a.status == "succeeded")
                            c(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
                        else
                            c(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError)
                };
                OSF.DDA.DispIdHost.addEventSupport(a._pseudoDocument, a._eventDispatch);
                a._pseudoDocument.addHandlerAsync(Microsoft.Office.WebExtension.EventType.AppCommandInvoked, a._processAppCommandInvocation, d)
            }
            ;
            f.prototype._verifyManifestCallback = function(a) {
                var b = {
                    callback: d,
                    errorCode: OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCallback
                };
                a = a.trim();
                try {
                    var c = this._getCallbackFunc(a);
                    if (typeof c != j)
                        return b
                } catch (e) {
                    return b
                }
                return {
                    callback: c,
                    errorCode: OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess
                }
            }
            ;
            f.prototype._getUseAssociatedActionsOnly = function() {
                var a = this;
                if (a._useAssociatedActionsOnly == d) {
                    a._useAssociatedActionsOnly = g;
                    try {
                        if (window["useAssociatedActionsOnly"] === e)
                            a._useAssociatedActionsOnly = e;
                        else
                            a._useAssociatedActionsOnly = OSF._OfficeAppFactory.getLoadScriptHelper().getUseAssociatedActionsOnlyDefined()
                    } catch (b) {}
                }
                return a._useAssociatedActionsOnly
            }
            ;
            f.prototype._getCallbackFuncFromWindow = function(f) {
                for (var a = f.split("."), b = window, c = 0; c < a.length - 1; c++)
                    if (b[a[c]] && (typeof b[a[c]] == i || typeof b[a[c]] == j))
                        b = b[a[c]];
                    else
                        return d;
                var e = b[a[a.length - 1]];
                return e
            }
            ;
            f.prototype._getCallbackFuncFromActionAssociateTable = function(b) {
                var a = b.toUpperCase();
                return Office.actions._association.mappings[a]
            }
            ;
            f.prototype._getCallbackFunc = function(i) {
                var a = d
                  , c = g
                  , h = g
                  , b = this._getUseAssociatedActionsOnly();
                a = this._getCallbackFuncFromActionAssociateTable(i);
                if (a)
                    c = e;
                else if (!b) {
                    a = this._getCallbackFuncFromWindow(i);
                    if (a)
                        h = e
                }
                if (!f.isTelemetrySubmitted) {
                    f.isTelemetrySubmitted = e;
                    try {
                        if (OTel && oteljs && Microsoft.Office.WebExtension.sendTelemetryEvent) {
                            var j = OTel.OTelLogger.getHost() == "Outlook" ? .1 : .2;
                            Math.random() < j && OTel.OTelLogger.onTelemetryLoaded(function() {
                                var a = [oteljs.makeBooleanDataField("UseAction", b === e), oteljs.makeBooleanDataField("UseAssociateTable", c), oteljs.makeBooleanDataField("UseGlobal", h)];
                                Microsoft.Office.WebExtension.sendTelemetryEvent({
                                    eventName: "Office.Extensibility.OfficeJs.AppCommandDefinition",
                                    dataFields: a,
                                    eventFlags: {
                                        dataCategories: oteljs.DataCategories.ProductServiceUsage,
                                        diagnosticLevel: oteljs.DiagnosticLevel.NecessaryServiceDataEvent
                                    }
                                })
                            })
                        }
                    } catch (k) {}
                }
                return a
            }
            ;
            f.prototype._invokeAppCommandCompletedMethod = function(a, b, c) {
                this._pseudoDocument.appCommandInvocationCompletedAsync(a, b, c)
            }
            ;
            f.prototype._constructEventObjectForCallback = function(b) {
                var g = this
                  , a = new c;
                try {
                    var f = JSON.parse(b.eventObjStr);
                    this._translateEventObjectInternal(f, a);
                    Object.defineProperty(a, "completed", {
                        value: function(c) {
                            a.completedContext = c;
                            var d = JSON.stringify(a);
                            g._invokeAppCommandCompletedMethod(b.appCommandId, OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess, d)
                        },
                        enumerable: e
                    })
                } catch (h) {
                    a = d
                }
                return a
            }
            ;
            f.prototype._translateEventObjectInternal = function(f, c) {
                for (var a in f) {
                    if (!f.hasOwnProperty(a))
                        continue;
                    var b = f[a];
                    if (typeof b == i && b != d) {
                        OSF.OUtil.defineEnumerableProperty(c, a, {
                            value: {}
                        });
                        this._translateEventObjectInternal(b, c[a])
                    } else
                        Object.defineProperty(c, a, {
                            value: b,
                            enumerable: e,
                            writable: e
                        })
                }
            }
            ;
            f.prototype._constructObjectByTemplate = function(c, j) {
                var b = {};
                if (!c || !j)
                    return b;
                for (var a in c)
                    if (c.hasOwnProperty(a)) {
                        b[a] = d;
                        if (j[a] != d) {
                            var f = c[a]
                              , g = j[a]
                              , e = typeof g;
                            if (typeof f == i && f != d)
                                b[a] = this._constructObjectByTemplate(f, g);
                            else if (e == "number" || e == h || e == "boolean")
                                b[a] = g
                        }
                    }
                return b
            }
            ;
            f.instance = function() {
                if (f._instance == d)
                    f._instance = new f;
                return f._instance
            }
            ;
            f.isTelemetrySubmitted = g;
            f._instance = d;
            return f
        }();
        b.AppCommandManager = e;
        var d = function() {
            function a(b, c, d) {
                var a = this;
                a.type = Microsoft.Office.WebExtension.EventType.AppCommandInvoked;
                a.appCommandId = b;
                a.callbackName = c;
                a.eventObjStr = d
            }
            a.create = function(c) {
                return new a(c[b.AppCommandInvokedEventEnums.AppCommandId],c[b.AppCommandInvokedEventEnums.CallbackName],c[b.AppCommandInvokedEventEnums.EventObjStr])
            }
            ;
            return a
        }();
        b.AppCommandInvokedEventArgs = d;
        var c = function() {
            function a() {}
            return a
        }();
        b.AppCommandCallbackEventArgs = c;
        b.AppCommandInvokedEventEnums = {
            AppCommandId: "appCommandId",
            CallbackName: "callbackName",
            EventObjStr: "eventObjStr"
        }
    }
    )(b = a.AppCommand || (a.AppCommand = {}))
}
)(OfficeExt || (OfficeExt = {}));
OfficeExt.AppCommand.AppCommandManager.initializeOsfDda();
(function(a) {
    var b;
    (function(c) {
        function b() {
            if (OSF.DDA.SafeArray) {
                var b = OSF.DDA.SafeArray.Delegate.ParameterMap;
                b.define({
                    type: OSF.DDA.MethodDispId.dispidAppCommandInvocationCompletedMethod,
                    toHost: [{
                        name: Microsoft.Office.WebExtension.Parameters.Id,
                        value: 0
                    }, {
                        name: Microsoft.Office.WebExtension.Parameters.Status,
                        value: 1
                    }, {
                        name: Microsoft.Office.WebExtension.Parameters.AppCommandInvocationCompletedData,
                        value: 2
                    }]
                });
                b.define({
                    type: OSF.DDA.EventDispId.dispidAppCommandInvokedEvent,
                    fromHost: [{
                        name: OSF.DDA.EventDescriptors.AppCommandInvokedEvent,
                        value: b.self
                    }],
                    isComplexType: true
                });
                b.define({
                    type: OSF.DDA.EventDescriptors.AppCommandInvokedEvent,
                    fromHost: [{
                        name: a.AppCommand.AppCommandInvokedEventEnums.AppCommandId,
                        value: 0
                    }, {
                        name: a.AppCommand.AppCommandInvokedEventEnums.CallbackName,
                        value: 1
                    }, {
                        name: a.AppCommand.AppCommandInvokedEventEnums.EventObjStr,
                        value: 2
                    }],
                    isComplexType: true
                })
            }
        }
        c.registerDdaFacade = b
    }
    )(b = a.AppCommand || (a.AppCommand = {}))
}
)(OfficeExt || (OfficeExt = {}));
OSF.DDA.AsyncMethodNames.addNames({
    CloseContainerAsync: "closeContainer"
});
(function(b) {
    var a = function() {
        function a() {}
        return a
    }();
    b.Container = a
}
)(OfficeExt || (OfficeExt = {}));
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.CloseContainerAsync,
    requiredArguments: [],
    supportedOptions: [],
    privateStateCallbacks: []
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidCloseContainerMethod,
    fromHost: [],
    toHost: []
});
Microsoft.Office.WebExtension.AccountTypeFilter = {
    NoFilter: "noFilter",
    AAD: "aad",
    MSA: "msa"
};
OSF.DDA.AsyncMethodNames.addNames({
    GetAccessTokenAsync: "getAccessTokenAsync"
});
OSF.DDA.Auth = function() {}
;
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.GetAccessTokenAsync,
    requiredArguments: [],
    supportedOptions: [{
        name: Microsoft.Office.WebExtension.Parameters.ForceConsent,
        value: {
            types: ["boolean"],
            defaultValue: false
        }
    }, {
        name: Microsoft.Office.WebExtension.Parameters.ForceAddAccount,
        value: {
            types: ["boolean"],
            defaultValue: false
        }
    }, {
        name: Microsoft.Office.WebExtension.Parameters.AuthChallenge,
        value: {
            types: ["string"],
            defaultValue: ""
        }
    }, {
        name: Microsoft.Office.WebExtension.Parameters.AllowConsentPrompt,
        value: {
            types: ["boolean"],
            defaultValue: false
        }
    }, {
        name: Microsoft.Office.WebExtension.Parameters.ForMSGraphAccess,
        value: {
            types: ["boolean"],
            defaultValue: false
        }
    }, {
        name: Microsoft.Office.WebExtension.Parameters.AllowSignInPrompt,
        value: {
            types: ["boolean"],
            defaultValue: false
        }
    }, {
        name: Microsoft.Office.WebExtension.Parameters.EnableNewHosts,
        value: {
            types: ["number"],
            defaultValue: 0
        }
    }, {
        name: Microsoft.Office.WebExtension.Parameters.AccountTypeFilter,
        value: {
            "enum": Microsoft.Office.WebExtension.AccountTypeFilter,
            defaultValue: Microsoft.Office.WebExtension.AccountTypeFilter.NoFilter
        }
    }],
    checkCallArgs: function(c) {
        var b = true, a, e = OSF._OfficeAppFactory.getInitializationHelper()._appContext;
        if (e && e._wopiHostOriginForSingleSignOn) {
            var h = OSF.OUtil.Guid.generateNewGuid();
            window.parent.parent.postMessage('{"MessageId":"AddinTrustedOrigin","AddinTrustId":"' + h + '"}', e._wopiHostOriginForSingleSignOn);
            c[Microsoft.Office.WebExtension.Parameters.AddinTrustId] = h
        }
        if (window.Office.context.requirements.isSetSupported("JsonPayloadSSO")) {
            for (var g = (a = {},
            a[Microsoft.Office.WebExtension.Parameters.ForceConsent] = false,
            a[Microsoft.Office.WebExtension.Parameters.ForceAddAccount] = false,
            a[Microsoft.Office.WebExtension.Parameters.AuthChallenge] = b,
            a[Microsoft.Office.WebExtension.Parameters.AllowConsentPrompt] = b,
            a[Microsoft.Office.WebExtension.Parameters.ForMSGraphAccess] = b,
            a[Microsoft.Office.WebExtension.Parameters.AllowSignInPrompt] = b,
            a[Microsoft.Office.WebExtension.Parameters.EnableNewHosts] = b,
            a[Microsoft.Office.WebExtension.Parameters.AccountTypeFilter] = b,
            a), i = {}, f = 0, j = Object.keys(g); f < j.length; f++) {
                var d = j[f];
                if (g[d])
                    i[d] = c[d];
                delete c[d]
            }
            c[Microsoft.Office.WebExtension.Parameters.JsonPayload] = JSON.stringify(i)
        }
        return c
    },
    onSucceeded: function(a) {
        var b = a[Microsoft.Office.WebExtension.Parameters.Data];
        return b
    }
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidGetAccessTokenMethod,
    toHost: [{
        name: Microsoft.Office.WebExtension.Parameters.JsonPayload,
        value: 0
    }, {
        name: Microsoft.Office.WebExtension.Parameters.ForceConsent,
        value: 0
    }, {
        name: Microsoft.Office.WebExtension.Parameters.ForceAddAccount,
        value: 1
    }, {
        name: Microsoft.Office.WebExtension.Parameters.AuthChallenge,
        value: 2
    }, {
        name: Microsoft.Office.WebExtension.Parameters.AllowConsentPrompt,
        value: 3
    }, {
        name: Microsoft.Office.WebExtension.Parameters.ForMSGraphAccess,
        value: 4
    }, {
        name: Microsoft.Office.WebExtension.Parameters.AllowSignInPrompt,
        value: 5
    }],
    fromHost: [{
        name: Microsoft.Office.WebExtension.Parameters.Data,
        value: OSF.DDA.SafeArray.Delegate.ParameterMap.self
    }]
});
OSF.DDA.AsyncMethodNames.addNames({
    GetNestedAppAuthContextAsync: "getAuthContextAsync"
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.GetNestedAppAuthContextAsync,
    requiredArguments: [],
    supportedOptions: [],
    onSucceeded: function(d) {
        var a = d[Microsoft.Office.WebExtension.Parameters.JsonData]
          , f = a.userObjectId || ""
          , h = a.tenantId || ""
          , b = a.userPrincipalName || ""
          , e = a.authorityType || ""
          , c = a.authorityBaseUrl || ""
          , g = a.loginHint || b;
        return {
            userObjectId: f,
            tenantId: h,
            userPrincipalName: b,
            authorityType: e,
            authorityBaseUrl: c,
            loginHint: g
        }
    }
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidGetNestedAppAuthContextMethod,
    toHost: [],
    fromHost: [{
        name: Microsoft.Office.WebExtension.Parameters.JsonData,
        value: OSF.DDA.SafeArray.Delegate.ParameterMap.self
    }]
});
OSF.DDA.AsyncMethodNames.addNames({
    OpenBrowserWindow: "openBrowserWindow"
});
OSF.DDA.OpenBrowser = function() {}
;
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.OpenBrowserWindow,
    requiredArguments: [{
        name: Microsoft.Office.WebExtension.Parameters.Url,
        types: ["string"]
    }],
    supportedOptions: [{
        name: Microsoft.Office.WebExtension.Parameters.Reserved,
        value: {
            types: ["number"],
            defaultValue: 0
        }
    }],
    privateStateCallbacks: []
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidOpenBrowserWindow,
    toHost: [{
        name: Microsoft.Office.WebExtension.Parameters.Reserved,
        value: 0
    }, {
        name: Microsoft.Office.WebExtension.Parameters.Url,
        value: 1
    }]
});
OSF.DDA.AsyncMethodNames.addNames({
    ExecuteFeature: "executeFeatureAsync",
    QueryFeature: "queryFeatureAsync"
});
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
    FeatureProperties: "FeatureProperties",
    TcidEnabled: "TcidEnabled",
    TcidVisible: "TcidVisible"
});
OSF.DDA.ExecuteFeature = function() {}
;
OSF.DDA.QueryFeature = function() {}
;
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.ExecuteFeature,
    requiredArguments: [{
        name: Microsoft.Office.WebExtension.Parameters.Tcid,
        types: ["number"]
    }],
    privateStateCallbacks: []
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.QueryFeature,
    requiredArguments: [{
        name: Microsoft.Office.WebExtension.Parameters.Tcid,
        types: ["number"]
    }],
    privateStateCallbacks: []
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.PropertyDescriptors.FeatureProperties,
    fromHost: [{
        name: OSF.DDA.PropertyDescriptors.TcidEnabled,
        value: 0
    }, {
        name: OSF.DDA.PropertyDescriptors.TcidVisible,
        value: 1
    }],
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidExecuteFeature,
    toHost: [{
        name: Microsoft.Office.WebExtension.Parameters.Tcid,
        value: 0
    }]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidQueryFeature,
    fromHost: [{
        name: OSF.DDA.PropertyDescriptors.FeatureProperties,
        value: OSF.DDA.SafeArray.Delegate.ParameterMap.self
    }],
    toHost: [{
        name: Microsoft.Office.WebExtension.Parameters.Tcid,
        value: 0
    }]
});
OSF.DDA.AsyncMethodNames.addNames({
    ExecuteRichApiRequestAsync: "executeRichApiRequestAsync"
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync,
    requiredArguments: [{
        name: Microsoft.Office.WebExtension.Parameters.Data,
        types: ["object"]
    }],
    supportedOptions: []
});
OSF.OUtil.setNamespace("RichApi", OSF.DDA);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidExecuteRichApiRequestMethod,
    toHost: [{
        name: Microsoft.Office.WebExtension.Parameters.Data,
        value: 0
    }],
    fromHost: [{
        name: Microsoft.Office.WebExtension.Parameters.Data,
        value: OSF.DDA.SafeArray.Delegate.ParameterMap.self
    }]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    RichApiMessage: "richApiMessage"
});
OSF.DDA.RichApiMessageEventArgs = function(f, e) {
    var b = e[Microsoft.Office.WebExtension.Parameters.Data]
      , d = [];
    if (b)
        for (var c = 0; c < b.length; c++) {
            var a = b[c];
            if (a.toArray)
                a = a.toArray();
            d.push({
                messageCategory: a[0],
                messageType: a[1],
                targetId: a[2],
                message: a[3],
                id: a[4],
                isRemoteOverride: a[5]
            })
        }
    OSF.OUtil.defineEnumerableProperties(this, {
        type: {
            value: Microsoft.Office.WebExtension.EventType.RichApiMessage
        },
        entries: {
            value: d
        }
    })
}
;
(function(b) {
    var a = function() {
        function a() {
            var a = this;
            a._eventDispatch = null;
            a._registerHandlers = [];
            a._eventDispatch = new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.RichApiMessage]);
            OSF.DDA.DispIdHost.addEventSupport(a, a._eventDispatch)
        }
        a.prototype.register = function(c) {
            var a = this
              , b = a;
            if (!a._registerWithHostPromise)
                a._registerWithHostPromise = new Office.Promise(function(a, c) {
                    b.addHandlerAsync(Microsoft.Office.WebExtension.EventType.RichApiMessage, function(a) {
                        b._registerHandlers.forEach(function(b) {
                            b && b(a)
                        })
                    }, function(b) {
                        if (b.status == "failed")
                            c(b.error);
                        else
                            a()
                    })
                }
                );
            return a._registerWithHostPromise.then(function() {
                b._registerHandlers.push(c)
            })
        }
        ;
        return a
    }();
    b.RichApiMessageManager = a
}
)(OfficeExt || (OfficeExt = {}));
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidRichApiMessageEvent,
    toHost: [{
        name: Microsoft.Office.WebExtension.Parameters.Data,
        value: 0
    }],
    fromHost: [{
        name: Microsoft.Office.WebExtension.Parameters.Data,
        value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData
    }]
});
window.OfficeRuntime = window.OfficeRuntime || {};
window.OfficeRuntime.auth = {
    getAccessToken: function(b) {
        var a = window.Promise ? window.Promise : window.Office.Promise;
        return new a(function(d, a) {
            try {
                window.Office.context.auth.getAccessTokenAsync(b || {}, function(b) {
                    if (b.status === "succeeded")
                        d(b.value);
                    else
                        a(b.error)
                })
            } catch (c) {
                a(c)
            }
        }
        )
    },
    getAuthContext: function() {
        var a = window.Promise ? window.Promise : window.Office.Promise;
        return new a(function(c, a) {
            try {
                window.Office.context.auth.getAuthContextAsync(function(b) {
                    if (b.status === "succeeded")
                        c(b.value);
                    else
                        a(b.error)
                })
            } catch (b) {
                a(b)
            }
        }
        )
    }
};
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
    DialogParentMessageReceivedEvent: "DialogParentMessageReceivedEvent"
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
    DialogParentMessageReceived: "dialogParentMessageReceived",
    DialogParentEventReceived: "dialogParentEventReceived"
});
OSF.DialogParentMessageEventDispatch = new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived, Microsoft.Office.WebExtension.EventType.DialogParentEventReceived]);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidDialogParentMessageReceivedEvent,
    fromHost: [{
        name: OSF.DDA.EventDescriptors.DialogParentMessageReceivedEvent,
        value: OSF.DDA.SafeArray.Delegate.ParameterMap.self
    }],
    isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDescriptors.DialogParentMessageReceivedEvent,
    fromHost: [{
        name: OSF.DDA.PropertyDescriptors.MessageType,
        value: 0
    }, {
        name: OSF.DDA.PropertyDescriptors.MessageContent,
        value: 1
    }, {
        name: OSF.DDA.PropertyDescriptors.MessageOrigin,
        value: 2
    }],
    isComplexType: true
});
OSF.DDA.UI.EnableMessageChildDialogAPI = true;
var OfficeJsClient_OutlookWin32;
(function(a) {
    function c(a) {
        if (a.get_isDialog())
            a.ui = new OSF.DDA.UI.ChildUI;
        else {
            a.ui = new OSF.DDA.UI.ParentUI;
            OSF.DDA.DispIdHost.addAsyncMethods(a.ui, [OSF.DDA.AsyncMethodNames.CloseContainerAsync])
        }
        a.auth = new OSF.DDA.Auth;
        OSF.DDA.DispIdHost.addAsyncMethods(a.auth, [OSF.DDA.AsyncMethodNames.GetAccessTokenAsync, OSF.DDA.AsyncMethodNames.GetNestedAppAuthContextAsync]);
        OSF.DDA.OpenBrowser && OSF.DDA.DispIdHost.addAsyncMethods(a.ui, [OSF.DDA.AsyncMethodNames.OpenBrowserWindow])
    }
    a.prepareApiSurface = c;
    function b() {
        var a = OfficeExt.AppCommand.AppCommandManager.instance();
        a.initializeAndChangeOnce()
    }
    a.prepareRightAfterWebExtensionInitialize = b
}
)(OfficeJsClient_OutlookWin32 || (OfficeJsClient_OutlookWin32 = {}));
OSF.InitializationHelper.prototype.prepareRightAfterWebExtensionInitialize = function() {
    OfficeJsClient_OutlookWin32.prepareRightAfterWebExtensionInitialize()
}
;
OSF.InitializationHelper.prototype.prepareApiSurface = function(e) {
    var t = new OSF.DDA.License(e.get_eToken());
    e.get_appName() == OSF.AppName.OutlookWebApp ? (OSF.WebApp._UpdateLinksForHostAndXdmInfo(),
    this.initWebDialog(e),
    this.initWebAuth(e),
    OSF._OfficeAppFactory.setContext(new OSF.DDA.OutlookContext(e,this._settings,t,e.appOM)),
    OSF._OfficeAppFactory.setHostFacade(new OSF.DDA.DispIdHost.Facade(OSF.DDA.WAC.getDelegateMethods,OSF.DDA.WAC.Delegate.ParameterMap))) : (OfficeJsClient_OutlookWin32.prepareApiSurface(e),
    OSF._OfficeAppFactory.setContext(new OSF.DDA.OutlookContext(e,this._settings,t,e.appOM,OSF.DDA.OfficeTheme ? OSF.DDA.OfficeTheme.getOfficeTheme : null,e.ui)),
    OSF._OfficeAppFactory.setHostFacade(new OSF.DDA.DispIdHost.Facade(OSF.DDA.DispIdHost.getClientDelegateMethods,OSF.DDA.SafeArray.Delegate.ParameterMap)))
}
,
OSF.DDA.SettingsManager = {
    SerializedSettings: "serializedSettings",
    DateJSONPrefix: "Date(",
    DataJSONSuffix: ")",
    serializeSettings: function(e) {
        var t = {};
        for (var n in e) {
            var r = e[n];
            try {
                r = JSON ? JSON.stringify(r, (function(e, t) {
                    return OSF.OUtil.isDate(this[e]) ? OSF.DDA.SettingsManager.DateJSONPrefix + this[e].getTime() + OSF.DDA.SettingsManager.DataJSONSuffix : t
                }
                )) : Sys.Serialization.JavaScriptSerializer.serialize(r),
                t[n] = r
            } catch (e) {}
        }
        return t
    },
    deserializeSettings: function(e) {
        var t = {};
        for (var n in e = e || {}) {
            var r = e[n];
            try {
                r = JSON ? JSON.parse(r, (function(e, t) {
                    var n;
                    return "string" === typeof t && t && t.length > 6 && t.slice(0, 5) === OSF.DDA.SettingsManager.DateJSONPrefix && t.slice(-1) === OSF.DDA.SettingsManager.DataJSONSuffix && (n = new Date(parseInt(t.slice(5, -1)))) ? n : t
                }
                )) : Sys.Serialization.JavaScriptSerializer.deserialize(r, !0),
                t[n] = r
            } catch (e) {}
        }
        return t
    }
},
OSF.InitializationHelper.prototype.loadAppSpecificScriptAndCreateOM = function(e, t, n) {
    var r = "object" === typeof r ? r : {};
    r.OutlookAppOm = function(e) {
        var t = {};
        function n(r) {
            if (t[r])
                return t[r].exports;
            var i = t[r] = {
                i: r,
                l: !1,
                exports: {}
            };
            return e[r].call(i.exports, i, i.exports, n),
            i.l = !0,
            i.exports
        }
        return n.m = e,
        n.c = t,
        n.d = function(e, t, r) {
            n.o(e, t) || Object.defineProperty(e, t, {
                enumerable: !0,
                get: r
            })
        }
        ,
        n.r = function(e) {
            "undefined" !== typeof Symbol && Symbol.toStringTag && Object.defineProperty(e, Symbol.toStringTag, {
                value: "Module"
            }),
            Object.defineProperty(e, "__esModule", {
                value: !0
            })
        }
        ,
        n.t = function(e, t) {
            if (1 & t && (e = n(e)),
            8 & t)
                return e;
            if (4 & t && "object" === typeof e && e && e.__esModule)
                return e;
            var r = Object.create(null);
            if (n.r(r),
            Object.defineProperty(r, "default", {
                enumerable: !0,
                value: e
            }),
            2 & t && "string" != typeof e)
                for (var i in e)
                    n.d(r, i, function(t) {
                        return e[t]
                    }
                    .bind(null, i));
            return r
        }
        ,
        n.n = function(e) {
            var t = e && e.__esModule ? function() {
                return e.default
            }
            : function() {
                return e
            }
            ;
            return n.d(t, "a", t),
            t
        }
        ,
        n.o = function(e, t) {
            return Object.prototype.hasOwnProperty.call(e, t)
        }
        ,
        n.p = "/",
        n(n.s = 2)
    }([function(e, t) {
        e.exports = OSF
    }
    , function(e, t) {
        e.exports = Microsoft
    }
    , function(e, t, n) {
        "use strict";
        function r(e) {
            return null === e || void 0 === e
        }
        n.r(t);
        var i, a = window;
        function o(e) {
            return i[e]
        }
        var s, c = "", d = "", l = null, u = !1;
        function f(e) {
            var t, n, r = document.createElement("script");
            return r.type = "text/javascript",
            r.src = null !== (n = null === (t = OSF.TrustedTypesPolicy) || void 0 === t ? void 0 : t.createScriptURL(e)) && void 0 !== n ? n : e,
            r
        }
        function m() {
            u = !0,
            r(s) || !r(l.readyState) && (r(l.readyState) || "loaded" !== l.readyState && "complete" !== l.readyState) || (l.onload = null,
            l.onreadystatechange = null,
            "undefined" !== typeof a._u && (i = a._u.ExtensibilityStrings),
            s())
        }
        function p() {
            if (!u) {
                var e = document.getElementsByTagName("head")[0]
                  , t = d + "en-us/outlook_strings.js";
                l.onload = null,
                l.onreadystatechange = null,
                (l = f(t)).onload = m,
                l.onreadystatechange = m,
                e.appendChild(l)
            }
        }
        function y(e, t, n) {
            var r = n.substring(0, t)
              , i = r.lastIndexOf("/", r.length - 2);
            return -1 === i && (i = r.lastIndexOf("\\", r.length - 2)),
            -1 !== i && r.length > i + 1 && (e = r.substring(0, i + 1)),
            e
        }
        var v, g = function() {
            function e() {}
            return e.success = 0,
            e.noResponseDictionary = -900,
            e.noErrorCodeForStandardInvokeMethod = -901,
            e.genericProxyError = -902,
            e.genericLegacyApiError = -903,
            e.genericUnknownError = -904,
            e
        }(), h = function(e) {
            switch (e) {
            case 402:
            case 401:
            case 400:
            case 403:
                return !0;
            default:
                return !1
            }
        };
        !function(e) {
            e[e.noError = 0] = "noError",
            e[e.errorInRequest = -1] = "errorInRequest",
            e[e.errorHandlingRequest = -2] = "errorHandlingRequest",
            e[e.errorInResponse = -3] = "errorInResponse",
            e[e.errorHandlingResponse = -4] = "errorHandlingResponse",
            e[e.errorHandlingRequestAccessDenied = -5] = "errorHandlingRequestAccessDenied",
            e[e.errorHandlingMethodCallTimedout = -6] = "errorHandlingMethodCallTimedout"
        }(v || (v = {}));
        var A = n(0)
          , T = !1;
        function S(e) {
            return T || (b(9e3, "AttachmentSizeExceeded", o("l_AttachmentExceededSize_Text")),
            b(9001, "NumberOfAttachmentsExceeded", o("l_ExceededMaxNumberOfAttachments_Text")),
            b(9002, "InternalFormatError", o("l_InternalFormatError_Text")),
            b(9003, "InvalidAttachmentId", o("l_InvalidAttachmentId_Text")),
            b(9004, "InvalidAttachmentPath", o("l_InvalidAttachmentPath_Text")),
            b(9005, "CannotAddAttachmentBeforeUpgrade", o("l_CannotAddAttachmentBeforeUpgrade_Text")),
            b(9006, "AttachmentDeletedBeforeUploadCompletes", o("l_AttachmentDeletedBeforeUploadCompletes_Text")),
            b(9007, "AttachmentUploadGeneralFailure", o("l_AttachmentUploadGeneralFailure_Text")),
            b(9008, "AttachmentToDeleteDoesNotExist", o("l_DeleteAttachmentDoesNotExist_Text")),
            b(9009, "AttachmentDeleteGeneralFailure", o("l_AttachmentDeleteGeneralFailure_Text")),
            b(9010, "InvalidEndTime", o("l_InvalidEndTime_Text")),
            b(9011, "HtmlSanitizationFailure", o("l_HtmlSanitizationFailure_Text")),
            b(9012, "NumberOfRecipientsExceeded", o("l_NumberOfRecipientsExceeded_Text").replace("{0}", 500)),
            b(9013, "NoValidRecipientsProvided", o("l_NoValidRecipientsProvided_Text")),
            b(9014, "CursorPositionChanged", o("l_CursorPositionChanged_Text")),
            b(9016, "InvalidSelection", o("l_InvalidSelection_Text")),
            b(9017, "AccessRestricted", ""),
            b(9018, "GenericTokenError", ""),
            b(9019, "GenericSettingsError", ""),
            b(9020, "GenericResponseError", ""),
            b(9021, "SaveError", o("l_SaveError_Text")),
            b(9022, "MessageInDifferentStoreError", o("l_MessageInDifferentStoreError_Text")),
            b(9023, "DuplicateNotificationKey", o("l_DuplicateNotificationKey_Text")),
            b(9024, "NotificationKeyNotFound", o("l_NotificationKeyNotFound_Text")),
            b(9025, "NumberOfNotificationsExceeded", o("l_NumberOfNotificationsExceeded_Text")),
            b(9026, "PersistedNotificationArrayReadError", o("l_PersistedNotificationArrayReadError_Text")),
            b(9027, "PersistedNotificationArraySaveError", o("l_PersistedNotificationArraySaveError_Text")),
            b(9028, "CannotPersistPropertyInUnsavedDraftError", o("l_CannotPersistPropertyInUnsavedDraftError_Text")),
            b(9029, "CanOnlyGetTokenForSavedItem", o("l_CallSaveAsyncBeforeToken_Text")),
            b(9030, "APICallFailedDueToItemChange", o("l_APICallFailedDueToItemChange_Text")),
            b(9031, "InvalidParameterValueError", o("l_InvalidParameterValueError_Text")),
            b(9032, "ApiCallNotSupportedByExtensionPoint", o("l_API_Not_Supported_By_ExtensionPoint_Error_Text")),
            b(9033, "SetRecurrenceOnInstanceError", o("l_Recurrence_Error_Instance_SetAsync_Text")),
            b(9034, "InvalidRecurrenceError", o("l_Recurrence_Error_Properties_Invalid_Text")),
            b(9035, "RecurrenceZeroOccurrences", o("l_RecurrenceErrorZeroOccurrences_Text")),
            b(9036, "RecurrenceMaxOccurrences", o("l_RecurrenceErrorMaxOccurrences_Text")),
            b(9037, "RecurrenceInvalidTimeZone", o("l_RecurrenceInvalidTimeZone_Text")),
            b(9038, "InsufficientItemPermissionsError", o("l_Insufficient_Item_Permissions_Text")),
            b(9039, "RecurrenceUnsupportedAlternateCalendar", o("l_RecurrenceUnsupportedAlternateCalendar_Text")),
            b(9040, "HTTPRequestFailure", o("l_Olk_Http_Error_Text")),
            b(9041, "NetworkError", o("l_Internet_Not_Connected_Error_Text")),
            b(9042, "InternalServerError", o("l_Internal_Server_Error_Text")),
            b(9043, "AttachmentTypeNotSupported", o("l_AttachmentNotSupported_Text")),
            b(9044, "InvalidCategory", o("l_Invalid_Category_Error_Text")),
            b(9045, "DuplicateCategory", o("l_Duplicate_Category_Error_Text")),
            b(9046, "ItemNotSaved", o("l_Item_Not_Saved_Error_Text")),
            b(9047, "MissingExtendedPermissionsForAPIError", o("l_Missing_Extended_Permissions_For_API")),
            b(9048, "TokenAccessDenied", o("l_TokenAccessDeniedWithoutItemContext_Text")),
            b(9049, "ItemNotFound", o("l_ItemNotFound_Text")),
            b(9050, "KeyNotFound", o("l_KeyNotFound_Text")),
            b(9051, "SessionObjectMaxLengthExceeded", o("l_SessionDataObjectMaxLengthExceeded_Text").replace("{0}", 5e4)),
            b(9052, "AttachmentResourceNotFound", o("l_Attachment_Resource_Not_Found")),
            b(9053, "AttachmentResourceUnAuthorizedAccess", o("l_Attachment_Resource_UnAuthorizedAccess")),
            b(9054, "AttachmentDownloadFailed", o("l_Attachment_Download_Failed_Generic_Error")),
            b(9055, "APINotSupportedForSharedFolders", o("l_API_Not_Supported_For_Shared_Folders_Error")),
            b(9057, "RoamingSettingsExceededSize", o("l_RoamingSettingsExceededSize_Text")),
            b(9058, "NativeLabelingNotEnabled", o("l_NativeLabelingNotEnabled_Text")),
            b(9059, "SensitivityUnableToSetParent", o("l_SensitivityUnableToSetParent_Text")),
            b(9060, "UserHasNoLabelsToSet", o("l_UserHasNoLabelsToSet_Text")),
            b(9061, "FailedToGetLabelsCatalog", o("l_FailedToGetLabelsCatalog_Text")),
            b(9062, "FailedToGetLabel", o("l_FailedToGetLabel_Text")),
            b(9063, "FailedToSetLabel", o("l_FailedToSetLabel_Text")),
            b(9064, "ExceededMaxNumberOfSelectedItems", o("I_ExceededMaxNumberOfSelectedItems_Text")),
            b(9065, "MAMServiceNotAvailable", o("l_MAMServiceNotAvailable_Text")),
            b(9066, "InvalidOpenLocationInput", o("l_InvalidOpenLocationInput_Text")),
            b(9067, "InvalidSaveLocationInput", o("l_InvalidSaveLocationInput_Text")),
            T = !0),
            A.DDA.ErrorCodeManager.getErrorArgs(e)
        }
        function b(e, t, n) {
            A.DDA.ErrorCodeManager.addErrorMessage(e, {
                name: t,
                message: n
            })
        }
        var D, C = !1, x = new Map;
        function w(e) {
            return k(),
            x.get(e)
        }
        function E(e) {
            return k(),
            Boolean(x.get(e))
        }
        function k() {
            if (!C) {
                x.clear();
                var e = ao("nativeFlights");
                void 0 != e && (Object.keys(e).forEach((function(t) {
                    x.set(t, e[t])
                }
                )),
                C = !0)
            }
        }
        var O = function() {
            return D
        }
          , I = function(e) {
            return (D = new M).parameterBlobSupported = !0,
            D
        }
          , M = function() {
            function e() {
                this._parameterBlobSupported = !0,
                this._itemNumber = 0,
                this._itemNumberForLoadedItem = 0,
                D = this
            }
            return Object.defineProperty(e.prototype, "parameterBlobSupported", {
                set: function(e) {
                    this._parameterBlobSupported = e
                },
                enumerable: !0,
                configurable: !0
            }),
            e.prototype.setActionsDefinition = function(e) {
                this._actionsDefinition = e
            }
            ,
            e.prototype.setCurrentItemNumber = function(e) {
                e > 0 && (this._itemNumber = e)
            }
            ,
            Object.defineProperty(e.prototype, "itemNumber", {
                get: function() {
                    return this._itemNumber
                },
                enumerable: !0,
                configurable: !0
            }),
            e.prototype.setItemNumberForLoadedItem = function(e) {
                e > 0 && (this._itemNumberForLoadedItem = e)
            }
            ,
            Object.defineProperty(e.prototype, "itemNumberForLoadedItem", {
                get: function() {
                    return this._itemNumberForLoadedItem
                },
                enumerable: !0,
                configurable: !0
            }),
            Object.defineProperty(e.prototype, "actionsDefinition", {
                get: function() {
                    return this._actionsDefinition
                },
                enumerable: !0,
                configurable: !0
            }),
            e.prototype.updateOutlookExecuteParameters = function(e, t, n) {
                var r = e;
                if (this._parameterBlobSupported) {
                    if (this._itemNumber > 0 && (E("MultiSelectV2") && n && this._itemNumber != this._itemNumberForLoadedItem ? (t.itemNumber = this._itemNumberForLoadedItem.toString(),
                    t.isNoUI = !0) : t.itemNumber = this._itemNumber.toString()),
                    null != this._actionsDefinition && (t.actions = this.actionsDefinition),
                    0 === Object.keys(t).length)
                        return r;
                    null == r && (r = []),
                    r.push(JSON.stringify(t))
                }
                return r
            }
            ,
            e
        }()
          , _ = n(0)
          , N = function(e) {
            return _._OfficeAppFactory.getHostInfo().hostPlatform == e
        }
          , F = n(0)
          , P = function(e, t) {
            if (0 == e.length)
                return null;
            var n = j(e);
            W(e);
            var r = n > 0
              , i = 0;
            return O() && (i = O().itemNumber),
            B(e, r && i > 0 && n > i && !t)
        }
          , R = function(e, t, n) {
            var r = null
              , i = {};
            switch (e) {
            case 12:
                i.isRest = t.isRest;
                break;
            case 4:
                r = [JSON.stringify(t.customProperties)];
                break;
            case 5:
                r = new Array(t.body);
                break;
            case 8:
            case 9:
            case 179:
            case 180:
                r = new Array(t.itemId);
                break;
            case 7:
            case 177:
                r = new Array(U(t.requiredAttendees),U(t.optionalAttendees),t.start,t.end,t.location,U(t.resources),t.subject,t.body);
                break;
            case 44:
            case 178:
                r = [U(t.toRecipients), U(t.ccRecipients), U(t.bccRecipients), t.subject, t.htmlBody, t.attachments];
                break;
            case 43:
                r = [t.ewsIdOrEmail];
                break;
            case 45:
                r = [t.module, t.queryString];
                break;
            case 40:
                r = [t.extensionId, t.consentState];
                break;
            case 11:
            case 10:
            case 184:
            case 183:
                r = [t.htmlBody];
                break;
            case 31:
            case 30:
            case 182:
            case 181:
                r = [t.htmlBody, t.attachments];
                break;
            case 23:
            case 13:
            case 38:
            case 29:
                r = [t.data, t.coercionType];
                break;
            case 37:
                r = N("ios") || N("android") ? [t.coercionType, t.bodyMode] : [t.coercionType];
                break;
            case 28:
                r = [t.coercionType];
                break;
            case 17:
                r = [t.subject];
                break;
            case 15:
                r = [t.recipientField];
                break;
            case 22:
            case 21:
                r = [t.recipientField, L(t.recipientArray)];
                break;
            case 19:
                r = [t.itemId, t.name];
                break;
            case 16:
                r = [t.uri, t.name, t.isInline];
                break;
            case 148:
                r = [t.base64String, t.name, t.isInline];
                break;
            case 20:
                r = [t.attachmentIndex];
                break;
            case 25:
                r = [t.TimeProperty, t.time];
                break;
            case 24:
                r = [t.TimeProperty];
                break;
            case 27:
                r = [t.location];
                break;
            case 33:
            case 35:
                r = [t.key, t.type, t.persistent, t.message, t.icon],
                O().setActionsDefinition(t.actions);
                break;
            case 36:
                r = [t.key];
                break;
            default:
                i = t || {}
            }
            return 1 !== e && (r = O().updateOutlookExecuteParameters(r, i, n)),
            r
        }
          , U = function(e) {
            return null != e ? e.join(";") : ""
        }
          , L = function(e) {
            var t = [];
            if (null == e)
                return t;
            for (var n = 0; n < e.length; n++) {
                var r = [e[n].address, e[n].name];
                t.push(r)
            }
            return t
        }
          , j = function(e) {
            var t = 0;
            if (e.length > 2) {
                var n = JSON.parse(e[2]);
                n && "object" === typeof n && (t = n.itemNumber)
            }
            return t
        }
          , W = function(e) {
            if (e.length > 2) {
                var t = JSON.parse(e[2]);
                t && "object" === typeof t && t.itemNumberForLoadedItem && O().setItemNumberForLoadedItem(t.itemNumberForLoadedItem)
            }
        }
          , B = function(e, t) {
            var n = null
              , r = JSON.parse(e[0]);
            if ("number" === typeof r)
                n = J(e, t);
            else {
                if (!r || "object" !== typeof r)
                    throw new Error("Return data type from host must be Object or Number");
                n = H(e, t)
            }
            return n
        }
          , H = function(e, t) {
            var n = JSON.parse(e[0]);
            if (t)
                n.error = !0,
                n.errorCode = 9030;
            else if (e.length > 1 && 0 !== e[1]) {
                if (n.error = !0,
                n.errorCode = e[1],
                e.length > 2) {
                    var r = JSON.parse(e[2]);
                    n.diagnostics = r.Diagnostics
                }
                e.length >= 5 && (n.errorMessage = e[3],
                n.errorName = e[4])
            } else
                n.error = !1;
            return n
        }
          , J = function(e, t) {
            var n = {
                error: !0
            };
            return n.errorCode = e[0],
            n
        };
        var z = n(0);
        function q(e, t, n, r, i, a, o) {
            G(e, r, (function(e, r) {
                if (n) {
                    var o = void 0
                      , s = !0;
                    if ("object" === typeof r && null !== r) {
                        if (void 0 !== r.wasSuccessful && (s = r.wasSuccessful),
                        void 0 !== r.error || void 0 !== r.errorCode || void 0 !== r.data)
                            if (r.error) {
                                var c = r.errorCode;
                                o = V(void 0, z.DDA.AsyncResultEnum.ErrorCode.Failed, c, t)
                            } else {
                                o = V(i ? i(r.data) : r.data, z.DDA.AsyncResultEnum.ErrorCode.Success, 0, t)
                            }
                        a && (o = a(r, t, e)),
                        o || e === v.noError || (o = V(void 0, z.DDA.AsyncResultEnum.ErrorCode.Failed, 9002, t)),
                        o || e !== v.noError || !1 !== s || (o = V(void 0, z.DDA.AsyncResultEnum.ErrorCode.Failed, z.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupported, t)),
                        n(o)
                    }
                }
            }
            ), o)
        }
        function V(e, t, n, r, i, a) {
            var o, s = {};
            if (s[z.DDA.AsyncResultEnum.Properties.Value] = e,
            s[z.DDA.AsyncResultEnum.Properties.Context] = r,
            z.DDA.AsyncResultEnum.ErrorCode.Success !== t) {
                o = {};
                var c;
                c = S(n),
                o[z.DDA.AsyncResultEnum.ErrorProperties.Name] = a || c.name,
                o[z.DDA.AsyncResultEnum.ErrorProperties.Message] = i || c.message,
                o[z.DDA.AsyncResultEnum.ErrorProperties.Code] = n
            }
            return new z.DDA.AsyncResult(s,o)
        }
        var Y, G = function(e, t, n, r) {
            Z(e, t, n, r)
        }, Z = function(e, t, n, r) {
            if (z.AppName.OutlookWebApp !== so() && h(e))
                n(v.errorHandlingRequest, null);
            else {
                var i = performance && performance.now()
                  , a = function(t, r) {
                    K(t, r, e, i),
                    n && n(t, r)
                };
                if (z.AppName.OutlookWebApp === so()) {
                    var o = {
                        ApiParams: t,
                        MethodData: E("MultiSelectV2") ? {
                            ControlId: z._OfficeAppFactory.getId(),
                            DispatchId: e,
                            isLoadedItem: r
                        } : {
                            ControlId: z._OfficeAppFactory.getId(),
                            DispatchId: e
                        }
                    };
                    1 === e ? z._OfficeAppFactory.getClientEndPoint().invoke("GetInitialData", a, o) : z._OfficeAppFactory.getClientEndPoint().invoke("ExecuteMethod", a, o)
                } else
                    !function(e, t, n, r) {
                        var i = R(e, t, r);
                        F.ClientHostController.execute(e, i, (function(e, t) {
                            var i = e.toArray()
                              , a = P(i, r);
                            null != n && n(t, a)
                        }
                        ))
                    }(e, t, a, r)
            }
        }, K = function(e, t, n, i) {
            if (z.AppTelemetry) {
                var a = function(e, t) {
                    if (t) {
                        if ("error"in t)
                            return t.error ? "errorCode"in t ? t.errorCode : g.noErrorCodeForStandardInvokeMethod : g.success;
                        if ("wasProxySuccessful"in t)
                            return t.wasProxySuccessful ? g.success : g.genericProxyError;
                        if ("wasSuccessful"in t)
                            return t.wasSuccessful ? g.success : g.genericLegacyApiError
                    }
                    return r(e) ? g.genericUnknownError : e
                }(e, t)
                  , o = performance && performance.now();
                z.AppTelemetry.onMethodDone(n, null, Math.round(o - i), a)
            }
        }, $ = function() {
            var e = ao("permissionLevel");
            return r(e) ? -1 : e
        };
        function Q(e, t) {
            var n = new Error(e);
            if (n.message = e || "",
            t)
                for (var r in t)
                    n[r] = t[r];
            return n
        }
        function X(e, t) {
            var n = "Sys.ArgumentException: " + (t || "Value does not fall within the expected range.");
            return e && (n += "\n" + "Parameter name: {0}".replace("{0}", e)),
            Q(n, {
                name: "Sys.ArgumentException",
                paramName: e
            })
        }
        function ee(e, t) {
            var n = "Sys.ArgumentNullException: " + (t || "Value cannot be null.");
            return e && (n += "\n" + "Parameter name: {0}".replace("{0}", e)),
            Q(n, {
                name: "Sys.ArgumentNullException",
                paramName: e
            })
        }
        function te(e, t, n) {
            var r = "Sys.ArgumentOutOfRangeException: " + (n || "Specified argument was out of the range of valid values.");
            return e && (r += "\n" + "Parameter name: {0}".replace("{0}", e)),
            "undefined" !== typeof t && null !== t && (r += "\n" + "Actual value was {0}.".replace("{0}", t)),
            Q(r, {
                name: "Sys.ArgumentOutOfRangeException",
                paramName: e,
                actualValue: t
            })
        }
        function ne(e, t, n, r) {
            var i = "Sys.ArgumentTypeException: ";
            return i += r || (t && n ? "Object of type '{0}' cannot be converted to type '{1}'.".replace("{0}", t.getName ? t.getName() : t).replace("{1}", n.getName ? n.getName() : n) : "Object cannot be converted to the required type."),
            e && (i += "\n" + "Parameter name: {0}".replace("{0}", e)),
            Q(i, {
                name: "Sys.ArgumentTypeException",
                paramName: e,
                actualType: t,
                expectedType: n
            })
        }
        function re(e, t) {
            if (-1 == $())
                throw function(e) {
                    return Q("Invalid operation ({0}) when Office.context.mailbox.item is null.".replace("{0}", e))
                }(t);
            if ($() < e)
                throw Q(o("l_ElevatedPermissionNeededForMethod_Text").replace("{0}", t))
        }
        function ie(e, t, n) {
            var r = {};
            if (n && (r = function(e) {
                var t = {};
                if (1 === e.length || 2 === e.length)
                    return "function" !== typeof e[0] ? t : (t.callback = e[0],
                    2 === e.length && (t.asyncContext = e[1]),
                    t);
                return t
            }(e)).callback)
                return r;
            if (1 === e.length)
                if ("function" === typeof e[0])
                    r.callback = e[0];
                else {
                    if ("object" !== typeof e[0])
                        throw ne();
                    r.options = e[0]
                }
            else if (2 === e.length) {
                if ("object" !== typeof e[0])
                    throw X("options");
                if ("function" !== typeof e[1])
                    throw X("callback");
                r.callback = e[1],
                r.options = e[0]
            } else if (0 !== e.length)
                throw Q("Sys.ParameterCountException: " + (o("l_ParametersNotAsExpected_Text") || "Parameter count mismatch."), {
                    name: "Sys.ParameterCountException"
                });
            if (t && !r.callback)
                throw ee("callback");
            return r.options && r.options.asyncContext && (r.asyncContext = r.options.asyncContext),
            r
        }
        !function(e) {
            e[e.to = 0] = "to",
            e[e.cc = 1] = "cc",
            e[e.bcc = 2] = "bcc",
            e[e.requiredAttendees = 0] = "requiredAttendees",
            e[e.optionalAttendees = 1] = "optionalAttendees"
        }(Y || (Y = {}));
        function ae(e, t, n, r) {
            if (e < t || e > n)
                throw te(String(r))
        }
        var oe, se, ce = {
            EntityType: {
                MeetingSuggestion: "meetingSuggestion",
                TaskSuggestion: "taskSuggestion",
                Address: "address",
                EmailAddress: "emailAddress",
                Url: "url",
                PhoneNumber: "phoneNumber",
                Contact: "contact",
                FlightReservations: "flightReservations",
                ParcelDeliveries: "parcelDeliveries"
            },
            ItemType: {
                Message: "message",
                Appointment: "appointment"
            },
            ResponseType: {
                None: "none",
                Organizer: "organizer",
                Tentative: "tentative",
                Accepted: "accepted",
                Declined: "declined"
            },
            RecipientType: {
                Other: "other",
                DistributionList: "distributionList",
                User: "user",
                ExternalUser: "externalUser"
            },
            AttachmentType: {
                File: "file",
                Item: "item",
                Cloud: "cloud",
                Base64: "base64"
            },
            AttachmentStatus: {
                Added: "added",
                Removed: "removed"
            },
            AttachmentContentFormat: {
                Base64: "base64",
                Url: "url",
                Eml: "eml",
                ICalendar: "iCalendar"
            },
            BodyType: {
                Text: "text",
                Html: "html"
            },
            ItemNotificationMessageType: {
                ProgressIndicator: "progressIndicator",
                InformationalMessage: "informationalMessage",
                ErrorMessage: "errorMessage",
                InsightMessage: "insightMessage"
            },
            Folder: {
                Inbox: "inbox",
                Junk: "junk",
                DeletedItems: "deletedItems"
            },
            ComposeType: {
                Forward: "forward",
                NewMail: "newMail",
                Reply: "reply"
            }
        }, de = {
            Text: "text",
            Html: "html"
        };
        function le(e) {
            if (null === e || void 0 === e)
                throw ee(e);
            if (e !== ce.RestVersion.v1_0 && e !== ce.RestVersion.v2_0 && e !== ce.RestVersion.Beta)
                throw X(e)
        }
        function ue(e, t) {
            if (null === e || void 0 === e)
                throw ee(e);
            return le(t),
            e.replace(new RegExp("[/]","g"), "-").replace(new RegExp("[+]","g"), "_")
        }
        function fe(e, t) {
            if (null === e || void 0 === e)
                throw ee(e);
            return le(t),
            e.replace(new RegExp("[-]","g"), "/").replace(new RegExp("[_]","g"), "+")
        }
        function me(e, t) {
            if (!Array.isArray(e))
                throw ne("name");
            ae(e.length, 0, 100, "{0}.length".replace("{0}", t))
        }
        function pe(e, t) {
            for (var n = e, r = [], i = 0; i < n.length; i++)
                if ("object" === typeof n[i]) {
                    if (ye(n[i]),
                    r[i] = n[i].emailAddress,
                    "string" !== typeof r[i])
                        throw X("{0}[{1}]".replace(t, String(i)))
                } else {
                    if ("string" !== typeof n[i])
                        throw X("{0}[{1}]".replace(t, String(i)));
                    r[i] = n[i]
                }
            return r
        }
        function ye(e) {
            if (!r(e.displayName) && "string" === typeof e.displayName && e.displayName.length > 255)
                throw te("displayName");
            if (!r(e.emailAddress) && "string" === typeof e.emailAddress && e.emailAddress.length > 571)
                throw te("emailAddress");
            if (!r(e.appointmentResponse) && "string" !== typeof e.appointmentResponse)
                throw te("appointmentResponse");
            if (!r(e.recipientType) && "string" !== typeof e.recipientType)
                throw te("recipientType")
        }
        function ve(e) {
            if ("string" !== typeof e)
                throw ne("itemId");
            !function(e) {
                if (r(e) || "" === e)
                    throw ee("itemId")
            }(e)
        }
        function ge(e) {
            return ao("isRestIdSupported") ? ue(e, ce.RestVersion.v1_0) : fe(e, ce.RestVersion.v1_0)
        }
        ce.UserProfileType = {
            Office365: "office365",
            OutlookCom: "outlookCom",
            Enterprise: "enterprise"
        },
        ce.RestVersion = {
            v1_0: "v1.0",
            v2_0: "v2.0",
            Beta: "beta"
        },
        ce.ModuleType = {
            Addins: "addins"
        },
        ce.ActionType = {
            ShowTaskPane: "showTaskPane",
            ExecuteFunction: "executeFunction"
        },
        ce.SendModeOverride = {
            PromptUser: "promptUser"
        },
        ce.Days = {
            Mon: "mon",
            Tue: "tue",
            Wed: "wed",
            Thu: "thu",
            Fri: "fri",
            Sat: "sat",
            Sun: "sun",
            Weekday: "weekday",
            WeekendDay: "weekendDay",
            Day: "day"
        },
        ce.WeekNumber = {
            First: "first",
            Second: "second",
            Third: "third",
            Fourth: "fourth",
            Last: "last"
        },
        ce.RecurrenceType = {
            Daily: "daily",
            Weekday: "weekday",
            Weekly: "weekly",
            Monthly: "monthly",
            Yearly: "yearly"
        },
        ce.Month = {
            Jan: "jan",
            Feb: "feb",
            Mar: "mar",
            Apr: "apr",
            May: "may",
            Jun: "jun",
            Jul: "jul",
            Aug: "aug",
            Sep: "sep",
            Oct: "oct",
            Nov: "nov",
            Dec: "dec"
        },
        ce.DelegatePermissions = {
            Read: 1,
            Write: 2,
            DeleteOwn: 4,
            DeleteAll: 8,
            EditOwn: 16,
            EditAll: 32
        },
        ce.TimeZone = {
            AfghanistanStandardTime: "Afghanistan Standard Time",
            AlaskanStandardTime: "Alaskan Standard Time",
            AleutianStandardTime: "Aleutian Standard Time",
            AltaiStandardTime: "Altai Standard Time",
            ArabStandardTime: "Arab Standard Time",
            ArabianStandardTime: "Arabian Standard Time",
            ArabicStandardTime: "Arabic Standard Time",
            ArgentinaStandardTime: "Argentina Standard Time",
            AstrakhanStandardTime: "Astrakhan Standard Time",
            AtlanticStandardTime: "Atlantic Standard Time",
            AUSCentralStandardTime: "AUS Central Standard Time",
            AusCentralWStandardTime: "Aus Central W. Standard Time",
            AUSEasternStandardTime: "AUS Eastern Standard Time",
            AzerbaijanStandardTime: "Azerbaijan Standard Time",
            AzoresStandardTime: "Azores Standard Time",
            BahiaStandardTime: "Bahia Standard Time",
            BangladeshStandardTime: "Bangladesh Standard Time",
            BelarusStandardTime: "Belarus Standard Time",
            BougainvilleStandardTime: "Bougainville Standard Time",
            CanadaCentralStandardTime: "Canada Central Standard Time",
            CapeVerdeStandardTime: "Cape Verde Standard Time",
            CaucasusStandardTime: "Caucasus Standard Time",
            CenAustraliaStandardTime: "Cen. Australia Standard Time",
            CentralAmericaStandardTime: "Central America Standard Time",
            CentralAsiaStandardTime: "Central Asia Standard Time",
            CentralBrazilianStandardTime: "Central Brazilian Standard Time",
            CentralEuropeStandardTime: "Central Europe Standard Time",
            CentralEuropeanStandardTime: "Central European Standard Time",
            CentralPacificStandardTime: "Central Pacific Standard Time",
            CentralStandardTime: "Central Standard Time",
            CentralStandardTime_Mexico: "Central Standard Time (Mexico)",
            ChathamIslandsStandardTime: "Chatham Islands Standard Time",
            ChinaStandardTime: "China Standard Time",
            CubaStandardTime: "Cuba Standard Time",
            DatelineStandardTime: "Dateline Standard Time",
            EAfricaStandardTime: "E. Africa Standard Time",
            EAustraliaStandardTime: "E. Australia Standard Time",
            EEuropeStandardTime: "E. Europe Standard Time",
            ESouthAmericaStandardTime: "E. South America Standard Time",
            EasterIslandStandardTime: "Easter Island Standard Time",
            EasternStandardTime: "Eastern Standard Time",
            EasternStandardTime_Mexico: "Eastern Standard Time (Mexico)",
            EgyptStandardTime: "Egypt Standard Time",
            EkaterinburgStandardTime: "Ekaterinburg Standard Time",
            FijiStandardTime: "Fiji Standard Time",
            FLEStandardTime: "FLE Standard Time",
            GeorgianStandardTime: "Georgian Standard Time",
            GMTStandardTime: "GMT Standard Time",
            GreenlandStandardTime: "Greenland Standard Time",
            GreenwichStandardTime: "Greenwich Standard Time",
            GTBStandardTime: "GTB Standard Time",
            HaitiStandardTime: "Haiti Standard Time",
            HawaiianStandardTime: "Hawaiian Standard Time",
            IndiaStandardTime: "India Standard Time",
            IranStandardTime: "Iran Standard Time",
            IsraelStandardTime: "Israel Standard Time",
            JordanStandardTime: "Jordan Standard Time",
            KaliningradStandardTime: "Kaliningrad Standard Time",
            KamchatkaStandardTime: "Kamchatka Standard Time",
            KoreaStandardTime: "Korea Standard Time",
            LibyaStandardTime: "Libya Standard Time",
            LineIslandsStandardTime: "Line Islands Standard Time",
            LordHoweStandardTime: "Lord Howe Standard Time",
            MagadanStandardTime: "Magadan Standard Time",
            MagallanesStandardTime: "Magallanes Standard Time",
            MarquesasStandardTime: "Marquesas Standard Time",
            MauritiusStandardTime: "Mauritius Standard Time",
            MidAtlanticStandardTime: "Mid-Atlantic Standard Time",
            MiddleEastStandardTime: "Middle East Standard Time",
            MontevideoStandardTime: "Montevideo Standard Time",
            MoroccoStandardTime: "Morocco Standard Time",
            MountainStandardTime: "Mountain Standard Time",
            MountainStandardTime_Mexico: "Mountain Standard Time (Mexico)",
            MyanmarStandardTime: "Myanmar Standard Time",
            NCentralAsiaStandardTime: "N. Central Asia Standard Time",
            NamibiaStandardTime: "Namibia Standard Time",
            NepalStandardTime: "Nepal Standard Time",
            NewZealandStandardTime: "New Zealand Standard Time",
            NewfoundlandStandardTime: "Newfoundland Standard Time",
            NorfolkStandardTime: "Norfolk Standard Time",
            NorthAsiaEastStandardTime: "North Asia East Standard Time",
            NorthAsiaStandardTime: "North Asia Standard Time",
            NorthKoreaStandardTime: "North Korea Standard Time",
            OmskStandardTime: "Omsk Standard Time",
            PacificSAStandardTime: "Pacific SA Standard Time",
            PacificStandardTime: "Pacific Standard Time",
            PacificStandardTime_Mexico: "Pacific Standard Time (Mexico)",
            PakistanStandardTime: "Pakistan Standard Time",
            ParaguayStandardTime: "Paraguay Standard Time",
            RomanceStandardTime: "Romance Standard Time",
            RussiaTimeZone10: "Russia Time Zone 10",
            RussiaTimeZone11: "Russia Time Zone 11",
            RussiaTimeZone3: "Russia Time Zone 3",
            RussianStandardTime: "Russian Standard Time",
            SAEasternStandardTime: "SA Eastern Standard Time",
            SAPacificStandardTime: "SA Pacific Standard Time",
            SAWesternStandardTime: "SA Western Standard Time",
            SaintPierreStandardTime: "Saint Pierre Standard Time",
            SakhalinStandardTime: "Sakhalin Standard Time",
            SamoaStandardTime: "Samoa Standard Time",
            SaratovStandardTime: "Saratov Standard Time",
            SEAsiaStandardTime: "SE Asia Standard Time",
            SingaporeStandardTime: "Singapore Standard Time",
            SouthAfricaStandardTime: "South Africa Standard Time",
            SriLankaStandardTime: "Sri Lanka Standard Time",
            SudanStandardTime: "Sudan Standard Time",
            SyriaStandardTime: "Syria Standard Time",
            TaipeiStandardTime: "Taipei Standard Time",
            TasmaniaStandardTime: "Tasmania Standard Time",
            TocantinsStandardTime: "Tocantins Standard Time",
            TokyoStandardTime: "Tokyo Standard Time",
            TomskStandardTime: "Tomsk Standard Time",
            TongaStandardTime: "Tonga Standard Time",
            TransbaikalStandardTime: "Transbaikal Standard Time",
            TurkeyStandardTime: "Turkey Standard Time",
            TurksAndCaicosStandardTime: "Turks And Caicos Standard Time",
            UlaanbaatarStandardTime: "Ulaanbaatar Standard Time",
            USEasternStandardTime: "US Eastern Standard Time",
            USMountainStandardTime: "US Mountain Standard Time",
            UTC: "UTC",
            UTCPLUS12: "UTC+12",
            UTCPLUS13: "UTC+13",
            UTCMINUS02: "UTC-02",
            UTCMINUS08: "UTC-08",
            UTCMINUS09: "UTC-09",
            UTCMINUS11: "UTC-11",
            VenezuelaStandardTime: "Venezuela Standard Time",
            VladivostokStandardTime: "Vladivostok Standard Time",
            WAustraliaStandardTime: "W. Australia Standard Time",
            WCentralAfricaStandardTime: "W. Central Africa Standard Time",
            WEuropeStandardTime: "W. Europe Standard Time",
            WMongoliaStandardTime: "W. Mongolia Standard Time",
            WestAsiaStandardTime: "West Asia Standard Time",
            WestBankStandardTime: "West Bank Standard Time",
            WestPacificStandardTime: "West Pacific Standard Time",
            YakutskStandardTime: "Yakutsk Standard Time"
        },
        ce.LocationType = {
            Custom: "custom",
            Room: "room"
        },
        ce.AppointmentSensitivityType = {
            Normal: "normal",
            Personal: "personal",
            Private: "private",
            Confidential: "confidential"
        },
        ce.CategoryColor = {
            None: "None",
            Preset0: "Preset0",
            Preset1: "Preset1",
            Preset2: "Preset2",
            Preset3: "Preset3",
            Preset4: "Preset4",
            Preset5: "Preset5",
            Preset6: "Preset6",
            Preset7: "Preset7",
            Preset8: "Preset8",
            Preset9: "Preset9",
            Preset10: "Preset10",
            Preset11: "Preset11",
            Preset12: "Preset12",
            Preset13: "Preset13",
            Preset14: "Preset14",
            Preset15: "Preset15",
            Preset16: "Preset16",
            Preset17: "Preset17",
            Preset18: "Preset18",
            Preset19: "Preset19",
            Preset20: "Preset20",
            Preset21: "Preset21",
            Preset22: "Preset22",
            Preset23: "Preset23",
            Preset24: "Preset24"
        },
        ce.MoveSpamItemTo = {
            DeletedItemsFolder: "deletedItemsFolder",
            CustomFolder: "customFolder",
            JunkFolder: "junkFolder",
            NoMove: "noMove"
        },
        ce.SaveLocation = {
            OnedriveForBusiness: 1,
            SharePoint: 2,
            Box: 4,
            Dropbox: 8,
            GoogleDrive: 16,
            Local: 32,
            AccountDocument: 64,
            PhotoLibrary: 128,
            Other: 1 << 31
        },
        ce.OpenLocation = {
            OnedriveForBusiness: 1,
            SharePoint: 2,
            Camera: 4,
            Local: 8,
            AccountDocument: 16,
            PhotoLibrary: 32,
            Other: 1 << 31
        },
        ce.BodyMode = {
            FullBody: 0,
            HostConfig: 1
        },
        (se = oe || (oe = {}))[se.camera = 0] = "camera",
        se[se.microphone = 1] = "microphone",
        se[se.geolocation = 2] = "geolocation";
        var he = function() {
            for (var e = 0, t = 0, n = arguments.length; t < n; t++)
                e += arguments[t].length;
            var r = Array(e)
              , i = 0;
            for (t = 0; t < n; t++)
                for (var a = arguments[t], o = 0, s = a.length; o < s; o++,
                i++)
                    r[i] = a[o];
            return r
        };
        function Ae(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            Se.apply(void 0, he([9, e], t))
        }
        function Te(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            Se.apply(void 0, he([180, e], t))
        }
        function Se(e, t) {
            for (var n = [], r = 2; r < arguments.length; r++)
                n[r - 2] = arguments[r];
            re(1, "mailbox.displayAppointmentForm");
            var i = ie(n, !1, !1)
              , a = {
                itemId: t
            };
            be(a),
            q(e, i.asyncContext, i.callback, {
                itemId: ge(a.itemId)
            }, void 0, void 0, void 0)
        }
        function be(e) {
            ve(e.itemId)
        }
        var De = function() {
            for (var e = 0, t = 0, n = arguments.length; t < n; t++)
                e += arguments[t].length;
            var r = Array(e)
              , i = 0;
            for (t = 0; t < n; t++)
                for (var a = arguments[t], o = 0, s = a.length; o < s; o++,
                i++)
                    r[i] = a[o];
            return r
        };
        function Ce(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            we.apply(void 0, De([8, e], t))
        }
        function xe(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            we.apply(void 0, De([179, e], t))
        }
        function we(e, t) {
            for (var n = [], r = 2; r < arguments.length; r++)
                n[r - 2] = arguments[r];
            re(1, "mailbox.displayMessageForm");
            var i = ie(n, !1, !1)
              , a = {
                itemId: t
            };
            Ee(a),
            q(e, i.asyncContext, i.callback, {
                itemId: ge(a.itemId)
            }, void 0, void 0, void 0)
        }
        function Ee(e) {
            ve(e.itemId)
        }
        function ke(e, t, n, r) {
            if ("string" !== typeof e)
                throw X(String(r));
            ae(e.length, t, n, r)
        }
        var Oe = function(e) {
            return e instanceof Date || "[object Date]" == Object.prototype.toString.call(e)
        }
          , Ie = function() {
            for (var e = 0, t = 0, n = arguments.length; t < n; t++)
                e += arguments[t].length;
            var r = Array(e)
              , i = 0;
            for (t = 0; t < n; t++)
                for (var a = arguments[t], o = 0, s = a.length; o < s; o++,
                i++)
                    r[i] = a[o];
            return r
        };
        function Me(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            return Ne.apply(void 0, Ie([7, e], t))
        }
        function _e(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            return Ne.apply(void 0, Ie([177, e], t))
        }
        function Ne(e, t) {
            for (var n = [], r = 2; r < arguments.length; r++)
                n[r - 2] = arguments[r];
            re(1, "mailbox.displayNewAppointmentForm");
            var i = ie(n, !1, !1);
            Fe(t);
            var a = Pe(t);
            q(e, i.asyncContext, i.callback, a, void 0, void 0, void 0)
        }
        function Fe(e) {
            if (r(e.requiredAttendees) || me(e.requiredAttendees, "requiredAttendees"),
            r(e.optionalAttendees) || me(e.optionalAttendees, "optionalAttendees"),
            r(e.location) || ke(e.location, 0, 255, "location"),
            r(e.body) || ke(e.body, 0, 32768, "body"),
            r(e.subject) || ke(e.subject, 0, 255, "subject"),
            !r(e.start)) {
                if (!Oe(e.start))
                    throw X("start");
                if (!r(e.end)) {
                    if (!Oe(e.end))
                        throw X("end");
                    if (e.end && e.start && e.end < e.start)
                        throw X("end", o("l_InvalidEventDates_Text"))
                }
            }
        }
        function Pe(e) {
            var t = null
              , n = null;
            if (r(e.requiredAttendees) || (t = pe(e.requiredAttendees, "requiredAttendees")),
            r(e.optionalAttendees) || (n = pe(e.optionalAttendees, "optionalAttendees")),
            !r(e.start)) {
                var i = e.start;
                e.start = i.getTime()
            }
            if (!r(e.end)) {
                var a = e.end;
                e.end = a.getTime()
            }
            var o = JSON.parse(JSON.stringify(e));
            return (t || n) && (r(e.requiredAttendees) || (o.requiredAttendees = t),
            r(e.optionalAttendees) || (o.optionalAttendees = n)),
            o
        }
        var Re = function() {
            return (Re = Object.assign || function(e) {
                for (var t, n = 1, r = arguments.length; n < r; n++)
                    for (var i in t = arguments[n])
                        Object.prototype.hasOwnProperty.call(t, i) && (e[i] = t[i]);
                return e
            }
            ).apply(this, arguments)
        };
        function Ue(e) {
            var t = [];
            return null != e && "object" === typeof e ? (Array.isArray(null === e || void 0 === e ? void 0 : e.attachments) && (t = e.attachments.map((function(e) {
                return e.type == ce.AttachmentType.File && e.url ? fetch(e.url).then((function(e) {
                    return e.blob()
                }
                )).then((function(t) {
                    return new Promise((function(n, r) {
                        var i = new FileReader;
                        i.onloadend = function() {
                            var t;
                            n(Re(Re({}, e), {
                                type: ce.AttachmentType.Base64,
                                base64file: (t = i.result,
                                t.replace("data:", "").replace(/^.+,/, "")),
                                url: void 0
                            }))
                        }
                        ,
                        i.onerror = r,
                        i.readAsDataURL(t)
                    }
                    ))
                }
                )) : Promise.resolve(e)
            }
            ))),
            Promise.all(t).then((function(t) {
                return e.attachments = t,
                e
            }
            )).catch((function(e) {
                return Promise.reject(new Error(o("l_AttachmentUploadGeneralFailure_Text")))
            }
            ))) : Promise.resolve(e)
        }
        var Le = function() {
            for (var e = 0, t = 0, n = arguments.length; t < n; t++)
                e += arguments[t].length;
            var r = Array(e)
              , i = 0;
            for (t = 0; t < n; t++)
                for (var a = arguments[t], o = 0, s = a.length; o < s; o++,
                i++)
                    r[i] = a[o];
            return r
        };
        function je(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            return Be.apply(void 0, Le([44, e], t))
        }
        function We(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            return Be.apply(void 0, Le([178, e], t))
        }
        function Be(e, t) {
            for (var n = [], r = 2; r < arguments.length; r++)
                n[r - 2] = arguments[r];
            re(1, "mailbox.displayNewMessageForm");
            var i = ie(n, !1, !1)
              , a = function(e, t, n) {
                var r = 9007;
                e instanceof Error && "Sys.ArgumentOutOfRangeException" === e.name && (r = 9e3);
                var i = V(void 0, OSF.DDA.AsyncResultEnum.ErrorCode.Failed, r, n, "");
                t && setTimeout((function() {
                    return t(i)
                }
                ), 0)
            };
            if (He(t),
            N("web"))
                Ue(t).then((function(n) {
                    var r = Je(n || t);
                    q(e, i.asyncContext, i.callback, null === r || void 0 === r ? t : r, void 0, void 0, void 0)
                }
                )).catch((function(e) {
                    return a(e, i.callback, i.asyncContext)
                }
                ));
            else {
                var o = Je(t);
                q(e, i.asyncContext, i.callback, null === o || void 0 === o ? t : o, void 0, void 0, void 0)
            }
        }
        function He(e) {
            null !== e && null !== e && (r(e.toRecipients) || me(e.toRecipients, "toRecipients"),
            r(e.ccRecipients) || me(e.ccRecipients, "ccRecipients"),
            r(e.bccRecipients) || me(e.bccRecipients, "bccRecipients"),
            r(e.htmlBody) || ke(e.htmlBody, 0, 32768, "htmlBody"),
            r(e.subject) || ke(e.subject, 0, 255, "subject"))
        }
        function Je(e) {
            var t = JSON.parse(JSON.stringify(e));
            if (!r(e)) {
                e.toRecipients && (t.toRecipients = pe(e.toRecipients, "toRecipients")),
                e.ccRecipients && (t.ccRecipients = pe(e.ccRecipients, "ccRecipients")),
                e.bccRecipients && (t.bccRecipients = pe(e.bccRecipients, "bccRecipients"));
                var n = function(e) {
                    var t = [];
                    e.attachments && ze(t = e.attachments);
                    return t
                }(e);
                e.attachments && (t.attachments = qe(n))
            }
            return t
        }
        function ze(e) {
            if (!r(e) && !Array.isArray(e))
                throw X("attachments")
        }
        function qe(e) {
            for (var t = [], n = 0; n < e.length; n++) {
                if ("object" !== typeof e[n])
                    throw X("attachments");
                var r = e[n];
                Ve(r),
                t.push(Ye(r))
            }
            return t
        }
        function Ve(e) {
            if ("object" !== typeof e)
                throw X("attachments");
            if (!e.type || !e.name)
                throw X("attachments");
            if (!e.url && !e.itemId && !e.base64file)
                throw X("attachments")
        }
        function Ye(e) {
            var t, n = null;
            if (t = N("win32"),
            e.type === ce.AttachmentType.File) {
                var r = e.url
                  , i = e.name
                  , a = !!e.isInline;
                !function(e, t) {
                    if ("string" !== typeof e && "string" !== typeof t)
                        throw X("attachments");
                    if (e.length > 2048)
                        throw te("attachments", e.length, o("l_AttachmentUrlTooLong_Text"));
                    Ge(t)
                }(r, i),
                n = [ce.AttachmentType.File, i, r, a]
            } else if (e.type === ce.AttachmentType.Item) {
                var s = ge(e.itemId)
                  , c = e.name;
                !function(e, t) {
                    if ("string" !== typeof e || "string" !== typeof t)
                        throw X("attachments");
                    if (e.length > 200)
                        throw te("attachments", e.length, o("l_AttachmentItemIdTooLong_Text"));
                    Ge(t)
                }(s, c),
                n = [ce.AttachmentType.Item, c, s]
            } else {
                if (!E("ReplyFormBase64Support") && t || e.type !== ce.AttachmentType.Base64)
                    throw X("attachments");
                var d = e.base64file
                  , l = e.name;
                a = !!e.isInline;
                !function(e, t) {
                    if ("string" !== typeof e || "string" !== typeof t)
                        throw X("attachments");
                    if (e.length > 27892122)
                        throw te("attachments", e.length, o("l_AttachmentExceededSize_Text"));
                    Ge(t)
                }(d, l),
                n = [ce.AttachmentType.Base64, l, d, a]
            }
            return n
        }
        function Ge(e) {
            if (e.length > 255)
                throw te("attachments", e.length, o("l_AttachmentNameTooLong_Text"))
        }
        var Ze = n(0);
        function Ke(e, t, n) {
            var r = void 0;
            return so() === Ze.AppName.Outlook && void 0 !== e.error && void 0 !== e.errorCode && e.error && 9030 === e.errorCode ? r = V(void 0, Ze.DDA.AsyncResultEnum.ErrorCode.Failed, e.errorCode, t, e.errorMessage, e.errorName) : n && n !== v.noError ? (r = V(void 0, Ze.DDA.AsyncResultEnum.ErrorCode.Failed, 9017, t, o("l_InternalProtocolError_Text").replace("{0}", n))) && (r.diagnostics = {
                InvokeCodeResult: n
            }) : (r = e.wasSuccessful ? V(e.token, Ze.DDA.AsyncResultEnum.ErrorCode.Success, 0, t) : V(void 0, Ze.DDA.AsyncResultEnum.ErrorCode.Failed, e.errorCode, t, e.errorMessage, e.errorName),
            e.diagnostics && (r.diagnostics = e.diagnostics)),
            r
        }
        function $e() {
            for (var e = [], t = 0; t < arguments.length; t++)
                e[t] = arguments[t];
            re(1, "mailbox.getCallbackTokenAsync");
            var n = ie(e, !0, !0)
              , r = !1;
            if (n.options && n.options.isRest && (r = !0),
            oo() && (!r || $() < 3))
                throw Q(o("l_TokenAccessDeniedWithoutItemContext_Text"));
            q(12, n.asyncContext, n.callback, {
                isRest: r
            }, void 0, Ke, void 0)
        }
        function Qe() {
            for (var e = [], t = 0; t < arguments.length; t++)
                e[t] = arguments[t];
            re(1, "mailbox.getUserIdentityToken");
            var n = ie(e, !0, !0);
            q(2, n.asyncContext, n.callback, void 0, void 0, Ke, void 0)
        }
        var Xe = function(e) {
            if (!r(e)) {
                var t = ao("hostVersion").split(".")
                  , n = e.split(".")
                  , i = 0;
                if (t.length >= 4 && n.length >= 4) {
                    for (var a = 0; a < 4; a++) {
                        var o = parseInt(t[a])
                          , s = parseInt(n[a]);
                        if (isNaN(o) || isNaN(s) || o < s)
                            return !1;
                        if (o > s)
                            return !0;
                        i++
                    }
                    return 4 == i
                }
            }
            return !1
        }
          , et = n(0);
        function tt(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            re(3, "mailbox.makeEwsRequest");
            var r = ie(t, !0, !0);
            if (null === e || void 0 === e)
                throw ee("data");
            if ("string" !== typeof e)
                throw ne("data", typeof e, "string");
            if (so() == et.AppName.Outlook && "win32" == et._OfficeAppFactory.getHostInfo().hostPlatform && Xe("16.0.16224.10000")) {
                if (e.length > 5242880)
                    throw X("data", o("l_NewEwsRequestOversized_Text"))
            } else if (e.length > 1e6)
                throw X("data", o("l_EwsRequestOversized_Text"));
            q(5, r.asyncContext, r.callback, {
                body: e
            }, void 0, nt, void 0)
        }
        function nt(e, t, n) {
            return n && n !== v.noError ? V(void 0, et.DDA.AsyncResultEnum.ErrorCode.Failed, 9017, t, o("l_InternalProtocolError_Text").replace("{0}", n)) : !1 === e.wasProxySuccessful ? V(void 0, et.DDA.AsyncResultEnum.ErrorCode.Failed, 9020, t, e.errorMessage) : V(e.body, et.DDA.AsyncResultEnum.ErrorCode.Success, 0, t)
        }
        var rt = function(e, t) {
            var n = Object.keys(t)
              , r = n.map((function(e) {
                return {
                    value: t[e],
                    writable: !1
                }
            }
            ))
              , i = {};
            return n.forEach((function(e, t) {
                i[e] = r[t]
            }
            )),
            Object.defineProperties(e, i)
        }
          , it = n(0)
          , at = function() {
            switch (so()) {
            case it.AppName.Outlook:
                return "Outlook";
            case it.AppName.OutlookWebApp:
                return ot() ? "newOutlookWindows" : "OutlookWebApp";
            case it.AppName.OutlookIOS:
                return "OutlookIOS";
            case it.AppName.OutlookAndroid:
                return "OutlookAndroid";
            default:
                return
            }
        };
        var ot = function() {
            return 0 != (it._OfficeAppFactory.getHostInfo().flags & it.HostInfoFlags.IsMonarch)
        };
        var st = ce.CategoryColor
          , ct = [st.None, st.Preset0, st.Preset1, st.Preset2, st.Preset3, st.Preset4, st.Preset5, st.Preset6, st.Preset7, st.Preset8, st.Preset9, st.Preset10, st.Preset11, st.Preset12, st.Preset13, st.Preset14, st.Preset15, st.Preset16, st.Preset17, st.Preset18, st.Preset19, st.Preset20, st.Preset21, st.Preset22, st.Preset23, st.Preset24];
        function dt(e) {
            if (!e)
                throw X("categoryDetails");
            if (!Array.isArray(e))
                throw ne("categoryDetails", typeof e, typeof []);
            if (0 === e.length)
                throw X("categoryDetails");
            e.forEach(lt)
        }
        function lt(e) {
            if (!e)
                throw X("categoryDetails");
            if (!e.color || !e.displayName)
                throw X("categoryDetails");
            if ("string" !== typeof e.color)
                throw ne("categoryDetails.color", typeof e.color, "string");
            if ("string" !== typeof e.displayName)
                throw ne("categoryDetails.displayName", typeof e.displayName, "string");
            if (e.displayName.length > 255)
                throw te("categoryDetails.displayName", e.displayName.length);
            if (-1 === ct.indexOf(e.color))
                throw X("categoryDetails.color")
        }
        function ut(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            re(3, "masterCategories.addAsync");
            var r = ie(t, !1, !1)
              , i = {
                categoryDetails: e
            };
            dt(e),
            q(161, r.asyncContext, r.callback, i, void 0, void 0, void 0)
        }
        function ft() {
            for (var e = [], t = 0; t < arguments.length; t++)
                e[t] = arguments[t];
            re(3, "masterCategories.getAsync");
            var n = ie(e, !0, !1);
            q(160, n.asyncContext, n.callback, void 0, void 0, void 0, void 0)
        }
        function mt(e) {
            if (!e)
                throw X("categories");
            if (!Array.isArray(e))
                throw ne("categories", typeof e, typeof Array);
            if (0 === e.length)
                throw X("categories");
            e.forEach(pt)
        }
        function pt(e) {
            if (!e)
                throw X("categories");
            if ("string" !== typeof e)
                throw ne("categories", typeof e, "string");
            if (e.length > 255)
                throw te("categories", e.length)
        }
        function yt(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            re(3, "masterCategories.removeAsync");
            var r = ie(t, !1, !1)
              , i = {
                categories: e
            };
            mt(e),
            q(162, r.asyncContext, r.callback, i, void 0, void 0, void 0)
        }
        function vt() {
            q(42, void 0, void 0, void 0, void 0, void 0, void 0)
        }
        var gt, ht = function() {
            return ao("itemType")
        };
        !function(e) {
            e[e.Message = 1] = "Message",
            e[e.Appointment = 2] = "Appointment",
            e[e.MeetingRequest = 3] = "MeetingRequest",
            e[e.MessageCompose = 4] = "MessageCompose",
            e[e.AppointmentCompose = 5] = "AppointmentCompose",
            e[e.ItemLess = 6] = "ItemLess"
        }(gt || (gt = {}));
        var At = n(0);
        function Tt(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(1, "item.getInitializationContext");
                var r = ie(t, !0, !1);
                N("win32") && Xe("16.0.17215.10000") ? q(99, r.asyncContext, r.callback, void 0, void 0, St, e) : q(99, r.asyncContext, r.callback, void 0, void 0, void 0, void 0)
            }
        }
        function St(e, t, n) {
            return n && n !== v.noError ? V(void 0, At.DDA.AsyncResultEnum.ErrorCode.Failed, 9017, t, o("l_InternalProtocolError_Text").replace("{0}", n)) : e.wasSuccessful ? "" === e.data ? V(void 0, At.DDA.AsyncResultEnum.ErrorCode.Success, 0, t) : V(JSON.parse(e.data), At.DDA.AsyncResultEnum.ErrorCode.Success, 0, t) : V(void 0, At.DDA.AsyncResultEnum.ErrorCode.Failed, e.errorCode, t, e.errorMessage, e.errorName)
        }
        var bt;
        function Dt(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            re(1, "item.saveCustomProperties");
            var r = this && this.isLoadedItem
              , i = ie(t, !1, !0);
            Ct(e),
            q(4, i.asyncContext, i.callback, {
                customProperties: e
            }, void 0, void 0, r)
        }
        function Ct(e) {
            if (JSON.stringify(e).length > 2500)
                throw te("customProperties")
        }
        !function(e) {
            e[e.NonTransmittable = 0] = "NonTransmittable"
        }(bt || (bt = {}));
        var xt = function() {
            for (var e = 0, t = 0, n = arguments.length; t < n; t++)
                e += arguments[t].length;
            var r = Array(e)
              , i = 0;
            for (t = 0; t < n; t++)
                for (var a = arguments[t], o = 0, s = a.length; o < s; o++,
                i++)
                    r[i] = a[o];
            return r
        }
          , wt = function() {
            function e(e, t) {
                if (this.isLoadedItem = !1,
                r(e) && ee("data"),
                Array.isArray(e)) {
                    var n = e;
                    if (!(n.length > bt.NonTransmittable))
                        throw X("data");
                    e = n[bt.NonTransmittable]
                } else
                    this.rawData = e;
                this.isLoadedItem = 1 == t
            }
            return e.prototype.get = function(e) {
                var t = this.rawData[e];
                if ("string" === typeof t) {
                    var n = t;
                    if (n.length > "Date(".length + ")".length && n.startsWith("Date(") && n.endsWith(")")) {
                        var i = n.substring("Date(".length, n.length - 1)
                          , a = parseInt(i);
                        if (!isNaN(a)) {
                            var o = new Date(a);
                            r(o) || (t = o)
                        }
                    }
                }
                return t
            }
            ,
            e.prototype.set = function(e, t) {
                Oe(t) && (t = "Date(" + t.getTime() + ")"),
                this.rawData[e] = t
            }
            ,
            e.prototype.remove = function(e) {
                delete this.rawData[e]
            }
            ,
            e.prototype.saveAsync = function() {
                for (var e = [], t = 0; t < arguments.length; t++)
                    e[t] = arguments[t];
                Dt.apply(void 0, xt([this.rawData], e))
            }
            ,
            e.prototype.getAll = function() {
                var e = this
                  , t = {};
                return Object.keys(this.rawData).forEach((function(n) {
                    t[n] = e.get(n)
                }
                )),
                t
            }
            ,
            e
        }()
          , Et = n(0);
        function kt(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                var r = ie(t, !0, !0);
                q(3, r.asyncContext, r.callback, void 0, void 0, Ot, e)
            }
        }
        function Ot(e, t, n) {
            if ("undefined" !== typeof n && n !== v.noError)
                return V(void 0, Et.DDA.AsyncResultEnum.ErrorCode.Failed, 9017, t, o("l_InternalProtocolError_Text").replace("{0}", n));
            if (e.wasSuccessful) {
                var r = JSON.parse(e.customProperties)
                  , i = this && this.isLoadedItem;
                return V(new wt(r,i), Et.DDA.AsyncResultEnum.ErrorCode.Success, 0, t)
            }
            return V(void 0, Et.DDA.AsyncResultEnum.ErrorCode.Failed, 9020, t, e.errorMessage)
        }
        var It, Mt, _t = n(0);
        function Nt(e, t) {
            t.options && "string" === typeof t.options.coercionType ? e.coercionType = Ft(t.options.coercionType) : e.coercionType = Mt.Text
        }
        function Ft(e) {
            return e === de.Html ? Mt.Html : e === de.Text ? Mt.Text : void 0
        }
        function Pt(e) {
            e.callback && e.callback(V(void 0, _t.DDA.AsyncResultEnum.ErrorCode.Failed, 1e3, e.asyncContext))
        }
        function Rt(e) {
            if (e !== ce.BodyMode.FullBody && e !== ce.BodyMode.HostConfig)
                throw X("bodyMode")
        }
        function Ut(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                re(1, "body.getAsync");
                var i = ie(n, !0, !1)
                  , a = ce.BodyMode.FullBody;
                i.options && void 0 !== i.options.bodyMode && (Rt(i.options.bodyMode),
                a = i.options.bodyMode);
                var o = {
                    coercionType: Ft(t),
                    bodyMode: a
                };
                if (void 0 === o.coercionType)
                    throw X("coercionType");
                q(37, i.asyncContext, i.callback, o, void 0, void 0, e)
            }
        }
        function Lt(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(1, "body.getTypeAsync");
                var r = ie(t, !0, !1);
                q(14, r.asyncContext, r.callback, void 0, void 0, void 0, e)
            }
        }
        !function(e) {
            e[e.Compose = 0] = "Compose",
            e[e.Read = 1] = "Read",
            e[e.ReadUI = 2] = "ReadUI"
        }(It || (It = {})),
        function(e) {
            e[e.Text = 0] = "Text",
            e[e.Html = 3] = "Html"
        }(Mt || (Mt = {}));
        function jt(e) {
            if ("string" !== typeof e.data)
                throw ne("data", typeof e.data, "string");
            if (e.data.length > 1e6)
                throw te("data", e.data.length)
        }
        function Wt(e) {
            if ("string" !== typeof e.data)
                throw ne("data", typeof e.data, "string");
            if (e.data.length > 12e4)
                throw te("data", e.data.length)
        }
        var Bt = "setUIAsync";
        function Ht(e) {
            throw Q("The feature {0}, is only enabled on the beta api endpoint".replace("{0}", e), {
                name: "Sys.FeatureNotEnabled"
            })
        }
        var Jt, zt = n(0);
        function qt(e, t) {
            var n = V(void 0, zt.DDA.AsyncResultEnum.ErrorCode.Failed, 5e3, e, "");
            t && setTimeout((function() {
                t && t(n)
            }
            ), 0)
        }
        function Vt(e, t) {
            return function(n) {
                for (var r = [], i = 1; i < arguments.length; i++)
                    r[i - 1] = arguments[i];
                if (38 == e)
                    re(2, "body.setAsync");
                else {
                    if (206 != e)
                        throw "Unexpected dispid";
                    Ht(Bt),
                    re(2, "display.body.setAsync")
                }
                var a = ie(r, !1, !1);
                if (t)
                    qt(a.asyncContext, a.callback);
                else {
                    var o = ce.BodyMode.FullBody;
                    a.options && void 0 !== a.options.bodyMode && (Rt(a.options.bodyMode),
                    o = a.options.bodyMode);
                    var s = {
                        data: n,
                        bodyMode: o
                    };
                    jt(s),
                    Nt(s, a),
                    void 0 !== s.coercionType ? q(e, a.asyncContext, a.callback, s, void 0, void 0, t) : Pt(a)
                }
            }
        }
        function Yt(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                re(2, "body.prependAsync");
                var i = ie(n, !1, !1);
                if (e)
                    qt(i.asyncContext, i.callback);
                else {
                    var a = {
                        data: t
                    };
                    jt(a),
                    Nt(a, i),
                    void 0 !== a.coercionType ? q(23, i.asyncContext, i.callback, a, void 0, void 0, e) : Pt(i)
                }
            }
        }
        function Gt(e, t, n, i, a) {
            re(2, t);
            var o = ie(i, !1, !1);
            if (a)
                qt(o.asyncContext, o.callback);
            else {
                var s = {
                    appendTxt: n
                };
                r(s.appendTxt) ? s.appendTxt = "" : function(e) {
                    if ("string" !== typeof e.appendTxt)
                        throw ne("data", typeof e.appendTxt, "string");
                    if (e.appendTxt.length > 5e3)
                        throw te("data", e.appendTxt.length)
                }(s),
                Nt(s, o),
                void 0 !== s.coercionType ? q(e, o.asyncContext, o.callback, s, void 0, void 0, a) : Pt(o)
            }
        }
        function Zt(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                Gt(100, "body.appendOnSendAsync", t, n, e)
            }
        }
        function Kt(e, t) {
            return function(n) {
                for (var r = [], i = 1; i < arguments.length; i++)
                    r[i - 1] = arguments[i];
                var a = ie(r, !1, !1);
                if (E("MultiSelectV2") && t)
                    qt(a.asyncContext, a.callback);
                else {
                    re(2, "body.setSelectedDataAsync");
                    var o = {
                        data: n
                    };
                    jt(o),
                    Nt(o, a),
                    void 0 !== o.coercionType ? q(e, a.asyncContext, a.callback, o, void 0, void 0, void 0) : Pt(a)
                }
            }
        }
        function $t(e) {
            return function(t) {
                for (var n = [], i = 1; i < arguments.length; i++)
                    n[i - 1] = arguments[i];
                re(2, "item.body.setSignatureAsync");
                var a = ie(n, !1, !1);
                if (e)
                    qt(a.asyncContext, a.callback);
                else {
                    var o = {
                        data: t
                    };
                    r(o.data) ? o.data = "" : Wt(o),
                    Nt(o, a),
                    void 0 !== o.coercionType ? q(173, a.asyncContext, a.callback, o, void 0, void 0, e) : Pt(a)
                }
            }
        }
        function Qt(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                Gt(197, "body.prependOnSendAsync", t, n, e)
            }
        }
        function Xt(e, t) {
            return !E("MultiSelectV2") || r(t) ? ao(e) : t && t[e]
        }
        function en(e, t, n) {
            var r = rt({}, {});
            if (e == It.Compose)
                rt(r, {
                    appendOnSendAsync: Zt(t),
                    getTypeAsync: Lt(t),
                    prependAsync: Yt(t),
                    setAsync: Vt(38, t),
                    setSelectedDataAsync: Kt(13, t),
                    setSignatureAsync: $t(t),
                    prependOnSendAsync: Qt(t),
                    getAsync: Ut(t)
                });
            else if (e == It.Read)
                rt(r, {
                    getAsync: Ut(t),
                    type: Xt("bodyType", n)
                });
            else {
                if (e != It.ReadUI)
                    throw "Unexpected ItemSurfaceType";
                rt(r, {
                    setAsync: Vt(206, t)
                })
            }
            return r
        }
        function tn(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(1, "item.getAllInternetHeadersAsync");
                var r = ie(t, !0, !1);
                q(168, r.asyncContext, r.callback, void 0, void 0, void 0, e)
            }
        }
        function nn(e, t) {
            if (r(t) || "" === t)
                throw ee(e);
            if ("string" !== typeof t)
                throw ne(e, typeof t, "string")
        }
        !function(e) {
            e[e.informationalMessage = 0] = "informationalMessage",
            e[e.progressIndicator = 1] = "progressIndicator",
            e[e.errorMessage = 2] = "errorMessage",
            e[e.insightMessage = 3] = "insightMessage"
        }(Jt || (Jt = {}));
        var rn = n(0);
        function an(e) {
            if (nn("key", e),
            e.length > 32)
                throw te("key", e.length)
        }
        function on(e) {
            if (nn("type", e.type),
            e.type === ce.ItemNotificationMessageType.InformationalMessage) {
                if (nn("icon", e.icon),
                e.icon.length > 32)
                    throw te("icon", e.icon.length);
                if (r(e.persistent))
                    throw ee("persistent");
                if ("boolean" !== typeof e.persistent)
                    throw ne("persistent", typeof e.persistent, "boolean");
                if (!r(e.actions))
                    throw X("actions", o("l_ActionsDefinitionWrongNotificationMessageError_Text"))
            } else if (e.type === ce.ItemNotificationMessageType.InsightMessage)
                !function(e) {
                    if (nn("icon", e.icon),
                    e.icon.length > 32)
                        throw te("icon", e.icon.length);
                    if ((!N("win32") || !Xe("16.0.14620.10000")) && !r(e.persistent))
                        throw X("persistent");
                    if (r(e.actions))
                        throw ee("actions");
                    !function(e) {
                        var t = function(e) {
                            var t = null;
                            if (!Array.isArray(e))
                                throw X("actions");
                            if (1 === e.length)
                                t = e[0];
                            else if (e.length > 1)
                                throw X("actions", o("l_ActionsDefinitionMultipleActionsError_Text"));
                            return t
                        }(e);
                        if (r(t))
                            return;
                        (function(e) {
                            if (r(e.actionType))
                                throw ee("actionType");
                            var t = e.actionType
                              , n = ["showTaskPane"];
                            rn.AppName.OutlookWebApp === so() && n.push("executeFunction");
                            if (-1 === n.indexOf(t))
                                throw X("actionType", o("l_InvalidActionType_Text"));
                            if (r(e.commandId) || "string" !== typeof e.commandId || "" === e.commandId)
                                throw X("commandId", o("l_InvalidCommandIdError_Text"))
                        }
                        )(t),
                        function(e) {
                            if (r(e.actionText) || "" === e.actionText || "string" !== typeof e.actionText)
                                throw ee("actionText");
                            if (e.actionText.length > 30)
                                throw te("actionText", e.actionText.length)
                        }(t)
                    }(e.actions)
                }(e);
            else {
                if (!r(e.icon))
                    throw X("icon");
                if (!r(e.persistent))
                    throw X("persistent");
                if (!r(e.actions))
                    throw X("actions", o("l_ActionsDefinitionWrongNotificationMessageError_Text"))
            }
            if (nn("message", e.message),
            e.message.length > 150)
                throw te("message", e.message.length)
        }
        function sn(e) {
            return function(t, n) {
                for (var i = [], a = 2; a < arguments.length; a++)
                    i[a - 2] = arguments[a];
                re(1, "notificationMessages.addAsync");
                var o = ie(i, !1, !1);
                if (e)
                    qt(o.asyncContext, o.callback);
                else {
                    an(t),
                    on(n);
                    var s, c, d = Jt[n.type];
                    if (r(d))
                        throw X("type");
                    s = N("win32") && Xe("16.0.17215.10000"),
                    c = N("win32") && Xe("16.0.17803.10000");
                    var l, u = w("notificationActionsPassByValue"), f = c && (!0 === u || void 0 === u), m = n.message, p = n.icon, y = n.persistent;
                    null === (l = n.actions && f ? JSON.parse(JSON.stringify(n.actions)) : n.actions) || void 0 === l || l.forEach((function(e) {
                        if (s)
                            try {
                                e.contextData = JSON.stringify(e.contextData)
                            } catch (t) {
                                e.contextData = void 0
                            }
                        else
                            void 0 === e.contextData || null !== e.contextData && "" !== e.contextData || (e.contextData = "{}")
                    }
                    ));
                    var v = {
                        key: t,
                        message: m,
                        type: d,
                        icon: p,
                        persistent: y,
                        actions: l
                    };
                    q(33, o.asyncContext, o.callback, v, void 0, void 0, e)
                }
            }
        }
        var cn = n(0);
        function dn(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(1, "notificationMessages.getAllAsync");
                var r, i = ie(t, !0, !1);
                r = N("win32") && Xe("16.0.17215.10000"),
                q(34, i.asyncContext, i.callback, void 0, void 0, r ? ln : void 0, e)
            }
        }
        function ln(e, t, n) {
            if (n && n !== v.noError)
                return V(void 0, cn.DDA.AsyncResultEnum.ErrorCode.Failed, 9017, t, o("l_InternalProtocolError_Text").replace("{0}", n));
            if (e.wasSuccessful) {
                for (var r = [], i = 0; i < e.data.length; i++) {
                    var a = e.data[i];
                    a.action && void 0 !== a.action.contextData && (a.action.contextData = JSON.parse(a.action.contextData)),
                    r.push(a)
                }
                return V(r, cn.DDA.AsyncResultEnum.ErrorCode.Success, 0, t)
            }
            return V(void 0, cn.DDA.AsyncResultEnum.ErrorCode.Failed, e.errorCode, t, e.errorMessage, e.errorName)
        }
        function un(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                re(1, "notificationMessages.removeAsync");
                var i = ie(n, !1, !1);
                if (e)
                    qt(i.asyncContext, i.callback);
                else {
                    an(t);
                    var a = {
                        key: t
                    };
                    q(36, i.asyncContext, i.callback, a, void 0, void 0, e)
                }
            }
        }
        function fn(e) {
            return function(t, n) {
                for (var i = [], a = 2; a < arguments.length; a++)
                    i[a - 2] = arguments[a];
                re(1, "notificationMessages.replaceAsync");
                var o = ie(i, !1, !1);
                if (e)
                    qt(o.asyncContext, o.callback);
                else {
                    an(t),
                    on(n);
                    var s, c, d = Jt[n.type];
                    if (r(d))
                        throw X("type");
                    s = N("win32") && Xe("16.0.17215.10000"),
                    c = N("win32") && Xe("16.0.17803.10000");
                    var l, u = w("notificationActionsPassByValue"), f = c && (!0 === u || void 0 === u), m = n.message, p = n.icon, y = n.persistent;
                    null === (l = n.actions && f ? JSON.parse(JSON.stringify(n.actions)) : n.actions) || void 0 === l || l.forEach((function(e) {
                        if (s)
                            try {
                                e.contextData = JSON.stringify(e.contextData)
                            } catch (t) {
                                e.contextData = void 0
                            }
                        else
                            void 0 === e.contextData || null !== e.contextData && "" !== e.contextData || (e.contextData = "{}")
                    }
                    ));
                    var v = {
                        key: t,
                        message: m,
                        type: d,
                        icon: p,
                        persistent: y,
                        actions: l
                    };
                    q(35, o.asyncContext, o.callback, v, void 0, void 0, e)
                }
            }
        }
        function mn(e) {
            return rt({}, {
                addAsync: sn(e),
                getAllAsync: dn(e),
                removeAsync: un(e),
                replaceAsync: fn(e)
            })
        }
        function pn(e) {
            r(e) || ae(e.length, 0, 32768, "htmlBody")
        }
        function yn(e) {
            var t = "";
            return e.htmlBody && (!function(e) {
                if ("string" !== typeof e)
                    throw ne("htmlBody", typeof e, "string");
                if (r(e))
                    throw ee("htmlBody");
                ae(e.length, 0, 32768, "htmlBody")
            }(e.htmlBody),
            t = e.htmlBody),
            t
        }
        function vn(e) {
            var t = [];
            return e.attachments && ze(t = e.attachments),
            t
        }
        function gn(e) {
            var t = [];
            return r(e.options) || (t[0] = e.options),
            r(e.callback) || (t[t.length] = e.callback),
            t
        }
        var hn = function() {
            for (var e = 0, t = 0, n = arguments.length; t < n; t++)
                e += arguments[t].length;
            var r = Array(e)
              , i = 0;
            for (t = 0; t < n; t++)
                for (var a = arguments[t], o = 0, s = a.length; o < s; o++,
                i++)
                    r[i] = a[o];
            return r
        };
        function An(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                Dn.apply(void 0, hn([!1, !1, t, e], n))
            }
        }
        function Tn(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                Dn.apply(void 0, hn([!0, !1, t, e], n))
            }
        }
        function Sn(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                Dn.apply(void 0, hn([!1, !0, t, e], n))
            }
        }
        function bn(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                Dn.apply(void 0, hn([!0, !0, t, e], n))
            }
        }
        function Dn(e, t, n, i) {
            for (var a = [], o = 4; o < arguments.length; o++)
                a[o - 4] = arguments[o];
            var s = ie(gn(n), !1, !1);
            if (E("MultiSelectV2") && i && !t)
                qt(s.asyncContext, s.callback);
            else {
                var c;
                re(1, "mailbox.displayReplyForm"),
                (r(s) || void 0 === s.options && void 0 === s.callback) && (s = ie(a, !1, !1));
                var d = {
                    formData: n
                }
                  , l = null
                  , u = null;
                if ("string" === typeof d.formData)
                    c = e ? t ? 184 : 11 : t ? 183 : 10,
                    pn(d.formData),
                    q(c, s.asyncContext, s.callback, {
                        htmlBody: d.formData
                    }, void 0, void 0, i);
                else {
                    if ("object" !== typeof d.formData)
                        throw X();
                    l = yn(d.formData);
                    var f = function(e, t) {
                        q(e ? t ? 182 : 31 : t ? 181 : 30, s.asyncContext, s.callback, {
                            htmlBody: l,
                            attachments: u
                        }, void 0, void 0, i)
                    };
                    N("web") ? Ue(d.formData).then((function(n) {
                        u = qe(vn(n)),
                        f(e, t)
                    }
                    )).catch((function(e) {
                        var t = 9007;
                        e instanceof Error && "Sys.ArgumentOutOfRangeException" === e.name && (t = 9e3);
                        var n = V(void 0, OSF.DDA.AsyncResultEnum.ErrorCode.Failed, t, s.asyncContext, "");
                        s.callback && setTimeout((function() {
                            s.callback && s.callback(n)
                        }
                        ), 0)
                    }
                    )) : (u = qe(vn(d.formData)),
                    f(e, t))
                }
            }
        }
        function Cn(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                re(2, "categories.addAsync");
                var i = ie(n, !1, !1)
                  , a = {
                    categories: t
                };
                mt(t),
                q(158, i.asyncContext, i.callback, a, void 0, void 0, e)
            }
        }
        function xn(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(1, "categories.getAsync");
                var r = ie(t, !0, !1);
                q(157, r.asyncContext, r.callback, void 0, void 0, void 0, e)
            }
        }
        function wn(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                re(2, "categories.removeAsync");
                var i = ie(n, !1, !1)
                  , a = {
                    categories: t
                };
                mt(t),
                q(159, i.asyncContext, i.callback, a, void 0, void 0, e)
            }
        }
        function En(e) {
            return rt({}, {
                addAsync: Cn(e),
                removeAsync: wn(e),
                getAsync: xn(e)
            })
        }
        function kn(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                re(1, "item.getAttachmentContentAsync");
                var i = ie(n, !0, !1)
                  , a = {
                    id: t
                };
                On(a),
                q(150, i.asyncContext, i.callback, a, void 0, void 0, e)
            }
        }
        function On(e) {
            nn("attachmentId", e.id)
        }
        var In = ce.Folder;
        function Mn(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                re(3, "item.move");
                var i = ie(n, !1, !1)
                  , a = {
                    destinationFolder: t
                };
                _n(t),
                q(101, i.asyncContext, i.callback, a, void 0, void 0, e)
            }
        }
        function _n(e) {
            if (e !== In.Inbox && e !== In.Junk && e !== In.DeletedItems)
                throw X("destinationFolder")
        }
        var Nn = ce.ResponseType
          , Fn = ce.RecipientType
          , Pn = [Nn.None, Nn.Organizer, Nn.Tentative, Nn.Accepted, Nn.Declined]
          , Rn = [Fn.Other, Fn.DistributionList, Fn.User, Fn.ExternalUser]
          , Un = function(e) {
            var t = e.appointmentResponse
              , n = e.recipientType
              , r = {
                emailAddress: e.address,
                displayName: e.name
            };
            return "number" === typeof e.appointmentResponse && (r.appointmentResponse = t < Pn.length ? Pn[t] : Nn.None),
            "number" === typeof e.recipientType && (r.recipientType = n < Rn.length ? Rn[n] : Fn.Other),
            r
        };
        function Ln(e) {
            return Un({
                name: e.Name || "",
                address: e.UserId || ""
            })
        }
        function jn(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(1, "delayDeliveryTime.getAsync");
                var r = ie(t, !0, !1);
                q(166, r.asyncContext, r.callback, void 0, void 0, void 0, e)
            }
        }
        function Wn(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                re(2, "delayDeliveryTime.setAsync");
                var i = ie(n, !1, !1);
                e ? qt(i.asyncContext, i.callback) : (Bn(t),
                q(167, i.asyncContext, i.callback, {
                    time: t.getTime()
                }, void 0, void 0, e))
            }
        }
        function Bn(e) {
            if (r(e))
                throw ee("dateTime", "You cannot conduct to a null dateTime");
            if (!Oe(e))
                throw ne("dateTime", typeof e, typeof Date);
            if (isNaN(e.getTime()))
                throw X("dateTime");
            ae(e.getTime(), -864e13, 864e13, "dateTime")
        }
        function Hn(e, t) {
            var n = rt({}, {
                getAsync: jn(t)
            });
            return e && rt(n, {
                setAsync: Wn(t)
            }),
            n
        }
        function Jn(e, t) {
            for (var n = e.length - 1; n >= 0; n--) {
                for (var r = !1, i = n - 1; i >= 0; i--)
                    if (t(e[n], e[i])) {
                        r = !0;
                        break
                    }
                r && e.splice(n, 1)
            }
            return e
        }
        var zn = function(e, t) {
            return e === t
        }
          , qn = function(e, t) {
            return e === t || !(!e || !t) && e.meetingString === t.meetingString
        }
          , Vn = function(e, t) {
            return e === t || !(!e || !t) && e.taskString === t.taskString
        }
          , Yn = function(e, t) {
            return e === t || !(!e || !t) && e.contactString === t.contactString
        };
        function Gn() {
            return !!ao("entities") && void 0 !== ao("entities").IsLegacyExtraction && ao("entities").IsLegacyExtraction
        }
        var Zn, Kn = new Date("0001-01-01T00:00:00Z");
        function $n(e, t) {
            if (!t)
                return e;
            var n = null;
            try {
                var i = new Date(t.getFullYear(),t.getMonth(),t.getDate(),0,0,0,0)
                  , a = function(e) {
                    var t = 0;
                    if (null == e)
                        return;
                    t = Gn() ? Xn(e) : Qn(e);
                    var n = (e.getTime() - t - Kn.getTime()) / 864e5;
                    if (n < 0)
                        return;
                    if (n >= 1 << 18)
                        return;
                    var r = n >> 15;
                    switch (n &= 32767,
                    r) {
                    case 0:
                        return function(e) {
                            var t = 0
                              , n = 0
                              , r = 0
                              , i = e >> 12 & 7;
                            if (4 == (4 & i)) {
                                if (t = e >> 5 & 127,
                                2 == (2 & i)) {
                                    if (1 == (1 & i))
                                        return;
                                    n = e >> 1 & 15
                                }
                            } else
                                2 == (2 & i) && (n = e >> 8 & 15),
                                1 == (1 & i) && (r = e >> 3 & 31);
                            return function(e, t, n) {
                                return {
                                    day: e,
                                    month: t,
                                    year: n % 100
                                }
                            }(r, n, t)
                        }(n);
                    case 1:
                        return function(e) {
                            var t = 15 & e
                              , n = function(e, t) {
                                var n = 1 << t - 1
                                  , r = (1 << t) - 1;
                                return (e & n) == n ? -(1 + (e ^ r)) : e
                            }(63 & (e >>= 4), 6)
                              , r = 7 & (e >>= 6)
                              , i = 3 & (e >>= 3);
                            try {
                                return function(e, t, n, r) {
                                    return {
                                        modifier: e,
                                        offset: t,
                                        unit: n,
                                        tag: r
                                    }
                                }(i, n, r, t)
                            } catch (e) {
                                return
                            }
                        }(n);
                    default:
                        return
                    }
                }(e);
                if (!a)
                    return e;
                var o = a;
                if (o.day && o.month && void 0 !== o.year)
                    n = function(e, t) {
                        var n, i = t.year, a = 0 == t.month ? e.getMonth() : t.month - 1, o = t.day;
                        if (0 == o)
                            return e;
                        r(i) ? (n = new Date(e.getFullYear(),a,o)).getTime() < e.getTime() && (n = new Date(e.getFullYear() + 1,a,o)) : n = new Date(i < 50 ? 2e3 + i : 1900 + i,a,o);
                        if (n.getMonth() != a)
                            return e;
                        return n
                    }(i, a);
                else {
                    var s = a;
                    n = void 0 !== s.modifier && void 0 !== s.offset && void 0 !== s.tag && void 0 !== s.unit ? function(e, t) {
                        var n;
                        switch (t.unit) {
                        case 0:
                            return (n = new Date(e.getFullYear(),e.getMonth(),e.getDate())).setDate(n.getDate() + t.offset),
                            n;
                        case 5:
                            return function(e, t, n) {
                                if (t > -5 && t < 5) {
                                    var r = 7 * t + ((n + 6) % 7 + 1 - e.getDay());
                                    return e.setDate(e.getDate() + r),
                                    e
                                }
                                return (r = (n - e.getDay()) % 7) < 0 && (r += 7),
                                e.setDate(e.getDate() + r),
                                e
                            }(e, t.offset, t.tag);
                        case 2:
                            var r = 1;
                            switch (t.modifier) {
                            case 1:
                                break;
                            case 2:
                                r = 16;
                                break;
                            default:
                                0 == t.offset && (r = e.getDate())
                            }
                            return (n = new Date(e.getFullYear(),e.getMonth(),r)).setMonth(n.getMonth() + t.offset),
                            n.getTime() < e.getTime() && n.setDate(n.getDate() + e.getDate() - 1),
                            n;
                        case 1:
                            if ((n = new Date(e.getFullYear(),e.getMonth(),e.getDate())).setDate(e.getDate() + 7 * t.offset),
                            1 == t.modifier || 0 == t.modifier)
                                return n.setDate(n.getDate() + 1 - n.getDay()),
                                n.getTime() < e.getTime() ? e : n;
                            if (2 == t.modifier)
                                return n.setDate(n.getDate() + 5 - n.getDay()),
                                n;
                            break;
                        case 4:
                            return function(e, t) {
                                var n, r, i;
                                if (n = e,
                                t.tag <= 0 || t.tag > 12 || t.offset <= 0 || t.offset > 5)
                                    return e;
                                var a = (12 + t.tag - n.getMonth() - 1) % 12;
                                if (r = new Date(n.getFullYear(),n.getMonth() + a,1),
                                1 == t.modifier)
                                    return 1 == t.offset && 6 != r.getDay() && 0 != r.getDay() ? r : ((i = new Date(r.getFullYear(),r.getMonth(),r.getDate())).setDate(i.getDate() + (1 - r.getDay() + 7) % 7),
                                    6 != r.getDay() && 0 != r.getDay() && 1 != r.getDay() && i.setDate(i.getDate() - 7),
                                    i.setDate(i.getDate() + 7 * (t.offset - 1)),
                                    i.getMonth() + 1 != t.tag ? e : i);
                                var o = 1 - (i = new Date(r.getFullYear(),r.getMonth(),(s = r.getMonth(),
                                c = r.getFullYear(),
                                32 - new Date(c,s,32).getDate()))).getDay();
                                return o > 0 && (o -= 7),
                                i.setDate(i.getDate() + o),
                                i.setDate(i.getDate() + 7 * (1 - t.offset)),
                                i.getMonth() + 1 != t.tag ? 6 != r.getDay() && 0 != r.getDay() ? r : e : i;
                                var s, c
                            }(e, t);
                        case 3:
                            if (t.offset > 0)
                                return new Date(e.getFullYear() + t.offset,0,1)
                        }
                        return e
                    }(i, a) : i
                }
                return isNaN(n.getTime()) ? t : (n.setMilliseconds(n.getMilliseconds() + (Gn() ? Xn(e) : Qn(e))),
                n)
            } catch (e) {
                return t
            }
        }
        function Qn(e) {
            var t = 0;
            return t += 3600 * e.getHours(),
            t += 60 * e.getMinutes(),
            t += e.getSeconds(),
            t *= 1e3,
            t += e.getMilliseconds()
        }
        function Xn(e) {
            var t = 0;
            return t += 3600 * e.getUTCHours(),
            t += 60 * e.getUTCMinutes(),
            t += e.getUTCSeconds(),
            t *= 1e3,
            t += e.getUTCMilliseconds()
        }
        function er(e) {
            for (var t = ao("timeZoneOffsets"), n = 0; n < t.length; n++) {
                var r = t[n]
                  , i = parseInt(r.start)
                  , a = parseInt(r.end);
                if (e.getTime() - i >= 0 && e.getTime() - a < 0)
                    return parseInt(r.offset)
            }
            throw X("input", o("l_InvalidDate_Text"))
        }
        function tr(e) {
            var t = function(e) {
                var t = new Date(e.year,e.month,e.date,e.hours,e.minutes,e.seconds,null === e.milliseconds ? 0 : e.milliseconds);
                if (isNaN(t.getTime()))
                    throw X("input", o("l_InvalidDate_Text"));
                return t
            }(e);
            if (!r(ao("timeZoneOffsets"))) {
                var n = er(t);
                t.setUTCMinutes(t.getUTCMinutes() - n),
                n = e.timezoneOffset ? e.timezoneOffset : -1 * t.getTimezoneOffset(),
                t.setUTCMinutes(t.getUTCMinutes() + n)
            }
            return t
        }
        function nr(e) {
            return {
                month: e.getMonth(),
                date: e.getDate(),
                year: e.getFullYear(),
                hours: e.getHours(),
                minutes: e.getMinutes(),
                seconds: e.getSeconds(),
                milliseconds: e.getMilliseconds()
            }
        }
        function rr(e) {
            return r(e) ? {
                addresses: [],
                emailAddresses: [],
                urls: [],
                taskSuggestions: [],
                meetingSuggestions: [],
                phoneNumbers: [],
                contacts: [],
                flightReservations: [],
                parcelDelivery: []
            } : {
                addresses: ir(e[Zn.address]),
                emailAddresses: ar(e[Zn.emailAddress]),
                urls: or(e[Zn.url]),
                taskSuggestions: sr(e[Zn.taskSuggestion]),
                meetingSuggestions: cr(e[Zn.meetingSuggestion]),
                phoneNumbers: lr(e[Zn.phoneNumber]),
                contacts: ur(e[Zn.contact]),
                flightReservations: fr(e[Zn.flightReservations]),
                parcelDelivery: fr(e[Zn.parcelDeliveries])
            }
        }
        !function(e) {
            e.meetingSuggestion = "MeetingSuggestions",
            e.taskSuggestion = "TaskSuggestions",
            e.address = "Addresses",
            e.emailAddress = "EmailAddresses",
            e.url = "Urls",
            e.phoneNumber = "PhoneNumbers",
            e.contact = "Contacts",
            e.flightReservations = "FlightReservations",
            e.parcelDeliveries = "ParcelDeliveries"
        }(Zn || (Zn = {}));
        var ir = function(e) {
            return Jn(e || [], zn)
        }
          , ar = function(e) {
            return 0 === $() ? [] : e || []
        }
          , or = function(e) {
            return e || []
        }
          , sr = function(e) {
            if (0 === $())
                return [];
            var t = e || [];
            return Jn(t = t.map((function(e) {
                return {
                    assignees: (e.Assignees || []).map(Ln),
                    taskString: e.TaskString
                }
            }
            )), Vn)
        }
          , cr = function(e) {
            if (0 === $())
                return [];
            var t = e || [];
            return Jn(t = t.map((function(e) {
                var t = "" !== e.StartTime ? dr(e.StartTime) : void 0
                  , n = "" !== e.EndTime ? dr(e.EndTime) : void 0;
                return {
                    meetingString: e.MeetingString,
                    attendees: (e.Attendees || []).map(Ln),
                    location: e.Location,
                    subject: e.Subject,
                    start: void 0 !== e.StartTime ? t : void 0,
                    end: void 0 !== e.EndTime ? n : void 0
                }
            }
            )), qn)
        };
        function dr(e) {
            var t = $n(new Date(e), new Date(ao("dateTimeSent")));
            return t.getTime() !== new Date(e).getTime() ? tr(nr(t)) : new Date(e)
        }
        var lr = function(e) {
            return (e || []).map((function(e) {
                return {
                    phoneString: e.PhoneString,
                    originalPhoneString: e.OriginalPhoneString,
                    type: e.Type
                }
            }
            ))
        }
          , ur = function(e) {
            if (0 === $())
                return [];
            var t = e || [];
            return Jn(t = t.map((function(e) {
                return {
                    personName: e.PersonName,
                    businessName: e.BusinessName,
                    phoneNumbers: lr(e.PhoneNumbers || []),
                    emailAddresses: e.EmailAddresses || [],
                    urls: e.Urls || [],
                    addresses: e.Addresses || [],
                    contactString: e.ContactString
                }
            }
            )), Yn)
        }
          , fr = function(e) {
            return 0 === $() ? [] : e || []
        }
          , mr = {
            meetingSuggestion: 1,
            taskSuggestion: 1,
            address: 0,
            emailAddress: 1,
            url: 0,
            phoneNumber: 0,
            contact: 1,
            flightReservations: 1,
            parcelDeliveries: 1
        }
          , pr = {
            meetingSuggestion: "meetingSuggestions",
            taskSuggestion: "taskSuggestions",
            address: "addresses",
            emailAddress: "emailAddresses",
            url: "urls",
            phoneNumber: "phoneNumbers",
            contact: "contacts",
            flightReservations: "flightReservations",
            parcelDeliveries: "parcelDeliveries"
        }
          , yr = function(e) {
            return function() {
                return rr(Xt("entities", e))
            }
        }
          , vr = function(e) {
            return function(t) {
                var n = rr(Xt("entities", e));
                re(void 0 !== mr[t] ? mr[t] : 1, t);
                var r = pr[t];
                return void 0 === r ? null : n[r]
            }
        }
          , gr = function(e) {
            return function(t) {
                return function(e, t) {
                    re(1, "item.getFilteredEntitiesByName");
                    var n = Object.keys(e).map((function(n) {
                        return e[n][t] ? {
                            entityType: n,
                            name: t,
                            entities: e[n][t]
                        } : void 0
                    }
                    )).filter((function(e) {
                        return void 0 !== e
                    }
                    ));
                    if (0 === n.length)
                        return null;
                    var r = n[0];
                    switch (r.entityType) {
                    case Zn.meetingSuggestion:
                        return cr(r.entities);
                    case Zn.address:
                        return ir(r.entities);
                    case Zn.contact:
                        return ur(r.entities);
                    case Zn.emailAddress:
                        return ar(r.entities);
                    case Zn.phoneNumber:
                        return lr(r.entities);
                    case Zn.taskSuggestion:
                        return sr(r.entities);
                    case Zn.url:
                        return or(r.entities);
                    default:
                        return fr(r.entities)
                    }
                }(Xt("filteredEntities", e), t)
            }
        }
          , hr = function(e) {
            return function() {
                return Xt("regExMatches", e)
            }
        }
          , Ar = function(e) {
            return function(t) {
                return (Xt("regExMatches", e) || {})[t]
            }
        }
          , Tr = function(e) {
            return function() {
                return rr(Xt("selectedEntities", e))
            }
        }
          , Sr = function(e) {
            return function() {
                return Xt("selectedRegExMatches", e)
            }
        };
        function br(e) {
            var t = [];
            if (0 === $())
                return [];
            if (e)
                for (var n = 0; n < e.length; n++)
                    if (e[n]) {
                        var r = Dr(e[n]);
                        t.push(r)
                    }
            return t
        }
        function Dr(e) {
            if (null !== e.attachmentType || void 0 !== e.attachmentType)
                switch (e.attachmentType) {
                case 0:
                    e.attachmentType = ce.AttachmentType.File;
                    break;
                case 1:
                    e.attachmentType = ce.AttachmentType.Item;
                    break;
                case 2:
                    e.attachmentType = ce.AttachmentType.Cloud
                }
            return e
        }
        function Cr(e) {
            return JSON.parse(JSON.stringify(e))
        }
        function xr(e) {
            return e < 0 && (e = 1),
            e < 10 ? "0" + e.toString() : e.toString()
        }
        function wr(e, t, n) {
            if (!Er(e, t, n))
                throw X("seriesTime", o("l_InvalidDate_Text"))
        }
        function Er(e, t, n) {
            return !(e < 1601 || t < 1 || t > 12 || n < 1 || n > 31)
        }
        var kr = function() {
            function e() {
                this.startYear = 0,
                this.startMonth = 0,
                this.startDay = 0,
                this.endYear = 0,
                this.endMonth = 0,
                this.endDay = 0,
                this.startTimeMinutes = 0,
                this.durationMinutes = 0
            }
            return e.prototype.getDuration = function() {
                return this.durationMinutes
            }
            ,
            e.prototype.getEndTime = function() {
                var e = this.startTimeMinutes + this.durationMinutes
                  , t = e % 60;
                return "T" + xr(Math.floor(e / 60) % 24) + ":" + xr(t) + ":00.000"
            }
            ,
            e.prototype.getEndDate = function() {
                return 0 === this.endYear && 0 === this.endMonth && 0 === this.endDay ? null : this.endYear.toString() + "-" + xr(this.endMonth) + "-" + xr(this.endDay)
            }
            ,
            e.prototype.getStartDate = function() {
                return this.startYear.toString() + "-" + xr(this.startMonth) + "-" + xr(this.startDay)
            }
            ,
            e.prototype.getStartTime = function() {
                var e = this.startTimeMinutes % 60;
                return "T" + xr(Math.floor(this.startTimeMinutes / 60)) + ":" + xr(e) + ":00.000"
            }
            ,
            e.prototype.setDuration = function(e) {
                if (!(e >= 0))
                    throw X(void 0, o("l_InvalidTime_Text"));
                this.durationMinutes = e
            }
            ,
            e.prototype.setEndDate = function(e, t, n) {
                null === e || r(t) || null === n ? null !== e ? this.setDateHelper(!1, e) : null == e && (this.endYear = 0,
                this.endMonth = 0,
                this.endDay = 0) : this.setDateHelper(!1, e, t, n)
            }
            ,
            e.prototype.setStartDate = function(e, t, n) {
                null === e || r(t) || null === n ? null !== e && this.setDateHelper(!0, e) : this.setDateHelper(!0, e, t, n)
            }
            ,
            e.prototype.setStartTime = function(e, t) {
                if (r(e) || r(t)) {
                    if (!r(e)) {
                        var n = e
                          , i = "2017-01-15" + n + "Z";
                        if (!new RegExp("^T[0-2]\\d:[0-5]\\d:[0-5]\\d\\.\\d{3}$").test(n))
                            throw X(void 0, o("l_InvalidTime_Text"));
                        var a = new Date(i);
                        if (r(a) || isNaN(a.getUTCHours()) || isNaN(a.getUTCMinutes()))
                            throw X(void 0, o("l_InvalidTime_Text"));
                        this.startTimeMinutes = 60 * a.getUTCHours() + a.getUTCMinutes()
                    }
                } else {
                    var s = 60 * e + t;
                    if (!(s >= 0))
                        throw X(void 0, o("l_InvalidTime_Text"));
                    this.startTimeMinutes = s
                }
            }
            ,
            e.prototype.isValid = function() {
                return !!Er(this.startYear, this.startMonth, this.startDay) && (!(0 !== this.endDay && 0 !== this.endMonth && 0 !== this.endYear && !Er(this.endYear, this.endMonth, this.endDay)) && !(this.startTimeMinutes < 0 || this.durationMinutes <= 0))
            }
            ,
            e.prototype.exportToSeriesTimeJson = function() {
                var e = {};
                return e.startYear = this.startYear,
                e.startMonth = this.startMonth,
                e.startDay = this.startDay,
                0 === this.endYear && 0 === this.endMonth && 0 === this.endDay ? e.noEndDate = !0 : (e.endYear = this.endYear,
                e.endMonth = this.endMonth,
                e.endDay = this.endDay),
                e.startTimeMin = this.startTimeMinutes,
                this.durationMinutes > 0 && (e.durationMin = this.durationMinutes),
                e
            }
            ,
            e.prototype.importFromSeriesTimeJsonObject = function(e) {
                this.startYear = e.startYear,
                this.startMonth = e.startMonth,
                this.startDay = e.startDay,
                null != e.noEndDate && "boolean" === typeof e.noEndDate ? (this.endYear = 0,
                this.endMonth = 0,
                this.endDay = 0) : (this.endYear = e.endYear,
                this.endMonth = e.endMonth,
                this.endDay = e.endDay),
                this.startTimeMinutes = e.startTimeMin,
                this.durationMinutes = e.durationMin
            }
            ,
            e.prototype.setDateHelper = function(e, t, n, i) {
                var a = 0
                  , s = 0
                  , c = 0;
                if (null === t || r(n) || null === i) {
                    if (null !== t) {
                        var d = t;
                        !function(e) {
                            if (!new RegExp("^\\d{4}-(?:[0]\\d|1[0-2])-(?:[0-2]\\d|3[01])$").test(e))
                                throw X("seriesTime", o("l_InvalidDate_Text"))
                        }(d);
                        var l = new Date(d);
                        null === l || isNaN(l.getUTCFullYear()) || isNaN(l.getUTCMonth()) || isNaN(l.getUTCDate()) || (wr(l.getUTCFullYear(), l.getUTCMonth() + 1, l.getUTCDate()),
                        a = l.getUTCFullYear(),
                        s = l.getUTCMonth() + 1,
                        c = l.getUTCDate())
                    }
                } else
                    wr(t, n + 1, i),
                    a = t,
                    s = n + 1,
                    c = i;
                0 !== a && 0 !== s && 0 !== c && (e ? (this.startYear = a,
                this.startMonth = s,
                this.startDay = c) : (this.endYear = a,
                this.endMonth = s,
                this.endDay = c))
            }
            ,
            e.prototype.isEndAfterStart = function() {
                if (0 === this.endYear && 0 === this.endMonth && 0 === this.endDay)
                    return !0;
                var e = new Date;
                e.setFullYear(this.startYear),
                e.setMonth(this.startMonth - 1),
                e.setDate(this.startDay);
                var t = new Date;
                return t.setFullYear(this.endYear),
                t.setMonth(this.endMonth - 1),
                t.setDate(this.endDay),
                t >= e
            }
            ,
            e
        }();
        function Or(e) {
            if (r(e) || r(e.seriesTimeJson))
                return e;
            var t = {
                recurrenceType: "",
                recurrenceProperties: null,
                recurrenceTimeZone: null
            }
              , n = new kr;
            return r(e.recurrenceProperties) || (t.recurrenceProperties = Cr(e.recurrenceProperties)),
            t.recurrenceType = e.recurrenceType,
            r(e.recurrenceTimeZone) || (t.recurrenceTimeZone = Cr(e.recurrenceTimeZone)),
            n.importFromSeriesTimeJsonObject(e.seriesTimeJson),
            t.seriesTime = n,
            t
        }
        function Ir(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(1, "getAsFileAsync");
                var r = ie(t, !0, !1);
                q(204, r.asyncContext, r.callback, void 0, void 0, void 0, e)
            }
        }
        function Mr(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(1, "subject.getAsync");
                var r = ie(t, !0, !1);
                q(18, r.asyncContext, r.callback, void 0, void 0, void 0, e)
            }
        }
        function _r(e) {
            if ("string" !== typeof e.subject)
                throw ne("subject", typeof e.subject, "string");
            ae(e.subject.length, 0, 255, "subject")
        }
        function Nr(e, t) {
            return function(n) {
                for (var r = [], i = 1; i < arguments.length; i++)
                    r[i - 1] = arguments[i];
                if (17 == e)
                    re(2, "subject.setAsync");
                else {
                    if (207 != e)
                        throw "Unexpected dispid";
                    Ht(Bt),
                    re(2, "display.subject.setAsync")
                }
                var a = ie(r, !1, !1);
                if (t)
                    qt(a.asyncContext, a.callback);
                else {
                    var o = {
                        subject: n
                    };
                    _r(o),
                    q(e, a.asyncContext, a.callback, o, void 0, void 0, t)
                }
            }
        }
        function Fr(e, t) {
            var n = rt({}, {
                getAsync: Mr(t)
            });
            return rt(n, e ? {
                setAsync: Nr(207, t)
            } : {
                setAsync: Nr(17, t)
            }),
            n
        }
        function Pr() {
            for (var e = [], t = 0; t < arguments.length; t++)
                e[t] = arguments[t];
            re(2, "item.unloadAsync");
            var n = ie(e, !1, !1);
            q(212, n.asyncContext, n.callback, void 0, void 0, void 0, void 0)
        }
        function Rr(e, t) {
            var n = Xt("sender", t)
              , r = Xt("from", t)
              , i = Xt("dateTimeCreated", t)
              , a = Xt("dateTimeModified", t)
              , o = Xt("end", t)
              , s = Xt("start", t)
              , c = rt({}, {
                attachments: br(Xt("attachments", t)),
                bcc: (Xt("bcc", t) || []).map(Un),
                body: en(It.Read, e, t),
                categories: En(e),
                cc: (Xt("cc", t) || []).map(Un),
                conversationId: Xt("conversationId", t),
                dateTimeCreated: i ? new Date(i) : void 0,
                dateTimeModified: a ? new Date(a) : void 0,
                display: rt({}, {
                    body: en(It.ReadUI),
                    subject: Fr(!0)
                }),
                end: o ? new Date(o) : void 0,
                from: r ? Un(r) : void 0,
                getAllInternetHeadersAsync: tn(e),
                internetMessageId: Xt("internetMessageId", t),
                itemClass: Xt("itemClass", t),
                itemId: Xt("id", t),
                itemType: "message",
                location: Xt("location", t),
                move: Mn(e),
                normalizedSubject: Xt("normalizedSubject", t),
                notificationMessages: mn(e),
                recurrence: Or(Xt("recurrence", t)),
                seriesId: Xt("seriesId", t),
                sender: n ? Un(n) : void 0,
                start: s ? new Date(s) : void 0,
                subject: Xt("subject", t),
                to: (Xt("to", t) || []).map(Un),
                displayReplyForm: An(e),
                displayReplyFormAsync: Sn(e),
                displayReplyAllForm: Tn(e),
                displayReplyAllFormAsync: bn(e),
                getAttachmentContentAsync: kn(e),
                getEntities: yr(t),
                getEntitiesByType: vr(t),
                getFilteredEntitiesByName: gr(t),
                getInitializationContextAsync: Tt(e),
                getRegExMatches: hr(t),
                getRegExMatchesByName: Ar(t),
                getSelectedEntities: Tr(t),
                getSelectedRegExMatches: Sr(t),
                loadCustomPropertiesAsync: kt(e),
                delayDeliveryTime: Hn(!1, e),
                isAllDayEvent: Xt("isAllDayEvent", t),
                sensitivity: Xt("sensitivity", t),
                getAsFileAsync: Ir(e)
            });
            return E("MultiSelectV2") && e && rt(c, {
                isLoadedItem: e,
                unloadAsync: Pr
            }),
            c
        }
        function Ur(e) {
            if (r(e) || "" === e || "string" !== typeof e)
                throw X("attachmentName");
            ae(e.length, 0, 255, "attachmentName")
        }
        function Lr(e) {
            return function(t, n) {
                for (var r = [], i = 2; i < arguments.length; i++)
                    r[i - 2] = arguments[i];
                re(2, "item.addBase64FileAttachmentAsync");
                var a = ie(r, !1, !1);
                if (e)
                    qt(a.asyncContext, a.callback);
                else {
                    var o = !1;
                    a.options && (o = !!a.options.isInline);
                    var s = {
                        base64String: t,
                        name: n,
                        isInline: o,
                        __timeout__: 6e5
                    };
                    jr(s),
                    q(148, a.asyncContext, a.callback, s, void 0, void 0, e)
                }
            }
        }
        function jr(e) {
            nn("base64Encoded", e.base64String),
            ae(e.base64String.length, 0, 27892122, "base64File"),
            Ur(e.name)
        }
        var Wr = function() {
            for (var e = 0, t = 0, n = arguments.length; t < n; t++)
                e += arguments[t].length;
            var r = Array(e)
              , i = 0;
            for (t = 0; t < n; t++)
                for (var a = arguments[t], o = 0, s = a.length; o < s; o++,
                i++)
                    r[i] = a[o];
            return r
        };
        function Br(e) {
            return function(t, n) {
                for (var r = [], i = 2; i < arguments.length; i++)
                    r[i - 2] = arguments[i];
                re(2, "item.addFileAttachmentAsync");
                var a = ie(r, !1, !1);
                if (e)
                    qt(a.asyncContext, a.callback);
                else {
                    var s = !1;
                    a.options && (s = !!a.options.isInline);
                    var c = n
                      , d = {
                        uri: t,
                        name: c,
                        isInline: s,
                        __timeout__: 6e5
                    };
                    if (Hr(d),
                    N("web"))
                        try {
                            fetch(t).then((function(e) {
                                return e.blob()
                            }
                            )).then((function(e) {
                                var t = new FileReader;
                                t.onloadend = function() {
                                    try {
                                        var e = t.result;
                                        Lr().apply(void 0, Wr([Jr(e), n], r))
                                    } catch (e) {
                                        if (e instanceof Error && "Sys.ArgumentOutOfRangeException" === e.name) {
                                            var i = V(void 0, OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9e3, a.asyncContext, "");
                                            return void (a.callback && setTimeout((function() {
                                                a.callback && a.callback(i)
                                            }
                                            ), 0))
                                        }
                                    }
                                }
                                ,
                                t.readAsDataURL(e)
                            }
                            )).catch((function(e) {
                                var t = V(void 0, OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9007, a.asyncContext, "");
                                a.callback && setTimeout((function() {
                                    a.callback && a.callback(t)
                                }
                                ), 0)
                            }
                            ))
                        } catch (e) {
                            var l = V(void 0, OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9007, a.asyncContext, "");
                            throw a.callback && setTimeout((function() {
                                a.callback && a.callback(l)
                            }
                            ), 0),
                            Q(o("l_AttachmentUploadGeneralFailure_Text"))
                        }
                    else
                        q(16, a.asyncContext, a.callback, d, void 0, void 0, e)
                }
            }
        }
        function Hr(e) {
            nn("uri", e.uri),
            ae(e.uri.length, 0, 2048, "uri"),
            Ur(e.name)
        }
        var Jr = function(e) {
            return e.replace("data:", "").replace(/^.+,/, "")
        };
        function zr(e, t) {
            for (var n = [], r = 2; r < arguments.length; r++)
                n[r - 2] = arguments[r];
            re(2, "item.addItemAttachmentAsync");
            var i = this && this.isLoadedItem
              , a = ie(n, !1, !1);
            if (i)
                qt(a.asyncContext, a.callback);
            else {
                var o = {
                    itemId: e,
                    name: t
                };
                qr(o),
                q(19, a.asyncContext, a.callback, {
                    itemId: ge(o.itemId),
                    name: o.name,
                    __timeout__: 6e5
                }, void 0, void 0, void 0)
            }
        }
        function qr(e) {
            nn("itemId", e.itemId),
            nn("attachmentName", e.name),
            ae(e.itemId.length, 0, 200, "itemId"),
            ae(e.name.length, 0, 255, "attachmentName")
        }
        function Vr() {
            q(41, void 0, void 0, void 0, void 0, void 0, void 0)
        }
        function Yr(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(1, "item.getAttachmentsAsync");
                var r = ie(t, !0, !1);
                q(149, r.asyncContext, r.callback, void 0, br, void 0, e)
            }
        }
        function Gr(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            re(2, "item.getSelectedDataAsync");
            var r = ie(t, !0, !1)
              , i = {
                coercionType: Ft(e)
            };
            if (void 0 === i.coercionType)
                throw X("coercionType");
            q(28, r.asyncContext, r.callback, i, void 0, void 0, void 0)
        }
        function Zr(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                var i = ie(n, !0, !1);
                if (e)
                    qt(i.asyncContext, i.callback);
                else {
                    re(2, "item.getSelectedDataAsync");
                    var a = {
                        coercionType: Ft(t)
                    };
                    if (void 0 === a.coercionType)
                        throw X("coercionType");
                    q(28, i.asyncContext, i.callback, a, void 0, void 0, e)
                }
            }
        }
        function Kr(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                re(2, "item.removeAttachmentAsync");
                var i = ie(n, !1, !1);
                if (e)
                    qt(i.asyncContext, i.callback);
                else {
                    var a = {
                        attachmentIndex: t
                    };
                    $r(a),
                    q(20, i.asyncContext, i.callback, a, void 0, void 0, e)
                }
            }
        }
        function $r(e) {
            nn("attachmentId", e.attachmentIndex),
            ae(e.attachmentIndex.length, 0, 200, "attachmentId")
        }
        function Qr(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(2, "item.saveAsync");
                var r = ie(t, !1, !1);
                q(32, r.asyncContext, r.callback, void 0, void 0, void 0, e)
            }
        }
        function Xr(e) {
            if (!Array.isArray(e.recipientArray))
                throw X("recipients");
            if (e.recipientArray.length > 100)
                throw te("recipients", e.recipientArray.length);
            var t = e.recipientArray.map((function(e) {
                if (r(e))
                    throw X("recipients");
                if ("string" === typeof e)
                    return ei(e, e),
                    ti(e, e);
                if ("object" === typeof e)
                    return ei(e.displayName, e.emailAddress),
                    ti(e.displayName, e.emailAddress);
                throw X("recipients")
            }
            ));
            e.recipientArray = t
        }
        function ei(e, t) {
            if (!e && !t)
                throw X("recipients");
            if ("string" === typeof e && e.length > 255)
                throw te("recipients", e.length, o("l_DisplayNameTooLong_Text"));
            if ("string" === typeof t && t.length > 571)
                throw te("recipients", t.length, o("l_EmailAddressTooLong_Text"));
            if ("string" !== typeof e && "string" !== typeof t)
                throw X("recipients")
        }
        function ti(e, t) {
            return {
                address: t,
                name: e
            }
        }
        function ni(e, t) {
            return function(n) {
                for (var r = [], i = 1; i < arguments.length; i++)
                    r[i - 1] = arguments[i];
                re(2, e + ".addAsync");
                var a = ie(r, !1, !1);
                if (t)
                    qt(a.asyncContext, a.callback);
                else {
                    var o = {
                        recipientField: Y[e],
                        recipientArray: n
                    };
                    Xr(o),
                    q(22, a.asyncContext, a.callback, o, void 0, void 0, t)
                }
            }
        }
        function ri(e, t) {
            return function() {
                for (var n = [], r = 0; r < arguments.length; r++)
                    n[r] = arguments[r];
                re(1, e + ".getAsync");
                var i = ie(n, !0, !1);
                q(15, i.asyncContext, i.callback, {
                    recipientField: Y[e]
                }, ii, void 0, t)
            }
        }
        function ii(e) {
            return null === e || void 0 === e ? [] : e.map((function(e) {
                return Un(e)
            }
            ))
        }
        function ai(e, t) {
            return function(n) {
                for (var r = [], i = 1; i < arguments.length; i++)
                    r[i - 1] = arguments[i];
                re(2, e + ".setAsync");
                var a = ie(r, !1, !1);
                if (t)
                    qt(a.asyncContext, a.callback);
                else {
                    var o = {
                        recipientField: Y[e],
                        recipientArray: n
                    };
                    Xr(o),
                    q(21, a.asyncContext, a.callback, o, void 0, void 0, t)
                }
            }
        }
        function oi(e, t) {
            return rt({}, {
                addAsync: ni(e, t),
                getAsync: ri(e, t),
                setAsync: ai(e, t)
            })
        }
        function si(e, t) {
            return function() {
                for (var n = [], r = 0; r < arguments.length; r++)
                    n[r] = arguments[r];
                re(1, e + ".getAsync");
                var i = ie(n, !0, !1);
                q(107, i.asyncContext, i.callback, void 0, ci, void 0, t)
            }
        }
        function ci(e) {
            return r(e) ? null : Un(e)
        }
        function di(e, t) {
            return rt({}, {
                getAsync: si(e, t)
            })
        }
        function li(e) {
            if (r(e))
                throw X("internetHeaders");
            if (!Array.isArray(e))
                throw ne("internetHeaders", typeof e, "Array");
            if (0 === e.length)
                throw X("internetHeaders");
            for (var t = 0, n = e; t < n.length; t++) {
                nn("internetHeaders", n[t])
            }
        }
        function ui(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            re(2, "internetHeaders.removeAsync");
            var r = this && this.isLoadedItem
              , i = ie(t, !1, !1);
            if (r)
                qt(i.asyncContext, i.callback);
            else {
                var a = {
                    internetHeaderKeys: e
                };
                fi(a),
                q(153, i.asyncContext, i.callback, a, void 0, void 0, r)
            }
        }
        function fi(e) {
            li(e.internetHeaderKeys)
        }
        function mi(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            re(1, "internetHeaders.getAsync");
            var r = this && this.isLoadedItem
              , i = ie(t, !0, !1)
              , a = {
                internetHeaderKeys: e
            };
            pi(a),
            q(151, i.asyncContext, i.callback, a, void 0, void 0, r)
        }
        function pi(e) {
            li(e.internetHeaderKeys)
        }
        var yi;
        function vi(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            re(2, "internetHeaders.setAsync");
            var r = this && this.isLoadedItem
              , i = ie(t, !1, !1);
            if (r)
                qt(i.asyncContext, i.callback);
            else {
                var a = {
                    internetHeaderNameValuePairs: e
                };
                gi(a),
                q(152, i.asyncContext, i.callback, a, void 0, void 0, r)
            }
        }
        function gi(e) {
            if (r(e.internetHeaderNameValuePairs))
                throw ee("internetHeaders");
            var t = Object.keys(e.internetHeaderNameValuePairs);
            if (0 === t.length)
                throw X("internetHeaders");
            for (var n = 0, i = t; n < i.length; n++) {
                var a = i[n]
                  , o = e.internetHeaderNameValuePairs[a];
                if (nn("internetHeaders", a),
                "string" !== typeof o)
                    throw ne("internetHeaders", typeof o, "string");
                ae(a.length + o.length, 0, 998, a)
            }
        }
        function hi(e, t) {
            var n = rt({}, {
                isLoadedItem: t,
                getAsync: mi
            });
            return e && rt(n, {
                removeAsync: ui,
                setAsync: vi
            }),
            n
        }
        function Ai(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(1, "item.getItemIdAsync");
                var r = ie(t, !0, !1);
                q(164, r.asyncContext, r.callback, void 0, void 0, void 0, e)
            }
        }
        function Ti(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(1, "item.getComposeTypeAsync");
                var r = ie(t, !0, !1);
                q(174, r.asyncContext, r.callback, void 0, void 0, void 0, e)
            }
        }
        function Si(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(1, "isClientSignatureEnabledAsync");
                var r = ie(t, !0, !1);
                q(175, r.asyncContext, r.callback, void 0, void 0, void 0, e)
            }
        }
        function bi(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(2, "disableClientSignatureAsync");
                var r = ie(t, !0, !1);
                e ? qt(r.asyncContext, r.callback) : q(176, r.asyncContext, r.callback, void 0, void 0, void 0, e)
            }
        }
        function Di(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            re(2, "sessionData.getAsync");
            var r = ie(t, !0, !1)
              , i = {
                name: e
            };
            xi(i),
            q(186, r.asyncContext, r.callback, i, void 0, void 0, void 0)
        }
        function Ci(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                var i = ie(n, !0, !1);
                if (e)
                    qt(i.asyncContext, i.callback);
                else {
                    re(2, "sessionData.getAsync");
                    var a = {
                        name: t
                    };
                    xi(a),
                    q(186, i.asyncContext, i.callback, a, void 0, void 0, e)
                }
            }
        }
        function xi(e) {
            nn("name", e.name)
        }
        function wi(e, t) {
            for (var n = [], r = 2; r < arguments.length; r++)
                n[r - 2] = arguments[r];
            re(2, "sessionData.setAsync");
            var i = ie(n, !1, !1)
              , a = {
                name: e,
                value: t
            };
            ki(a),
            q(185, i.asyncContext, i.callback, a, void 0, void 0, void 0)
        }
        function Ei(e) {
            return function(t, n) {
                for (var r = [], i = 2; i < arguments.length; i++)
                    r[i - 2] = arguments[i];
                var a = ie(r, !1, !1);
                if (e)
                    qt(a.asyncContext, a.callback);
                else {
                    re(2, "sessionData.setAsync");
                    var o = {
                        name: t,
                        value: n
                    };
                    ki(o),
                    q(185, a.asyncContext, a.callback, o, void 0, void 0, e)
                }
            }
        }
        function ki(e) {
            nn("name", e.name),
            function(e, t) {
                if (r(t))
                    throw ee(e);
                if ("string" !== typeof t)
                    throw ne(e, typeof t, "string")
            }("value", e.value)
        }
        function Oi() {
            for (var e = [], t = 0; t < arguments.length; t++)
                e[t] = arguments[t];
            re(2, "sessionData.getAllAsync");
            var n = ie(e, !0, !1);
            q(187, n.asyncContext, n.callback, void 0, void 0, void 0, void 0)
        }
        function Ii(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                var r = ie(t, !0, !1);
                e ? qt(r.asyncContext, r.callback) : (re(2, "sessionData.getAllAsync"),
                q(187, r.asyncContext, r.callback, void 0, void 0, void 0, e))
            }
        }
        function Mi() {
            for (var e = [], t = 0; t < arguments.length; t++)
                e[t] = arguments[t];
            re(2, "sessionData.clearAsync");
            var n = ie(e, !1, !1);
            q(188, n.asyncContext, n.callback, void 0, void 0, void 0, void 0)
        }
        function _i(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                var r = ie(t, !1, !1);
                e ? qt(r.asyncContext, r.callback) : (re(2, "sessionData.clearAsync"),
                q(188, r.asyncContext, r.callback, void 0, void 0, void 0, e))
            }
        }
        function Ni(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            re(2, "sessionData.removeAsync");
            var r = ie(t, !1, !1)
              , i = {
                name: e
            };
            Pi(i),
            q(189, r.asyncContext, r.callback, i, void 0, void 0, void 0)
        }
        function Fi(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                var i = ie(n, !1, !1);
                if (e)
                    qt(i.asyncContext, i.callback);
                else {
                    re(2, "sessionData.removeAsync");
                    var a = {
                        name: t
                    };
                    Pi(a),
                    q(189, i.asyncContext, i.callback, a, void 0, void 0, e)
                }
            }
        }
        function Pi(e) {
            nn("name", e.name)
        }
        function Ri(e) {
            return E("MultiSelectV2") ? rt({}, {
                getAsync: Ci(e),
                setAsync: Ei(e),
                getAllAsync: Ii(e),
                clearAsync: _i(e),
                removeAsync: Fi(e)
            }) : rt({}, {
                getAsync: Di,
                setAsync: wi,
                getAllAsync: Oi,
                clearAsync: Mi,
                removeAsync: Ni
            })
        }
        function Ui(e) {
            if (!e)
                throw X("sensitivityLabel");
            if ("string" != typeof e && !e.id)
                throw X("sensitivityLabel.id", void 0);
            if ("string" != typeof e && e.id && null != (e = e).children)
                throw X(void 0, o("l_SensitivityUnableToSetParent_Text"))
        }
        function Li(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                re(2, "sensitivityLabel.setAsync");
                var i = ie(n, !1, !1);
                if (e)
                    qt(i.asyncContext, i.callback);
                else {
                    Ui(t);
                    var a = ji(t)
                      , o = {
                        sensitivityLabelID: a
                    };
                    q(200, i.asyncContext, i.callback, o, void 0, void 0, e)
                }
            }
        }
        function ji(e) {
            return "string" != typeof e && void 0 !== e.id ? e.id : e
        }
        function Wi(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(2, "sensitivityLabel.getAsync");
                var r = ie(t, !0, !1);
                q(201, r.asyncContext, r.callback, void 0, void 0, void 0, e)
            }
        }
        function Bi(e) {
            return rt({}, {
                getAsync: Wi(e),
                setAsync: Li(e)
            })
        }
        function Hi() {
            for (var e = [], t = 0; t < arguments.length; t++)
                e[t] = arguments[t];
            re(2, "item.closeAsync");
            var n = ie(e, !1, !1)
              , r = !1;
            n.options && (r = !!n.options.discardItem);
            var i = {
                discardItem: r
            };
            q(203, n.asyncContext, n.callback, i, void 0, void 0, void 0)
        }
        function Ji(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(1, "item.getItemClassAsync");
                var r = ie(t, !0, !1);
                q(210, r.asyncContext, r.callback, void 0, void 0, void 0, e)
            }
        }
        function zi(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(1, "item.getConversationIndexAsync");
                var r = ie(t, !0, !1);
                q(213, r.asyncContext, r.callback, void 0, void 0, void 0, e)
            }
        }
        function qi(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                var r = ie(t, !1, !1);
                !e && E("UILessCommandsV2") ? (re(2, "item.sendAsync"),
                q(215, r.asyncContext, r.callback, void 0, void 0, void 0, e)) : qt(r.asyncContext, r.callback)
            }
        }
        function Vi(e, t) {
            var n = rt({}, {
                bcc: oi("bcc", e),
                body: en(It.Compose, e, t),
                categories: En(e),
                cc: oi("cc", e),
                conversationId: Xt("conversationId", t),
                from: di("from", e),
                internetHeaders: hi(!0, e),
                itemType: "message",
                notificationMessages: mn(e),
                seriesId: Xt("seriesId", t),
                subject: Fr(!1, e),
                to: oi("to", e),
                addFileAttachmentAsync: Br(e),
                addFileAttachmentFromBase64Async: Lr(e),
                addItemAttachmentAsync: zr,
                close: Vr,
                closeAsync: Hi,
                getAttachmentsAsync: Yr(e),
                getAttachmentContentAsync: kn(e),
                getInitializationContextAsync: Tt(e),
                getItemIdAsync: Ai(e),
                getSelectedDataAsync: E("MultiSelectV2") ? Zr(e) : Gr,
                loadCustomPropertiesAsync: kt(e),
                removeAttachmentAsync: Kr(e),
                saveAsync: Qr(e),
                setSelectedDataAsync: Kt(29, e),
                delayDeliveryTime: Hn(!0, e),
                getComposeTypeAsync: Ti(e),
                isClientSignatureEnabledAsync: Si(e),
                disableClientSignatureAsync: bi(e),
                sessionData: Ri(e),
                sensitivityLabel: Bi(e),
                getItemClassAsync: Ji(e),
                inReplyTo: Xt("inReplyTo", t),
                getConversationIndexAsync: zi(e),
                sendAsync: qi(e)
            });
            return E("MultiSelectV2") && e && rt(n, {
                isLoadedItem: e,
                unloadAsync: Pr
            }),
            n
        }
        function Yi(e) {
            if (r(e))
                throw ee("locationIdentifier");
            if (!Array.isArray(e))
                throw ne("locationIdentifier", typeof e, "Array");
            if (0 === e.length)
                throw X("locationIdentifier");
            for (var t = 0, n = e; t < n.length; t++) {
                Gi(n[t])
            }
        }
        function Gi(e) {
            if (r(e) || r(e.id) || r(e.type))
                throw ee("locationIdentifier");
            if (e.type !== ce.LocationType.Room && e.type !== ce.LocationType.Custom)
                throw X("type");
            !function(e, t) {
                if ("" === e)
                    throw X("id");
                if (t === ce.LocationType.Room && e.length > 571)
                    throw X("id")
            }(e.id, e.type)
        }
        function Zi(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            re(2, "enhancedLocations.addAsync");
            var r = ie(t, !1, !1)
              , i = {
                enhancedLocations: e
            };
            Ki(i),
            q(155, r.asyncContext, r.callback, i, void 0, void 0, void 0)
        }
        function Ki(e) {
            Yi(e.enhancedLocations)
        }
        function $i() {
            for (var e = [], t = 0; t < arguments.length; t++)
                e[t] = arguments[t];
            re(1, "enhancedLocations.getAsync");
            var n = ie(e, !0, !1);
            q(154, n.asyncContext, n.callback, void 0, void 0, void 0, void 0)
        }
        function Qi(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            re(2, "enhancedLocations.removeAsync");
            var r = ie(t, !1, !1)
              , i = {
                enhancedLocations: e
            };
            Xi(i),
            q(156, r.asyncContext, r.callback, i, void 0, void 0, void 0)
        }
        function Xi(e) {
            Yi(e.enhancedLocations)
        }
        function ea(e) {
            var t = rt({}, {
                getAsync: $i
            });
            return e && rt(t, {
                addAsync: Zi,
                removeAsync: Qi
            }),
            t
        }
        function ta(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(1, e + ".getAsync");
                var r = ie(t, !0, !1);
                q(24, r.asyncContext, r.callback, {
                    TimeProperty: yi[e]
                }, na, void 0, void 0)
            }
        }
        function na(e) {
            return new Date(e)
        }
        !function(e) {
            e[e.start = 1] = "start",
            e[e.end = 2] = "end"
        }(yi || (yi = {}));
        function ra(e) {
            return function(t) {
                for (var n = [], r = 1; r < arguments.length; r++)
                    n[r - 1] = arguments[r];
                re(2, e + ".setAsync");
                var i = ie(n, !1, !1)
                  , a = {
                    date: t
                };
                ia(a),
                q(25, i.asyncContext, i.callback, {
                    TimeProperty: yi[e],
                    time: a.date.getTime()
                }, void 0, void 0, void 0)
            }
        }
        function ia(e) {
            if (!Oe(e.date))
                throw ne("dateTime", typeof e.date, typeof Date);
            if (isNaN(e.date.getTime()))
                throw X("dateTime");
            if (e.date.getTime() < -864e13 || e.date.getTime() > 864e13)
                throw te("dateTime")
        }
        function aa(e) {
            return rt({}, {
                getAsync: ta(e),
                setAsync: ra(e)
            })
        }
        function oa() {
            for (var e = [], t = 0; t < arguments.length; t++)
                e[t] = arguments[t];
            re(1, "location.getAsync");
            var n = ie(e, !0, !1);
            q(26, n.asyncContext, n.callback, void 0, void 0, void 0, void 0)
        }
        function sa(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            re(2, "location.setAsync");
            var r = ie(t, !1, !1)
              , i = {
                location: e
            };
            ca(i),
            q(27, r.asyncContext, r.callback, i, void 0, void 0, void 0)
        }
        function ca(e) {
            if (r(e.location))
                throw ee("location");
            if ("string" !== typeof e.location)
                throw ne("location", typeof e.location, "string");
            ae(e.location.length, 0, 255, "location")
        }
        function da() {
            for (var e = [], t = 0; t < arguments.length; t++)
                e[t] = arguments[t];
            re(1, "recurrenceProperties.getAsync");
            var n = ie(e, !0, !1);
            q(103, n.asyncContext, n.callback, void 0, la, void 0, void 0)
        }
        function la(e) {
            if (null !== e && null !== e.seriesTimeJson) {
                var t = new kr;
                t.importFromSeriesTimeJsonObject(e.seriesTimeJson),
                delete e.seriesTimeJson,
                e.seriesTime = t
            }
            return e
        }
        function ua(e) {
            if (!r(e)) {
                if (r((e = e).recurrenceType))
                    throw ee("recurrenceType");
                if (r(e.seriesTime))
                    throw ee("seriesTime");
                if (!(e.seriesTime instanceof kr) || !e.seriesTime.isValid())
                    throw X("seriesTime");
                if (!e.seriesTime.isEndAfterStart())
                    throw X("seriesTime", o("l_InvalidEventDates_Text"));
                if (function(e) {
                    if (e !== ce.RecurrenceType.Daily && e !== ce.RecurrenceType.Weekly && e !== ce.RecurrenceType.Weekday && e !== ce.RecurrenceType.Yearly && e !== ce.RecurrenceType.Monthly)
                        throw X("recurrenceType")
                }(e.recurrenceType),
                e.recurrenceType !== ce.RecurrenceType.Weekday && r(e.recurrenceProperties))
                    throw ee("recurrenceType");
                if (!r(e.recurrenceTimeZone)) {
                    if (r(e.recurrenceTimeZone.name))
                        throw ee("name");
                    if ("string" !== typeof e.recurrenceTimeZone.name)
                        throw ne("name", typeof e.recurrenceTimeZone.name, "string")
                }
                e.recurrenceType === ce.RecurrenceType.Daily ? fa(e.recurrenceProperties) : e.recurrenceType === ce.RecurrenceType.Weekly ? function(e) {
                    if (fa(e),
                    r(e.days))
                        throw ne("days");
                    if (!Array.isArray(e.days))
                        throw ne("days");
                    if (function(e) {
                        for (var t = 0; t < e.length; t++)
                            if (!ma(e[t], !1))
                                throw X("days")
                    }(e.days),
                    !r(e.firstDayOfWeek)) {
                        if ("string" !== typeof e.firstDayOfWeek)
                            throw ne("firstDayOfWeek");
                        if (!ma(e.firstDayOfWeek, !1))
                            throw X("firstDayOfWeek")
                    }
                }(e.recurrenceProperties) : e.recurrenceType === ce.RecurrenceType.Monthly ? function(e) {
                    if (r(e.interval))
                        throw ee("interval");
                    if ("number" !== typeof e.interval)
                        throw ne("interval", typeof e.interval, "number");
                    if (r(e.dayOfMonth)) {
                        if (r(e.dayOfWeek) || r(e.weekNumber))
                            throw X(void 0, o("l_Recurrence_Error_Properties_Invalid_Text"));
                        if ("string" !== typeof e.dayOfWeek)
                            throw ne("dayOfWeek", typeof e.dayOfWeek, "string");
                        if (!ma(e.dayOfWeek, !0))
                            throw X("dayOfWeek");
                        if ("string" !== typeof e.weekNumber)
                            throw ne("weekNumber", typeof e.weekNumber, "string");
                        pa(e.weekNumber)
                    } else {
                        if ("number" !== typeof e.dayOfMonth)
                            throw ne("dayOfMonth", typeof e.dayOfMonth, "number");
                        ya(e.dayOfMonth)
                    }
                }(e.recurrenceProperties) : e.recurrenceType === ce.RecurrenceType.Yearly && function(e) {
                    if (r(e.interval))
                        throw ee("interval");
                    if ("number" !== typeof e.interval)
                        throw ne("interval", typeof e.interval, "number");
                    if (r(e.month))
                        throw ee("month");
                    if ("string" !== typeof e.month)
                        throw ne("month", typeof e.month, "string");
                    if (function(e) {
                        if (e !== ce.Month.Jan && e !== ce.Month.Feb && e !== ce.Month.Mar && e !== ce.Month.Apr && e !== ce.Month.May && e !== ce.Month.Jun && e !== ce.Month.Jul && e !== ce.Month.Aug && e !== ce.Month.Sep && e !== ce.Month.Oct && e !== ce.Month.Nov && e !== ce.Month.Dec)
                            throw X("month")
                    }(e.month),
                    r(e.dayOfMonth)) {
                        if (r(e.weekNumber) || r(e.dayOfWeek))
                            throw X(void 0, o("l_Recurrence_Error_Properties_Invalid_Text"));
                        if ("string" !== typeof e.dayOfWeek)
                            throw ne("dayOfWeek", typeof e.dayOfWeek, "string");
                        if (!ma(e.dayOfWeek, !0))
                            throw X("dayOfWeek");
                        if ("string" !== typeof e.weekNumber)
                            throw ne("weekNumber", typeof e.weekNumber, "string");
                        pa(e.weekNumber)
                    } else {
                        if ("number" !== typeof e.dayOfMonth)
                            throw ne("dayOfMonth", typeof e.dayOfMonth, "number");
                        ya(e.dayOfMonth)
                    }
                }(e.recurrenceProperties)
            }
        }
        function fa(e) {
            if (r(e.interval))
                throw ee("interval");
            if ("number" !== typeof e.interval)
                throw ne("interval", typeof e.interval, "number");
            if (e.interval <= 0)
                throw X("interval")
        }
        function ma(e, t) {
            var n = e === ce.Days.Mon || e === ce.Days.Tue || e === ce.Days.Wed || e === ce.Days.Thu || e === ce.Days.Fri || e === ce.Days.Sat || e === ce.Days.Sun;
            return t && (e === ce.Days.WeekendDay || e === ce.Days.Weekday || e === ce.Days.Day) || n
        }
        function pa(e) {
            if (e !== ce.WeekNumber.First && e !== ce.WeekNumber.Second && e !== ce.WeekNumber.Third && e !== ce.WeekNumber.Fourth && e !== ce.WeekNumber.Last)
                throw X("weekNumber")
        }
        function ya(e) {
            if (e < 1 || e > 31)
                throw X("dayOfMonth")
        }
        function va(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            re(2, "recurrenceProperties.setAsync");
            var i = xa().seriesId;
            if (!r(i) && i.length > 0)
                throw X(void 0, o("l_Recurrence_Error_Instance_SetAsync_Text"));
            ua(e);
            var a = ie(t, !1, !1)
              , s = ga(e)
              , c = {
                recurrenceData: s
            };
            q(104, a.asyncContext, a.callback, c, void 0, void 0, void 0)
        }
        function ga(e) {
            if (null !== e && null !== e.seriesTime && e.seriesTime instanceof kr)
                return {
                    recurrenceProperties: e.recurrenceProperties,
                    recurrenceTimeZone: e.recurrenceTimeZone,
                    recurrenceType: e.recurrenceType,
                    seriesTimeJson: e.seriesTime.exportToSeriesTimeJson()
                };
            return e
        }
        function ha(e) {
            var t = rt({}, {
                getAsync: da
            });
            return e && rt(t, {
                setAsync: va
            }),
            t
        }
        function Aa() {
            for (var e = [], t = 0; t < arguments.length; t++)
                e[t] = arguments[t];
            re(1, "isAllDayEvent.getAsync");
            var n = ie(e, !0, !1);
            q(169, n.asyncContext, n.callback, void 0, void 0, void 0, void 0)
        }
        function Ta(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            re(2, "isAllDayEvent.setAsync");
            var r = ie(t, !0, !1)
              , i = {
                isAllDayEvent: e
            };
            Sa(i),
            q(170, r.asyncContext, r.callback, i, void 0, void 0, void 0)
        }
        function Sa(e) {
            if (r(e.isAllDayEvent))
                throw ee("isAllDayEvent");
            if ("boolean" !== typeof e.isAllDayEvent)
                throw ne("isAllDayEvent", typeof e.isAllDayEvent, "boolean")
        }
        function ba(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            re(2, "sensitivity.setAsync");
            var r = ie(t, !0, !1)
              , i = {
                sensitivity: e
            };
            Da(i),
            q(172, r.asyncContext, r.callback, i, void 0, void 0, void 0)
        }
        function Da(e) {
            nn("sensitivity", e.sensitivity),
            function(e) {
                if (e !== ce.AppointmentSensitivityType.Normal && e !== ce.AppointmentSensitivityType.Personal && e !== ce.AppointmentSensitivityType.Private && e !== ce.AppointmentSensitivityType.Confidential)
                    throw X("sensitivity")
            }(e.sensitivity)
        }
        function Ca() {
            for (var e = [], t = 0; t < arguments.length; t++)
                e[t] = arguments[t];
            re(1, "sensitivity.getAsync");
            var n = ie(e, !0, !1);
            q(171, n.asyncContext, n.callback, void 0, void 0, void 0, void 0)
        }
        function xa() {
            return rt({}, {
                body: en(It.Compose),
                categories: En(),
                end: aa("end"),
                enhancedLocation: ea(!0),
                itemType: "appointment",
                location: rt({}, {
                    getAsync: oa,
                    setAsync: sa
                }),
                notificationMessages: mn(),
                optionalAttendees: oi("optionalAttendees"),
                organizer: di("organizer"),
                recurrence: ha(!0),
                requiredAttendees: oi("requiredAttendees"),
                seriesId: ao("seriesId"),
                start: aa("start"),
                subject: Fr(!1),
                addFileAttachmentAsync: Br(),
                addFileAttachmentFromBase64Async: Lr(),
                addItemAttachmentAsync: zr,
                close: Vr,
                getAttachmentsAsync: Yr(),
                getAttachmentContentAsync: kn(),
                getInitializationContextAsync: Tt(),
                getItemIdAsync: Ai(),
                getSelectedDataAsync: E("MultiSelectV2") ? Zr() : Gr,
                loadCustomPropertiesAsync: kt(),
                removeAttachmentAsync: Kr(),
                saveAsync: Qr(),
                setSelectedDataAsync: Kt(29),
                isAllDayEvent: rt({}, {
                    getAsync: Aa,
                    setAsync: Ta
                }),
                sensitivity: rt({}, {
                    getAsync: Ca,
                    setAsync: ba
                }),
                isClientSignatureEnabledAsync: Si(),
                disableClientSignatureAsync: bi(),
                sensitivityLabel: Bi(),
                sessionData: Ri(),
                sendAsync: qi()
            })
        }
        var wa, Ea = n(0), ka = n(1);
        function Oa(e) {
            var t = {
                consentState: e,
                extensionId: ao("extensionId")
            };
            !function(e) {
                if (e !== wa.Consented && e !== wa.NotConsented && e !== wa.NotResponded)
                    throw te("consentState")
            }(e),
            q(40, void 0, void 0, t, void 0, void 0, void 0)
        }
        function Ia(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            var r = ie(t, !1, !1)
              , i = {
                module: e
            };
            Ma(e),
            e === ce.ModuleType.Addins && (r.options && r.options.queryString ? i.queryString = r.options.queryString : i.queryString = ""),
            q(45, r.asyncContext, r.callback, i, void 0, void 0, void 0)
        }
        function Ma(e) {
            if (r(e))
                throw ee("module");
            if ("" === e)
                throw X("module");
            if (e !== ce.ModuleType.Addins)
                throw X("module")
        }
        function _a(e) {
            if (r(e))
                throw ee("data");
            q(402, void 0, void 0, e, void 0, void 0, void 0)
        }
        function Na(e) {
            if (r(e))
                throw ee("data");
            q(401, void 0, void 0, e, void 0, void 0, void 0)
        }
        function Fa(e) {
            if (r(e))
                throw ee("data");
            q(400, void 0, void 0, e, void 0, void 0, void 0)
        }
        function Pa(e, t, n, r) {
            return q(403, void 0, void 0, {
                launchUrl: e
            }, void 0, void 0, void 0),
            window
        }
        function Ra(e) {
            if (r(e))
                throw ee("data");
            q(163, void 0, void 0, {
                telemetryData: e
            }, void 0, void 0, void 0)
        }
        function Ua(e) {
            if (r(e))
                throw ee("data");
            q(193, void 0, void 0, {
                telemetryData: e
            }, void 0, void 0, void 0)
        }
        !function(e) {
            e[e.NotResponded = 0] = "NotResponded",
            e[e.NotConsented = 1] = "NotConsented",
            e[e.Consented = 2] = "Consented"
        }(wa || (wa = {}));
        var La = function() {
            return (La = Object.assign || function(e) {
                for (var t, n = 1, r = arguments.length; n < r; n++)
                    for (var i in t = arguments[n])
                        Object.prototype.hasOwnProperty.call(t, i) && (e[i] = t[i]);
                return e
            }
            ).apply(this, arguments)
        };
        function ja(e) {
            if (!Oe(e))
                throw X("timeValue");
            var t = new Date(e.getTime())
              , n = -1 * t.getTimezoneOffset();
            return r(ao("timeZoneOffsets")) || (t.setUTCMinutes(t.getUTCMinutes() - n),
            n = er(t),
            t.setUTCMinutes(t.getUTCMinutes() + n)),
            La({
                timezoneOffset: n
            }, nr(t))
        }
        function Wa(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            var r = ie(t, !1, !1)
              , i = {
                ewsIdOrEmail: e
            };
            Ba(i),
            q(43, r.asyncContext, r.callback, {
                ewsIdOrEmail: e.trim()
            }, void 0, void 0, void 0)
        }
        function Ba(e) {
            if (r(e.ewsIdOrEmail))
                throw ee("ewsIdOrEmail");
            if (function(e) {
                if ("string" !== typeof e)
                    throw X("ewsIdOrEmail")
            }(e.ewsIdOrEmail),
            "" === e.ewsIdOrEmail)
                throw X("ewsIdOrEmail", "ewsIdOrEmail cannot be empty.")
        }
        function Ha(e) {
            return function() {
                for (var t = [], n = 0; n < arguments.length; n++)
                    t[n] = arguments[n];
                re(1, "item.getSharedPropertiesAsync");
                var r = ie(t, !0, !1);
                q(108, r.asyncContext, r.callback, void 0, void 0, void 0, e)
            }
        }
        var Ja = function(e, t) {
            e && ao("isFromSharedFolder") && ht() !== gt.ItemLess && rt(e, {
                getSharedPropertiesAsync: Ha(t)
            })
        };
        function za() {
            for (var e = [], t = 0; t < arguments.length; t++)
                e[t] = arguments[t];
            re(3, "getSelectedItemsAsync");
            var n = ie(e, !0, !1);
            q(196, n.asyncContext, n.callback, void 0, void 0, void 0, void 0)
        }
        var qa = n(0);
        function Va(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            re(2, "mailbox.loadItemByIdAsync");
            var r = ie(t, !1, !1)
              , i = {
                itemId: e
            };
            Ya(i),
            q(211, r.asyncContext, r.callback, {
                itemId: ge(i.itemId)
            }, void 0, Ga, void 0)
        }
        function Ya(e) {
            ve(e.itemId)
        }
        function Ga(e, t, n) {
            if (e.error)
                return V(void 0, qa.DDA.AsyncResultEnum.ErrorCode.Failed, e.errorCode, t, e.errorMessage, e.errorName);
            if (n && n !== v.noError)
                return V(void 0, qa.DDA.AsyncResultEnum.ErrorCode.Failed, 9002, t);
            var r = !0;
            if (void 0 != e.wasSuccessful && (r = e.wasSuccessful),
            r) {
                var i = void 0
                  , a = void 0
                  , s = qa.DDA.AsyncResultEnum.ErrorCode.Success
                  , c = 0
                  , d = void 0;
                return void 0 != e.data && ((i = JSON.parse(e.data)).itemType == gt.Message || i.itemType == gt.MeetingRequest ? a = Rr(!0, i) : i.itemType == gt.MessageCompose ? a = Vi(!0, i) : (s = qa.DDA.AsyncResultEnum.ErrorCode.Failed,
                c = 9016,
                d = o("l_InvalidSelection_Text")),
                Ja(a, !0)),
                V(a, s, c, t, d)
            }
            return V(void 0, qa.DDA.AsyncResultEnum.ErrorCode.Failed, 5e3, t)
        }
        var Za = n(0)
          , Ka = function() {
            var e, t = void 0;
            switch (ht()) {
            case gt.Message:
                t = Rr();
                break;
            case gt.MessageCompose:
                t = Vi();
                break;
            case gt.Appointment:
                t = function() {
                    var e = ao("organizer")
                      , t = ao("dateTimeCreated")
                      , n = ao("dateTimeModified")
                      , r = ao("end")
                      , i = ao("start");
                    return rt({}, {
                        attachments: br(ao("attachments")),
                        body: en(It.Read),
                        categories: En(),
                        dateTimeCreated: t ? new Date(t) : void 0,
                        dateTimeModified: n ? new Date(n) : void 0,
                        end: r ? new Date(r) : void 0,
                        enhancedLocation: ea(!1),
                        itemClass: ao("itemClass"),
                        itemId: ao("id"),
                        itemType: "appointment",
                        location: ao("location"),
                        normalizedSubject: ao("normalizedSubject"),
                        notificationMessages: mn(),
                        optionalAttendees: (ao("cc") || []).map(Un),
                        organizer: e ? Un(e) : void 0,
                        recurrence: Or(ao("recurrence")),
                        requiredAttendees: (ao("to") || []).map(Un),
                        start: i ? new Date(i) : void 0,
                        seriesId: ao("seriesId"),
                        subject: ao("subject"),
                        displayReplyForm: An(void 0),
                        displayReplyFormAsync: Sn(void 0),
                        displayReplyAllForm: Tn(void 0),
                        displayReplyAllFormAsync: bn(void 0),
                        getAttachmentContentAsync: kn(),
                        getEntities: yr(),
                        getEntitiesByType: vr(),
                        getFilteredEntitiesByName: gr(),
                        getInitializationContextAsync: Tt(),
                        getRegExMatches: hr(),
                        getRegExMatchesByName: Ar(),
                        getSelectedEntities: Tr(),
                        getSelectedRegExMatches: Sr(),
                        loadCustomPropertiesAsync: kt(),
                        isAllDayEvent: ao("isAllDayEvent"),
                        sensitivity: ao("sensitivity")
                    })
                }();
                break;
            case gt.AppointmentCompose:
                t = xa();
                break;
            case gt.MeetingRequest:
                t = Rr();
                break;
            default:
                return
            }
            return e = t,
            Ea.DDA.DispIdHost.addEventSupport(e, new Ea.EventDispatch([ka.Office.WebExtension.EventType.RecipientsChanged, ka.Office.WebExtension.EventType.AppointmentTimeChanged, ka.Office.WebExtension.EventType.AttachmentsChanged, ka.Office.WebExtension.EventType.EnhancedLocationsChanged, ka.Office.WebExtension.EventType.InfobarClicked, ka.Office.WebExtension.EventType.RecurrenceChanged, ka.Office.WebExtension.EventType.SensitivityLabelChanged, ka.Office.WebExtension.EventType.InitializationContextChanged])),
            Ja(t, !1),
            t
        };
        function $a() {
            for (var e = [], t = 0; t < arguments.length; t++)
                e[t] = arguments[t];
            re(2, "sensitivityLabelsCatalog.getAsync");
            var n = ie(e, !0, !1);
            q(199, n.asyncContext, n.callback, void 0, void 0, void 0, void 0)
        }
        function Qa() {
            for (var e = [], t = 0; t < arguments.length; t++)
                e[t] = arguments[t];
            re(2, "sensitivityLabelsCatalog.getIsEnabledAsync");
            var n = ie(e, !0, !1);
            q(202, n.asyncContext, n.callback, void 0, void 0, void 0, void 0)
        }
        function Xa() {
            return rt({}, {
                getAsync: $a,
                getIsEnabledAsync: Qa
            })
        }
        function eo(e) {
            if (!e)
                throw ee("permissions");
            if (!Array.isArray(e))
                throw ne("permissions", typeof e, typeof Array);
            if (0 === e.length)
                throw ee("permissions");
            e.forEach((function(e) {
                if (!(e in oe))
                    throw te("permissions", e)
            }
            ))
        }
        function to(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            re(2, "devicePermission.requestPermissionsAsync");
            var r = ie(t, !0, !1);
            eo(e);
            var i = {
                permissions: e
            };
            q(214, r.asyncContext, r.callback, i, void 0, void 0, void 0)
        }
        function no() {
            return rt({}, {
                requestPermissionsAsync: to
            })
        }
        var ro, io = n(0), ao = function(e) {
            return ro && ro.getInitialDataProp(e)
        }, oo = function() {
            return !ro || !ro.item
        }, so = function() {
            return ro && ro.getAppName()
        }, co = function() {
            function e(e, t, n) {
                var i = this;
                this.displayName = "mailbox",
                this.stringLoadedCallback = function() {
                    i.appReadyCallback && (i.officeAppContext.get_isDialog() ? setTimeout((function() {
                        return i.appReadyCallback()
                    }
                    ), 0) : G(1, void 0, i.onInitialDataResponse))
                }
                ,
                this.initialize = function(e) {
                    if (null === e || void 0 === e)
                        I(),
                        i.initialData = null,
                        i.item = null;
                    else {
                        i.initialData = e,
                        i.initialData.permissionLevel = lo(),
                        i.item = Ka();
                        so() !== io.AppName.Outlook || function(e) {
                            var t = 0
                              , n = 0;
                            return r(e) || (t = e.indexOf("."),
                            n = parseInt(e.substring(0, t))),
                            n >= 16
                        }(ao("hostVersion")) || function(e, t) {
                            var n = !1;
                            try {
                                var r = JSON.parse(t.get_requirementMatrix()).Mailbox.split(".")
                                  , i = e.split(".");
                                (parseInt(r[0]) > parseInt(i[0]) || parseInt(r[0]) === parseInt(i[0]) && parseInt(r[1]) >= parseInt(i[1])) && (n = !0)
                            } catch (e) {}
                            return n
                        }("1.5", i.officeAppContext),
                        I(),
                        "undefined" !== typeof e.itemNumber && O().setCurrentItemNumber(e.itemNumber)
                    }
                }
                ,
                this.exposeDevicePermissionApi = function() {
                    if (so() == io.AppName.OutlookWebApp) {
                        window.OfficeCore || (window.OfficeCore = {}),
                        window.OfficeCore.DevicePermissionType = {
                            camera: 0,
                            microphone: 1,
                            geolocation: 2
                        },
                        i.officeAppContext.devicePermission = no
                    }
                }
                ,
                this.onInitialDataResponse = function(e, t) {
                    if (!e || e === v.noError) {
                        var n;
                        i.initialize(t),
                        rt(n = i, {
                            ewsUrl: ao("ewsUrl"),
                            restUrl: ao("restUrl"),
                            displayAppointmentForm: Ae,
                            displayAppointmentFormAsync: Te,
                            displayMessageForm: Ce,
                            displayMessageFormAsync: xe,
                            displayPersonaCardAsync: Wa,
                            getCallbackTokenAsync: $e,
                            getUserIdentityTokenAsync: Qe,
                            logTelemetry: Ra,
                            logCustomerContentTelemetry: Ua,
                            makeEwsRequestAsync: tt,
                            masterCategories: rt({}, {
                                addAsync: ut,
                                getAsync: ft,
                                removeAsync: yt
                            }),
                            navigateToModuleAsync: Ia,
                            diagnostics: rt({}, {
                                hostName: at(),
                                hostVersion: ao("hostVersion"),
                                OWAView: ao("owaView")
                            }),
                            userProfile: rt({}, {
                                accountType: ao("userProfileType"),
                                displayName: ao("userDisplayName"),
                                emailAddress: ao("userEmailAddress"),
                                timeZone: ao("userTimeZone")
                            }),
                            convertToEwsId: fe,
                            convertToLocalClientTime: ja,
                            convertToRestId: ue,
                            convertToUtcClientTime: tr,
                            getSelectedItemsAsync: za,
                            RegisterConsentAsync: Oa,
                            GetIsRead: function() {
                                return ao("isRead")
                            },
                            GetEndPointUrl: function() {
                                return ao("endNodeUrl")
                            },
                            GetConsentMetaData: function() {
                                return ao("consentMetadata")
                            },
                            GetMarketplaceContentMarket: function() {
                                return ao("marketplaceContentMarket")
                            },
                            GetMarketplaceAssetId: function() {
                                return ao("marketplaceAssetId")
                            },
                            GetExtensionId: function() {
                                return ao("extensionId")
                            },
                            CloseApp: vt,
                            recordDataPoint: _a,
                            recordTrace: Na,
                            trackCtq: Fa
                        }),
                        ht() !== gt.MessageCompose && ht() !== gt.AppointmentCompose && rt(n, {
                            displayNewAppointmentForm: Me,
                            displayNewMessageForm: je,
                            displayNewAppointmentFormAsync: _e,
                            displayNewMessageFormAsync: We
                        }),
                        E("MultiSelectV2") && rt(n, {
                            loadItemByIdAsync: Va
                        }),
                        so() === Za.AppName.OutlookWebApp && ao("openWindowOpen") && (window.open = Pa);
                        var r = i.officeAppContext;
                        r.sensitivityLabelsCatalog = Xa,
                        i.exposeDevicePermissionApi(),
                        r.urls = ao("urls"),
                        setTimeout((function() {
                            return i.appReadyCallback()
                        }
                        ), 0)
                    }
                }
                ,
                this.officeAppContext = e,
                this.targetWindow = window,
                this.appReadyCallback = n,
                ro = this,
                function(e) {
                    var t;
                    s = e;
                    for (var n = document.getElementsByTagName("script"), r = 0; r < n.length; r++) {
                        var i = n.item(r);
                        if (i && i.src) {
                            var a = i.src || "";
                            if (t = (a = a.toLowerCase()).indexOf("office_strings.js"),
                            a && t > 0) {
                                c = a.replace("office_strings.js", "outlook_strings.js"),
                                d = y(d, t, a);
                                break
                            }
                            if (t = a.indexOf("office_strings.debug.js"),
                            a && t > 0) {
                                c = a.replace("office_strings.debug.js", "outlook_strings.js"),
                                d = y(d, t, a);
                                break
                            }
                        }
                    }
                    if (c) {
                        var o = document.getElementsByTagName("head")[0];
                        (l = f(c)).onload = m,
                        l.onreadystatechange = m,
                        window.setTimeout(p, 2e3),
                        o.appendChild(l)
                    }
                }(this.stringLoadedCallback)
            }
            return e.prototype.getAppName = function() {
                return this.officeAppContext.get_appName()
            }
            ,
            e.prototype.getInitialDataProp = function(e) {
                return this.initialData && this.initialData[e]
            }
            ,
            e.prototype.setCurrentItemNumber = function(e) {
                O().setCurrentItemNumber(e)
            }
            ,
            e.addAdditionalArgs = function(e, t) {
                return t
            }
            ,
            e.shouldRunInitialDataResponse = function() {
                return !0
            }
            ,
            e
        }(), lo = function() {
            var e = ao("permissionLevel");
            if (void 0 === e)
                return 0;
            switch (e) {
            case 1:
                return 1;
            case 3:
                return 2;
            case 2:
                return 3;
            default:
                return 0
            }
        }, uo = n(0);
        function fo(e) {
            for (var t = [], n = 1; n < arguments.length; n++)
                t[n - 1] = arguments[n];
            var r = ie(t, !1, !1)
              , i = uo.DDA.SettingsManager.serializeSettings(e);
            if (JSON.stringify(i).length > 32768) {
                var a = V(void 0, uo.DDA.AsyncResultEnum.ErrorCode.Failed, 9057, r.asyncContext, "");
                r.callback && setTimeout((function() {
                    r.callback && r.callback(a)
                }
                ), 0)
            } else
                uo.AppName.OutlookWebApp === so() ? mo(r, i) : po(r, i)
        }
        function mo(e, t) {
            q(404, e.asyncContext, e.callback, [t], void 0, void 0, void 0)
        }
        function po(e, t) {
            var n = -1
              , r = null;
            try {
                var i = JSON.stringify(t)
                  , a = {};
                a.SettingsKey = i,
                uo.DDA.ClientSettingsManager.write(a)
            } catch (e) {
                r = e
            }
            var o = void 0;
            null != r ? (n = 9019,
            o = V(void 0, uo.DDA.AsyncResultEnum.ErrorCode.Failed, n, e.asyncContext, r)) : (n = 0,
            o = V(void 0, uo.DDA.AsyncResultEnum.ErrorCode.Success, n, e.asyncContext)),
            e.callback && e.callback(o)
        }
        var yo = function() {
            for (var e = 0, t = 0, n = arguments.length; t < n; t++)
                e += arguments[t].length;
            var r = Array(e)
              , i = 0;
            for (t = 0; t < n; t++)
                for (var a = arguments[t], o = 0, s = a.length; o < s; o++,
                i++)
                    r[i] = a[o];
            return r
        }
          , vo = n(0)
          , go = function() {
            function e(e) {
                this.rawData = e,
                this.settingsData = null
            }
            return e.prototype.getSettingsData = function() {
                return null == this.settingsData && (this.settingsData = this.convertFromRawSettings(this.rawData),
                this.rawData = null),
                this.settingsData
            }
            ,
            e.prototype.get = function(e) {
                return this.getSettingsData()[e]
            }
            ,
            e.prototype.set = function(e, t) {
                this.getSettingsData()[e] = t
            }
            ,
            e.prototype.remove = function(e) {
                delete this.getSettingsData()[e]
            }
            ,
            e.prototype.saveAsync = function() {
                for (var e = [], t = 0; t < arguments.length; t++)
                    e[t] = arguments[t];
                fo.apply(void 0, yo([this.getSettingsData()], e))
            }
            ,
            e.prototype.convertFromRawSettings = function(e) {
                if (null == e)
                    return {};
                if (so() !== vo.AppName.OutlookWebApp) {
                    var t = e.SettingsKey;
                    if (t)
                        return vo.DDA.SettingsManager.deserializeSettings(t)
                }
                return e
            }
            ,
            e
        }()
          , ho = {
            toItemRead: function(e) {
                var t = ht();
                if (t === gt.Message || t === gt.Appointment || t === gt.MeetingRequest)
                    return e;
                throw ne()
            },
            toItemCompose: function(e) {
                var t = ht();
                if (t === gt.MessageCompose || t === gt.AppointmentCompose)
                    return e;
                throw ne()
            },
            toMessage: function(e) {
                return ho.toMessageRead(e)
            },
            toMessageRead: function(e) {
                if (ht() === gt.Message || ht() === gt.MeetingRequest)
                    return e;
                throw ne()
            },
            toMessageCompose: function(e) {
                if (ht() === gt.MessageCompose)
                    return e;
                throw ne()
            },
            toMeetingRequest: function(e) {
                if (ht() === gt.MeetingRequest)
                    return e;
                throw ne()
            },
            toAppointment: function(e) {
                if (ht() === gt.Appointment)
                    return e;
                throw ne()
            },
            toAppointmentRead: function(e) {
                if (ht() === gt.Appointment)
                    return e;
                throw ne()
            },
            toAppointmentCompose: function(e) {
                if (ht() === gt.AppointmentCompose)
                    return e;
                throw ne()
            }
        }
          , Ao = {
            SeriesTimeJsonConverter: function(e) {
                if (null !== e && "object" === typeof e && null !== e.seriesTimeJson) {
                    var t = new kr;
                    t.importFromSeriesTimeJsonObject(e.seriesTimeJson),
                    delete e.seriesTimeJson,
                    e.seriesTime = t
                }
                return e
            },
            CreateAttachmentDetails: function(e) {
                return Dr(e),
                e
            }
        };
        OSF = "object" === typeof OSF ? OSF : {},
        OSF.DDA = OSF.DDA || {},
        OSF.DDA.Settings = go,
        OSF = "object" === typeof OSF ? OSF : {},
        OSF.DDA = OSF.DDA || {},
        OSF.DDA.OutlookAppOm = co,
        Office = "object" === typeof Office ? Office : {},
        Office.cast = Office.cast || {},
        Office.cast.item = ho,
        Microsoft.Office.WebExtension.MailboxEnums = ce,
        Microsoft.Office.WebExtension.CoercionType = de,
        Microsoft.Office.WebExtension.SeriesTime = kr,
        Microsoft.Office.WebExtension.OutlookBase = Ao,
        Microsoft.Office.WebExtension.DevicePermissionType = oe;
        t.default = co;
        var To = window;
        To.$h = "object" === typeof $h ? $h : {},
        To.$h.Message = $h.Message || {},
        To.$h.Appointment = $h.Appointment || {},
        To.$h.Message.isInstanceOfType = function(e) {
            return e && "message" === e.itemType
        }
        ,
        To.$h.Appointment.isInstanceOfType = function(e) {
            return e && "appointment" === e.itemType
        }
    }
    ]).default,
    OSF.DDA.ErrorCodeManager.initializeErrorMessages(Strings.OfficeOM),
    e.get_appName() == OSF.AppName.Outlook && OSF.DDA.RichApi && OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync && (OSF.DDA.DispIdHost.addAsyncMethods(OSF.DDA.RichApi, [OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync]),
    OSF.DDA.RichApi.richApiMessageManager = new OfficeExt.RichApiMessageManager),
    e.get_appName() == OSF.AppName.OutlookWebApp || e.get_appName() == OSF.AppName.OutlookIOS || e.get_appName() == OSF.AppName.OutlookAndroid ? this._settings = this._initializeSettings(e, !1) : "mac" == OSF._OfficeAppFactory.getHostInfo().hostPlatform && e.get_appName() == OSF.AppName.Outlook ? this._settings = this.initializeMacSettings(e, !1) : this._settings = this._initializeSettings(!1),
    e.appOM = new OSF.DDA.OutlookAppOm(e,this._webAppState.wnd,t),
    e.get_appName() != OSF.AppName.Outlook && e.get_appName() != OSF.AppName.OutlookWebApp && e.get_appName() != OSF.AppName.OutlookIOS && e.get_appName() != OSF.AppName.OutlookAndroid || OSF.DDA.DispIdHost.addEventSupport(e.appOM, new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.ItemChanged, Microsoft.Office.WebExtension.EventType.OfficeThemeChanged, Microsoft.Office.WebExtension.EventType.SelectedItemsChanged, Microsoft.Office.WebExtension.EventType.DragAndDropEvent]))
}
;
var OfficeFirstPartyAuth, __assign = this && this.__assign || function() {
    return (__assign = Object.assign || function(e) {
        for (var t, n = 1, r = arguments.length; n < r; n++)
            for (var i in t = arguments[n])
                Object.prototype.hasOwnProperty.call(t, i) && (e[i] = t[i]);
        return e
    }
    ).apply(this, arguments)
}
;
function exposeOfficeRuntimeThroughOfficeNamespace(e, t) {
    var n, r;
    "undefined" === typeof e && "undefined" !== typeof window && (e = null === window || void 0 === window ? void 0 : window.OfficeRuntime),
    "undefined" === typeof e && (e = {}),
    "undefined" !== typeof t && (t.storage = t.storage || (null === e || void 0 === e ? void 0 : e.storage),
    t.auth = t.auth || (null === e || void 0 === e ? void 0 : e.auth),
    t.getAccessToken = t.getAccessToken || (null === (n = null === e || void 0 === e ? void 0 : e.auth) || void 0 === n ? void 0 : n.getAccessToken),
    t.addin = t.addin || (null === e || void 0 === e ? void 0 : e.addin),
    t.isSetSupported = t.isSetSupported || (null === (r = null === e || void 0 === e ? void 0 : e.apiInformation) || void 0 === r ? void 0 : r.isSetSupported),
    t.license = t.license || (null === e || void 0 === e ? void 0 : e.license),
    t.message = t.message || (null === e || void 0 === e ? void 0 : e.message))
}
!function(e) {
    !function(e) {
        var t, n;
        !function(e) {
            e[e.None = 0] = "None",
            e[e.Auto = 1] = "Auto",
            e[e.Force = 2] = "Force"
        }(t = e.PopupOptions || (e.PopupOptions = {})),
        function(e) {
            e[e.UnsupportedUserIdentity = 13003] = "UnsupportedUserIdentity",
            e[e.UserAborted = 13004] = "UserAborted",
            e[e.InteractionRequired = 13005] = "InteractionRequired",
            e[e.ClientError = 13006] = "ClientError",
            e[e.ServerError = 13007] = "ServerError",
            e[e.NotAvailable = 13012] = "NotAvailable",
            e[e.InternalError = 5001] = "InternalError",
            e[e.InvalidApiArguments = 5013] = "InvalidApiArguments"
        }(n = e.AuthErrorCode || (e.AuthErrorCode = {}));
        var r = {
            ACCOUNT_UNAVAILABLE: n.UnsupportedUserIdentity,
            USER_CANCEL: n.UserAborted,
            USER_INTERACTION_REQUIRED: n.InteractionRequired,
            PERSISTENT_ERROR: n.ClientError,
            NO_NETWORK: n.ServerError,
            TRANSIENT_ERROR: n.ServerError,
            NESTED_APP_AUTH_UNAVAILABLE: n.NotAvailable
        }
          , i = {
            POPUP_WINDOW_ERROR: n.ClientError,
            USER_CANCELLED: n.UserAborted
        }
          , a = "access_token"
          , o = "xms_cc"
          , s = 0
          , c = !1
          , d = !1
          , l = void 0
          , u = null
          , f = void 0;
        e.clientCapabilities = [],
        e.upnCheck = !0,
        e.timeout = void 0,
        e.msal = "https://alcdn.msauth.net/browser-1p/2.28.1/js/msal-browser-1p.min.js",
        e.debugging = !1,
        e.delay = 0,
        e.delayMsal = 0,
        e.useMsal3 = void 0;
        var m = {}
          , p = function(e) {
            try {
                var t = "string" === typeof e ? e : e.data
                  , i = JSON.parse(t);
                if (i.requestId) {
                    var a = i.requestId;
                    if (m.hasOwnProperty(a)) {
                        var o = n.InternalError
                          , s = m[a]
                          , c = s[0]
                          , d = s[1];
                        delete m[a];
                        var l = i.token;
                        if (l && !0 === i.success && l.access_token && "number" == typeof l.expires_in)
                            return void c({
                                accessToken: l.access_token,
                                idToken: l.id_token,
                                expiresOn: new Date(Date.now() + 1e3 * l.expires_in)
                            });
                        var u = i.error;
                        if (u) {
                            var f = u.status;
                            r[f] && (o = r[f])
                        }
                        d({
                            code: o
                        })
                    }
                }
            } catch (e) {}
        }
          , y = {
            code: n.NotAvailable
        };
        function v(t) {
            if (0 === e.clientCapabilities.length)
                return t;
            var n = {};
            if (t)
                try {
                    n = JSON.parse(t)
                } catch (e) {}
            return n.hasOwnProperty(a) || (n[a] = {}),
            n[a][o] = {
                values: e.clientCapabilities
            },
            JSON.stringify(n)
        }
        function g(e, t) {
            var r = e.clientId || f
              , a = e.correlationId || OSF.OUtil.Guid.generateNewGuid()
              , o = Date.now()
              , u = function(e) {
                var n = function(e, n) {
                    var i = Date.now() - o;
                    !function(e, t, n, r, i, a) {
                        if (s > 0 && !n)
                            return;
                        s++,
                        "undefined" !== typeof OTel && OTel.OTelLogger.onTelemetryLoaded((function() {
                            var i = [oteljs.makeStringDataField("NestedClientId", e), oteljs.makeStringDataField("CorrelationId", t), oteljs.makeBooleanDataField("Popup", n), oteljs.makeInt64DataField("Duration", r), oteljs.makeInt64DataField("ErrorCode", a ? a.code : 0), oteljs.makeBooleanDataField("BridgeAvailable", h())];
                            OTel.OTelLogger.sendTelemetryEvent({
                                eventName: "Office.Extensibility.OfficeJs.NestedAppAuth.GetAccessToken",
                                dataFields: i,
                                eventFlags: {
                                    dataCategories: oteljs.DataCategories.ProductServiceUsage,
                                    diagnosticLevel: oteljs.DiagnosticLevel.NecessaryServiceDataEvent
                                }
                            })
                        }
                        ))
                    }(r, a, t, i, 0, n)
                };
                return e.then((function(e) {
                    n(0, null)
                }
                )).catch((function(e) {
                    n(0, e)
                }
                )),
                e
            };
            return h() ? u(new Promise((function(n, i) {
                var o = OSF.OUtil.Guid.generateNewGuid()
                  , s = e.scopes.join(" ")
                  , d = {
                    messageType: "NestedAppAuthRequest",
                    method: t ? "GetTokenPopup" : "GetToken",
                    requestId: o,
                    clientLibrary: "officejs",
                    sendTime: Date.now(),
                    tokenParams: {
                        clientId: r,
                        scope: s,
                        correlationId: a,
                        claims: v(e.claims)
                    }
                };
                c || (nestedAppAuthBridge.addEventListener("message", p),
                c = !0),
                m[o] = [n, i],
                nestedAppAuthBridge.postMessage(JSON.stringify(d))
            }
            ))) : u(l ? l.then((function() {
                if (!d)
                    return Promise.reject(y);
                var o = e.scopes.join(" ");
                return o = o.replace(/(\/.default)$/, ""),
                OSF.WebAuth.getToken(o, e.scopes, r, a, t, e.claims).then((function(e) {
                    return {
                        accessToken: e.Token,
                        expiresOn: e.MsalResult ? e.MsalResult.expiresOn : void 0
                    }
                }
                )).catch((function(e) {
                    var t = n.InternalError
                      , r = void 0;
                    if (e)
                        if (r = e.ErrorMessage,
                        e.MsalResult && "InteractionRequiredAuthError" === e.MsalResult.name)
                            t = n.InteractionRequired;
                        else if (e.ErrorCode) {
                            var a = e.ErrorCode.toUpperCase();
                            i[a] && (t = i[a])
                        }
                    return Promise.reject({
                        code: t,
                        description: r
                    })
                }
                ))
            }
            )) : Promise.reject(y))
        }
        function h() {
            return "undefined" !== typeof nestedAppAuthBridge
        }
        e.isBridgeAvailable = h,
        e.load = function(t, n, r, i) {
            return f = t,
            i && (e.clientCapabilities = i),
            l || (l = new Promise((function(a, o) {
                if (h())
                    a();
                else if (Office && Office.context && Office.context.auth && OSF.WebAuth)
                    try {
                        Office.context.auth.getAuthContextAsync((function(s) {
                            if ("succeeded" === s.status) {
                                if (!(u = s.value))
                                    return void o(y);
                                OSF.WebAuth.config = {
                                    authFlow: "authcode",
                                    authVersion: e.authVersion ? e.authVersion : null,
                                    msal: e.msal,
                                    delayWebAuth: e.delay,
                                    delayMsal: e.delayMsal,
                                    debugging: e.debugging,
                                    useMsal3: e.useMsal3,
                                    authority: e.authorityOverride ? e.authorityOverride : u.authorityBaseUrl,
                                    idp: "msa" === u.authorityType.toLowerCase() ? "msa" : "aad",
                                    appIds: [t],
                                    redirectUri: n || null,
                                    upn: u.userPrincipalName,
                                    prefetch: r,
                                    telemetryInstance: "otel",
                                    enableUpnCheck: e.upnCheck,
                                    enableConsoleLogging: e.debugging,
                                    checkActiveAccount: !0,
                                    tenantId: u.tenantId,
                                    timeout: e.timeout,
                                    clientCapabilities: i
                                },
                                OSF.WebAuth.load().then((function(e) {
                                    d = !0,
                                    a()
                                }
                                )).catch((function(e) {
                                    o(__assign({}, y, {
                                        description: e instanceof Event ? e.type : void 0
                                    }))
                                }
                                ))
                            } else
                                o(y)
                        }
                        ))
                    } catch (e) {
                        o(y)
                    }
                else
                    o(y)
            }
            )))
        }
        ,
        e.getAccessToken = function(e) {
            var r = null == e.popup ? t.None : e.popup;
            if (r === t.Auto && !e.directUserActionCallback)
                throw {
                    code: n.InvalidApiArguments
                };
            return g(e, r === t.Force).catch((function(i) {
                if (i.code == n.InteractionRequired && r === t.Auto && e.directUserActionCallback)
                    return e.directUserActionCallback().then((function(t) {
                        if (t)
                            return g(e, !0);
                        throw {
                            code: n.UserAborted
                        }
                    }
                    )).catch((function() {
                        throw {
                            code: n.UserAborted
                        }
                    }
                    ));
                throw i
            }
            ))
        }
    }(e.NestedAppAuth || (e.NestedAppAuth = {}))
}(OfficeFirstPartyAuth || (OfficeFirstPartyAuth = {})),
exposeOfficeRuntimeThroughOfficeNamespace("undefined" !== typeof OfficeRuntime && OfficeRuntime || void 0, "undefined" !== typeof Office && Office || void 0),
"undefined" !== typeof OSFPerformance && (OSFPerformance.hostInitializationEnd = OSFPerformance.now(),
OSFPerformance.totalJSHeapSize = OSFPerformance.getTotalJSHeapSize(),
OSFPerformance.usedJSHeapSize = OSFPerformance.getUsedJSHeapSize(),
OSFPerformance.jsHeapSizeLimit = OSFPerformance.getJSHeapSizeLimit());
