/*!
 * muffin-v2-polyfills
 * Companion extensions for atom-websdk v2.x projects.
 * Assumes window.Muffin is already loaded via the v2 SDK script tag.
 */

(function () {

    if (!window.Muffin) {
        console.warn("muffin-v2-polyfills: window.Muffin not found — load atom-websdk first.");
        return;
    }

    if (window.Muffin._v2PolyfillsApplied) return; // idempotent

    // ─── Internal helpers (used by polyfill methods below) ───────────────────

    function _delayTime(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    function _generateRandomString(length) {
        var result = '';
        var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
        for (var i = 0; i < length; i++) {
            result += chars.charAt(Math.floor(Math.random() * chars.length));
        }
        return result;
    }

    function _isCssURL(url) {
        return url.split('.').pop().toLowerCase() === 'css';
    }

    function _requireScript(src) {
        return new Promise(function (resolve) {
            var s = document.createElement('script');
            s.src = src;
            s.onreadystatechange = function () {
                if (s.readyState === 'loaded' || s.readyState === 'complete') {
                    s.onreadystatechange = null;
                    resolve(s);
                }
            };
            s.onload = function () { resolve(s); };
            document.documentElement.firstChild.appendChild(s);
        });
    }

    function _requireCss(href) {
        return new Promise(function (resolve) {
            var l = document.createElement('link');
            l.href = href;
            l.setAttribute('rel', 'stylesheet');
            l.onreadystatechange = function () {
                if (l.readyState === 'loaded' || l.readyState === 'complete') {
                    l.onreadystatechange = null;
                    resolve(l);
                }
            };
            l.onload = function () { resolve(l); };
            document.documentElement.firstChild.appendChild(l);
        });
    }

    function _getUserAgentName() {
        var ua = window.navigator.userAgent;
        if (ua.indexOf('Chrome')  !== -1) return 'chrome';
        if (ua.indexOf('Firefox') !== -1) return 'firefox';
        if (ua.indexOf('Edge')    !== -1) return 'edge';
        if (ua.indexOf('Safari')  !== -1) return 'safari';
        if (ua.indexOf('Opera')   !== -1 || ua.indexOf('OPR') !== -1) return 'opera';
        if (ua.indexOf('MSIE')    !== -1 || ua.indexOf('Trident') !== -1) return 'ie';
        return 'unknown';
    }

    function _checkIfTouchDevice() {
        return 'ontouchstart' in window || navigator.maxTouchPoints;
    }

    // ─── Muffin.Service ───────────────────────────────────────────────────────

    Muffin.Service = class MuffinService {
        static name = null;
        static lockedInterfaces = [];
        static lockedSubscriptions = [];
        static _cache = new Map();
        static _defaultTTL = 60 * 1000;
        static _cacheCleanerInterval = null;

        static initCacheCleaner() {
            if (!this._cacheCleanerInterval) {
                this._cacheCleanerInterval = setInterval(() => {
                    const now = Date.now();
                    for (const [key, { expiry }] of this._cache.entries()) {
                        if (expiry < now) this._cache.delete(key);
                    }
                }, 30 * 1000);
            }
        }

        static getCached(key) {
            const entry = this._cache.get(key);
            if (!entry) return null;
            if (entry.expiry < Date.now()) { this._cache.delete(key); return null; }
            return entry.value;
        }

        static setCached(key, value, ttl = this._defaultTTL) {
            this._cache.set(key, { value, expiry: Date.now() + ttl });
            this.initCacheCleaner();
        }

        static clearCache(key) {
            if (key) this._cache.delete(key);
            else this._cache.clear();
        }

        static async lockInterface(interfaceName, throttle = 500) {
            while (this.lockedInterfaces.includes(interfaceName)) {
                await new Promise(r => setTimeout(r, throttle));
            }
            this.lockedInterfaces.push(interfaceName);
        }

        static unlockInterface(interfaceName) {
            const idx = this.lockedInterfaces.indexOf(interfaceName);
            if (idx > -1) this.lockedInterfaces.splice(idx, 1);
        }

        static async lockSubscription(subscriptionInterface, throttle = 200) {
            while (this.lockedSubscriptions.includes(subscriptionInterface)) {
                await new Promise(r => setTimeout(r, throttle));
            }
            this.lockedSubscriptions.push(subscriptionInterface);
        }

        static unlockSubscription(subscriptionInterface) {
            const idx = this.lockedSubscriptions.indexOf(subscriptionInterface);
            if (idx > -1) this.lockedSubscriptions.splice(idx, 1);
        }

        constructor() {
            this.pubKey = _generateRandomString(6);
            if (this.constructor.name) {
                this.interface = Muffin.PostOffice.getOrCreateInterface(this.constructor.name);
            }
        }
    };

    // ─── DOMComponent prototype extensions ───────────────────────────────────

    var proto = Muffin.DOMComponent.prototype;

    proto.getElement = function (filter) {
        var node = this._getDomNode();
        return node ? node.querySelector(filter) : null;
    };

    proto.getElements = function (filter) {
        var node = this._getDomNode();
        return node ? node.querySelectorAll(filter) : [];
    };

    proto.toggleBtnBusyState = function (btnEl, state) {
        state = state || 'busy';
        if (state === 'busy') {
            btnEl.classList.add('busy');
            btnEl.setAttribute('disabled', true);
        } else {
            btnEl.classList.remove('busy');
            btnEl.removeAttribute('disabled');
        }
    };

    proto.toggleRootAttr = function (attrName, attrValue) {
        var node = this._getDomNode();
        if (!node) return;
        if (attrValue) node.setAttribute(attrName, attrValue);
        else node.removeAttribute(attrName);
    };

    proto.callParent = function (srcEl, ev) {
        var args = [srcEl, ev, {}, {}];
        if (srcEl.hasAttribute('pass-data')) {
            var d = {};
            srcEl.getAttribute('pass-data').split(',').forEach(function (k) {
                var key = k.trim(); d[key] = this.data[key];
            }, this);
            args[2] = d;
        }
        if (srcEl.hasAttribute('pass-uivars')) {
            var u = {};
            srcEl.getAttribute('pass-uivars').split(',').forEach(function (k) {
                var key = k.trim(); u[key] = this.uiVars[key];
            }, this);
            args[3] = u;
        }
        var parent = this.getParent();
        parent[srcEl.dataset.function].call(parent, ...args);
    };

    proto.callGrandParent = function (srcEl, ev) {
        var args = [srcEl, ev, {}, {}];
        if (srcEl.hasAttribute('pass-data')) {
            var d = {};
            srcEl.getAttribute('pass-data').split(',').forEach(function (k) {
                var key = k.trim(); d[key] = this.data[key];
            }, this);
            args[2] = d;
        }
        if (srcEl.hasAttribute('pass-uivars')) {
            var u = {};
            srcEl.getAttribute('pass-uivars').split(',').forEach(function (k) {
                var key = k.trim(); u[key] = this.uiVars[key];
            }, this);
            args[3] = u;
        }
        var gp = this.getParent().getParent();
        gp[srcEl.dataset.function].call(gp, ...args);
    };

    proto.awaitChildLoad = function (childName, timeout) {
        timeout = timeout || 3000;
        return new Promise(function (resolve, reject) {
            if (this.composedScope[childName]) return resolve(this.composedScope[childName]);
            this.interface.on('child-composed', function (child) {
                if (child === childName) resolve(this.composedScope[childName]);
            }.bind(this));
            setTimeout(function () { reject('Error: child loading timed out'); }, timeout);
        }.bind(this));
    };

    proto.initSubscriptions = function (subscriptionsArr) {
        return Promise.all(subscriptionsArr.map(function (sub) {
            return Muffin.WebInterface.subscribe(
                sub.host + '|||' + sub.interface,
                this.interface.name,
                sub.localInterfaceEvent
            );
        }, this));
    };

    proto.copyToClipboard = function (srcEl) {
        navigator.clipboard.writeText(srcEl.dataset.text);
        this.notifyUser("<i class='fas fa-check-circle'></i> Copied");
    };

    proto.notifyUser = function (msgTxt, duration, interfaceSuffix) {
        duration = duration || 4000;
        interfaceSuffix = interfaceSuffix || ':::notify-foreground';
        var text = (msgTxt instanceof HTMLElement) ? msgTxt.dataset.msg : msgTxt;
        PostOffice.publishToInterface('NotificationManager' + interfaceSuffix, {
            msgTxt: text,
            duration: duration
        });
    };

    proto.notifyUserDebounce = function (msgTxt, duration, interfaceSuffix, delay) {
        duration = duration || 4000;
        interfaceSuffix = interfaceSuffix || ':::notify-foreground';
        delay = delay || 1000;
        if (!this._notifyDebounceTimer) this._notifyDebounceTimer = null;
        clearTimeout(this._notifyDebounceTimer);
        this._notifyDebounceTimer = setTimeout(function () {
            PostOffice.publishToInterface('NotificationManager' + interfaceSuffix, {
                msgTxt: msgTxt, duration: duration
            });
        }, delay);
    };

    proto.getSessionUser = async function () {
        try {
            var res = await Muffin.AccountManagerService.getAccountDetails();
            var name = res.user_details.name;
            return { firstName: name.split(' ')[0], fullName: name, email: res.user_details.email };
        } catch (e) {
            console.error('ERROR: Failed to get Session User - ', e);
            return {};
        }
    };

    proto.toggleSurface = function (surfaceName, state, toggleClass) {
        state = state || 'switch';
        toggleClass = toggleClass || '_active';
        var node = this._getDomNode();
        if (!node) return;
        var el = node.querySelector("[surface='" + surfaceName + "']");
        if (!el) { console.warn('WARN: Surface not found - ', surfaceName); return; }
        var ops = { show: 'add', hide: 'remove', switch: 'toggle' };
        el.classList[ops[state] || 'toggle'](toggleClass);
    };

    proto.isSurfaceActive = function (surfaceName, toggleClass) {
        toggleClass = toggleClass || '_active';
        var node = this._getDomNode();
        if (!node) return false;
        var el = node.querySelector("[surface='" + surfaceName + "']");
        return el ? el.classList.contains(toggleClass) : false;
    };

    proto.toggleTargetSurface = async function (srcEl) {
        var beforeHook = srcEl.getAttribute('before-toggle-hook');
        var afterHook  = srcEl.getAttribute('after-toggle-hook');
        if (beforeHook) { this[beforeHook](srcEl); await _delayTime(10); }
        this.toggleSurface(srcEl.dataset.target, srcEl.dataset.state, srcEl.dataset['toggle-class']);
        if (afterHook) setTimeout(function () { this[afterHook](srcEl); }.bind(this), 10);
    };

    proto.toggleTargetTab = function (srcEl) {
        var targetTab   = srcEl.dataset.targettab;
        var tabGroup    = srcEl.dataset.tabgroup;
        if (!targetTab || !this._getDomNode()) return;

        var parentControlled = srcEl.hasAttribute ? srcEl.hasAttribute('parent-controlled') : false;
        var activeToggleEl;

        this._getDomNode()
            .querySelectorAll('[data-tabgroup="' + tabGroup + '"][data-targettab]')
            .forEach(function (el) {
                if (el.dataset.targettab === targetTab) {
                    el.classList.add('_active');
                    this.toggleSurface(el.dataset.targettab, 'show');
                    if (!el.hasAttribute('disabled')) activeToggleEl = el;
                } else {
                    el.classList.remove('_active');
                    this.toggleSurface(el.dataset.targettab, 'hide');
                }
            }.bind(this));

        if (!parentControlled &&
            targetTab === this.uiVars.activeTab &&
            tabGroup  === this.uiVars.activeTabGroup) return;

        this.uiVars.activeTab      = targetTab;
        this.uiVars.activeTabGroup = tabGroup;

        if (activeToggleEl) {
            var onOpen = activeToggleEl.getAttribute('on-open');
            if (onOpen && this[onOpen]) this[onOpen].call(this, this.uiVars.activeTab, srcEl);
        }
    };

    proto._getMuffinScope = function (targetEl) {
        if (!targetEl || !targetEl.hasAttribute('muffin-scope')) return this;
        var muffinScope = targetEl.getAttribute('muffin-scope');
        var scopeEl = this.getElement("[data-component='" + muffinScope + "']");
        return (scopeEl && scopeEl.constructedFrom) ? scopeEl.constructedFrom : this;
    };

    proto.renderSelectively = async function (options) {
        options = options || {};
        if (!this._getDomNode()) return;
        var _this = this;
        var linkedElms = this._getDomNode().querySelectorAll('[data-uivar]');

        [...linkedElms].forEach(function (el) {
            var renderFunc            = el.getAttribute('render-func');
            var postRenderFunc        = el.getAttribute('post-render');
            var skipEventHandlersExcept = el.getAttribute('skip-event-handlers-except');
            var renderEventRequired   = el.getAttribute('render-event');

            if (renderEventRequired && options.event !== renderEventRequired) return;

            var uiVarsToPass = el.dataset.uivar === '*' ? _this.uiVars : _this.uiVars[el.dataset.uivar];

            var updatedMarkup;
            if (renderFunc) {
                updatedMarkup = _this.constructor[renderFunc]
                    ? _this.constructor[renderFunc].call(_this.constructor, uiVarsToPass, _this.uid)
                    : el[el.dataset.relation];
            } else {
                updatedMarkup = uiVarsToPass;
            }

            el[el.dataset.relation] = updatedMarkup;

            var containers = skipEventHandlersExcept
                ? Array.from(el.querySelectorAll(skipEventHandlersExcept))
                : [el];

            containers.forEach(function (container) {
                [['on-click', 'onclick'], ['on-input', 'oninput'], ['on-change', 'onchange'],
                 ['on-contextmenu', 'oncontextmenu'], ['on-longpress', 'onlongpress'],
                 ['on-longpress-end', 'onlongpressend']
                ].forEach(function (pair) {
                    var attr = pair[0], evProp = pair[1];
                    container.querySelectorAll('[' + attr + ']').forEach(function (targetEl) {
                        var scope = _this._getMuffinScope(targetEl);
                        var methodName = targetEl.getAttribute(attr);
                        targetEl[evProp] = function (ev) {
                            if (evProp === 'oncontextmenu') ev.preventDefault();
                            scope[methodName].call(scope, targetEl, ev);
                        };
                    });
                });
            });

            if (postRenderFunc) {
                setTimeout(function () { _this[postRenderFunc].call(_this, el); }, 100);
            }
        });
    };

    proto._loadAllDependencies = async function () {
        if (!this.constructor.dependencies) return;
        await Promise.all(this.constructor.dependencies.map(function (src) {
            return _isCssURL(src) ? _requireCss(src) : _requireScript(src);
        }));
    };

    proto.gotoRoute = function (srcEl) {
        if (!Muffin._router) {
            console.warn('WARN: gotoRoute called but Muffin._router not found.');
            return;
        }
        var target = srcEl.dataset.target;
        if (!target) { console.warn('WARN: gotoRoute — no data-target specified.'); return; }

        if (target.startsWith('https://') || target.startsWith('http://')) {
            var linkTarget = srcEl.getAttribute('target');
            if (linkTarget === '_blank') window.open(target, '_blank');
            else window.location.href = target;
        } else if (target.startsWith('event:::')) {
            var payload = { target: target.split('event:::')[1], params: {} };
            for (var i = 0; i < srcEl.attributes.length; i++) {
                var attr = srcEl.attributes[i];
                if (attr.name.startsWith('event-')) payload.params[attr.name.slice(6)] = attr.value;
            }
            if (payload.target) PostOffice.sockets.global.dispatchMessage('goto-route-event', payload);
        } else {
            var routeParams = srcEl.getAttribute('route-params');
            Muffin._router.go(target, routeParams);
        }

        var onAfter = srcEl.getAttribute('on-after');
        if (onAfter) this[onAfter].call(this, srcEl, target);
    };

    // ─── Router extension ─────────────────────────────────────────────────────

    Muffin.Router.prototype.updateHistory = function (title, url) {
        window.history.pushState({ name: title, url: url }, title, url);
    };

    // ─── HTMLElement extension ────────────────────────────────────────────────

    HTMLElement.prototype.addOneTimeEventListener = function (eventName, cb) {
        var self = this;
        function handler(ev) {
            self.removeEventListener(eventName, handler);
            return cb.call(self, ev);
        }
        this.addEventListener(eventName, handler);
    };

    // ─── Array / String extensions ────────────────────────────────────────────

    if (!Array.prototype.splitIntoMultipleArrays) {
        Object.defineProperty(Array.prototype, 'splitIntoMultipleArrays', {
            value: function (chunkSize) {
                var result = [];
                for (var i = 0; i < this.length; i += chunkSize) result.push(this.slice(i, i + chunkSize));
                return result;
            },
            enumerable: false
        });
    }

    if (!String.prototype.ellipsify) {
        String.prototype.ellipsify = function (threshold, fromStart, fromEnd) {
            threshold = threshold || 9; fromStart = fromStart || 6; fromEnd = fromEnd || 6;
            return this.length > threshold
                ? this.substr(0, fromStart) + '......' + this.substr(-fromEnd)
                : String(this);
        };
    }

    // ─── Muffin globals ───────────────────────────────────────────────────────

    Muffin.userAgent    = _getUserAgentName();
    Muffin.isTouchDevice = _checkIfTouchDevice();

    Muffin._v2PolyfillsApplied = true;

})();
