(function () {
    const XHR = XMLHttpRequest.prototype;
    const open = XHR.open;
    const send = XHR.send;
    const setRequestHeader = XHR.setRequestHeader;

    // Aggressive Catch: Capture ALL POST requests
    function isTarget(url, method) {
        if (method === 'POST') {
            // Only log specific hits to avoid spam
            if (url.includes('list') || url.includes('fakturpajak')) {
                console.log("Interceptor TARGET MATCH:", url);
            }
            return true;
        }
        return false;
    }

    // Intercept XHR
    XHR.open = function (method, url) {
        this._method = method;
        this._url = url;
        this._requestHeaders = {};
        return open.apply(this, arguments);
    };

    XHR.setRequestHeader = function (header, value) {
        this._requestHeaders[header] = value;
        return setRequestHeader.apply(this, arguments);
    };

    XHR.send = function (postData) {
        this.addEventListener('load', function () {
            if (isTarget(this._url, this._method)) {
                try {
                    // Try to parse response
                    let responseData = null;
                    if (this.responseText && this.getResponseHeader('content-type') && this.getResponseHeader('content-type').includes('json')) {
                        try {
                            responseData = JSON.parse(this.responseText);
                        } catch (e) { }
                    } else if (this.responseText) {
                        try { responseData = JSON.parse(this.responseText); } catch (ev) { }
                    }

                    const layout = {
                        type: 'TAX_INTERCEPT',
                        url: this._url,
                        payload: postData,
                        auth: this._requestHeaders['Authorization'] || this._requestHeaders['authorization'],
                        response: responseData // CAPTURED RESPONSE
                    };
                    window.postMessage(layout, '*');
                } catch (e) { }
            }
        });
        return send.apply(this, arguments);
    };

    // Intercept Fetch
    const originalFetch = window.fetch;
    window.fetch = async function (...args) {
        const [resource, config] = args;
        const response = await originalFetch.apply(this, args);

        const url = (typeof resource === 'string') ? resource : resource.url;
        const method = (config && config.method) ? config.method : 'GET';

        if (isTarget(url, method)) {
            try {
                // Find auth
                let auth = null;
                if (config && config.headers) {
                    if (config.headers instanceof Headers) {
                        auth = config.headers.get('Authorization');
                    } else {
                        const keys = Object.keys(config.headers);
                        const authKey = keys.find(k => k.toLowerCase() === 'authorization');
                        if (authKey) auth = config.headers[authKey];
                    }
                }

                // Clone response to read body
                const clone = response.clone();
                clone.json().then(responseData => {
                    const layout = {
                        type: 'TAX_INTERCEPT',
                        url: url,
                        payload: config.body,
                        auth: auth,
                        response: responseData // CAPTURED RESPONSE
                    };
                    window.postMessage(layout, '*');
                }).catch(err => {
                    // If not json, maybe just send headers?
                    // For now we only care about JSON lists.
                });

            } catch (e) { }
        }
        return response;
    };
})();
