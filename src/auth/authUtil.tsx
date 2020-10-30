// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as msal from '@azure/msal-browser';

/**
 * Scopes you add here will be prompted for user consent during sign-in.
 * By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
 * For more information about OIDC scopes, visit:
 * https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#openid-connect-scopes
 */
const loginRequestScopes = {
    scopes: ['User.Read'] // optional Array<string>
} as msal.SilentRequest;
/**
 * Add here the scopes to request when obtaining an access token for MS Graph API. For more information, see:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/resources-and-scopes.md
 */
const tokenRequest = {
    // PrintJob.ReadWrite.All user.read openid profile offline_access
    scopes: ['User.Read', 'Printer.ReadWrite.All', 'PrintJob.ReadWrite.All', 'offline_access'],
    forceRefresh: false // Set this to "true" to skip a cached token and go to the server to get a new token
};

const msalConfig = {
    auth: {
        clientId: '<Application_Client_ID>',
        authority: 'https://login.microsoftonline.com/<Directory_Tenant_ID>', // default: https://login.microsoftonline.com/common
        redirectUri: 'http://localhost:3000/',
        knownAuthorities: [],
    },
    cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false // Set this to "true" if you are having issues on IE11 or Edge
    },
    system: {
        loggerOptions: {
            loggerCallback: (level: msal.LogLevel, message: string, containsPii: boolean): void => {
                if (containsPii) {
                    return;
                }
                switch (level) {
                    case msal.LogLevel.Error:
                        console.error(message);
                        return;
                    case msal.LogLevel.Info:
                        console.info(message);
                        return;
                    case msal.LogLevel.Verbose:
                        console.debug(message);
                        return;
                    case msal.LogLevel.Warning:
                        console.warn(message);
                        return;
                }
            },
            piiLoggingEnabled: false
        },
        windowHashTimeout: 60000,
        iframeHashTimeout: 6000,
        loadFrameTimeout: 0
    }
};

export const msalInstance = new msal.PublicClientApplication(msalConfig);

export const login = async (): Promise<any> => {
    return msalInstance.loginPopup(loginRequestScopes);
}

export const logout = async () => {
    msalInstance.logout();
}

export const tokenResponse = async () => await msalInstance.acquireTokenSilent(loginRequestScopes).catch(async (error) => {
    if (error instanceof msal.InteractionRequiredAuthError) {
        // fallback to interaction when silent call fails
        return await msalInstance.acquireTokenPopup(loginRequestScopes).catch(error => {
            console.error(error);
            throw error;
        });
    } else {
        throw error;
    }
});
