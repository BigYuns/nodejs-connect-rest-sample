/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

module.exports = {
  creds: {
    redirectUrl: 'http://localhost:3000/token',
    clientID: '980996a8-d0cb-40f2-8f82-bf447ee7241b',
    clientSecret: 'fLP432:}+[hpbbghXJUCP55',
    identityMetadata: 'https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration',
    allowHttpForRedirectUrl: true, // For development only
    responseType: 'code',
    validateIssuer: false, // For development only
    responseMode: 'query',
    scope: ['User.Read', 'Mail.Send', 'Files.ReadWrite', 'People.Read', 'Sites.ReadWrite.All' ]
  }
};
