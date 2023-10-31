/*
 This application demonstrates how to issue a call to a protected web API
 using the client credentials flow.  A request will be issued to
 Microsoft Graph using the application's own identity.
*/

// Microsoft Authentication Library (MSAL) for Node.js
const msal = require('@azure/msal-node')
const graph = require('@microsoft/microsoft-graph-client');

// Node.js Express Framework
const express = require('express')

// Used to make the HTTP GET request to the Graph API
const https = require('https')

// MSAL configs
const msalConfig = {
  auth: {
    // 'Application (client) ID' of app registration in Azure portal - this value is a GUID
    clientId: '',

    // Client secret 'Value' (not the ID) from 'Client secrets' in app registration in Azure portal
    clientSecret: '',

    // Full directory URL, in the form of https://login.microsoftonline.com/<tenant>
    authority: '',

    // 'Object ID' of app registration in Azure portal - this value is a GUID
    clientObjectId: ''
  }
}

// Initialize MSAL
const msalClient = new msal.ConfidentialClientApplication(msalConfig)

msal.OnBehalfOfClient

// In a client credentials flow, the scope is always in the format '<resource>/.default'
const tokenRequest = {
  scopes: ['https://graph.microsoft.com/.default']
}

const client = graph.Client.init({
  authProvider: async (done) => {
    try {
      done(null, '');
    } catch (err) {
      console.log('err', JSON.stringify(err, Object.getOwnPropertyNames(err)));
      done(err, null);
    }
  }
})

// Initialize Express
const app = express()

app.get('/api', async () => {
  const response = await client.api('/users').get();
  console.log('response', response);
  return true;
})

app.listen(8080, () => console.log('\nListening here:\nhttp://localhost:8080/'))