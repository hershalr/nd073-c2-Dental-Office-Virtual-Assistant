// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const path = require('path');
const dotenv = require('dotenv');
dotenv.config({ path: path.join(__dirname, '.env') });

const express = require('express');
const {
    CloudAdapter,
    ConfigurationBotFrameworkAuthentication
} = require('botbuilder');
const { DentaBot } = require('./bot');

console.log('[startup] Node', process.version);
console.log('[startup] AppId set:', !!process.env.MicrosoftAppId);
console.log('[startup] AppPassword set:', !!process.env.MicrosoftAppPassword);
console.log('[startup] AppType:', process.env.MicrosoftAppType || '(default MultiTenant)');
console.log('[startup] AppTenantId set:', !!process.env.MicrosoftAppTenantId);
console.log('[startup] QnA host set:', !!process.env.QnAEndpointHostName);
console.log('[startup] QnA kb set:', !!process.env.QnAKnowledgebaseId);
console.log('[startup] Luis project set:', !!process.env.LuisAppId);
console.log('[startup] Scheduler set:', !!process.env.SchedulerEndpoint);

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(process.env);
const adapter = new CloudAdapter(botFrameworkAuthentication);

adapter.onTurnError = async (context, error) => {
    console.error('[onTurnError]', error);
    try {
        await context.sendActivity('The bot encountered an error or bug.');
        await context.sendActivity(String(error && error.message ? error.message : error));
    } catch (sendErr) {
        console.error('[onTurnError] failed to send error activity', sendErr);
    }
};

let myBot;
try {
    const configuration = {
        QnAConfiguration: {
            knowledgeBaseId: process.env.QnAKnowledgebaseId,
            endpointKey: process.env.QnAAuthKey,
            host: (process.env.QnAEndpointHostName || '').replace(/\/$/, '')
        },
        LuisConfiguration: {
            applicationId: process.env.LuisAppId,
            endpointKey: process.env.LuisAPIKey,
            endpoint: process.env.LuisAPIHostName
        },
        SchedulerConfiguration: {
            SchedulerEndpoint: process.env.SchedulerEndpoint
        }
    };
    myBot = new DentaBot(configuration, {});
    console.log('[startup] DentaBot constructed OK');
} catch (err) {
    console.error('[startup] DentaBot construction FAILED', err);
    throw err;
}

const server = express();
server.use(express.json());
server.use(express.urlencoded({ extended: true }));

server.get('/', (_req, res) => {
    res.status(200).send('Contoso Dentistry bot is running.');
});

server.post('/api/messages', async (req, res) => {
    console.log('[api/messages] hit', {
        hasAuth: !!(req.headers && req.headers.authorization),
        activityType: req.body && req.body.type,
        text: req.body && req.body.text
    });
    try {
        await adapter.process(req, res, async (context) => {
            console.log('[api/messages] authenticated turn', context.activity && context.activity.type);
            await myBot.run(context);
        });
        console.log('[api/messages] process finished');
    } catch (err) {
        console.error('[api/messages] ERROR', err);
        if (!res.headersSent) {
            res.status(500).send('Bot error');
        }
    }
});

const port = process.env.port || process.env.PORT || 3978;
server.listen(port, () => {
    console.log(`[startup] listening on ${port}`);
});
