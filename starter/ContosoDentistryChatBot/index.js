// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const path = require('path');
const dotenv = require('dotenv');
const ENV_FILE = path.join(__dirname, '.env');
dotenv.config({ path: ENV_FILE });

const restify = require('restify');
const {
    CloudAdapter,
    ConfigurationServiceClientCredentialFactory,
    createBotFrameworkAuthenticationFromConfiguration
} = require('botbuilder');
const { DentaBot } = require('./bot');

const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.listen(process.env.port || process.env.PORT || 3978, () => {
    console.log(`\n${server.name} listening to ${server.url}`);
});

// Single Tenant Azure Bot requires App Type + Tenant Id (lab botreg is Single Tenant).
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
    MicrosoftAppId: process.env.MicrosoftAppId,
    MicrosoftAppPassword: process.env.MicrosoftAppPassword,
    MicrosoftAppType: process.env.MicrosoftAppType || 'SingleTenant',
    MicrosoftAppTenantId: process.env.MicrosoftAppTenantId
});

const botFrameworkAuthentication = createBotFrameworkAuthenticationFromConfiguration(null, credentialsFactory);
const adapter = new CloudAdapter(botFrameworkAuthentication);

const onTurnErrorHandler = async (context, error) => {
    console.error(`\n [onTurnError] unhandled error: ${error}`);
    console.error(error && error.stack ? error.stack : error);
    await context.sendActivity('The bot encountered an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};

adapter.onTurnError = onTurnErrorHandler;

const QnAConfiguration = {
    knowledgeBaseId: process.env.QnAKnowledgebaseId,
    endpointKey: process.env.QnAAuthKey,
    host: process.env.QnAEndpointHostName
};

const LuisConfiguration = {
    applicationId: process.env.LuisAppId,
    endpointKey: process.env.LuisAPIKey,
    endpoint: process.env.LuisAPIHostName
};

const SchedulerConfiguration = {
    SchedulerEndpoint: process.env.SchedulerEndpoint
};

const configuration = {
    QnAConfiguration,
    LuisConfiguration,
    SchedulerConfiguration
};

const myBot = new DentaBot(configuration, {});

server.post('/api/messages', async (req, res) => {
    await adapter.process(req, res, (context) => myBot.run(context));
});
