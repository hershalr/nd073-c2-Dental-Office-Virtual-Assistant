// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, MessageFactory } = require('botbuilder');
const { QnAMaker } = require('botbuilder-ai');
const DentistScheduler = require('./dentistscheduler');
const IntentRecognizer = require('./intentrecognizer');

class DentaBot extends ActivityHandler {
    constructor(configuration, qnaOptions) {
        super();
        if (!configuration) throw new Error('[DentaBot]: Missing parameter. configuration is required');

        this.qnaMaker = new QnAMaker(configuration.QnAConfiguration, qnaOptions);
        this.scheduler = new DentistScheduler(configuration.SchedulerConfiguration);
        this.intentRecognizer = new IntentRecognizer(configuration.LuisConfiguration);

        this.onMessage(async (context, next) => {
            const qnaResults = await this.qnaMaker.getAnswers(context);
            const luisResult = await this.intentRecognizer.executeLuisQuery(context);

            let topIntent = 'None';
            let topScore = 0;
            if (luisResult && luisResult.intents) {
                Object.keys(luisResult.intents).forEach((intentName) => {
                    const score = luisResult.intents[intentName].score || 0;
                    if (score > topScore) {
                        topScore = score;
                        topIntent = intentName;
                    }
                });
            }

            if (topIntent === 'GetAvailability' && topScore > 0.5) {
                const availability = await this.scheduler.getAvailability();
                await context.sendActivity(availability);
            } else if (topIntent === 'ScheduleAppointment' && topScore > 0.5) {
                const time = this.intentRecognizer.getTimeEntity(luisResult);
                if (time) {
                    const confirmation = await this.scheduler.scheduleAppointment(time);
                    await context.sendActivity(confirmation);
                } else {
                    await context.sendActivity('I can help schedule an appointment. Please include a time, for example: "Book me at 2pm".');
                }
            } else if (qnaResults && qnaResults[0]) {
                await context.sendActivity(qnaResults[0].answer);
            } else {
                await context.sendActivity('I can answer Contoso Dentistry FAQs or help you check availability and schedule an appointment.');
            }

            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            const welcomeText = 'Welcome to Contoso Dentistry! Ask about our office FAQs, say "What appointments are available?", or "Schedule an appointment at 2pm".';
            for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
                }
            }
            await next();
        });
    }
}

module.exports.DentaBot = DentaBot;
