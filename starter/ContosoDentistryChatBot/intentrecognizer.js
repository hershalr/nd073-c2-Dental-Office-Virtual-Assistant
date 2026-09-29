/**
 * Calls Conversational Language Understanding (CLU) using the same env names
 * the starter maps for LUIS: applicationId=project, endpointKey=key, endpoint=host.
 * LuisRecognizer cannot call CLU after LUIS retirement.
 */
class CluIntentRecognizer {
    constructor(config) {
        this.isConfigured = !!(config && config.applicationId && config.endpointKey && config.endpoint);
        if (!this.isConfigured) {
            return;
        }
        this.projectName = config.applicationId;
        this.endpointKey = config.endpointKey;
        this.endpoint = config.endpoint.replace(/\/$/, '');
        this.deploymentName = process.env.LuisDeploymentName || 'production';
        this.apiVersion = process.env.CluApiVersion || '2023-04-01';
    }

    async executeLuisQuery(context) {
        if (!this.isConfigured) {
            return { intents: {}, entities: {} };
        }

        const utterance = (context.activity && context.activity.text) ? context.activity.text : '';
        const url = `${this.endpoint}/language/:analyze-conversations?api-version=${this.apiVersion}`;

        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Ocp-Apim-Subscription-Key': this.endpointKey
            },
            body: JSON.stringify({
                kind: 'Conversation',
                analysisInput: {
                    conversationItem: {
                        id: '1',
                        participantId: '1',
                        text: utterance
                    }
                },
                parameters: {
                    projectName: this.projectName,
                    deploymentName: this.deploymentName,
                    stringIndexType: 'TextElement_V8'
                }
            })
        });

        if (!response.ok) {
            const detail = await response.text();
            throw new Error(`CLU prediction failed (${response.status}): ${detail}`);
        }

        const payload = await response.json();
        const prediction = (payload.result && payload.result.prediction) ? payload.result.prediction : {};
        const topIntent = prediction.topIntent || 'None';
        const intents = {};

        if (prediction.intents && typeof prediction.intents === 'object' && !Array.isArray(prediction.intents)) {
            Object.keys(prediction.intents).forEach((name) => {
                const item = prediction.intents[name] || {};
                intents[name] = { score: typeof item.confidenceScore === 'number' ? item.confidenceScore : (item.score || 0) };
            });
        } else if (Array.isArray(prediction.intents)) {
            prediction.intents.forEach((item) => {
                intents[item.category || item.intent] = {
                    score: typeof item.confidenceScore === 'number' ? item.confidenceScore : (item.score || 0)
                };
            });
        }

        if (!intents[topIntent]) {
            intents[topIntent] = { score: 1 };
        }

        return {
            text: utterance,
            intents,
            entities: {
                datetime: this._mapDatetimeEntities(prediction.entities || [])
            },
            cluEntities: prediction.entities || []
        };
    }

    _mapDatetimeEntities(entities) {
        const datetimeEntities = entities.filter((entity) => {
            const category = (entity.category || entity.extraInformation || '').toString().toLowerCase();
            return category.includes('datetime') || category.includes('date') || category.includes('time');
        });

        if (!datetimeEntities.length) {
            return undefined;
        }

        return datetimeEntities.map((entity) => ({
            timex: [entity.text || (entity.resolutions && entity.resolutions[0] && entity.resolutions[0].value) || '']
        }));
    }

    getTimeEntity(result) {
        const datetimeEntity = result.entities && result.entities.datetime;
        if (datetimeEntity && datetimeEntity[0] && datetimeEntity[0].timex && datetimeEntity[0].timex[0]) {
            return datetimeEntity[0].timex[0];
        }

        if (result.cluEntities && result.cluEntities.length) {
            const hit = result.cluEntities.find((entity) => {
                const category = (entity.category || '').toLowerCase();
                return category.includes('datetime') || category.includes('date') || category.includes('time');
            });
            if (hit && hit.text) {
                return hit.text;
            }
        }

        return undefined;
    }
}

module.exports = CluIntentRecognizer;
