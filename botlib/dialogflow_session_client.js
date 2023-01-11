/**
 * Copyright 2019 Google Inc. All Rights Reserved.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *    http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
/**
 * @fileoverview Contacts dialogflow and returns response.
 */
const { SessionsClient } = require('@google-cloud/dialogflow-cx');
const jsonToProto = require('./json_to_proto')
module.exports = class DialogflowSessionClient {

  constructor(projectId){
    this.projectId=projectId;
    this.agentId="c946912f-53b4-4098-b7f8-5abb0e17dd9a";
    this.location="us-central1";
    this.sessionClient = new SessionsClient({ apiEndpoint: `${this.location}-dialogflow.googleapis.com` });
  }

  constructRequest(text, sessionPath, payload) {
    return {
      session: sessionPath,
      queryInput: {
        text: {
          text: text
        },
        languageCode: 'en'
      },
      queryParams: {
        payload: jsonToProto.jsonToStructProto(payload)
      }
    };
  }

  constructRequestWithEvent(eventName, sessionPath) {
    return {
      session: sessionPath,
      queryInput: {
        event: {
          event: eventName
        },
        languageCode: 'en'
      }
    };
  }
  //This function calls Dialogflow DetectIntent API to retrieve the response
  //https://cloud.google.com/dialogflow/docs/reference/rest/v2/projects.agent.sessions/detectIntent
  async detectIntentHelper(detectIntentRequest) {
    let [response] = await this.sessionClient.detectIntent(detectIntentRequest);
    return (response.queryResult);
  }

  async detectIntent(text, sessionId, payload) {
    const sessionPath = this.sessionClient.projectLocationAgentSessionPath(
          this.projectId,
          location,
          agentId,
          sessionId
      )
    const request = this.constructRequest(text, sessionPath, payload);
    return await this.detectIntentHelper(request);
  }

  async detectIntentWithEvent(eventName, sessionId) {
    const sessionPath = this.sessionClient.projectLocationAgentSessionPath(
      this.projectId,
      location,
      agentId,
      sessionId
  )
    const request = this.constructRequestWithEvent(
        eventName, sessionPath);
    return await this.detectIntentHelper(request);
  }
};
