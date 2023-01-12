/**
 * Copyright 2019 Google Inc. All Rights Reserved.
 * Adapted by @olivomarco to connect to Microsoft Teams.
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
 const {ActivityTypes,
  CardFactory,
  MessageFactory,
  BotFrameworkAdapter} = require('botbuilder');
const protoToJson = require('./botlib/proto_to_json.js');
const dialogflowSessionClient =
    require('./botlib/dialogflow_session_client.js');
const filterResponses = require('./botlib/filter_responses.js');
const express = require('express');
const app = express();

//For authenticating dialogflow_session_client.js, create a Service Account and
// download its key file. Set the environmental variable
// GOOGLE_APPLICATION_CREDENTIALS to the key file's location.
//See https://dialogflow.com/docs/reference/v2-auth-setup and
// https://cloud.google.com/dialogflow/docs/setup for details.

const projectId = process.env.GOOGLE_PROJECT_ID;
const appId = process.env.MICROSOFT_APP_ID;
const appPassword = process.env.MICROSOFT_APP_PASSWORD;

const sessionClient = new dialogflowSessionClient(projectId);

// Create bot adapter, which defines how the bot sends and receives messages.
let adapter = new BotFrameworkAdapter({
  appId: appId,
  appPassword: appPassword,
  channelAuthTenant: process.env.MICROSOFT_TENANT_ID
});

app.post('/', (req, res) => {
  // Use the adapter to process the incoming web request into a TurnContext object.
  adapter.processActivity(req, res, async (turnContext) => {
    console.log("Logging*******************************************",JSON.stringify(turnContext,null,1));
    if (isMessage(turnContext)) {
      console.log("In is Message*********")
      const utterance = getMessageText(turnContext);
      const senderId = turnContext.activity.from.id;
      const payload = turnContext.activity;
      console.log("Before detect intent*********", utterance)
      const responses = (await sessionClient.detectIntent(
          utterance, senderId, null)).responseMessages;

      console.log("After detect intent", JSON.stringify(responses,null,1))
      const replies = await convertToTeamsMessage(turnContext, responses);
      console.log("MS replies*********", replies)
      await turnContext.sendActivities(replies);
    } else if(isMemberAdded(turnContext)) {
      console.log("Is Member Added****************8", JSON.stringify(turnContext,null,1))
      for (let idx in turnContext.activity.membersAdded) {
        if (turnContext.activity.membersAdded[idx].id !==
            turnContext.activity.recipient.id) {
          const result = await sessionClient.detectIntentWithEvent('TEAMS_WELCOME',
              projectId);
              console.log("After detect intent", JSON.stringify(result,null,1))
          const replies = await convertToTeamsMessage(turnContext,
              result.responseMessages);
              console.log("MS replies*******************", JSON.stringify(replies,null,1))
          await turnContext.sendActivity(replies);
        }
      }
    }
  });
});

function turnContextType(turnContext) {
  return turnContext.activity.type;
}

function isMessage(turnContext){
  return turnContextType(turnContext) === 'message';
}

function getMessageText(turnContext) {
  return turnContext.activity.text;
}

function isMemberAdded(turnContext){
  return Array.isArray(turnContext.activity.membersAdded);
}

async function convertToTeamsMessage(turnContext, responses){
  const replies = [];
  if (Array.isArray(responses)) {
    const filteredResponses = await filterResponses.filterResponses(responses, 'TEAMS');
    console.log("filtered response##########################################", JSON.stringify(filterResponses,null,1))
    filteredResponses.forEach((response)=> {
      let reply = {type: ActivityTypes.Message};
      switch (response.message) {
        case 'text': {
          reply.text = response.text.text[0];
        }
          break;

        case 'image': {
          reply.attachments = [(CardFactory.heroCard(
              '',
              CardFactory.images([response.image.imageUri])
          ))];
        }
          break;

        case 'card': {
          const buttons = response.card.buttons;
          let teamsButtons = [];
          if (Array.isArray(buttons) && buttons.length > 0) {
            buttons.forEach((button) => {
              if (button.postback.startsWith('http')) {
                teamsButtons.push({
                  type: 'openUrl',
                  title: button.text,
                  value: button.postback
                });
              } else {
                teamsButtons.push({
                  type: 'postBack',
                  title: button.text,
                  value: button.postback
                });
              }
            });
            reply.attachments = [(CardFactory.heroCard(
                response.card.title,
                response.card.subtitle,
                CardFactory.images([response.card.imageUri]),
                teamsButtons))];
          }
        }
          break;

        case 'quickReplies': {
          reply = MessageFactory.suggestedActions(
              response.quickReplies.quickReplies, response.quickReplies.title);
        }
          break;

        case 'payload': {
          console.log(response);
          const protoPayload = response.payload.fields.teams.structValue;
          reply = protoToJson.structProtoToJson(protoPayload);
        }
          break;

        default:
          break;
      }
      replies.push(reply);
    });
  }
  return replies;
}

module.exports = {
  convertToTeamsMessage,
  app
};
