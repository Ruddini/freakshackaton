const axios = require("axios");
const querystring = require("querystring");
const { TeamsActivityHandler, CardFactory, TurnContext } = require("botbuilder");
const rawWelcomeCard = require("./adaptiveCards/welcome.json");
const rawLearnCard = require("./adaptiveCards/learn.json");
const cardTools = require("@microsoft/adaptivecards-tools");
const rawMainCard = require("./adaptiveCards/main.json");
const calendarCard = require("./adaptiveCards/calendar.json");
const calendarReturn = require("./adaptiveCards/calendarCheck.json");
const rawMyCar = require("./adaptiveCards/mycar.json");
const rawSubmit = require("./adaptiveCards/submit.json");
const rawMyDeskLoc = require("./adaptiveCards/mydesk_location.json");
const rawMyDeskNo = require("./adaptiveCards/mydesk_num.json");
const rawMyDeskDate = require("./adaptiveCards/mydesk_date.json");
const rawExplainAcronym = require("./adaptiveCards/explainAcronym.json");
const rawExplained = require("./adaptiveCards/explained.json");
const exchange = require("./adaptiveCards/exchange.json");
const exchangeCheck = require("./adaptiveCards/exchangeCheck.json");
const offlineCard = require("./adaptiveCards/offline.json");



class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    // record the likeCount
    this.likeCountObj = { likeCount: 0 };

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      // Trigger command by IM text
      switch (txt) {
        case "?": {
          var mainCard;
          try{
            var test = await axios.get("http://bot-backend-sesi.azurewebsites.net/mycar/");
            console.log(test.statusCode);
            var mainCard = cardTools.AdaptiveCards.declareWithoutData(rawMainCard).render();
          }
          catch(err) {
            var mainCard = cardTools.AdaptiveCards.declareWithoutData(offlineCard).render();
          }
          const card = mainCard;
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }

        case "welcome": {
          const card = cardTools.AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        case "learn": {
          this.likeCountObj.likeCount = 0;
          const card = cardTools.AdaptiveCards.declare(rawLearnCard).render(this.likeCountObj);
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        /**
         * case "yourCommand": {
         *   await context.sendActivity(`Add your response here!`);
         *   break;
         * }
         */
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    // Listen to MembersAdded event, view https://docs.microsoft.com/en-us/microsoftteams/platform/resources/bot-v3/bots-notifications for more events
    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card = cardTools.AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
      }
      await next();
    });
  }

  // Invoked when an action is taken on an Adaptive Card. The Adaptive Card sends an event to the Bot and this
  // method handles that event.
  async onAdaptiveCardInvoke(context, invokeValue) {
    // The verb "userlike" is sent from the Adaptive Card defined in adaptiveCards/learn.json
    if (invokeValue.action.verb === "mycar") {
      var test = await axios.get("http://bot-backend-sesi.azurewebsites.net/mycar/");
      console.log(test.data);
      rawMyCar.actions[0].title=test.data[0];
      rawMyCar.actions[1].title=test.data[1];
      rawMyCar.actions[2].title=test.data[2];      
      const card = cardTools.AdaptiveCards.declare(rawMyCar).render();
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
      return { statusCode: 200 };
    }
    else if (invokeValue.action.verb === "submit") {
      const card = cardTools.AdaptiveCards.declare(rawSubmit).render();
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
      return { statusCode: 200 };
    }
    else if (invokeValue.action.verb === "main") {
      const card = cardTools.AdaptiveCards.declare(rawMainCard).render();
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
      return { statusCode: 200 };
    }
    else if (invokeValue.action.verb === "personstatus") {
      const card = cardTools.AdaptiveCards.declare(calendarCard).render();
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
      return { statusCode: 200 };
    }
    else if (invokeValue.action.verb === "calendarCheck") {
      var name = invokeValue.action.data.name_surname;
      var time = invokeValue.action.data.time;
      var test = await axios.get(`http://bot-backend-sesi.azurewebsites.net/meeting/free_time/${name}/${time}`);  
      calendarReturn.body[1].text=test.data;
      const card = cardTools.AdaptiveCards.declare(calendarReturn).render();
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
      return { statusCode: 200 };
    }
    else if (invokeValue.action.verb === "currency") {
      const card = cardTools.AdaptiveCards.declare(exchange).render();
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
      return { statusCode: 200 };
    }
    else if (invokeValue.action.verb === "exchange") {
      var exchangeFrom = invokeValue.action.data.exchangeFrom;
      var exchangeTo = invokeValue.action.data.exchangeTo;
      var amount = invokeValue.action.data.exchangeAmount;
      var test = await axios.get(`http://bot-backend-sesi.azurewebsites.net/currency/{currency}/?currency1=${exchangeFrom}&currency2=${exchangeTo}&amount=${amount}`);  
      exchangeCheck.body[1].text=`1 ${exchangeTo} = ${test.data['info']['rate']} ${exchangeFrom}`;
      exchangeCheck.body[3].text= `${amount} ${exchangeTo} = ${test.data['result']} ${exchangeFrom}`;
      const card = cardTools.AdaptiveCards.declare(exchangeCheck).render();
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
      return { statusCode: 200 };
    }
    else if (invokeValue.action.verb === "mydesk_date") {
      const card = cardTools.AdaptiveCards.declare(rawMyDeskDate).render();
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
      return { statusCode: 200 };
    }
    else if (invokeValue.action.verb === "mydesk_location") {
      const card = cardTools.AdaptiveCards.declare(rawMyDeskLoc).render();
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
      return { statusCode: 200 };
    }
    else if (invokeValue.action.verb === "mydesk_num") {
      const card = cardTools.AdaptiveCards.declare(rawMyDeskNo).render();
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
      return { statusCode: 200 };
    }
    else if (invokeValue.action.verb === "acronyms") {
      const card = cardTools.AdaptiveCards.declare(rawExplainAcronym).render();
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
      return { statusCode: 200 };
    }
    else if (invokeValue.action.verb === "explain") {
      var toExplain = invokeValue.action.data.acAcronym;
      var explained = await axios.get(`http://bot-backend-sesi.azurewebsites.net/shortcut/${toExplain}/`);
      rawExplained.body[0].text = toExplain;
      if (explained.data[0] != null) {
        rawExplained.body[1].text = explained.data[0];
        rawExplained.actions[0].url = explained.data[1];
      } else {
        rawExplained.body[1].text = "Unfortunately we don't know this yet, we noted your request and will try to come up with an explanation for it :)";
        rawExplained.actions[0].url = explained.data[1];
      }
      const card = cardTools.AdaptiveCards.declare(rawExplained).render();
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
      return { statusCode: 200 };
    }
    
    else if (invokeValue.action.verb === "mydesk_location") {
      const card = cardTools.AdaptiveCards.declare(rawMyDeskNo).render();
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
      return { statusCode: 200 };
    }
    else {
      const card = cardTools.AdaptiveCards.declare(rawMainCard).render();
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
      return { statusCode: 200 };
    }
  }

  // Message extension Code
  // Action.
  handleTeamsMessagingExtensionSubmitAction(context, action) {
    switch (action.commandId) {
      case "createCard":
        return createCardCommand(context, action);
      case "shareMessage":
        return shareMessageCommand(context, action);
      default:
        throw new Error("NotImplemented");
    }
  }

  // Search.
  async handleTeamsMessagingExtensionQuery(context, query) {
    const searchQuery = query.parameters[0].value;
    const response = await axios.get(
      `http://registry.npmjs.com/-/v1/search?${querystring.stringify({
        text: searchQuery,
        size: 8,
      })}`
    );

    const attachments = [];
    response.data.objects.forEach((obj) => {
      const heroCard = CardFactory.heroCard(obj.package.name);
      const preview = CardFactory.heroCard(obj.package.name);
      preview.content.tap = {
        type: "invoke",
        value: { name: obj.package.name, description: obj.package.description },
      };
      const attachment = { ...heroCard, preview };
      attachments.push(attachment);
    });

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      },
    };
  }

  async handleTeamsMessagingExtensionSelectItem(context, obj) {
    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [CardFactory.heroCard(obj.name, obj.description)],
      },
    };
  }

  // Link Unfurling.
  handleTeamsAppBasedLinkQuery(context, query) {
    const attachment = CardFactory.thumbnailCard("Thumbnail Card", query.url, [query.url]);

    const result = {
      attachmentLayout: "list",
      type: "result",
      attachments: [attachment],
    };

    const response = {
      composeExtension: result,
    };
    return response;
  }
}

function createCardCommand(context, action) {
  // The user has chosen to create a card by choosing the 'Create Card' context menu command.
  const data = action.data;
  const heroCard = CardFactory.heroCard(data.title, data.text);
  heroCard.content.subtitle = data.subTitle;
  const attachment = {
    contentType: heroCard.contentType,
    content: heroCard.content,
    preview: heroCard,
  };

  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: [attachment],
    },
  };
}

function shareMessageCommand(context, action) {
  // The user has chosen to share a message by choosing the 'Share Message' context menu command.
  let userName = "unknown";
  if (
    action.messagePayload &&
    action.messagePayload.from &&
    action.messagePayload.from.user &&
    action.messagePayload.from.user.displayName
  ) {
    userName = action.messagePayload.from.user.displayName;
  }

  // This Message Extension example allows the user to check a box to include an image with the
  // shared message.  This demonstrates sending custom parameters along with the message payload.
  let images = [];
  const includeImage = action.data.includeImage;
  if (includeImage === "true") {
    images = [
      "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQtB3AwMUeNoq4gUBGe6Ocj8kyh3bXa9ZbV7u1fVKQoyKFHdkqU",
    ];
  }
  const heroCard = CardFactory.heroCard(
    `${userName} originally sent this message:`,
    action.messagePayload.body.content,
    images
  );

  if (
    action.messagePayload &&
    action.messagePayload.attachment &&
    action.messagePayload.attachments.length > 0
  ) {
    // This sample does not add the MessagePayload Attachments.  This is left as an
    // exercise for the user.
    heroCard.content.subtitle = `(${action.messagePayload.attachments.length} Attachments not included)`;
  }

  const attachment = {
    contentType: heroCard.contentType,
    content: heroCard.content,
    preview: heroCard,
  };

  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: [attachment],
    },
  };
}

module.exports.TeamsBot = TeamsBot;
