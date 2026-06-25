const express = require('express');
const { BotFrameworkAdapter, ActivityHandler, TurnContext } = require('botbuilder');
const fetch = require('node-fetch');

const app = express();
app.use(express.json());

const adapter = new BotFrameworkAdapter({
  appId: process.env.MicrosoftAppId,
  appPassword: process.env.MicrosoftAppPassword
});

let conversationRef = null;

class CIBot extends ActivityHandler {
  constructor() {
    super();
    this.onConversationUpdate(async (context, next) => {
      conversationRef = TurnContext.getConversationReference(context.activity);
      console.log('Conversation reference kaydedildi');
      await next();
    });
    this.onMessage(async (context, next) => {
      conversationRef = TurnContext.getConversationReference(context.activity);
      const reply = context.activity.text;
      const user = context.activity.from.name;
      console.log(`Reply: ${reply} | User: ${user}`);
      await next();
    });
  }
}

const bot = new CIBot();

app.post('/api/messages', (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

app.post('/api/ci-alert', async (req, res) => {
  console.log('CI Alert geldi:', req.body);
  if (conversationRef) {
    await adapter.continueConversation(conversationRef, async (context) => {
      await context.sendActivity(`🔴 CI Failed\nBranch: ${req.body.branch}\nHata: ${req.body.error}\nKim: ${req.body.user}`);
    });
  } else {
    console.log('Conversation reference yok henüz');
  }
  res.sendStatus(200);
});

const port = process.env.PORT || 3978;
app.listen(port, () => console.log(`Bot ${port} portunda çalışıyor`));
