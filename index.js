const express = require('express');
const { BotFrameworkAdapter, ActivityHandler } = require('botbuilder');
const fetch = require('node-fetch');

const app = express();
app.use(express.json());

const adapter = new BotFrameworkAdapter({
  appId: 'e4dc299b-86f8-412f-b8b1-1d087bf9426e',
  appPassword: '967dedd9-cafc-4a91-9c7f-b748e135c3f8'
});

// Teams'ten reply gelince çalışır
class CIBot extends ActivityHandler {
  constructor() {
    super();
    this.onMessage(async (context, next) => {
      const reply = context.activity.text;
      const user = context.activity.from.name;
      const threadId = context.activity.conversation.id;

      // n8n'e gönder (sonra dolduracağız)
      console.log(`Reply: ${reply} | User: ${user}`);

      await next();
    });
  }
}

const bot = new CIBot();

// Teams'ten gelen mesajlar
app.post('/api/messages', (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

// CI'dan gelen alert
app.post('/api/ci-alert', async (req, res) => {
  console.log('CI Alert geldi:', req.body);
  res.sendStatus(200);
});

const port = process.env.PORT || 3978;
app.listen(port, () => console.log(`Bot ${port} portunda çalışıyor`));
