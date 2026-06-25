const express = require('express');
const { BotFrameworkAdapter, ActivityHandler, TurnContext } = require('botbuilder');
const { BlobServiceClient } = require('@azure/storage-blob');

const app = express();
app.use(express.json());

const adapter = new BotFrameworkAdapter({
  appId: process.env.MicrosoftAppId,
  appPassword: process.env.MicrosoftAppPassword
});

const CONTAINER = 'botstate';
const BLOB = 'convref.json';

async function saveRef(ref) {
  const client = BlobServiceClient.fromConnectionString(process.env.AzureWebJobsStorage);
  const container = client.getContainerClient(CONTAINER);
  await container.createIfNotExists();
  await container.getBlockBlobClient(BLOB).upload(JSON.stringify(ref), JSON.stringify(ref).length, { overwrite: true });
}

async function loadRef() {
  try {
    const client = BlobServiceClient.fromConnectionString(process.env.AzureWebJobsStorage);
    const data = await client.getContainerClient(CONTAINER).getBlockBlobClient(BLOB).downloadToBuffer();
    return JSON.parse(data.toString());
  } catch { return null; }
}

class CIBot extends ActivityHandler {
  constructor() {
    super();
    this.onConversationUpdate(async (context, next) => {
      await saveRef(TurnContext.getConversationReference(context.activity));
      await next();
    });
    this.onMessage(async (context, next) => {
      await saveRef(TurnContext.getConversationReference(context.activity));
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
  const ref = await loadRef();
  if (ref) {
    await adapter.continueConversation(ref, async (context) => {
      await context.sendActivity(`🔴 CI Failed\nBranch: ${req.body.branch}\nHata: ${req.body.error}\nKim: ${req.body.user}`);
    });
  }
  res.sendStatus(200);
});

const port = process.env.PORT || 3978;
app.listen(port, () => console.log(`Bot ${port} portunda çalışıyor`));
