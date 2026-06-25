const { BotFrameworkAdapter, ActivityHandler, TurnContext } = require('botbuilder');
const { BlobServiceClient } = require('@azure/storage-blob');

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
  await container.getBlockBlobClient(BLOB).upload(
    JSON.stringify(ref), JSON.stringify(ref).length, { overwrite: true }
  );
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

module.exports = async function(context, req) {
  if (req.method === 'POST' && context.bindingData.route === 'ci-alert') {
    const ref = await loadRef();
    if (ref) {
      await adapter.continueConversation(ref, async (turnContext) => {
        await turnContext.sendActivity(
          `🔴 CI Failed\nBranch: ${req.body.branch}\nHata: ${req.body.error}\nKim: ${req.body.user}`
        );
      });
    }
    context.res = { status: 200 };
  } else {
    await new Promise((resolve, reject) => {
      adapter.processActivity(req, context.res, async (turnContext) => {
        await bot.run(turnContext);
        resolve();
      });
    });
  }
};
