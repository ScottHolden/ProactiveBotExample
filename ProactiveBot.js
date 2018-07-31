var restify = require('restify');
var builder = require('botbuilder');
var azure = require('azure-storage');
var botbuilder_azure = require("botbuilder-azure");

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
   console.log('%s listening to %s', server.name, server.url); 
});
  
// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
    openIdMetadata: process.env.BotOpenIdMetadata 
});

// Create your bot with a function to receive messages from the user
var bot = new builder.UniversalBot(connector);

// Listen for messages from users 
server.post('/api/messages', connector.listen());

// Set up our state storage
var tableName = 'botdata';
var azureTableClient = new botbuilder_azure.AzureTableClient(tableName, process.env['AzureWebJobsStorage']);
var tableStorage = new botbuilder_azure.AzureBotStorage({ gzipData: false }, azureTableClient);
bot.set('storage', tableStorage);

// Setup a body parser
var plainTextParser = (req, res, next) => {
    req.rawBody = '';
    req.setEncoding('utf8');
    req.on('data', function(part) {
        req.rawBody += part;
    });
    req.on('end', next);
};

// Handle proactive calls via POST
server.post('/api/proactiveCall', plainTextParser, function(req, res, next) {
    // Here these are hard coded, but they could be stored in Table storage,
    //  Allowing for dynamic addition and removal
    //  This information can be gathered in a few different ways:
    //   - From 'session.message.address' when you recieve a message (as shown below)
    //   - Via the BotBuilder Teams package (using fetchChannelList)
    var serviceUrl = "https://smba.trafficmanager.net/amer/";
    var teamId = "19:9199b3396378zzzzzzzzzzzzzzzzzzzz@thread.skype";

    var address = {
        conversation: {
            isGroup: true,
            conversationType: 'channel',
            id: teamId
        },
        serviceUrl: serviceUrl 
    };
    
    var msg = new builder.Message().address(address);
    
    msg.text('Hello, this is a proactive notification.\n\n' + req.rawBody);
    msg.textLocale('en-US');
    
    bot.send(msg);
    
    res.send(201);
	next();
});

// Handle message from user by returning metadata
bot.dialog('/', function (session) {
    session.sendTyping();
    session.send('Group Conversation: ' +
                    session.message.address.conversation.isGroup +
                    '\n\nConversation Type: ' +
                    session.message.address.conversation.conversationType +
                    '\n\nTeam ID: ' +
                    session.message.address.conversation.id +
                    '\n\nService Url: ' +
                    session.message.address.serviceUrl);
});