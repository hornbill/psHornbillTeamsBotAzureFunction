# Azure Functions Hornbill Bot for Microsoft Teams Outbound Webhooks

## Introduction

Example PowerShell Azure Function to enable secure Hornbill integration from Microsoft Teams channels, using Teams Outbound Webhooks and Azure Functions HTTP Triggers.

Subscriptions to Azure Functions and Microsoft Teams are required.

## Reference

See the following documenation for more information on the target platforms:

- [Azure Functions HTTP Triggers](https://docs.microsoft.com/en-us/azure/azure-functions/functions-bindings-http-webhook)
- [Microsoft Teams Outbound Webhooks](https://docs.microsoft.com/en-us/microsoftteams/platform/webhooks-and-connectors/how-to/add-outgoing-webhook)

## Installation and Configuration

`run.ps1` should contain your Hornbill and Teams configuration, as so:

- `$InstanceName = "yourinstanceid"` : `yourinstanceid` should be replaced with the ID of your Hornbill instance;
- `$InstanceKey = "yourAPIkey"` : `yourAPIkey` should be replaced with the [Hornbill API key](https://wiki.hornbill.com/index.php/API_keys) generated for a Hornbill user that will be making the API calls from this integration;
- `$InstanceMicrosoftKey = 1` : `1` should be replaced with the primary key of a [Hornbill Keysafe Key](https://wiki.hornbill.com/index.php/Hornbill_KeySafe) of type `Microsoft`, which is used to authenticate requests to Azure;
- `$TeamsMapping` : This is a PowerShell object that is used to define which teams the bot is allowed to converse with, which teams in Hornbill the Microsoft Teams team is associated to, and the [security token for the Outbound Webhook](https://support.microsoft.com/en-gb/office/create-and-add-an-outgoing-webhook-in-teams-8e1a1648-982f-4511-b342-6d8492437207?ui=en-us&rs=en-gb&ad=gb). This PowerShell object should be in the following format:
  >        $TeamsMapping = @{
  >            "TEAMSTEAMID1" = @{
  >                "HornbillTeamID" = "HORNBILLTEAMID1"
  >                "WebhookToken" = "TEAMSWEBHOOKTOKEN1"
  >            },
  >            "TEAMSTEAMID2" = @{
  >                "HornbillTeamID" = "HORNBILLTEAMID2"
  >                "WebhookToken" = "TEAMSWEBHOOKTOKEN2"
  >            }
  >        }
  - Where:
    - `TEAMSTEAMID1` and `TEAMSTEAMID2` should be replaced by the ID's of the Microsoft Teams Teams that you wish to allow the bot to function within;
    - `HORNBILLTEAMID1` and `HORNBILLTEAMID2` should be replaced by the Hornbill Teams that you want to associate to the Microsoft Teams Teams, above;
    - `TEAMSWEBHOOKTOKEN1` and `TEAMSWEBHOOKTOKEN2` should be replaced by the security tokens that are generated when you create the [Microsoft Teams Outbound Webhooks](https://docs.microsoft.com/en-us/microsoftteams/platform/webhooks-and-connectors/how-to/add-outgoing-webhook).
  - So a real example may look like:
  >        $TeamsMapping = @{
  >            "19:d76da9c3c8ee44e38737ce1ae2cd4c06@thread.skype" = @{
  >                "HornbillTeamID" = "2ndLineSupport"
  >                "WebhookToken" = "6qwA0i7KP6LqZrvJaqoyj7S2F28yoy06H/axSmCbgQg="
  >            }
  >            "19:b12345a1c6b3495290d0adb0ef54c1b3@thread.tacv2" = @{
  >                "HornbillTeamID" = "3rdLineSupport"
  >                "WebhookToken" = "TFRzqQ3nTeJtsMX34CNb5t0Eo7uEUdjO2mJ+jd4CnXg="
  >            }
  >        }

Once your bot is configured, deployed to Azure Functions, and your Microsoft Teams Outgoing Webhooks are pointing to the newly deployed functions, then you are ready to talk to your bot from within Teams!

## Bot Commands

When tagging your Outgoing Webhooks in Teams channels, the following bot commands are made available to you:

### my requests

This command returns the newest 3 requests by date logged, that are owned by you in the team that you tagged the bot within.

Usage Example: `@Hornbill my requests`

- Alternatives:
  - `my new requests`
  - `my requests new`
  
### my old requests

This command returns the oldest 3 requests by date logged, that are owned by you in the team that you tagged the bot within.

Usage Example: `@Hornbill my old requests`

- Alternatives:
  - `my requests old`
  - `old my requests`

### team requests

This command returns the newest 3 requests by date logged, that are not owned by you but are owned by the team that you tagged the bot within.

Usage Example: `@Hornbill team requests`

- Alternatives:
  - `team new requests`
  - `team requests new`
  
### team old requests

This command returns the oldest 3 requests by date logged, that are not owned by you but are owned by the team that you tagged the bot within.

Usage Example: `@Hornbill team old requests`

- Alternatives:
  - `team requests old`
  - `old team requests`

### get request

This command returns the details of a specific request. The request must be owned by the team that you tagged the bot within.

Usage Example: `@Hornbill get request INC00012345`

This would return the details for request reference `INC0001234`.

### take request

This command will re-assign the stated request to you. The request must already be owned by the team that you tagged the bot within.

Usage Example: `@Hornbill take request INC00012345`

This would reassign request with reference `INC0001234` to you.

### update request

This command will perform a timeline update against the stated request. The request must already be owned by the team that you tagged the bot within.

Usage Example: `@Hornbill update request INC00012345 this is my update to be added to the timeline`

This would post an update to the timeline of the request with reference `INC0001234`, and the timeline update would be `this is my update to be added to the timeline`
