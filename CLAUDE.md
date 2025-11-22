# Context

The Graph API supports registering subscriptions so that a webhook is called when an event occurres in the microsoft ecosystem. n8n already has support for a few trigger nodes that make use of this system.
I want you to create more custom triggers for me to use with n8n.

## Development

Always use `bun` for scripts and package managing in this project. Do not use `npm` or other package managers.

## Files

### MicrosoftTeamsTrigger.node.ts

This file is copied directly from the n8n GitHub repo. It contains the trigger definitions for the microsoft teams integration. They already use the subscription system of Graph and are a great starting point.

### graph-subscription-docs/*.md

This folder contains documentation from the Microsoft Graph Docs about the subscription endpoints. Use this as context for the implementation in MicrosoftTeamsTrigger.node.ts.