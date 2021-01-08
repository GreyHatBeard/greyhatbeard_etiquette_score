# GreyHatBeard Etiquette Score

If you have ever written a set of documents that outline how people should do things but then found that they don't follow them, you should keep reading. We found the same thing but rather than be frustrated, we decided that there must be a better way. If people could be advised at the point where or when they were doing the wrong thing, they could be engaged with a Teachable Moment to help them learn right there and then. This is what the Etiquette Score brings.

Much like the productivity score or compliance score from Microsoft, it helps to guide people to the right thing by highlighting areas of improvement. The etiquette score does this by using the Microsoft Graph events to look at things that happen for a user. These are then analysed for a set of etiquette best practices as outlined in the [GreyHatBeard Etiquette of M365 Guide](https://greyhatbeard.github.io/m365-etiquette) to give users a score. A bot delivered via Microsoft Teams will engage with the user to help them understand any breaches and whether they want to make changes to their behaviour.

## How it works

The etiquette score has been built as a set of Azure Functions in Node JS that connect to Microsoft Graph. Subscriptions are made to key events in the graph that are handled via the functions where a set of rules are defined. When these are triggered, a score is updated and a message sent via a Bot.

## What rules exist?

The rules are based on what can be triggered via Graph events.

- Emails
- [Events](docs/events.md)
- Contacts
- Lists
- Users
- Groups
- Teams Calls
- Teams Chat

## How does it run?

Configure app registration and store details in secrets.ts
Run "ngrok http 7071 -host-header="localhost:7071"
Add the resulting ngrok url to secrets.ts (e.g. https://c3f375d89967.ngrok.io)
Run "npm run dev" from the functions folder

