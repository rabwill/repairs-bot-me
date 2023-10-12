import { default as axios } from "axios";
import * as querystring from "querystring";
import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
} from "botbuilder";

export class SearchApp extends TeamsActivityHandler {
  constructor() {
    super();
  }

  // Search.
  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<MessagingExtensionResponse> {
    const searchQuery = query.parameters[0].value;
    const response = await axios.get(`https://piercerepairsapi.azurewebsites.net/repairs?assignedTo=${searchQuery.toLowerCase()}`);
    const attachments = [];
    response.data.forEach((obj) => {
      const adaptiveCard = CardFactory.adaptiveCard({
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [
            {
                type: "Container",
                items: [
                    {
                        type: "TextBlock",
                        text: "Repairs",
                        size: "Large",
                        weight: "Bolder",
                        color: "Accent"
                    }
                ]
            },
            {
                type: "Container",
                items: [
                    {
                        type: "TextBlock",
                        text: "Title:",
                        weight: "Bolder"
                    },
                    {
                        type: "TextBlock",
                        text: `${obj.title}`,
                        wrap: true
                    }
                ]
            },
            {
                type: "Container",
                items: [
                    {
                        type: "TextBlock",
                        text: "Description:",
                        weight: "Bolder"
                    },
                    {
                        type: "TextBlock",
                        text: `${obj.description}`,
                        wrap: true
                    }
                ]
            },
            {
                type: "Container",
                items: [
                    {
                        type: "TextBlock",
                        text: "Assigned To:",
                        weight: "Bolder"
                    },
                    {
                        type: "TextBlock",
                        text: `${obj.assignedTo}`,
                        wrap: true
                    }
                ]
            },
            {
                type: "Container",
                items: [
                    {
                        type: "TextBlock",
                        text: "Date:",
                        weight: "Bolder"
                    },
                    {
                        type: "TextBlock",
                        text: `${obj.date}`,
                        wrap: true
                    }
                ]
            },
            {
                type: "ImageSet",
                images: [
                    {
                        type: "Image",
                        url: `${obj.image}`,
                        size: "Medium",
                        altText: "Sample Image"
                    }
                ]
            }
        ],
        actions: [
            {
                type: "Action.OpenUrl",
                title: "View Details",
                url: "/",
                style: "positive"
            }
        ]
    });
      const preview = CardFactory.heroCard(obj.title);
      const attachment = { ...adaptiveCard, preview };
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
}
