# n8n-nodes-teams-events

Custom n8n nodes for Microsoft Teams event triggers, including transcription notifications.

## Features

- **Microsoft Teams Transcription Trigger**: Receive notifications when new meeting transcriptions become available
  - Watch all meetings across the tenant
  - Watch meetings organized by a specific user
  - Watch a specific meeting

- **Microsoft Teams Transcript**: Retrieve and manage meeting transcriptions
  - List all transcripts for a specific meeting
  - List all transcripts for a user's meetings
  - Get transcript metadata
  - Get transcript content (VTT or plain text format)

## Installation

1. Install the package in your n8n instance:

```bash
npm install n8n-nodes-teams-events
```

2. Restart your n8n instance to load the new nodes

## Setup

### 1. Register an Azure AD Application

1. Go to the [Azure Portal](https://portal.azure.com/)
2. Navigate to **Azure Active Directory** > **App registrations** > **New registration**
3. Enter a name for your application
4. Set the redirect URI to your n8n webhook URL: `https://your-n8n-instance.com/rest/oauth2-credential/callback`
5. Click **Register**

### 2. Configure API Permissions

For the Microsoft Teams Transcription Trigger, you need to add the following API permissions:

#### Delegated Permissions (for user-scoped access):
- `OnlineMeetings.Read` - Read online meetings
- `OnlineMeetings.ReadWrite` - Read and create online meetings
- `OnlineMeetingTranscript.Read.All` - Read all transcripts for meetings
- `User.Read.All` - Read all users' profiles

#### Application Permissions (for tenant-wide access):
- `OnlineMeetingTranscript.Read.All` - Read all transcripts across the tenant
- `User.Read.All` - Read all users' profiles

**Important Notes:**
- Application permissions require admin consent
- For watching all meetings (`communications/onlineMeetings/getAllTranscripts`), you **must** use application permissions
- For watching specific user meetings, you can use either delegated or application permissions

### 3. Create Client Secret

1. In your Azure AD app registration, go to **Certificates & secrets**
2. Click **New client secret**
3. Add a description and set an expiration period
4. Copy the **Value** (not the Secret ID) - you'll need this for n8n

### 4. Configure n8n Credentials

1. In n8n, create new credentials for **Microsoft Teams OAuth2 API**
2. Enter the following information:
   - **Client ID**: Your Azure AD application's Application (client) ID
   - **Client Secret**: The secret value you copied
   - **Tenant ID**: Your Azure AD tenant ID (optional, defaults to 'common')
3. Click **Connect my account** and complete the OAuth flow

## Usage

### Microsoft Teams Transcription Trigger

This trigger node fires when a new transcription becomes available for Microsoft Teams meetings.

#### Trigger Options:

**1. New Transcript (All Meetings)**
- Watches for transcripts across all meetings in the tenant
- Requires `OnlineMeetingTranscript.Read.All` application permission
- Option to watch all meetings or a specific meeting

**2. New Transcript (User's Meetings)**
- Watches for transcripts from meetings organized by a specific user
- Can use delegated permissions
- Select user by list, ID, or email

#### Configuration:

1. Add the **Microsoft Teams Transcription Trigger** node to your workflow
2. Select the trigger type:
   - **New Transcript (All Meetings)**: For tenant-wide notifications
   - **New Transcript (User's Meetings)**: For user-specific notifications
3. Configure the watch settings:
   - For all meetings: Toggle "Watch All Meetings" on/off
   - For user meetings: Select the user
4. Save and activate your workflow

#### Webhook Data:

When a transcription becomes available, the trigger receives:

```json
{
  "subscriptionId": "subscription-uuid",
  "changeType": "created",
  "resource": "communications/onlineMeetings/{meetingId}/transcripts/{transcriptId}",
  "resourceData": {
    "@odata.type": "#Microsoft.Graph.callTranscript",
    "id": "transcript-id",
    ...
  },
  "tenantId": "tenant-uuid",
  "subscriptionExpirationDateTime": "2025-12-21T13:00:00Z"
}
```

#### Important Notes:

1. **HTTPS Required**: Microsoft Graph webhooks require HTTPS. Your n8n instance must be accessible via HTTPS.

2. **Subscription Expiration**: Subscriptions are created with a 30-day expiration (maximum for most resources). The node includes a lifecycle notification URL to handle subscription renewals.

3. **Validation Token**: When the webhook is first created, Microsoft Graph sends a validation request. The node automatically handles this validation.

4. **Transcript Availability**: Transcripts typically become available a few minutes after the meeting ends. The exact timing depends on meeting length and processing time.

5. **Required Permissions Summary**:

   | Scope | Permission | Type | Description |
   |-------|-----------|------|-------------|
   | Tenant-wide (all meetings) | `OnlineMeetingTranscript.Read.All` | Application | Read all transcripts |
   | User's meetings | `OnlineMeetingTranscript.Read.All` | Delegated or Application | Read user's transcripts |
   | List users | `User.Read.All` | Delegated or Application | Required for user selection |
   | List meetings | `OnlineMeetings.Read` | Delegated | Required for meeting selection |

### Microsoft Teams Transcript

This action node allows you to retrieve transcriptions from Microsoft Teams meetings.

#### Operations:

**1. Get All Transcripts**
- Retrieves all transcripts for a specific meeting
- Requires user and meeting selection
- Supports filtering and limiting results

**2. Get All User Transcripts**
- Retrieves all transcripts from meetings organized by a specific user
- Only requires user selection
- Supports filtering and limiting results
- Useful for bulk transcript retrieval

**3. Get Transcript**
- Retrieves metadata for a specific transcript
- Returns transcript details including:
  - Transcript ID
  - Meeting ID
  - Created date/time
  - End date/time
  - Meeting organizer
  - Content correlation ID

**4. Get Transcript Content**
- Downloads the actual transcript content
- Supports two formats:
  - **VTT (WebVTT)**: Includes timestamps and speaker information
  - **Text**: Plain text format without timestamps
- Returns content as binary data for further processing

#### Configuration:

##### Get All Transcripts
1. Add the **Microsoft Teams Transcript** node to your workflow
2. Select **Get All Transcripts** operation
3. Select the user who organized the meeting
4. Select the meeting
5. (Optional) Configure options:
   - **Limit**: Maximum number of transcripts to return
   - **Filter**: OData filter query (e.g., `meetingOrganizer/user/id eq 'user-id'`)

##### Get All User Transcripts
1. Select **Get All User Transcripts** operation
2. Select the user
3. (Optional) Configure options:
   - **Limit**: Maximum number of transcripts to return
   - **Filter**: OData filter query

##### Get Transcript
1. Select **Get Transcript** operation
2. Select the user who organized the meeting
3. Select the meeting
4. Enter the transcript ID

##### Get Transcript Content
1. Select **Get Transcript Content** operation
2. Select the user who organized the meeting
3. Select the meeting
4. Enter the transcript ID
5. Select the format (VTT or Text)
6. Specify the field name for the content (default: `data`)

#### Example Output:

**Get All Transcripts / Get Transcript:**
```json
{
  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('user-id')/onlineMeetings('meeting-id')/transcripts",
  "id": "MSMjMCMjZDAwYWU3NjUtNmM2Yi00NjQxLTgwMWQtMTkzMmFmMjEzNzdh",
  "meetingId": "MSo1N2Y5ZGFjYy03MWJmLTQ3NDMtYjQxMy01M2EdFGkdRWHJlQ",
  "callId": "af1a9d00-6c6b-4641-801d-1932af21377a",
  "createdDateTime": "2025-01-15T14:30:45.6463915Z",
  "endDateTime": "2025-01-15T15:00:45.6463915Z",
  "contentCorrelationId": "bc842d7a-2f6e-4b18-a141-94c1e4da1b7a",
  "transcriptContentUrl": "https://graph.microsoft.com/v1.0/users/user-id/onlineMeetings/meeting-id/transcripts/transcript-id/content",
  "meetingOrganizer": {
    "application": null,
    "device": null,
    "user": {
      "id": "user-id",
      "displayName": "John Doe",
      "tenantId": "tenant-id"
    }
  }
}
```

**Get Transcript Content:**
- Returns binary data that can be:
  - Saved to file using the "Write Binary File" node
  - Processed with text manipulation nodes
  - Sent via email or other communication channels

VTT format example:
```
WEBVTT

00:00:00.000 --> 00:00:03.120
<v John Doe>Welcome everyone to today's meeting.</v>

00:00:03.120 --> 00:00:06.450
<v Jane Smith>Thank you for having me.</v>
```

#### Use Cases:

1. **Automated Meeting Summary**
   - Trigger on new transcript → Get transcript content → Send to AI for summarization

2. **Transcript Archive**
   - Get all user transcripts → Download content → Store in cloud storage

3. **Compliance and Record-keeping**
   - Scheduled workflow to retrieve and archive all meeting transcripts

4. **Meeting Analytics**
   - Retrieve transcripts → Analyze content for keywords, sentiment, action items

#### Important Notes:

1. **Permissions Required**:
   - `OnlineMeetingTranscript.Read.All` (Application or Delegated)
   - `OnlineMeetings.Read` (Delegated) - for meeting selection
   - `User.Read.All` (Delegated or Application) - for user selection

2. **Transcript Availability**:
   - Transcripts are only available if recording/transcription was enabled during the meeting
   - It may take a few minutes after the meeting ends for transcripts to become available
   - Not all meetings will have transcripts

3. **Format Support**:
   - VTT format includes speaker identification and timestamps
   - Text format is plain text without metadata
   - The older DOCX format was deprecated as of May 31, 2023

4. **Rate Limits**:
   - Microsoft Graph API has rate limits that may affect bulk operations
   - Consider using the limit option for large datasets

## Graph API Subscription Details

### Resource Paths

The trigger uses the following Microsoft Graph API subscription resource paths:

- **Tenant-level (all meetings)**: `communications/onlineMeetings/getAllTranscripts`
- **User-level**: `users/{userId}/onlineMeetings/getAllTranscripts`
- **Specific meeting**: `communications/onlineMeetings/{meetingId}/transcripts`

### Change Types

The trigger subscribes to the `created` change type, which fires when a new transcription becomes available.

### Subscription Lifecycle

- **Initial Expiration**: 30 days (43,200 minutes)
- **Minimum Expiration**: 45 minutes (enforced by Microsoft Graph)
- **Lifecycle Notifications**: Automatically configured for subscriptions > 1 hour

## Troubleshooting

### "Invalid Notification URL" Error
- Ensure your n8n instance is accessible via HTTPS
- Verify the webhook URL is publicly accessible
- Check that your firewall allows incoming HTTPS connections

### "Insufficient Permissions" Error
- Verify you've added the required API permissions in Azure AD
- Ensure admin consent has been granted for application permissions
- Reconnect your credentials in n8n

### No Notifications Received
- Check that transcription is enabled in your Teams meetings
- Verify the subscription was created successfully (check workflow execution logs)
- Ensure the webhook URL is correct and accessible
- Note: Transcripts may take a few minutes to process after a meeting ends

### Subscription Already Exists
- The node automatically checks for existing subscriptions
- If a valid subscription exists, it will reuse it
- You can manually delete subscriptions via the Microsoft Graph API if needed

## Development

### Build

```bash
npm install
npm run build
```

### Lint

```bash
npm run lint
npm run lintfix
```

## License

MIT

## Author

Created for use with n8n - Workflow Automation Tool

## Resources

- [Microsoft Graph API Documentation](https://learn.microsoft.com/en-us/graph/)
- [Change Notifications for Transcripts](https://learn.microsoft.com/en-us/graph/teams-changenotifications-callrecording-and-calltranscript)
- [n8n Documentation](https://docs.n8n.io/)
- [n8n Community Nodes](https://docs.n8n.io/integrations/community-nodes/)
