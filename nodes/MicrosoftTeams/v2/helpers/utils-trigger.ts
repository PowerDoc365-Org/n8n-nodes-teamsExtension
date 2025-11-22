import type { IHookFunctions } from 'n8n-workflow';
import { NodeApiError } from 'n8n-workflow';

import type { CreateSubscriptionBody, SubscriptionResponse } from './types';
import { microsoftApiRequest } from '../transport';

/**
 * Creates a Graph API subscription for the webhook
 */
export async function createSubscription(
	this: IHookFunctions,
	notificationUrl: string,
	resource: string,
): Promise<SubscriptionResponse> {
	// Calculate expiration (1008 minutes = 7 days, maximum for most resources)
	const now = new Date();
	const expirationDateTime = new Date(now.getTime() + 1008 * 60 * 1000);

	const body: CreateSubscriptionBody = {
		changeType: 'created',
		notificationUrl,
		resource,
		expirationDateTime: expirationDateTime.toISOString(),
		clientState: 'n8n-webhook-secret',
	};

	console.log('body', body);

	// If expiration is more than 1 hour in the future, lifecycle notification URL is required
	const expirationDiff = expirationDateTime.getTime() - now.getTime();
	if (expirationDiff > 60 * 60 * 1000) {
		body.lifecycleNotificationUrl = notificationUrl;
	}

	try {
		const subscription = await microsoftApiRequest.call(this, 'POST', '/subscriptions', body);
		return subscription as SubscriptionResponse;
	} catch (error: any) {
		throw new NodeApiError(this.getNode(), {
			message: 'Failed to create subscription',
			description: error.message || 'Could not create Graph API subscription',
		});
	}
}

/**
 * Gets the resource path(s) for the Graph API subscription based on the event type
 */
export async function getResourcePath(
	this: IHookFunctions,
	event: string,
): Promise<string | string[]> {
	const watchAll = this.getNodeParameter('watchAllMeetings', false) as boolean;

	switch (event) {
		case 'newTranscript':
			if (watchAll) {
				// Watch all transcripts across the tenant
				return 'communications/onlineMeetings/getAllTranscripts';
			} else {
				// Watch transcripts for a specific meeting
				const meetingId = this.getNodeParameter('meetingId', '') as string;

				if (!meetingId) {
					throw new NodeApiError(this.getNode(), {
						message: 'Meeting ID is required',
						description: 'Please provide a meeting ID or enable "Watch All Meetings"',
					});
				}
				return `communications/onlineMeetings/${meetingId}/transcripts`;
			}

		case 'newUserTranscript':
			// Watch transcripts for meetings organized by a specific user
			const userId = this.getNodeParameter('userId.value', '') as string;
			if (!userId) {
				throw new NodeApiError(this.getNode(), {
					message: 'User ID is required',
					description: 'Please provide a user ID',
				});
			}
			return `users/${userId}/onlineMeetings/getAllTranscripts`;

		default:
			throw new NodeApiError(this.getNode(), {
				message: 'Invalid event type',
				description: `Event type "${event}" is not supported`,
			});
	}
}
