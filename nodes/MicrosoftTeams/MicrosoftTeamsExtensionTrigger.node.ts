import type {
	IExecuteFunctions,
	INodeType,
	INodeTypeDescription,
	IHookFunctions,
	IWebhookFunctions,
	IWebhookResponseData,
	IDataObject,
	ILoadOptionsFunctions,
	JsonObject,
	INodeExecutionData,
} from 'n8n-workflow';
import { NodeApiError, NodeConnectionTypes } from 'n8n-workflow';

import type { WebhookNotification, SubscriptionResponse } from './v2/helpers/types';
import {
	createSubscription,
	getResourcePath,
	isLifecycleNotification,
	renewSubscription,
} from './v2/helpers/utils-trigger';
import { listSearch } from './v2/methods';
import { microsoftApiRequest, microsoftApiRequestAllItems } from './v2/transport';

export class MicrosoftTeamsExtensionTrigger implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Microsoft Teams Extension Trigger',
		name: 'microsoftTeamsExtensionTrigger',
		icon: 'file:teams.svg',
		group: ['trigger'],
		version: 1,
		description: 'Triggers workflows off more resources from Microsoft Teams',
		subtitle: 'Teams Extension Trigger',
		defaults: {
			name: 'Microsoft Teams Extension Trigger',
		},
		credentials: [
			{
				name: 'microsoftTeamsOAuth2AppApi',
				required: true,
			},
		],
		inputs: [],
		outputs: [NodeConnectionTypes.Main],
		webhooks: [
			{
				name: 'default',
				httpMethod: 'POST',
				responseMode: 'onReceived',
				path: 'webhook',
			},
		],
		properties: [
			{
				displayName: 'Trigger On',
				name: 'event',
				type: 'options',
				default: 'newTranscript',
				options: [
					{
						name: 'New Transcript (All Meetings)',
						value: 'newTranscript',
						description:
							'Triggered when a transcript is available for any meeting in the tenant',
					},
					{
						name: "New Transcript (User's Meetings)",
						value: 'newUserTranscript',
						description:
							'Triggered when a transcript is available for meetings organized by a specific user',
					},
				],
				description: 'Select when to trigger the workflow',
			},
			{
				displayName:
					'Note: This trigger requires the Application credential (not delegated) with OnlineMeetingTranscript.Read.All permission',
				name: 'notice',
				type: 'notice',
				default: '',
				displayOptions: {
					show: {
						event: ['newTranscript'],
					},
				},
			},
			{
				displayName: 'Watch All Meetings',
				name: 'watchAllMeetings',
				type: 'boolean',
				default: true,
				description: 'Whether to watch for transcripts in all meetings across the tenant',
				displayOptions: {
					show: {
						event: ['newTranscript'],
					},
				},
			},
			{
				displayName: 'Meeting ID',
				name: 'meetingId',
				type: 'string',
				default: '',
				required: true,
				description: 'The ID of the specific meeting to watch',
				placeholder: 'e.g., MSo1N2Y5ZGFjYy03MWJmLTQ3NDMtYjQxMy01M2EdFGkdRWHJlQ',
				displayOptions: {
					show: {
						event: ['newTranscript'],
						watchAllMeetings: [false],
					},
				},
			},
			{
				displayName: 'User',
				name: 'userId',
				type: 'resourceLocator',
				default: {
					mode: 'list',
					value: '',
				},
				required: true,
				description: 'Select a user from the list or enter an ID',
				modes: [
					{
						displayName: 'From List',
						name: 'list',
						type: 'list',
						placeholder: 'Select a user...',
						typeOptions: {
							searchListMethod: 'getUsers',
							searchable: true,
						},
					},
					{
						displayName: 'By ID',
						name: 'id',
						type: 'string',
						placeholder: 'e.g., 48d31887-5fad-4d73-a9f5-3c356e68a038',
					},
					{
						displayName: 'By Email',
						name: 'email',
						type: 'string',
						placeholder: 'e.g., user@company.com',
						validation: [
							{
								type: 'regex',
								properties: {
									regex: '^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\\.[a-zA-Z0-9-.]+$',
									errorMessage: 'Not a valid email',
								},
							},
						],
					},
				],
				displayOptions: {
					show: {
						event: ['newUserTranscript'],
					},
				},
			},
		],
	};

	methods = {
		listSearch,
	};

	webhookMethods = {
		default: {
			async checkExists(this: IHookFunctions): Promise<boolean> {
				const event = this.getNodeParameter('event', 0) as string;
				const webhookUrl = this.getNodeWebhookUrl('default');
				const webhookData = this.getWorkflowStaticData('node');

				try {
					const subscriptions = (await microsoftApiRequestAllItems.call(
						this as unknown as ILoadOptionsFunctions,
						'value',
						'GET',
						'/subscriptions',
					)) as SubscriptionResponse[];

					const matchingSubscriptions = subscriptions.filter(
						(subscription) => subscription.notificationUrl === webhookUrl,
					);

					const now = new Date();
					const thresholdMs = 5 * 60 * 1000; // 5 minutes
					const validSubscriptions = matchingSubscriptions.filter((subscription) => {
						const expiration = new Date(subscription.expirationDateTime);
						return expiration.getTime() - now.getTime() > thresholdMs;
					});

					const resourcePaths = await getResourcePath.call(this, event);
					const requiredResources = Array.isArray(resourcePaths)
						? resourcePaths
						: [resourcePaths];

					const subscribedResources = validSubscriptions.map((sub) => sub.resource);
					const allResourcesSubscribed = requiredResources.every((resource) =>
						subscribedResources.includes(resource),
					);

					if (allResourcesSubscribed) {
						webhookData.subscriptionIds = validSubscriptions.map((sub) => sub.id);
						return true;
					}

					return false;
				} catch (error) {
					return false;
				}
			},

			async create(this: IHookFunctions): Promise<boolean> {
				const event = this.getNodeParameter('event', 0) as string;
				const webhookUrl = this.getNodeWebhookUrl('default');
				const webhookData = this.getWorkflowStaticData('node');

				if (!webhookUrl?.startsWith('https://')) {
					throw new NodeApiError(this.getNode(), {
						message: 'Invalid Notification URL',
						description: `The webhook URL "${webhookUrl}" is invalid. Microsoft Graph requires an HTTPS URL.`,
					});
				}

				const resourcePaths = await getResourcePath.call(this, event);
				const subscriptionIds: string[] = [];

				if (Array.isArray(resourcePaths)) {
					await Promise.all(
						resourcePaths.map(async (resource) => {
							const subscription = await createSubscription.call(
								this,
								webhookUrl,
								resource,
							);
							subscriptionIds.push(subscription.id);
							return subscription;
						}),
					);

					webhookData.subscriptionIds = subscriptionIds;
				} else {
					const subscription = await createSubscription.call(
						this,
						webhookUrl,
						resourcePaths,
					);
					webhookData.subscriptionIds = [subscription.id];
				}

				return true;
			},

			async delete(this: IHookFunctions): Promise<boolean> {
				const webhookData = this.getWorkflowStaticData('node');
				const storedIds = webhookData.subscriptionIds as string[] | undefined;

				if (!Array.isArray(storedIds)) {
					return false;
				}

				try {
					await Promise.all(
						storedIds.map(async (subscriptionId) => {
							try {
								await microsoftApiRequest.call(
									this as unknown as IExecuteFunctions,
									'DELETE',
									`/subscriptions/${subscriptionId}`,
								);
							} catch (error) {
								// Ignore 404 errors (subscription already deleted)
								if ((error as JsonObject).httpStatusCode !== 404) {
									throw error;
								}
							}
						}),
					);

					delete webhookData.subscriptionIds;
					return true;
				} catch (error) {
					return false;
				}
			},
		},
	};

	async webhook(this: IWebhookFunctions): Promise<IWebhookResponseData> {
		const req = this.getRequestObject();
		const res = this.getResponseObject();

		// Handle Microsoft Graph validation request
		if (req.query.validationToken) {
			res.status(200).send(req.query.validationToken);
			return { noWebhookResponse: true };
		}

		// Handle notifications
		const eventNotifications = req.body.value as WebhookNotification[];

		// Separate lifecycle notifications from change notifications
		const lifecycleNotifications: WebhookNotification[] = [];
		const changeNotifications: WebhookNotification[] = [];

		for (const notification of eventNotifications) {
			if (isLifecycleNotification(notification)) {
				lifecycleNotifications.push(notification);
			} else {
				changeNotifications.push(notification);
			}
		}

		// Handle lifecycle notifications (subscription expiration warnings)
		for (const lifecycleNotification of lifecycleNotifications) {
			console.log(
				`Received lifecycle notification for subscription ${lifecycleNotification.subscriptionId}`,
				lifecycleNotification.lifecycleEvent || 'expiration warning',
			);

			// Renew the subscription
			await renewSubscription.call(this, lifecycleNotification.subscriptionId);
		}

		// If there are no change notifications, just acknowledge receipt
		if (changeNotifications.length === 0) {
			return { noWebhookResponse: true };
		}

		// Return each change notification as a separate workflow execution
		const response: IWebhookResponseData = {
			workflowData: changeNotifications.map((event) => [
				{
					json: {
						subscriptionId: event.subscriptionId,
						changeType: event.changeType,
						resource: event.resource,
						resourceData: event.resourceData || {},
						tenantId: event.tenantId,
						subscriptionExpirationDateTime: event.subscriptionExpirationDateTime,
					} as IDataObject,
				} as INodeExecutionData,
			]),
		};

		return response;
	}
}
