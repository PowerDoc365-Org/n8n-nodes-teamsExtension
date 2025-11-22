import type { IDataObject } from 'n8n-workflow';

export interface WebhookNotification {
	subscriptionId: string;
	clientState?: string;
	changeType: 'created' | 'updated' | 'deleted';
	resource: string;
	subscriptionExpirationDateTime: string;
	resourceData?: IDataObject;
	tenantId?: string;
}

export interface SubscriptionResponse {
	id: string;
	resource: string;
	applicationId: string;
	changeType: string;
	clientState?: string;
	notificationUrl: string;
	lifecycleNotificationUrl?: string;
	expirationDateTime: string;
	creatorId: string;
	includeResourceData?: boolean;
	latestSupportedTlsVersion?: string;
	encryptionCertificate?: string;
	encryptionCertificateId?: string;
	notificationQueryOptions?: string;
	notificationUrlAppId?: string;
}

export interface CreateSubscriptionBody extends IDataObject {
	changeType: string;
	notificationUrl: string;
	resource: string;
	expirationDateTime: string;
	clientState?: string;
	lifecycleNotificationUrl?: string;
	includeResourceData?: boolean;
	encryptionCertificate?: string;
	encryptionCertificateId?: string;
}
