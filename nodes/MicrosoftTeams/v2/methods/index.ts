import type { ILoadOptionsFunctions, INodeListSearchResult } from 'n8n-workflow';
import { microsoftApiRequestAllItems } from '../transport';

async function getUsers(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const users = (await microsoftApiRequestAllItems.call(
		this,
		'value',
		'GET',
		'/users',
		{},
		filter ? { $filter: `startsWith(displayName,'${filter}') or startsWith(userPrincipalName,'${filter}')` } : {},
	)) as Array<{ id: string; displayName: string; userPrincipalName: string }>;

	return {
		results: users.map((user) => ({
			name: user.displayName || user.userPrincipalName,
			value: user.id,
			url: `https://admin.microsoft.com/#/users/:/UserDetails/${user.id}`,
		})),
	};
}

async function getMeetings(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	// Get userId from node parameters (required for application credentials)
	const userId = this.getNodeParameter('userId.value', 0) as string;

	if (!userId) {
		return { results: [] };
	}

	const meetings = (await microsoftApiRequestAllItems.call(
		this,
		'value',
		'GET',
		`/users/${userId}/onlineMeetings`,
		{},
		filter ? { $filter: `contains(subject,'${filter}')` } : {},
	)) as Array<{ id: string; subject: string; startDateTime: string }>;

	return {
		results: meetings.map((meeting) => ({
			name: `${meeting.subject} (${new Date(meeting.startDateTime).toLocaleString()})`,
			value: meeting.id,
		})),
	};
}

export const listSearch = {
	getUsers,
	getMeetings,
};
