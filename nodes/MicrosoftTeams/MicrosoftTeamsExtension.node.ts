import type {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
	IDataObject,
	ILoadOptionsFunctions,
	JsonObject,
} from 'n8n-workflow';
import { NodeConnectionTypes, NodeOperationError } from 'n8n-workflow';

import { microsoftApiRequest, microsoftApiRequestAllItems } from './v2/transport';
import { listSearch } from './v2/methods';

export class MicrosoftTeamsTranscript implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Microsoft Teams Extension',
		name: 'microsoftTeamsExtension',
		icon: 'file:teams.svg',
		group: ['transform'],
		version: 1,
		subtitle: '={{$parameter["operation"]}}',
		description: 'Use more resources from Microsoft Teams',
		defaults: {
			name: 'Microsoft Teams Extension',
		},
		inputs: [NodeConnectionTypes.Main],
		outputs: [NodeConnectionTypes.Main],
		credentials: [
			{
				name: 'microsoftTeamsOAuth2AppApi',
				required: true,
			},
		],
		properties: [
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				default: 'getAll',
				options: [
					{
						name: 'Get Meeting',
						value: 'getMeeting',
						description: 'Get details of a specific online meeting',
						action: 'Get a meeting',
					},
					{
						name: 'Get All Transcripts (User)',
						value: 'getAllGlobal',
						description: 'Get all transcripts from meetings organized by a user',
						action: 'Get all user transcripts',
					},
					{
						name: 'Get All Transcripts for Meeting',
						value: 'getAll',
						description: 'Get all transcripts for a specific meeting',
						action: 'Get all transcripts for a meeting',
					},
					{
						name: 'Get All User Transcripts',
						value: 'getAllUser',
						description: 'Get all transcripts from meetings organized by a user',
						action: 'Get all transcripts for a user',
					},
					{
						name: 'Get Transcript',
						value: 'get',
						description: 'Get a specific transcript metadata',
						action: 'Get a transcript',
					},
					{
						name: 'Get Transcript Content',
						value: 'getContent',
						description: 'Get the content of a specific transcript',
						action: 'Get transcript content',
					},
				],
			},

			// User field (for operations that need it)
			{
				displayName: 'User',
				name: 'userId',
				type: 'resourceLocator',
				default: {
					mode: 'list',
					value: '',
				},
				required: true,
				description: 'The user who organized the meeting',
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
			},

			// Meeting ID field (for getMeeting operation)
			{
				displayName: 'Meeting ID',
				name: 'meetingId',
				type: 'string',
				default: '',
				required: true,
				description: 'The ID of the meeting to retrieve',
				placeholder: 'e.g., MSo1N2Y5ZGFjYy03MWJmLTQ3NDMtYjQxMy01M2EdFGkdRWHJlQ',
				displayOptions: {
					show: {
						operation: ['getMeeting'],
					},
				},
			},

			// Meeting field
			{
				displayName: 'Meeting',
				name: 'meetingId',
				type: 'resourceLocator',
				default: {
					mode: 'list',
					value: '',
				},
				required: true,
				description: 'The meeting to get transcripts from',
				modes: [
					{
						displayName: 'From List',
						name: 'list',
						type: 'list',
						placeholder: 'Select a meeting...',
						typeOptions: {
							searchListMethod: 'getMeetings',
							searchable: true,
						},
					},
					{
						displayName: 'By ID',
						name: 'id',
						type: 'string',
						placeholder: 'e.g., MSo1N2Y5ZGFjYy03MWJmLTQ3NDMtYjQxMy01M2EdFGkdRWHJlQ',
					},
				],
				displayOptions: {
					show: {
						operation: ['getAll', 'get', 'getContent'],
					},
				},
			},

			// Transcript ID field
			{
				displayName: 'Transcript ID',
				name: 'transcriptId',
				type: 'string',
				default: '',
				required: true,
				description: 'The ID of the transcript',
				placeholder: 'e.g., MSMjMCMjZDAwYWU3NjUtNmM2Yi00NjQxLTgwMWQtMTkzMmFmMjEzNzdh',
				displayOptions: {
					show: {
						operation: ['get', 'getContent'],
					},
				},
			},

			// Additional options for getAll operations
			{
				displayName: 'Options',
				name: 'options',
				type: 'collection',
				placeholder: 'Add Option',
				default: {},
				displayOptions: {
					show: {
						operation: ['getAllGlobal', 'getAll', 'getAllUser'],
					},
				},
				options: [
					{
						displayName: 'Limit',
						name: 'limit',
						type: 'number',
						typeOptions: {
							minValue: 1,
						},
						default: 50,
						description: 'Max number of results to return',
					},
					{
						displayName: 'Filter',
						name: 'filter',
						type: 'string',
						default: '',
						description: 'OData filter query to filter transcripts',
						placeholder: 'e.g., meetingOrganizer/user/id eq \'user-id\'',
					},
				],
			},

			// Content format option
			{
				displayName: 'Format',
				name: 'format',
				type: 'options',
				default: 'vtt',
				options: [
					{
						name: 'VTT (WebVTT)',
						value: 'vtt',
						description: 'WebVTT format with timestamps',
					},
					{
						name: 'Text',
						value: 'text',
						description: 'Plain text format',
					},
				],
				description: 'The format of the transcript content',
				displayOptions: {
					show: {
						operation: ['getContent'],
					},
				},
			},

			// Return content as field option
			{
				displayName: 'Put Content in Field',
				name: 'binaryProperty',
				type: 'string',
				default: 'data',
				description: 'Name of the field to put the transcript content in',
				displayOptions: {
					show: {
						operation: ['getContent'],
					},
				},
			},
		],
	};

	methods = {
		listSearch,
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];
		const operation = this.getNodeParameter('operation', 0);

		for (let i = 0; i < items.length; i++) {
			try {
				if (operation === 'getMeeting') {
					// Get meeting details by ID
					const userId = this.getNodeParameter('userId.value', i) as string;
					const meetingId = this.getNodeParameter('meetingId', i) as string;

					const endpoint = `/users/${userId}/onlineMeetings/${meetingId}`;
					const meeting = await microsoftApiRequest.call(this, 'GET', endpoint);

					returnData.push({
						json: meeting as IDataObject,
						pairedItem: { item: i },
					});
				} else if (operation === 'getAllGlobal') {
					// Get all transcripts across the tenant (using user-based global endpoint)
					const userId = this.getNodeParameter('userId.value', i) as string;
					const options = this.getNodeParameter('options', i, {}) as IDataObject;

					const qs: IDataObject = {};
					if (options.filter) {
						qs.$filter = options.filter as string;
					}
					if (options.limit) {
						qs.$top = options.limit as number;
					}

					const endpoint = `/users/${userId}/onlineMeetings/getAllTranscripts`;

					let transcripts;
					if (options.limit) {
						transcripts = await microsoftApiRequest.call(this, 'GET', endpoint, {}, qs);
						transcripts = transcripts.value;
					} else {
						transcripts = await microsoftApiRequestAllItems.call(
							this,
							'value',
							'GET',
							endpoint,
							{},
							qs,
						);
					}

					for (const transcript of transcripts) {
						returnData.push({
							json: transcript as IDataObject,
							pairedItem: { item: i },
						});
					}
				} else if (operation === 'getAll') {
					// Get all transcripts for a specific meeting
					const userId = this.getNodeParameter('userId.value', i) as string;
					const meetingId = this.getNodeParameter('meetingId.value', i) as string;
					const options = this.getNodeParameter('options', i, {}) as IDataObject;

					const qs: IDataObject = {};
					if (options.filter) {
						qs.$filter = options.filter as string;
					}
					if (options.limit) {
						qs.$top = options.limit as number;
					}

					const endpoint = `/users/${userId}/onlineMeetings/${meetingId}/transcripts`;

					let transcripts;
					if (options.limit) {
						transcripts = await microsoftApiRequest.call(this, 'GET', endpoint, {}, qs);
						transcripts = transcripts.value;
					} else {
						transcripts = await microsoftApiRequestAllItems.call(
							this,
							'value',
							'GET',
							endpoint,
							{},
							qs,
						);
					}

					for (const transcript of transcripts) {
						returnData.push({
							json: transcript as IDataObject,
							pairedItem: { item: i },
						});
					}
				} else if (operation === 'getAllUser') {
					// Get all transcripts from meetings organized by a user
					const userId = this.getNodeParameter('userId.value', i) as string;
					const options = this.getNodeParameter('options', i, {}) as IDataObject;

					const qs: IDataObject = {};
					if (options.filter) {
						qs.$filter = options.filter as string;
					}
					if (options.limit) {
						qs.$top = options.limit as number;
					}

					const endpoint = `/users/${userId}/onlineMeetings/getAllTranscripts`;

					let transcripts;
					if (options.limit) {
						transcripts = await microsoftApiRequest.call(this, 'GET', endpoint, {}, qs);
						transcripts = transcripts.value;
					} else {
						transcripts = await microsoftApiRequestAllItems.call(
							this,
							'value',
							'GET',
							endpoint,
							{},
							qs,
						);
					}

					for (const transcript of transcripts) {
						returnData.push({
							json: transcript as IDataObject,
							pairedItem: { item: i },
						});
					}
				} else if (operation === 'get') {
					// Get a specific transcript metadata
					const userId = this.getNodeParameter('userId.value', i) as string;
					const meetingId = this.getNodeParameter('meetingId.value', i) as string;
					const transcriptId = this.getNodeParameter('transcriptId', i) as string;

					const endpoint = `/users/${userId}/onlineMeetings/${meetingId}/transcripts/${transcriptId}`;
					const transcript = await microsoftApiRequest.call(this, 'GET', endpoint);

					returnData.push({
						json: transcript as IDataObject,
						pairedItem: { item: i },
					});
				} else if (operation === 'getContent') {
					// Get transcript content
					const userId = this.getNodeParameter('userId.value', i) as string;
					const meetingId = this.getNodeParameter('meetingId.value', i) as string;
					const transcriptId = this.getNodeParameter('transcriptId', i) as string;
					const format = this.getNodeParameter('format', i) as string;
					const binaryProperty = this.getNodeParameter('binaryProperty', i) as string;

					const endpoint = `/users/${userId}/onlineMeetings/${meetingId}/transcripts/${transcriptId}/content`;

					// Request with format parameter
					const qs: IDataObject = {};
					if (format === 'text') {
						qs.$format = 'text/plain';
					} else {
						qs.$format = 'text/vtt';
					}

					const content = await microsoftApiRequest.call(
						this,
						'GET',
						endpoint,
						{},
						qs,
						undefined,
						{ Accept: format === 'text' ? 'text/plain' : 'text/vtt' },
					);

					// Return content as a field
					returnData.push({
						json: {
							transcriptId,
							meetingId,
							format,
						},
						binary: {
							[binaryProperty]: {
								data: Buffer.from(content).toString('base64'),
								mimeType: format === 'text' ? 'text/plain' : 'text/vtt',
								fileName: `transcript_${transcriptId}.${format === 'text' ? 'txt' : 'vtt'}`,
							},
						},
						pairedItem: { item: i },
					});
				}
			} catch (error) {
				if (this.continueOnFail()) {
					returnData.push({
						json: {
							error: (error as Error).message,
						},
						pairedItem: { item: i },
					});
					continue;
				}
				throw new NodeOperationError(this.getNode(), error as JsonObject, { itemIndex: i });
			}
		}

		return [returnData];
	}
}
