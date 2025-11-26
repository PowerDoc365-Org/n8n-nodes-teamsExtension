import type {
	IDataObject,
	IExecuteFunctions,
	IHookFunctions,
	ILoadOptionsFunctions,
	IWebhookFunctions,
	IHttpRequestMethods,
	IRequestOptions,
	JsonObject,
} from 'n8n-workflow';
import { NodeApiError } from 'n8n-workflow';

export async function microsoftApiRequest(
	this: IHookFunctions | IExecuteFunctions | ILoadOptionsFunctions | IWebhookFunctions,
	method: IHttpRequestMethods,
	resource: string,
	body: IDataObject = {},
	qs: IDataObject = {},
	uri?: string,
	headers: IDataObject = {},
): Promise<any> {
	const options: IRequestOptions = {
		headers: {
			'Content-Type': 'application/json',
			...headers,
		},
		method,
		body,
		qs,
		uri: uri || `https://graph.microsoft.com/v1.0${resource}`,
		json: true,
	};

	try {
		if (Object.keys(body).length === 0) {
			delete options.body;
		}

		return await this.helpers.requestWithAuthentication.call(
			this,
			'microsoftTeamsOAuth2AppApi',
			options,
		);
	} catch (error) {
		console.log('error', error);
		throw new NodeApiError(this.getNode(), error as JsonObject);
	}
}

export async function microsoftApiRequestAllItems(
	this: IHookFunctions | IExecuteFunctions | ILoadOptionsFunctions | IWebhookFunctions,
	propertyName: string,
	method: IHttpRequestMethods,
	endpoint: string,
	body: IDataObject = {},
	query: IDataObject = {},
): Promise<any> {
	const returnData: IDataObject[] = [];

	let responseData;
	let uri: string | undefined;
	query['$top'] = 100;

	do {
		responseData = await microsoftApiRequest.call(this, method, endpoint, body, query, uri);
		uri = responseData['@odata.nextLink'];
		if (uri?.includes('$top')) {
			delete query['$top'];
		}
		returnData.push.apply(returnData, responseData[propertyName] as IDataObject[]);
	} while (responseData['@odata.nextLink'] !== undefined);

	return returnData;
}
