import {
	IDataObject,
	IExecuteFunctions,
	ILoadOptionsFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
	NodeOperationError,
} from 'n8n-workflow';

import {
	cleanETag,
	createAssignmentsObject,
	formatDateTime,
	getUserIdByEmail,
	microsoftApiRequest,
	microsoftApiRequestAllItems,
	parseAssignments,
} from './GenericFunctions';
import { taskFields, taskOperations } from './TaskDescription';
import { planFields, planOperations } from './PlanDescription';
import { bucketFields, bucketOperations } from './BucketDescription';

export class MicrosoftPlanner implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Microsoft Planner',
		name: 'microsoftPlanner',
		icon: 'file:planner.svg',
		group: ['transform'],
		version: 1,
		subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
		description: 'Create and retrieve tasks in Microsoft Planner',
		defaults: {
			name: 'Microsoft Planner',
		},
		inputs: ['main'],
		outputs: ['main'],
		credentials: [
			{
				name: 'microsoftPlannerOAuth2Api',
				required: true,
			},
		],
		properties: [
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'Task',
						value: 'task',
					},
					{
						name: 'Plan',
						value: 'plan',
					},
					{
						name: 'Bucket',
						value: 'bucket',
					},
				],
				default: 'task',
			},
			// Grouped by resource so they appear nicely in the n8n UI
			...taskOperations,
			...taskFields,
			...planOperations,
			...planFields,
			...bucketOperations,
			...bucketFields,
		],
	};

	methods = {
		listSearch: {
			async getBuckets(this: ILoadOptionsFunctions) {
				try {
					const planId = this.getNodeParameter('planId', 0) as string;
					if (!planId) {
						return { results: [] };
					}

					const buckets = await microsoftApiRequestAllItems.call(
						this,
						'value',
						'GET',
						`/planner/plans/${planId}/buckets`,
					);

					if (!buckets || buckets.length === 0) {
						return { results: [] };
					}

					return {
						results: buckets.map((bucket: any) => ({
							name: bucket.name || bucket.id,
							value: bucket.id,
						})),
					};
				} catch (error) {
					console.error('Error loading buckets:', error);
					return { results: [] };
				}
			},

			async getTasks(this: ILoadOptionsFunctions) {
				try {
					const planId = this.getNodeParameter('planId', 0) as string;

					// Try to get bucketId - might be undefined or an object
					let bucketIdValue = '';
					try {
						const bucketId = this.getNodeParameter('bucketId', 0);
						if (typeof bucketId === 'string') {
							bucketIdValue = bucketId;
						} else if (bucketId && typeof bucketId === 'object' && 'value' in bucketId) {
							bucketIdValue = (bucketId as any).value;
						}
					} catch (error) {
						// bucketId might not exist yet, that's ok
					}

					let endpoint = '';
					if (bucketIdValue) {
						endpoint = `/planner/buckets/${bucketIdValue}/tasks`;
					} else if (planId) {
						endpoint = `/planner/plans/${planId}/tasks`;
					} else {
						return { results: [] };
					}

					const tasks = await microsoftApiRequestAllItems.call(
						this,
						'value',
						'GET',
						endpoint,
					);

					if (!tasks || tasks.length === 0) {
						return { results: [] };
					}

					return {
						results: tasks.map((task: any) => ({
							name: task.title || task.id,
							value: task.id,
						})),
					};
				} catch (error) {
					console.error('Error loading tasks:', error);
					return { results: [] };
				}
			},
		},
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: IDataObject[] = [];
		const resource = this.getNodeParameter('resource', 0) as string;
		const operation = this.getNodeParameter('operation', 0) as string;

		for (let i = 0; i < items.length; i++) {
			try {
				if (resource === 'task') {
					// ----------------------------------
					//         task:create
					// ----------------------------------
					if (operation === 'create') {
						const planId = this.getNodeParameter('planId', i) as string;
						const bucketIdParam = this.getNodeParameter('bucketId', i);
						const bucketId = typeof bucketIdParam === 'string'
							? bucketIdParam
							: (bucketIdParam as IDataObject).value as string;
						const title = this.getNodeParameter('title', i) as string;
						const additionalFields = this.getNodeParameter('additionalFields', i) as IDataObject;

						const body: IDataObject = {
							planId,
							bucketId,
							title,
						};

						if (additionalFields.priority !== undefined) {
							body.priority = additionalFields.priority;
						}

						const formattedDueDateTime = formatDateTime(additionalFields.dueDateTime as string);
						if (formattedDueDateTime) {
							body.dueDateTime = formattedDueDateTime;
						}

						const formattedStartDateTime = formatDateTime(additionalFields.startDateTime as string);
						if (formattedStartDateTime) {
							body.startDateTime = formattedStartDateTime;
						}

						if (additionalFields.percentComplete !== undefined) {
							body.percentComplete = additionalFields.percentComplete;
						}

						// Handle assignments
						if (additionalFields.assignments) {
							const emails = parseAssignments(additionalFields.assignments as string);
							const userIds: string[] = [];

							for (const email of emails) {
								const userId = await getUserIdByEmail.call(this, email);
								if (userId) {
									userIds.push(userId);
								} else {
									console.warn(`Could not find user ID for email: ${email}`);
								}
							}

							if (userIds.length > 0) {
								body.assignments = createAssignmentsObject(userIds);
							} else if (emails.length > 0) {
								console.warn('No valid user IDs found for assignment. Check if User.Read.All permission is granted.');
							}
						}

						const responseData = await microsoftApiRequest.call(
							this,
							'POST',
							'/planner/tasks',
							body,
						);

						// Add description if provided
						if (additionalFields.description) {
							const details = await microsoftApiRequest.call(
								this,
								'GET',
								`/planner/tasks/${responseData.id}/details`,
							);

							const eTag = cleanETag(details['@odata.etag']);

							await microsoftApiRequest.call(
								this,
								'PATCH',
								`/planner/tasks/${responseData.id}/details`,
								{
									description: additionalFields.description,
								},
								{},
								undefined,
								{
									'If-Match': eTag,
								},
							);

							responseData.description = additionalFields.description;
						}

						returnData.push(responseData);
					}

					// ----------------------------------
					//         task:get
					// ----------------------------------
					if (operation === 'get') {
						const taskIdParam = this.getNodeParameter('taskId', i);
						const taskId = typeof taskIdParam === 'string' ? taskIdParam : (taskIdParam as IDataObject).value as string;
						const additionalFields = this.getNodeParameter('additionalFields', i) as IDataObject;

						const responseData = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/tasks/${taskId}`,
						);

						if (additionalFields.includeDetails) {
							const details = await microsoftApiRequest.call(
								this,
								'GET',
								`/planner/tasks/${taskId}/details`,
							);
							responseData.details = details;
						}

						returnData.push(responseData);
					}

					// ----------------------------------
					//         task:getAll
					// ----------------------------------
					if (operation === 'getAll') {
						const returnAll = this.getNodeParameter('returnAll', i);
						const filterBy = this.getNodeParameter('filterBy', i) as string;
						const planId = this.getNodeParameter('planId', i) as string;

						let endpoint = '';

						if (filterBy === 'plan') {
							endpoint = `/planner/plans/${planId}/tasks`;
						} else if (filterBy === 'bucket') {
							const bucketIdParam = this.getNodeParameter('bucketId', i);
							const bucketIdValue = typeof bucketIdParam === 'string'
								? bucketIdParam
								: (bucketIdParam as IDataObject).value as string;
							endpoint = `/planner/buckets/${bucketIdValue}/tasks`;
						} else {
							throw new NodeOperationError(
								this.getNode(),
								'You must specify either a Plan ID or Bucket ID to retrieve tasks',
								{ itemIndex: i },
							);
						}

						if (returnAll) {
							const responseData = await microsoftApiRequestAllItems.call(
								this,
								'value',
								'GET',
								endpoint,
							);
							returnData.push(...responseData);
						} else {
							const limit = this.getNodeParameter('limit', i);
							const responseData = await microsoftApiRequest.call(this, 'GET', endpoint, {}, {
								$top: limit,
							});
							returnData.push(...(responseData.value as IDataObject[]));
						}
					}

					// ----------------------------------
					//         task:update
					// ----------------------------------
					if (operation === 'update') {
						const taskIdParam = this.getNodeParameter('taskId', i);
						const taskId = typeof taskIdParam === 'string' ? taskIdParam : (taskIdParam as IDataObject).value as string;
						const updateFields = this.getNodeParameter('updateFields', i) as IDataObject;

						// Get current task to retrieve eTag
						const currentTask = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/tasks/${taskId}`,
						);

						const eTag = cleanETag(currentTask['@odata.etag']);

						const body: IDataObject = {};

						if (updateFields.title) {
							body.title = updateFields.title;
						}

						if (updateFields.priority !== undefined) {
							body.priority = updateFields.priority;
						}

						const formattedDueDateTime = formatDateTime(updateFields.dueDateTime as string);
						if (formattedDueDateTime) {
							body.dueDateTime = formattedDueDateTime;
						}

						const formattedStartDateTime = formatDateTime(updateFields.startDateTime as string);
						if (formattedStartDateTime) {
							body.startDateTime = formattedStartDateTime;
						}

						if (updateFields.percentComplete !== undefined) {
							body.percentComplete = updateFields.percentComplete;
						}

						if (updateFields.bucketId) {
							body.bucketId = updateFields.bucketId;
						}

						// Handle assignments
						if (updateFields.assignments) {
							const emails = parseAssignments(updateFields.assignments as string);
							const userIds: string[] = [];

							for (const email of emails) {
								const userId = await getUserIdByEmail.call(this, email);
								if (userId) {
									userIds.push(userId);
								} else {
									console.warn(`Could not find user ID for email: ${email}`);
								}
							}

							if (userIds.length > 0) {
								body.assignments = createAssignmentsObject(userIds);
							} else if (emails.length > 0) {
								console.warn('No valid user IDs found for assignment. Check if User.Read.All permission is granted.');
							}
						}

						// Only send PATCH request if there are fields to update (excluding description)
						if (Object.keys(body).length > 0) {
							await microsoftApiRequest.call(
								this,
								'PATCH',
								`/planner/tasks/${taskId}`,
								body,
								{},
								undefined,
								{
									'If-Match': eTag,
								},
							);
						}

						// Update description if provided
						if (updateFields.description) {
							const details = await microsoftApiRequest.call(
								this,
								'GET',
								`/planner/tasks/${taskId}/details`,
							);

							const detailsETag = cleanETag(details['@odata.etag']);

							await microsoftApiRequest.call(
								this,
								'PATCH',
								`/planner/tasks/${taskId}/details`,
								{
									description: updateFields.description,
								},
								{},
								undefined,
								{
									'If-Match': detailsETag,
								},
							);
						}

						// Fetch updated task to return complete data
						const responseData = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/tasks/${taskId}`,
						);

						returnData.push(responseData);
					}

					// ----------------------------------
					//         task:delete
					// ----------------------------------
					if (operation === 'delete') {
						const taskIdParam = this.getNodeParameter('taskId', i);
						const taskId = typeof taskIdParam === 'string' ? taskIdParam : (taskIdParam as IDataObject).value as string;

						// Get current task to retrieve eTag
						const currentTask = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/tasks/${taskId}`,
						);

						const eTag = cleanETag(currentTask['@odata.etag']);

						await microsoftApiRequest.call(
							this,
							'DELETE',
							`/planner/tasks/${taskId}`,
							{},
							{},
							undefined,
							{
								'If-Match': eTag,
							},
						);

						returnData.push({ success: true, taskId });
					}


					// ----------------------------------
					//         task:getFiles
					// ----------------------------------
					if (operation === 'getFiles') {
						const taskIdParam = this.getNodeParameter('taskId', i);
						const taskId = typeof taskIdParam === 'string' ? taskIdParam : (taskIdParam as IDataObject).value as string;

						// Get task details
						const details = await microsoftApiRequest.call(this, 'GET', `/planner/tasks/${taskId}/details`);

						const references = details.references || {};
						const files = Object.keys(references).map((encodedUrl) => {
							// Decode the URL
							const url = decodeURIComponent(encodedUrl);
							return {
								url,
								alias: references[encodedUrl].alias,
								type: references[encodedUrl].type,
								previewPriority: references[encodedUrl].previewPriority,
								lastModifiedDateTime: references[encodedUrl].lastModifiedDateTime,
								lastModifiedBy: references[encodedUrl].lastModifiedBy,
							};
						});

						returnData.push({
							taskId,
							fileCount: files.length,
							files,
						});
					}
				}
				// ----------------------------------
				//         plan resource
				// ----------------------------------
				if (resource === 'plan') {
					// plan:create -> POST /planner/plans
					if (operation === 'create') {
						const owner = this.getNodeParameter('owner', i) as string;
						const title = this.getNodeParameter('title', i) as string;

						const body: IDataObject = { owner, title };

						const responseData = await microsoftApiRequest.call(
							this,
							'POST',
							'/planner/plans',
							body,
						);

						returnData.push(responseData);
					}

					// plan:get -> GET /planner/plans/{planId}
					if (operation === 'get') {
						const planId = this.getNodeParameter('planId', i) as string;

						const responseData = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/plans/${planId}`,
						);

						returnData.push(responseData);
					}

					// plan:getAll -> GET /planner/plans
					if (operation === 'getAll') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;

						if (returnAll) {
							const responseData = await microsoftApiRequestAllItems.call(
								this,
								'value',
								'GET',
								'/planner/plans',
							);
							returnData.push(...responseData);
						} else {
							const limit = this.getNodeParameter('limit', i) as number;
							const responseData = await microsoftApiRequest.call(this, 'GET', '/planner/plans', {}, {
								$top: limit,
							});
							returnData.push(...(responseData.value as IDataObject[]));
						}
					}

					// plan:getDetails -> GET /planner/plans/{planId}/details
					if (operation === 'getDetails') {
						const planId = this.getNodeParameter('planId', i) as string;

						const responseData = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/plans/${planId}/details`,
						);

						returnData.push(responseData);
					}

					// plan:update -> PATCH /planner/plans/{planId} with ETag
					if (operation === 'update') {
						const planId = this.getNodeParameter('planId', i) as string;
						const updateFields = this.getNodeParameter('updateFields', i) as IDataObject;

						// Get current plan to obtain ETag
						const currentPlan = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/plans/${planId}`,
						);
						const eTag = cleanETag(currentPlan['@odata.etag']);

						const body: IDataObject = {};
						if (updateFields.title) {
							body.title = updateFields.title;
						}

						if (Object.keys(body).length > 0) {
							await microsoftApiRequest.call(
								this,
								'PATCH',
								`/planner/plans/${planId}`,
								body,
								{},
								undefined,
								{
									'If-Match': eTag,
								},
							);
						}

						const responseData = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/plans/${planId}`,
						);

						returnData.push(responseData);
					}

					// plan:delete -> DELETE /planner/plans/{planId} with ETag
					if (operation === 'delete') {
						const planId = this.getNodeParameter('planId', i) as string;

						const currentPlan = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/plans/${planId}`,
						);
						const eTag = cleanETag(currentPlan['@odata.etag']);

						await microsoftApiRequest.call(
							this,
							'DELETE',
							`/planner/plans/${planId}`,
							{},
							{},
							undefined,
							{
								'If-Match': eTag,
							},
						);

						returnData.push({ success: true, planId });
					}

					// plan:countBuckets -> GET /planner/plans/{planId}/buckets/$count
					if (operation === 'countBuckets') {
						const planId = this.getNodeParameter('planId', i) as string;

						const responseData = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/plans/${planId}/buckets/$count`,
						);

						returnData.push({ planId, count: responseData });
					}

					// plan:countTasks -> GET /planner/plans/{planId}/tasks/$count
					if (operation === 'countTasks') {
						const planId = this.getNodeParameter('planId', i) as string;

						const responseData = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/plans/${planId}/tasks/$count`,
						);

						returnData.push({ planId, count: responseData });
					}

					// plan:updateDetails -> PATCH /planner/plans/{planId}/details with ETag
					if (operation === 'updateDetails') {
						const planId = this.getNodeParameter('planId', i) as string;
						const detailsJson = this.getNodeParameter('detailsJson', i) as string;

						let body: IDataObject;
						try {
							body = JSON.parse(detailsJson) as IDataObject;
						} catch (error) {
							throw new NodeOperationError(this.getNode(), 'Invalid JSON in Plan Details JSON field', {
								itemIndex: i,
							});
						}

						const currentDetails = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/plans/${planId}/details`,
						);
						const eTag = cleanETag(currentDetails['@odata.etag']);

						await microsoftApiRequest.call(
							this,
							'PATCH',
							`/planner/plans/${planId}/details`,
							body,
							{},
							undefined,
							{
								'If-Match': eTag,
							},
						);

						const updatedDetails = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/plans/${planId}/details`,
						);

						returnData.push(updatedDetails);
					}
				}

				// ----------------------------------
				//         bucket resource
				// ----------------------------------
				if (resource === 'bucket') {
					// bucket:create -> POST /planner/plans/{planId}/buckets
					if (operation === 'create') {
						const planId = this.getNodeParameter('planId', i) as string;
						const name = this.getNodeParameter('name', i) as string;

						const body: IDataObject = {
							name,
							planId,
						};

						const responseData = await microsoftApiRequest.call(
							this,
							'POST',
							`/planner/plans/${planId}/buckets`,
							body,
						);

						returnData.push(responseData);
					}

					// bucket:get -> GET /planner/buckets/{bucketId}
					if (operation === 'get') {
						const bucketId = this.getNodeParameter('bucketId', i) as string;

						const responseData = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/buckets/${bucketId}`,
						);

						returnData.push(responseData);
					}

					// bucket:getAll -> GET /planner/plans/{planId}/buckets
					if (operation === 'getAll') {
						const planId = this.getNodeParameter('planId', i) as string;
						const returnAll = this.getNodeParameter('returnAll', i, true) as boolean;

						const endpoint = `/planner/plans/${planId}/buckets`;

						if (returnAll) {
							const responseData = await microsoftApiRequestAllItems.call(
								this,
								'value',
								'GET',
								endpoint,
							);
							returnData.push(...responseData);
						} else {
							const limit = this.getNodeParameter('limit', i, 100) as number;
							const responseData = await microsoftApiRequest.call(this, 'GET', endpoint, {}, {
								$top: limit,
							});
							returnData.push(...(responseData.value as IDataObject[]));
						}
					}

					// bucket:update -> PATCH /planner/buckets/{bucketId} with ETag
					if (operation === 'update') {
						const bucketId = this.getNodeParameter('bucketId', i) as string;
						const updateFields = this.getNodeParameter('updateFields', i) as IDataObject;

						const currentBucket = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/buckets/${bucketId}`,
						);
						const eTag = cleanETag(currentBucket['@odata.etag']);

						const body: IDataObject = {};
						if (updateFields.name) {
							body.name = updateFields.name;
						}
						if (updateFields.orderHint) {
							body.orderHint = updateFields.orderHint;
						}

						if (Object.keys(body).length > 0) {
							await microsoftApiRequest.call(
								this,
								'PATCH',
								`/planner/buckets/${bucketId}`,
								body,
								{},
								undefined,
								{
									'If-Match': eTag,
								},
							);
						}

						const responseData = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/buckets/${bucketId}`,
						);

						returnData.push(responseData);
					}

					// bucket:countTasks -> GET /planner/buckets/{bucketId}/tasks/$count
					if (operation === 'countTasks') {
						const bucketId = this.getNodeParameter('bucketId', i) as string;

						const responseData = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/buckets/${bucketId}/tasks/$count`,
						);

						returnData.push({ bucketId, count: responseData });
					}
				}
			} catch (error) {
				if (this.continueOnFail()) {
					const errorMessage = error instanceof Error ? error.message : 'Unknown error';
					returnData.push({ error: errorMessage });
					continue;
				}
				throw error;
			}
		}

		return [this.helpers.returnJsonArray(returnData)];
	}
}
