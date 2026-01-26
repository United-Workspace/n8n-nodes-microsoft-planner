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
	encodeReferenceKey,
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

						// Add description or references if provided
						if (additionalFields.description || additionalFields.references) {
							const details = await microsoftApiRequest.call(
								this,
								'GET',
								`/planner/tasks/${responseData.id}/details`,
							);

							const eTag = cleanETag(details['@odata.etag']);
							const detailsBody: IDataObject = {};

							if (additionalFields.description) {
								detailsBody.description = additionalFields.description;
								responseData.description = additionalFields.description;
							}

							if (additionalFields.references) {
								const references = additionalFields.references as IDataObject;
								const referenceList = references.reference as IDataObject[];
								const referencesBody: IDataObject = {};
								for (const reference of referenceList) {
									const url = reference.url as string;
									const alias = reference.alias as string;
									const type = reference.type as string;
									// Use helper to encode URL and handle special characters like dots
									const encodedUrl = encodeReferenceKey(url);
									referencesBody[encodedUrl] = {
										'@odata.type': '#microsoft.graph.plannerExternalReference',
										alias,
										type,
									};
								}
								detailsBody.references = referencesBody;
							}

							await microsoftApiRequest.call(
								this,
								'PATCH',
								`/planner/tasks/${responseData.id}/details`,
								detailsBody,
								{},
								undefined,
								{
									'If-Match': eTag,
								},
							);
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
					// Always return all tasks for the given scope; Graph does not honor $top/$limit
					if (operation === 'getAll') {
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

						const responseData = await microsoftApiRequestAllItems.call(
							this,
							'value',
							'GET',
							endpoint,
						);
						returnData.push(...responseData);
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

						// Update description or references if provided
						if (updateFields.description || updateFields.references || updateFields.replaceAllReferences) {
							const details = await microsoftApiRequest.call(
								this,
								'GET',
								`/planner/tasks/${taskId}/details`,
							);

							const detailsETag = cleanETag(details['@odata.etag']);
							const detailsBody: IDataObject = {};

							if (updateFields.description) {
								detailsBody.description = updateFields.description;
							}

							const referencesBody: IDataObject = {};

							// 1. If replaceAllReferences is true, set all existing references to null
							if (updateFields.replaceAllReferences && details.references) {
								for (const key of Object.keys(details.references)) {
									referencesBody[key] = null;
								}
							}

							// 2. Add new/updated references
							if (updateFields.references) {
								const references = updateFields.references as IDataObject;
								const referenceList = references.reference as IDataObject[];

								for (const reference of referenceList) {
									const url = reference.url as string;
									const alias = reference.alias as string;
									const type = reference.type as string;
									// Use helper to encode URL and handle special characters like dots
									const encodedUrl = encodeReferenceKey(url);
									referencesBody[encodedUrl] = {
										'@odata.type': '#microsoft.graph.plannerExternalReference',
										alias,
										type,
									};
								}
							}

							if (Object.keys(referencesBody).length > 0) {
								detailsBody.references = referencesBody;
							}

							await microsoftApiRequest.call(
								this,
								'PATCH',
								`/planner/tasks/${taskId}/details`,
								detailsBody,
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

						// If description (task details) was updated, also return latest details
						if (Object.prototype.hasOwnProperty.call(updateFields, 'description') || Object.prototype.hasOwnProperty.call(updateFields, 'references') || Object.prototype.hasOwnProperty.call(updateFields, 'replaceAllReferences')) {
							const details = await microsoftApiRequest.call(
								this,
								'GET',
								`/planner/tasks/${taskId}/details`,
							);
							(responseData as IDataObject).details = details;
						}

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
						const additionalFields = this.getNodeParameter('additionalFields', i, {}) as IDataObject;

						const responseData = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/plans/${planId}`,
						);

						if (additionalFields.includeDetails) {
							const details = await microsoftApiRequest.call(
								this,
								'GET',
								`/planner/plans/${planId}/details`,
							);
							(responseData as IDataObject).details = details;
						}

						returnData.push(responseData);
					}

					// plan:getAll -> list plans scoped to current user or a specific group
					// Always return all plans; Graph does not honor $top/$limit for Planner.
					if (operation === 'getAll') {
						const scope = this.getNodeParameter('scope', i) as string;
						let endpoint = '';

						if (scope === 'my') {
							// List plans for the current user
							endpoint = '/me/planner/plans';
						} else if (scope === 'group') {
							// List plans owned by a specific Microsoft 365 group
							const groupId = this.getNodeParameter('groupId', i) as string;
							endpoint = `/groups/${groupId}/planner/plans`;
						} else {
							throw new NodeOperationError(this.getNode(), 'Invalid scope for plan getAll operation', {
								itemIndex: i,
							});
						}

						const responseData = await microsoftApiRequestAllItems.call(
							this,
							'value',
							'GET',
							endpoint,
						);
						returnData.push(...responseData);
					}

					// plan:update -> PATCH /planner/plans/{planId} with ETag, and optionally /details
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

						// Optionally update plan details if provided
						const detailsBody: IDataObject = {};
						if (updateFields.categoryDescriptions) {
							const categories = updateFields.categoryDescriptions as IDataObject;
							const cleanedCategories: IDataObject = {};
							for (const key of Object.keys(categories)) {
								const value = categories[key] as string | null | undefined;
								// Empty string in the UI means "unset" → send null to Graph
								if (value === '') {
									cleanedCategories[key] = null;
								} else if (value !== undefined) {
									cleanedCategories[key] = value;
								}
							}
							if (Object.keys(cleanedCategories).length > 0) {
								detailsBody.categoryDescriptions = cleanedCategories;
							}
						}

						if (updateFields.sharedWithJson) {
							let sharedWithParsed: IDataObject;
							try {
								sharedWithParsed = JSON.parse(updateFields.sharedWithJson as string) as IDataObject;
							} catch (error) {
								throw new NodeOperationError(this.getNode(), 'Invalid JSON in Shared With (plannerUserIds JSON) field', {
									itemIndex: i,
								});
							}
							if (Object.keys(sharedWithParsed).length > 0) {
								detailsBody.sharedWith = sharedWithParsed;
							}
						}

						if (Object.keys(detailsBody).length > 0) {
							const currentDetails = await microsoftApiRequest.call(
								this,
								'GET',
								`/planner/plans/${planId}/details`,
							);
							const detailsETag = cleanETag(currentDetails['@odata.etag']);

							await microsoftApiRequest.call(
								this,
								'PATCH',
								`/planner/plans/${planId}/details`,
								detailsBody,
								{},
								undefined,
								{
									'If-Match': detailsETag,
								},
							);
						}

						const responseData = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/plans/${planId}`,
						);

						// If plan details were updated, also include latest details in the response
						if (Object.keys(detailsBody).length > 0) {
							const details = await microsoftApiRequest.call(
								this,
								'GET',
								`/planner/plans/${planId}/details`,
							);
							(responseData as IDataObject).details = details;
						}

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
				}

				// ----------------------------------
				//         bucket resource
				// ----------------------------------
				if (resource === 'bucket') {
					// bucket:create -> POST /planner/buckets
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
							'/planner/buckets',
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
					// Always return all buckets for the plan; Graph does not honor $top/$limit for Planner.
					if (operation === 'getAll') {
						const planId = this.getNodeParameter('planId', i) as string;
						const endpoint = `/planner/plans/${planId}/buckets`;

						const responseData = await microsoftApiRequestAllItems.call(
							this,
							'value',
							'GET',
							endpoint,
						);
						returnData.push(...responseData);
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

					// bucket:delete -> DELETE /planner/buckets/{bucketId} with ETag
					if (operation === 'delete') {
						const bucketId = this.getNodeParameter('bucketId', i) as string;

						const currentBucket = await microsoftApiRequest.call(
							this,
							'GET',
							`/planner/buckets/${bucketId}`,
						);
						const eTag = cleanETag(currentBucket['@odata.etag']);

						await microsoftApiRequest.call(
							this,
							'DELETE',
							`/planner/buckets/${bucketId}`,
							{},
							{},
							undefined,
							{
								'If-Match': eTag,
							},
						);

						returnData.push({ success: true, bucketId });
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
