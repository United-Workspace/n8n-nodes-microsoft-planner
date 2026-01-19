import { INodeProperties } from 'n8n-workflow';

// ----------------------------------
//         Plan operations
// ----------------------------------
export const planOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['plan'],
			},
		},
		options: [
			{
				name: 'Create',
				value: 'create',
				description: 'Create a new plan',
				action: 'Create a plan',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Get a plan',
				action: 'Get a plan',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get many plans',
				action: 'Get many plans',
			},
			{
				name: "Get Details",
				value: 'getDetails',
				description: 'Get plan details (labels, metadata)',
				action: 'Get plan details',
			},
			{
				name: 'Update',
				value: 'update',
				description: 'Update a plan',
				action: 'Update a plan',
			},
			{
				name: 'Update Details',
				value: 'updateDetails',
				description: 'Update plan details (raw JSON)',
				action: 'Update plan details',
			},
			{
				name: 'Count Buckets',
				value: 'countBuckets',
				description: 'Count buckets in a plan',
				action: 'Count buckets',
			},
			{
				name: 'Count Tasks',
				value: 'countTasks',
				description: 'Count tasks in a plan',
				action: 'Count tasks',
			},
		],
		default: 'getAll',
	},
];

// ----------------------------------
//         Plan fields
// ----------------------------------
export const planFields: INodeProperties[] = [
	// ----------------------------------
	//         plan:create
	// ----------------------------------
	{
		displayName: 'Owner Group ID',
		name: 'owner',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['plan'],
				operation: ['create'],
			},
		},
		default: '',
		description: 'The ID of the Microsoft 365 group that owns the plan',
		placeholder: 'e.g. 02bd9fd6-8f93-4758-87c3-1fb73740a315',
	},
	{
		displayName: 'Title',
		name: 'title',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['plan'],
				operation: ['create'],
			},
		},
		default: '',
		description: 'Title of the plan',
	},

	// ----------------------------------
	//         plan:get
	// ----------------------------------
	{
		displayName: 'Plan ID',
		name: 'planId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['plan'],
				operation: ['get', 'update', 'delete'],
			},
		},
		default: '',
		description: 'The ID of the plan',
		placeholder: 'e.g. xqQg5FS2LkCp935s-FIFm2QAFkHM',
	},

	// ----------------------------------
	//         plan:getAll
	// ----------------------------------
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['plan'],
				operation: ['getAll'],
			},
		},
		default: false,
		description: 'Whether to return all results or only up to a given limit',
	},
	{
		displayName: 'Limit',
		name: 'limit',
		type: 'number',
		displayOptions: {
			show: {
				resource: ['plan'],
				operation: ['getAll'],
				returnAll: [false],
			},
		},
		typeOptions: {
			minValue: 1,
			maxValue: 500,
		},
		default: 100,
		description: 'Max number of results to return',
	},

	// ----------------------------------
	//         plan:update
	// ----------------------------------
	{
		displayName: 'Update Fields',
		name: 'updateFields',
		type: 'collection',
		placeholder: 'Add Field',
		default: {},
		displayOptions: {
			show: {
				resource: ['plan'],
				operation: ['update'],
			},
		},
		options: [
			{
				displayName: 'Title',
				name: 'title',
				type: 'string',
				default: '',
				description: 'Title of the plan',
			},
		],
	},

	// ----------------------------------
	//         plan:updateDetails
	// ----------------------------------
	{
		displayName: 'Plan Details JSON',
		name: 'detailsJson',
		type: 'string',
		required: true,
		typeOptions: {
			rows: 5,
		},
		displayOptions: {
			show: {
				resource: ['plan'],
				operation: ['updateDetails'],
			},
		},
		default: '',
		description: 'Raw JSON body to send to /planner/plans/{planId}/details (e.g. {"categoryDescriptions": {...}})',
	},
];
