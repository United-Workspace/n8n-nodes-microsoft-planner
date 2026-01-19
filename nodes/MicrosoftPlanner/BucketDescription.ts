import { INodeProperties } from 'n8n-workflow';

// ----------------------------------
//         Bucket operations
// ----------------------------------
export const bucketOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['bucket'],
			},
		},
		options: [
			{
				name: 'Create',
				value: 'create',
				description: 'Create a new bucket in a plan',
				action: 'Create a bucket',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Get a bucket',
				action: 'Get a bucket',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get many buckets in a plan',
				action: 'Get many buckets',
			},
			{
				name: 'Update',
				value: 'update',
				description: 'Update a bucket',
				action: 'Update a bucket',
			},
			{
				name: 'Count Tasks',
				value: 'countTasks',
				description: 'Count tasks in a bucket',
				action: 'Count tasks in bucket',
			},
		],
		default: 'getAll',
	},
];

// ----------------------------------
//         Bucket fields
// ----------------------------------
export const bucketFields: INodeProperties[] = [
	// ----------------------------------
	//         bucket:create
	// ----------------------------------
	{
		displayName: 'Plan ID',
		name: 'planId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['bucket'],
				operation: ['create', 'getAll'],
			},
		},
		default: '',
		description: 'The ID of the plan this bucket belongs to',
	},

	// ----------------------------------
	//         bucket:getAll
	// ----------------------------------
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['bucket'],
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
				resource: ['bucket'],
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
	{
		displayName: 'Name',
		name: 'name',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['bucket'],
				operation: ['create'],
			},
		},
		default: '',
		description: 'Name of the bucket',
	},

	// ----------------------------------
	//         bucket:get / bucket:update
	// ----------------------------------
	{
		displayName: 'Bucket ID',
		name: 'bucketId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['bucket'],
				operation: ['get', 'update'],
			},
		},
		default: '',
		description: 'The ID of the bucket',
	},
	{
		displayName: 'Update Fields',
		name: 'updateFields',
		type: 'collection',
		placeholder: 'Add Field',
		default: {},
		displayOptions: {
			show: {
				resource: ['bucket'],
				operation: ['update'],
			},
		},
		options: [
			{
				displayName: 'Name',
				name: 'name',
				type: 'string',
				default: '',
				description: 'Name of the bucket',
			},
			{
				displayName: 'Order Hint',
				name: 'orderHint',
				type: 'string',
				default: '',
				description: 'Used to sort buckets in the plan',
			},
		],
	},
];
