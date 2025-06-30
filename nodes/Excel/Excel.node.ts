import {
	NodeOperationError,
	type IExecuteFunctions,
	type INodeExecutionData,
	type INodeType,
	type INodeTypeDescription,
} from 'n8n-workflow';
import xlsx from 'xlsx';

enum Action {
	ADD_SHEET = 'addSheet',
	DELETE_SHEET = 'deleteSheet',
	LIST_SHEETS = 'listSheets',
}
export class Excel implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Excel',
		icon: 'file:excel.svg',
		name: 'excel',
		group: ['transform'],
		version: 1,
		subtitle: '={{ $parameter["operation"] }}',
		description: 'Excel Node',
		defaults: {
			name: 'Excel',
		},
		inputs: ['main'],
		outputs: ['main'],
		properties: [
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'Add Sheet',
						value: Action.ADD_SHEET,
					},
					{
						name: 'Delete Sheet',
						value: Action.DELETE_SHEET,
					},
					{
						name: 'List Sheets',
						value: Action.LIST_SHEETS,
					},
				],
				default: 'listSheets',
			},
			{
				displayName: 'Binary Property Name',
				name: 'binaryPropertyName',
				type: 'string',
				noDataExpression: false,
				default: 'data',
			},
			{
				displayName: 'Sheet Name',
				name: 'sheetName',
				type: 'string',
				noDataExpression: false,
				required: true,
				default: '',
				displayOptions: {
					show: {
						operation: [Action.ADD_SHEET, Action.DELETE_SHEET],
					},
				},
			},
			{
				displayName: 'Sheet Contents',
				name: 'sheetContents',
				type: 'json',
				noDataExpression: false,
				required: true,
				default: '[]',
				displayOptions: {
					show: {
						operation: [Action.ADD_SHEET],
					},
				},
			},
			{
				displayName: 'Include Hidden Sheets',
				name: 'includeHiddenSheets',
				type: 'boolean',
				noDataExpression: false,
				default: false,
				displayOptions: {
					show: {
						operation: [Action.LIST_SHEETS],
					},
				},
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();

		const returnData: INodeExecutionData[] = [];

		for (let i = 0; i < items.length; i++) {
			const operation = this.getNodeParameter('operation', i) as Action;
			const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i) as string;

			const binary = items[i].binary?.[binaryPropertyName];

			if (!binary) {
				throw new NodeOperationError(
					this.getNode(),
					`Binary property "${binaryPropertyName}" not found.`,
				);
			}

			const { data: binaryData, fileName, mimeType } = binary;
			const workbook = xlsx.read(Buffer.from(binaryData, 'base64'), { type: 'buffer' });

			switch (operation) {
				case Action.ADD_SHEET: {
					const sheetName = this.getNodeParameter('sheetName', i) as string;
					const sheetContents = this.getNodeParameter('sheetContents', i) as Record<string, any>[];

					xlsx.utils.book_append_sheet(workbook, xlsx.utils.json_to_sheet(sheetContents), sheetName);

					returnData.push({
						json: {},
						binary: {
							[binaryPropertyName]: {
								data: xlsx.write(workbook, { type: 'buffer' }).toString('base64'),
								mimeType,
								fileName,
							},
						},
					});

					break;
				}
				case Action.DELETE_SHEET: {
					const sheetName = this.getNodeParameter('sheetName', i) as string;

					workbook.SheetNames = workbook.SheetNames.filter(name => name !== sheetName);
					delete workbook.Sheets[sheetName];

					returnData.push({
						json: {},
						binary: {
							[binaryPropertyName]: {
								data: xlsx.write(workbook, { type: 'buffer' }).toString('base64'),
								mimeType,
								fileName,
							},
						},
					});

					break;
				}
				case Action.LIST_SHEETS: {
					const includeHiddenSheets = this.getNodeParameter('includeHiddenSheets', i) as boolean;

					const visibleSheetNames = includeHiddenSheets
						? workbook.SheetNames
						: workbook.SheetNames.filter(
								(sheetName) =>
									!workbook.Workbook?.Sheets?.find(({ name }) => name === sheetName)?.Hidden,
							);

					returnData.push({
						json: {
							sheetNames: visibleSheetNames,
						},
					});

					break;
				}
				default: {
					// Do nothing
				}
			}
		}

		return [returnData];
	}
}
