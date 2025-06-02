import {
	NodeOperationError,
	type IExecuteFunctions,
	type INodeExecutionData,
	type INodeType,
	type INodeTypeDescription,
} from 'n8n-workflow';
import xlsx from 'xlsx';

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
						name: 'List Sheets',
						value: 'listSheets',
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
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();

		const returnData: INodeExecutionData[] = [];

		for (let i = 0; i < items.length; i++) {
			const operation = this.getNodeParameter('operation', i) as string;
			const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i) as string;

			switch (operation) {
				case 'listSheets': {
					if (!items[i].binary?.[binaryPropertyName]) {
						throw new NodeOperationError(
							this.getNode(),
							`Binary property "${binaryPropertyName}" not found.`,
						);
					}

					const binaryData = items[i].binary![binaryPropertyName].data;
					const buffer = Buffer.from(binaryData, 'base64');

					const workbook = xlsx.read(buffer, { type: 'buffer' });

					const visibleSheetNames = workbook.SheetNames.filter(sheetName => {
						return !workbook.Workbook?.Sheets?.find(s => s.name === sheetName)?.Hidden;
					});

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
