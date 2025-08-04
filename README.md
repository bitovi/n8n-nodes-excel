# @bitovi/n8n-nodes-excel

This is an n8n community node that lets you work with Excel files in your n8n workflows.

The Excel node allows you to manipulate Excel workbooks by adding sheets, deleting sheets, and listing available sheets. It works with Excel files passed as binary data through your workflow.

[n8n](https://n8n.io/) is a [fair-code licensed](https://docs.n8n.io/reference/license/) workflow automation platform.

[Installation](#installation)  
[Operations](#operations)  
[Compatibility](#compatibility)  
[Usage](#usage)  
[Resources](#resources)  
[Version History](#version-history)  

## Installation

Follow the [installation guide](https://docs.n8n.io/integrations/community-nodes/installation/) in the n8n community nodes documentation.

### Manual Installation

1. Make sure to allow community nodes with `N8N_COMMUNITY_PACKAGES_ENABLED=true`
2. Once logged in to your N8N web UI, go to `/settings/community-nodes` 
3. Type `@bitovi/n8n-nodes-excel` and click install

## Operations

The Excel node supports the following operations:

### Add Sheet
Adds a new worksheet to an existing Excel workbook.

**Parameters:**
- **Binary Property Name**: Name of the binary property containing the Excel file (default: `data`)
- **Sheet Name**: Name for the new sheet (required)
- **Sheet Contents**: JSON array containing the data to populate the sheet (required)

### Delete Sheet
Removes a worksheet from an existing Excel workbook.

**Parameters:**
- **Binary Property Name**: Name of the binary property containing the Excel file (default: `data`)
- **Sheet Name**: Name of the sheet to delete (required)

### List Sheets
Returns a list of all sheet names in the Excel workbook.

**Parameters:**
- **Binary Property Name**: Name of the binary property containing the Excel file (default: `data`)
- **Include Hidden Sheets**: Whether to include hidden sheets in the list (default: `false`)

## Compatibility

- **Minimum n8n version**: 0.175.0
- **Node.js**: >=18.10
- **Tested with**: n8n 1.x

This node uses the `xlsx` library (v0.18.5) for Excel file manipulation, which supports a wide range of Excel formats including `.xlsx`, `.xls`, `.csv`, and more.

## Usage

### Basic Workflow Example

1. **Read Excel File**: Use a node like "Read Binary File" or "HTTP Request" to get your Excel file as binary data
2. **Excel Node**: Add the Excel node and configure your desired operation
3. **Process Results**: Use the output (modified Excel file or sheet names) in subsequent nodes

### Adding Data to a New Sheet

```json
[
  {
    "name": "John Doe",
    "email": "john@example.com",
    "age": 30
  },
  {
    "name": "Jane Smith", 
    "email": "jane@example.com",
    "age": 25
  }
]
```

### Working with Binary Data

The Excel node expects the Excel file to be available as binary data in the input. Make sure your previous node outputs the Excel file in binary format. The node will output the modified Excel file as binary data that can be saved or passed to other nodes.

**Note**: When using "List Sheets" operation, the output will be JSON data containing the sheet names array, not binary data.

## Resources

* [n8n community nodes documentation](https://docs.n8n.io/integrations/community-nodes/)
* [SheetJS Documentation](https://docs.sheetjs.com/) - The underlying library used for Excel manipulation
* [Excel File Format Documentation](https://support.microsoft.com/en-us/office/file-formats-that-are-supported-in-excel-0943ff2c-6014-4e8d-aaea-b83d51d46247)

## Version History

### v0.2.1
- Current stable version
- Supports Add Sheet, Delete Sheet, and List Sheets operations
- Uses xlsx library v0.18.5
- Node.js 18.10+ compatibility

## Need help or have questions?

Need guidance on leveraging AI agents or N8N for your business? Our [AI Agents workshop](https://hubs.ly/Q02X-9Qq0) will equip you with the knowledge and tools necessary to implement successful and valuable agentic workflows.

## License

[MIT](./LICENSE.md)
