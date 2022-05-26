'use strict';

goog.provide('Blockly.Blocks.Worksheet');  // Deprecated
goog.provide('Blockly.Constants.Worksheet');

goog.require('Blockly.Blocks');
goog.require('Blockly');

Blockly.defineBlocksWithJsonArray([
	{
		"type": "ws",
		"message0": "Get property or method %1 of  %2",
		"args0": [
			{
				"type": "input_value",
				"name": "NAME"
			},
			{
				"type": "field_input",
				"name": "SHEETNAME",
				"text": "Sheet1"
			}
		],
		"previousStatement": null,
		"nextStatement": null,
		"colour": 60,
		"tooltip": "",
		"helpUrl": ""
	},
	{
		"type": "activate",
		"message0": "Activate",
		"output": null,
		"colour": 0,
		"tooltip": "",
		"helpUrl": ""
	},
	{
		"type": "printout",
		"message0": "Print",
		"output": null,
		"colour": 0,
		"tooltip": "",
		"helpUrl": ""
	},
	{
		"type": "print_copies",
		"message0": "Print %1 copies, from page %2 to %3",
		"args0": [
			{
				"type": "input_value",
				"name": "COPY"
			},
			{
				"type": "input_value",
				"name": "START_PAGE"
			},
			{
				"type": "input_value",
				"name": "END_PAGE"
			}
		],
		"inputsInline": true,
		"output": null,
		"colour": 0,
		"tooltip": "",
		"helpUrl": ""
	},
	{
		"type": "active_sheet",
		"message0": "Active Sheet",
		"inputsInline": true,
		"output": null,
		"colour": 0,
		"tooltip": "",
		"helpUrl": ""
	},
	{
		"type": "range_get",
		"message0": "Get property %1 in Range %2",
		"args0": [
			{
				"type": "input_value",
				"name": "NAME"
			},
			{
				"type": "field_input",
				"name": "RANGE",
				"text": "A1:A2"
			}
		],
		"inputsInline": true,
		"output": null,
		"colour": 230,
		"tooltip": "",
		"helpUrl": ""
	},
	{
		"type": "range_set",
		"message0": "Set property %1 in Range %2 %3 to %4",
		"args0": [
			{
				"type": "input_value",
				"name": "NAME"
			},
			{
				"type": "field_input",
				"name": "RANGE",
				"text": "A1:A2"
			},
			{
				"type": "input_dummy"
			},
			{
				"type": "input_value",
				"name": "VALUE"
			}
		],
		"inputsInline": true,
		"previousStatement": null,
		"nextStatement": null,
		"colour": 230,
		"tooltip": "",
		"helpUrl": ""
	}, {
		"type": "range_get_ws",
		"message0": "Get property %1 in worksheet %2 from Range %3",
		"args0": [
			{
				"type": "input_value",
				"name": "NAME"
			},
			{
				"type": "input_value",
				"name": "WORKSHEET"
			},
			{
				"type": "field_input",
				"name": "RANGE",
				"text": "A1:A2"
			}
		],
		"inputsInline": false,
		"output": null,
		"colour": 230,
		"tooltip": "",
		"helpUrl": ""
	},
	{
		"type": "range_set_ws",
		"message0": "Set property %1 in worksheet %2 in Range %3 %4 to value %5",
		"args0": [
			{
				"type": "input_value",
				"name": "NAME"
			},
			{
				"type": "input_value",
				"name": "WORKSHEET",
			},
			{
				"type": "field_input",
				"name": "RANGE",
				"text": "A1:A2"
			},
			{
				"type": "input_dummy"
			},
			{
				"type": "input_value",
				"name": "VALUE"
			}
		],
		"inputsInline": false,
		"previousStatement": null,
		"nextStatement": null,
		"colour": 230,
		"tooltip": "",
		"helpUrl": ""
	},
]);