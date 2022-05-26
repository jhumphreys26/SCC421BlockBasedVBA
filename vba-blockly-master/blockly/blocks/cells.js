'use strict';

goog.provide('Blockly.Blocks.Cells');  // Deprecated
goog.provide('Blockly.Constants.Cells');

goog.require('Blockly.Blocks');
goog.require('Blockly');

Blockly.defineBlocksWithJsonArray([
	{
		"type": "get_cell",
		"message0": "Get property %1 of cell at Row %2 , Column %3",
		"args0": [
			{
				"type": "input_value",
				"name": "PROP"
			},
			{
				"type": "input_value",
				"name": "ROW"
			},
			{
				"type": "input_value",
				"name": "COL"
			}
		],


		"inputsInline": true,
		"output": null,
		"colour": 60,
		"tooltip": "",
		"helpUrl": ""
	},
	{
		"type": "get_cell_worksheet",
		"message0": "Get property %1 of cell at Row %2 Column %3 Worksheet %4",
		"args0": [
			{
				"type": "input_value",
				"name": "PROP"
			},
			{
				"type": "input_value",
				"name": "ROW"
			},
			{
				"type": "input_value",
				"name": "COL"
			},
			{
				"type": "input_value",
				"name": "SHEETNAME",
			}
		],
		"inputsInline": false,
		"output": null,
		"colour": 60,
		"tooltip": "",
		"helpUrl": "",
	},
	{
		"type": "value",
		"message0": "Value",
		"output": null,
		"colour": 0,
		"tooltip": "",
		"helpUrl": ""
	},
	{
		"type": "formula",
		"message0": "Formula",
		"output": null,
		"colour": 0,
		"tooltip": "",
		"helpUrl": ""
	},
	{
		"type": "formulaR1C1",
		"message0": "Formula(R1C1)",
		"output": null,
		"colour": 0,
		"tooltip": "",
		"helpUrl": ""
	},
	{
		"type": "set_cell",
		"message0": "Set property %1 of cell at Row %2 Column %3 to Value %4",
		"args0": [
			{
				"type": "input_value",
				"name": "PROP"
			},
			{
				"type": "input_value",
				"name": "ROW"
			},
			{
				"type": "input_value",
				"name": "COL"
			},
			{
				"type": "input_value",
				"name": "VAL"
			}
		],
		//In order to build statement blocks, needs to not have the output option
		"previousStatement": null,
		"nextStatement": null,
		"inputsInline": true,

		//"output": null,
		"colour": 60,
		"tooltip": "",
		"helpUrl": ""

	},
	{
		"type": "set_cell_worksheet",
		"message0": "Set property %1 of cell at Row %2 Column %3 in Worksheet %4 to Value %5",
		"args0": [
			{
				"type": "input_value",
				"name": "PROP"
			},
			//	{
			//		"type": "field_dropdown",
			//		"name": "PROP",
			//		"options": [
			//			[
			//				"Value",
			//				"VALUE"
			//			],
			//			[
			//				"Formula",
			//				"FORMULA"
			//			]
			//		]
			//	},
			{
				"type": "input_value",
				"name": "ROW"
			},
			{
				"type": "input_value",
				"name": "COL"
			},
			{
				"type": "input_value",
				"name": "WORKSHEET"
			},
			{
				"type": "input_value",
				"name": "VAL"
			}
		],
		"inputsInline": false,
		"colour": 60,
		"tooltip": "",
		"helpUrl": "",
		"previousStatement": null,
		"nextStatement": null
	},
	{
		"type": "font",
		"message0": "Font %1",
		"args0": [
			{
				"type": "field_dropdown",
				"name": "OPT",
				"options": [
					[
						"Name",
						"NAME"
					],
					[
						"Size",
						"SIZE"
					],
					[
						"Colour",
						"COLOUR"
					],
				]
			}
		],
		"inputsInline": true,
		"output": null,
		"colour": 0,
		"tooltip": "",
		"helpUrl": ""
	},
	{
		"type": "colour",
		"message0": "Color",
		"output": null,
		"colour": 0,
		"tooltip": "",
		"helpUrl": ""
	},
	{
		"type": "fill",
		"message0": "%1",
		"args0": [
			{
				"type": "field_colour",
				"name": "COL",
				"colour": "#ff0000"
			}
		],
		"output": null,
		"colour": 0,
		"tooltip": "",
		"helpUrl": ""
	},
	{
		"type": "wrap",
		"message0": "Wrap Text",
		"output": null,
		"colour": 0,
		"tooltip": "",
		"helpUrl": ""
	},
	{
		"type": "border",
		"message0": "Border %1",
		"args0": [
			{
				"type": "field_dropdown",
				"name": "OPT",
				"options": [
					[
						"xlDiagonalDown",
						"XLDIAGONALDOWN"
					],
					[
						"xlDiagonalUp",
						"XLDIAGONALUP"
					],
					[
						"xlEdgeBottom",
						"XLEDGEBOTTOM"
					],
					[
						"xlEdgeLeft",
						"XLEDGELEFT"
					],
					[
						"xlEdgeRight",
						"XLEDGERIGHT"
					],
					[
						"xlEdgeTop",
						"XLEDGETOP"
					],
					[
						"xlInsideHorizontal",
						"XLINSIDEHORIZONTAL"
					],
					[
						"xlInsideVertical",
						"XLINSIDEVERTICAL"
					],
				]
			}
		],
		"inputsInLine": true,
		"output": null,
		"colour": 0,
		"tooltip": "",
		"helpUrl": ""
	},
	{
		"type": "linestyle",
		"message0": "Line Style %1",
		"args0": [
			{
				"type": "field_dropdown",
				"name": "OPT",
				"options": [
					[
						"xlContinuous",
						"XLCONTINUOUS"
					],
					[
						"xldash",
						"XLDASH"
					],
					[
						"xlDashDot",
						"XLDASHDOT"
					],
					[
						"xlDashDotDot",
						"XLDASHDOTDOT"
					],
					[
						"xlDot",
						"XLDOT"
					],
					[
						"xlDouble",
						"XLDOUBLE"
					],
					[
						"XlLineStyleNone",
						"XLLINESTYLENONE"
					],
					[
						"xlSlantDashDot",
						"XLSLANTDASHDOT"
					]
				]


			}
		],
		"inputsInline": true,
		"output": null,
		"colour": 0,
		"tooltip": "",
		"helpUrl": ""
	},
	{
		"type": "rowheight",
		"message0": "Row Height",
		"output": null,
		"colour": 0,
		"tooltip": "",
		"helpUrl": ""
	},
	{
		"type": "columnwidth",
		"message0": "Column Width",
		"output": null,
		"colour": 0,
		"tooltip": "",
		"helpUrl": ""
	}
]);