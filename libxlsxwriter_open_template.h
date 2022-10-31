/**
 * File              : libxlsxwriter_open_template.h
 * Author            : Igor V. Sementsov <ig.kuzm@gmail.com>
 * Date              : 30.10.2022
 * Last Modified Date: 31.10.2022
 * Last Modified By  : Igor V. Sementsov <ig.kuzm@gmail.com>
 */

#ifndef _LXLSXWOT_H
#define _LXLSXWOT_H

#include <stdlib.h>
#include <stdio.h>
#include <stdbool.h>
#include "zip.h"
#include "ezxml.h"
#include "libxlsxwriter/include/xlsxwriter.h"

#ifdef __cplusplus
extern "C" {
#endif

	lxw_workbook *workbook_new_from_template(const char *filename, const char *template);


	/*
	 * IMP
	 */

	struct key_value{
		uint8_t value;
		char    key[32];	
	};

	const struct key_value border_styles[] = {
		{LXW_BORDER_THIN,                "none"             },
		{LXW_BORDER_THIN,                "thin"             },
		{LXW_BORDER_MEDIUM,              "medium"           },
		{LXW_BORDER_DASHED,              "dashed"           },
		{LXW_BORDER_DOTTED,              "dotted"           },
		{LXW_BORDER_THICK,               "thick"            },
		{LXW_BORDER_DOUBLE,              "double"           },
		{LXW_BORDER_HAIR,                "hair"             },
		{LXW_BORDER_MEDIUM_DASHED,       "mediumDashed"     },
		{LXW_BORDER_DASH_DOT,            "dashDot"          },
		{LXW_BORDER_MEDIUM_DASH_DOT,     "mediumDashDot"    },
		{LXW_BORDER_DASH_DOT_DOT,        "dashDotDot"       },
		{LXW_BORDER_MEDIUM_DASH_DOT_DOT, "mediumDashDotDot" },
		{LXW_BORDER_SLANT_DASH_DOT,      "slantDashDot"     },
	};

	const struct key_value pattern_types[] = {
		{LXW_PATTERN_NONE,           	 "none"            },
		{LXW_PATTERN_SOLID,           	 "solid"           },
		{LXW_PATTERN_MEDIUM_GRAY,        "mediumGray"      },
		{LXW_PATTERN_DARK_GRAY,          "darkGray"        },
		{LXW_PATTERN_LIGHT_GRAY,         "lightGray"       },
		{LXW_PATTERN_DARK_HORIZONTAL,    "darkHorizontal"  },
		{LXW_PATTERN_DARK_VERTICAL,      "darkVertical"    },
		{LXW_PATTERN_DARK_DOWN,          "darkDown"        },
		{LXW_PATTERN_DARK_UP,            "darkUp"          },
		{LXW_PATTERN_DARK_GRID,          "darkGrid"        },
		{LXW_PATTERN_DARK_TRELLIS,       "darkTrellis"     },
		{LXW_PATTERN_LIGHT_HORIZONTAL,   "lightHorizontal" },
		{LXW_PATTERN_LIGHT_VERTICAL,     "lightVertical"   },
		{LXW_PATTERN_LIGHT_DOWN,         "lightDown"       },
		{LXW_PATTERN_LIGHT_UP,           "lightUp"         },
		{LXW_PATTERN_LIGHT_GRID,         "lightGrid"       },
		{LXW_PATTERN_LIGHT_TRELLIS,      "lightTrellis"    },
		{LXW_PATTERN_GRAY_125,           "gray125"         },
		{LXW_PATTERN_GRAY_0625,          "gray0625"        },
	};

	const struct key_value horizontal_alignment[] = {
		{LXW_ALIGN_LEFT,                 "left"            },
        	{LXW_ALIGN_CENTER,               "center"          },
        	{LXW_ALIGN_RIGHT,                "right"           },
        	{LXW_ALIGN_FILL,                 "fill"            },
        	{LXW_ALIGN_JUSTIFY,              "justify"         },
        	{LXW_ALIGN_CENTER_ACROSS,        "centerContinuous"},
        	{LXW_ALIGN_DISTRIBUTED,          "distributed"     },
	};

	const struct key_value vertical_alignment[] = {
		{LXW_ALIGN_VERTICAL_TOP,         "top"             },
        	{LXW_ALIGN_VERTICAL_BOTTOM,      "bottom"          },
        	{LXW_ALIGN_VERTICAL_CENTER,      "center"          },
        	{LXW_ALIGN_VERTICAL_JUSTIFY,     "justify"         },
        	{LXW_ALIGN_VERTICAL_DISTRIBUTED, "distributed"     },
	};

	void parse_worksheet(lxw_workbook *wb, ezxml_t xml, lxw_worksheet *ws, ezxml_t sst, ezxml_t styles) {
		int i;
		
		/*! TODO: add sheet properties
		 *  \todo add sheet properties
		 */

		ezxml_t data = ezxml_child(xml, "sheetData");
		if (!data){
			printf("%s\n", "NO DATA IN SHEET");
		}
		//parse cols 
		ezxml_t cols = ezxml_child(data, "cols");
		ezxml_t col = ezxml_child(cols, "col");
		for(;col; col = col->next){
			const char * min = ezxml_attr(col, "min");
			if (min){
			}			

			const char * max = ezxml_attr(col, "max");
			if (max){
			}
			
			const char * bestFit = ezxml_attr(col, "bestFit");
			if (bestFit){
			}

			const char * customWidth = ezxml_attr(col, "customWidth");
			if (customWidth && atoi(customWidth)){
				const char * width = ezxml_attr(col, "width");
				if (width && max && min){
					worksheet_set_column(ws, atoi(min), atoi(max), atof(width), NULL);
				}			
			}			
		}

		//merge cells
		ezxml_t mergeCells = ezxml_child(xml, "mergeCells");
		ezxml_t mergeCell =  ezxml_child(mergeCells, "mergeCell");
		for (; mergeCell; mergeCell = mergeCell->next) {
			const char * ref = ezxml_attr(mergeCell, "ref");
			if (ref){
				worksheet_merge_range(ws, RANGE(ref), "", NULL);
			}
		}

		//parse row
		ezxml_t row = ezxml_child(data, "row");
		for(;row; row = row->next){
			//get row ref
			const char * r = ezxml_attr(row, "r");
			
			//get row spans
			const char * spans = ezxml_attr(row, "spans");
			if (spans){
			}

			const char * customHeight = ezxml_attr(row, "customHeight");
			if (customHeight && atoi(customHeight)){
				const char * ht = ezxml_attr(row, "ht");
				if (ht){
					worksheet_set_row(ws, atoi(r), atof(ht), NULL);
				}			
			}			
			//thick bottom
			const char * thickBot = ezxml_attr(row, "thickBot");
			if (thickBot){
			}
			//thick top
			const char * thickTop = ezxml_attr(row, "thickTop");
			if (thickTop){
			}
			//??
			const char * dyDescent = ezxml_attr(row, "x14ac:dyDescent");
			if (dyDescent){
			}

			//get cell
			ezxml_t cell = ezxml_child(row, "c");
			//int c = 0, l;
			for(; cell; cell = cell->next){
				//cell format
				lxw_format *format = workbook_add_format(wb);

				//cell ref 
				const char * r = ezxml_attr(cell, "r");
				
				//cell style
				const char * s = ezxml_attr(cell, "s");
				if (s) {
					ezxml_t xf = ezxml_get(styles, "cellXfs", 0, "xf", atoi(s), "");
					if (xf){
						const char * applyNumberFormat = ezxml_attr(xf, "applyNumberFormat");
						if (applyNumberFormat && atoi(applyNumberFormat)){
							const char * numFmtId = ezxml_attr(xf, "numFmtId");
							if (numFmtId){
								//set fmt
								format_set_num_format_index(format, atoi(numFmtId));
							}
						}
						const char * applyFill = ezxml_attr(xf, "applyFill");
						if (applyFill && atoi(applyFill)){
							const char * fillId = ezxml_attr(xf, "fillId");
							if (fillId){
								//set fill
								ezxml_t patternFill = ezxml_get(styles, "fills", 0, "fill", atoi(fillId), "patternFill", 0, "");
								if (patternFill){
									bool has_color = false;
									ezxml_t fgColor = ezxml_child(patternFill, "fgColor");
									if (fgColor){
										const char * rgb = ezxml_attr(fgColor, "rgb"); 
										if (rgb){
											//get hex from string
											int rgb_int = 0;
											sscanf(rgb, "%x", &rgb_int);
											//set color
											format_set_fg_color(format, rgb_int);
											has_color = true;
										}
										const char * fgColor_indexed = ezxml_attr(fgColor, "indexed");
										if (fgColor_indexed){
											format_set_color_indexed(format, atoi(fgColor_indexed));
											has_color = true;
										}
										/*
										const char * fgColor_theme = ezxml_attr(fgColor, "theme"); 
										if (fgColor_theme){
											format_set_theme(format, atoi(fgColor_theme));
											has_color = true;
										}
										*/										
									}
									ezxml_t bgColor = ezxml_child(patternFill, "bgColor");
									if (bgColor){
										const char * rgb = ezxml_attr(bgColor, "rgb"); 
										if (rgb){
											//get hex from string
											int rgb_int = 0;
											sscanf(rgb, "%x", &rgb_int);
											//set color
											format_set_bg_color(format, rgb_int);
											has_color = true;
										}
										
										const char * bgColor_indexed = ezxml_attr(bgColor, "indexed"); 
										if (bgColor_indexed){
											format_set_color_indexed(format, atoi(bgColor_indexed));
											has_color = true;
										}										
										/*
										const char * bgColor_theme = ezxml_attr(bgColor, "theme"); 
										if (bgColor_theme){
											format_set_theme(format, atoi(bgColor_theme));
											has_color = true;
										}
										*/										
									}
									const char * patternType = ezxml_attr(patternFill, "patternType");
									if (patternType){
										if (has_color){ //cant get color from theme and indexed yet
											for (i=0; i<sizeof(pattern_types)/sizeof(struct key_value); i++)
												if (strcmp(patternType, pattern_types[i].key) == 0){
													format_set_pattern(format, pattern_types[i].value);												
													break;
												}
										}
									}
								}
							}						
						}						
						const char * xfId = ezxml_attr(xf, "xfId");
						if (xfId){
							//set xf
						}						
						const char * applyProtection = ezxml_attr(xf, "applyProtection");
						if (applyProtection){
							//set xf
						}						
						const char * applyFont = ezxml_attr(xf, "applyFont");
						if (applyFont && atoi(applyFont)){
							//set applyFont
							const char * fontId = ezxml_attr(xf, "fontId");
							if (fontId){
								//set font
								ezxml_t font = ezxml_get(styles, "fonts", 0, "font", atoi(fontId), "");
								if (font){
									ezxml_t name = ezxml_child(font, "name");
									if (name){
										const char * val = ezxml_attr(name, "val");
										if(val)
											format_set_font_name(format, val);
									}
									ezxml_t color = ezxml_child(font, "color");
									if (color){
										const char * rgb = ezxml_attr(color, "rgb");
										if (rgb){
											//get hex from string
											int rgb_int = 0;
											sscanf(rgb, "%x", &rgb_int);
											//set font color
											format_set_font_color(format, rgb_int);
										}
									}									
									ezxml_t family = ezxml_child(font, "family");
									if (family){
										const char * val = ezxml_attr(family, "val");
										if(val)
											format_set_font_family(format, atoi(val));
									}									
									ezxml_t charset = ezxml_child(font, "charset");
									if (charset){
										const char * val = ezxml_attr(charset, "val");
										if(val)
											format_set_font_charset(format, atoi(val));
									}									
									ezxml_t sz = ezxml_child(font, "sz"); //size
									if (sz){
										const char * val = ezxml_attr(sz, "val");
										if(val)
											format_set_font_size(format, atof(val));
									}									
									ezxml_t b = ezxml_child(font, "b"); //bold
									if (b){
										format_set_bold(format);
									}									
									ezxml_t i = ezxml_child(font, "i"); //italic
									if (i){
										format_set_italic(format);
									}									
									ezxml_t outline = ezxml_child(font, "outline"); //outline
									if (outline){
										format_set_font_outline(format);
									}									
									ezxml_t strike = ezxml_child(font, "strike"); //strike
									if (strike){
										format_set_font_strikeout(format);
									}									
									ezxml_t shadow = ezxml_child(font, "shadow"); //shadow
									if (shadow){
										format_set_font_shadow(format);
									}									
								}
							}						
						}						
						const char * applyBorder = ezxml_attr(xf, "applyBorder");
						if (applyBorder && atoi(applyBorder)){
							//set applyBorder
							const char * borderId = ezxml_attr(xf, "borderId");
							if (borderId){
								//set border
								ezxml_t border = ezxml_get(styles, "borders", 0, "border", atoi(borderId), ""); 
								if (border){
									ezxml_t left = ezxml_child(border, "left");
									if (left) {
										const char * style = ezxml_attr(left, "style");
										if (style){
											for (i=0; i<sizeof(border_styles)/sizeof(struct key_value); i++)
												if (strcmp(style, border_styles[i].key) == 0){
													format_set_left(format, border_styles[i].value);
													break;
												}
										}
										ezxml_t color = ezxml_child(left, "color");
										if (color){
											const char * rgb = ezxml_attr(color, "rgb");
											if (rgb){
												//get hex from string
												int rgb_int = 0;
												sscanf(rgb, "%x", &rgb_int);
												//set font color
												format_set_left_color(format, rgb_int);
											}
										}
									}
									ezxml_t right = ezxml_child(border, "right");
									if (right) {
										const char * style = ezxml_attr(right, "style");
										if (style){
											for (i=0; i<sizeof(border_styles)/sizeof(struct key_value); i++)
												if (strcmp(style, border_styles[i].key) == 0){
													format_set_right(format, border_styles[i].value);
													break;
												}
										}										
										ezxml_t color = ezxml_child(right, "color");
										if (color){
											const char * rgb = ezxml_attr(color, "rgb");
											if (rgb){
												//get hex from string
												int rgb_int = 0;
												sscanf(rgb, "%x", &rgb_int);
												//set font color
												format_set_right_color(format, rgb_int);
											}
										}										
									}									
									ezxml_t top = ezxml_child(border, "top");
									if (top) {
										const char * style = ezxml_attr(top, "style");
										if (style){
											for (i=0; i<sizeof(border_styles)/sizeof(struct key_value); i++)
												if (strcmp(style, border_styles[i].key) == 0){
													format_set_top(format, border_styles[i].value);
													break;
												}
										}	
										ezxml_t color = ezxml_child(top, "color");
										if (color){
											const char * rgb = ezxml_attr(color, "rgb");
											if (rgb){
												//get hex from string
												int rgb_int = 0;
												sscanf(rgb, "%x", &rgb_int);
												//set font color
												format_set_top_color(format, rgb_int);
											}
										}										
									}																		
									ezxml_t bottom = ezxml_child(border, "bottom");
									if (bottom) {
										const char * style = ezxml_attr(bottom, "style");
										if (style){
											for (i=0; i<sizeof(border_styles)/sizeof(struct key_value); i++)
												if (strcmp(style, border_styles[i].key) == 0){
													format_set_bottom(format, border_styles[i].value);											
													break;
												}
										}	
										ezxml_t color = ezxml_child(bottom, "color");
										if (color){
											const char * rgb = ezxml_attr(color, "rgb");
											if (rgb){
												//get hex from string
												int rgb_int = 0;
												sscanf(rgb, "%x", &rgb_int);
												//set font color
												format_set_bottom_color(format, rgb_int);
											}
										}										
									}																										
									ezxml_t diagonal = ezxml_child(border, "diagonal");
									if (diagonal) {
									}									
								}
							}						
						}						
						const char * applyAlignment = ezxml_attr(xf, "applyAlignment");
						if (applyAlignment){
							//set applyAlignment
							ezxml_t alignment = ezxml_child(xf, "alignment");
							if (alignment){
								const char * horizontal = ezxml_attr(alignment, "horizontal");
								if (horizontal){
									for (i=0; i<sizeof(horizontal_alignment)/sizeof(struct key_value); i++)
										if (strcmp(horizontal, horizontal_alignment[i].key) == 0){
											format_set_align(format, horizontal_alignment[i].value);
											break;
										}
								}
								const char * vertical = ezxml_attr(alignment, "vertical");
								if (vertical){
									for (i=0; i<sizeof(vertical_alignment)/sizeof(struct key_value); i++)
										if (strcmp(vertical, vertical_alignment[i].key) == 0){
											format_set_align(format, vertical_alignment[i].value);
											break;
										}
								}
								const char * wrapText = ezxml_attr(alignment, "wrapText");
								if (wrapText && atoi(wrapText)){
									format_set_text_wrap(format);
								}
							}
						}						
					}
				}
				
				//cell type
				const char * type = ezxml_attr(cell, "t");
				if (type) {
					if (strcmp(type, "b") == 0){
						//boolean
						ezxml_t v = ezxml_child(cell, "v");
						if (v){
							worksheet_write_boolean(ws, CELL(r), atoi(v->txt), format);
						}
					} 
					else 
					if (strcmp(type, "e") == 0){
						//error
					} 
					else 
					if (strcmp(type, "inlineStr") == 0){
						//inline str
						ezxml_t t = ezxml_get(cell, "is", 0, "t", 0, "");
						if (t){
							const char * string = t->txt;
							if(string)
								worksheet_write_string(ws, CELL(r), string, format);
						}
					} 
					else 
					if (strcmp(type, "n") == 0){
						//number type
						ezxml_t v = ezxml_child(cell, "v");
						if (v){
							const char * value = v->txt;
							if(value)
								worksheet_write_number(ws, CELL(r), atoi(value), format);
						}						
					}
					else 
					if (strcmp(type, "s") == 0){
						//string sst
						ezxml_t v = ezxml_child(cell, "v");
						if (v){
							const char * value = v->txt;
							ezxml_t t = ezxml_get(sst, "si", atoi(value), "t", 0, "");
							if (t){
								const char * string = t->txt;
								if(string)
									worksheet_write_string(ws, CELL(r), string, format);
							}
						}						
					}
					else 
					if (strcmp(type, "str") == 0){
						//formula string
						ezxml_t f = ezxml_child(cell, "f");
						if (f){						
							const char * formula = f->txt;
							if(formula)
								worksheet_write_formula(ws, CELL(r), formula, format);
						}
					} 
				}
				else {
					ezxml_t v = ezxml_child(cell, "v");
					if (v){
						const char * value = v->txt;
						if(value)
							worksheet_write_number(ws, CELL(r), atoi(value), format);
					}
				}
				
			}
		}
		
	}	
	
	lxw_workbook *workbook_new_from_template(const char *filename, const char *template){
		//create output file
		lxw_workbook *wb = workbook_new(filename);

		//open template
		//open zip file
		struct zip_t * zip 
				= zip_open(template, ZIP_DEFAULT_COMPRESSION_LEVEL, 'r');

		//load sst file
		ezxml_t sst;
		{
			zip_entry_open(zip, "xl/sharedStrings.xml");
			void *buf; size_t bufsize;
			zip_entry_read(zip, &buf, &bufsize);
			sst = ezxml_parse_str((char*)buf, bufsize);
		}

		//load styles file
		ezxml_t styles;
		{
			zip_entry_open(zip, "xl/styles.xml");
			void *buf; size_t bufsize;
			zip_entry_read(zip, &buf, &bufsize);
			styles = ezxml_parse_str((char*)buf, bufsize);
		}		

		/*load workbook file*/
		zip_entry_open(zip, "xl/workbook.xml");
		void *buf; size_t bufsize;
		zip_entry_read(zip, &buf, &bufsize);
		ezxml_t workbook = 
				ezxml_parse_str((char*)buf, bufsize);
		/*! TODO: fill workbookproperties
		 *  \todo fill workbookproperties	
		 */

		//get sheets
		ezxml_t sheets = ezxml_child(workbook, "sheets");
		if (sheets){
			ezxml_t sheet;
			int i = 0;
			for (sheet = ezxml_child(sheets, "sheet"); sheet; sheet = sheet->next, ++i){
				//realloc array
				const char * sheetname = ezxml_attr(sheet, "name");
				lxw_worksheet *ws = workbook_add_worksheet(wb, sheetname);
		
				char sheetfilepath[64];
				sprintf(sheetfilepath, "xl/worksheets/sheet%d.xml", i+1); 
				
				//load sheet xml file
				zip_entry_open(zip, sheetfilepath);
				void *buf; size_t bufsize;
				zip_entry_read(zip, &buf, &bufsize);
				ezxml_t sheet_xml = 
						ezxml_parse_str((char*)buf, bufsize);
				if (sheet_xml){
					parse_worksheet(wb, sheet_xml, ws, sst, styles);	
				}
			}
		}
			
		

		

		return wb;	
	}

#ifdef __cplusplus
}
#endif

#endif // _LXLSXWOT_H	
