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

	struct border_style{
		uint8_t style;
		char name[32];	
	};

	const struct border_style border_styles[] = {
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

	void parse_worksheet(lxw_workbook *wb, ezxml_t xml, lxw_worksheet *ws, ezxml_t sst, ezxml_t styles) {
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
		int k = 0;
		for(; col; col = col->next, k++){
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
				if (width){
					worksheet_set_column(ws, k, k, atof(width), NULL);
				}			
			}			
		}

		//parse row
		ezxml_t row = ezxml_child(data, "row");
		int i = 0;
		for(; row; row = row->next, i++){
			//get row spans
			const char * spans = ezxml_attr(row, "spans");
			if (spans){
			}

			const char * customHeight = ezxml_attr(row, "customHeight");
			if (customHeight && atoi(customHeight)){
				const char * ht = ezxml_attr(row, "ht");
				if (ht){
					worksheet_set_row(ws, i, atof(ht), NULL);
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
			int c = 0;
			for(; cell; cell = cell->next, c++){
				//cell format
				lxw_format *format = workbook_add_format(wb);
				
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
											//set font color
											format_set_fg_color(format, rgb_int);
											has_color = true;
										}
										const char * fgColor_indexed = ezxml_attr(fgColor, "indexed"); 
										const char * fgColor_theme = ezxml_attr(fgColor, "theme"); 
									}
									ezxml_t bgColor = ezxml_child(patternFill, "bgColor");
									if (bgColor){
										const char * rgb = ezxml_attr(bgColor, "rgb"); 
										if (rgb){
											//get hex from string
											int rgb_int = 0;
											sscanf(rgb, "%x", &rgb_int);
											//set font color
											format_set_bg_color(format, rgb_int);
											has_color = true;
										}
										
										const char * bgColor_indexed = ezxml_attr(bgColor, "indexed"); 
										const char * bgColor_theme = ezxml_attr(bgColor, "theme"); 
									}
									const char * patternType = ezxml_attr(patternFill, "patternType");
									if (patternType){
										uint8_t pattern = 0;
										if (strcmp(patternType, "solid") == 0 ) pattern = LXW_PATTERN_SOLID;
										else if (strcmp(patternType, "mediumGray") == 0 ) pattern = LXW_PATTERN_MEDIUM_GRAY;
										else if (strcmp(patternType, "darkGray") == 0 ) pattern = LXW_PATTERN_DARK_GRAY;
										else if (strcmp(patternType, "lightGray") == 0 ) pattern = LXW_PATTERN_LIGHT_GRAY;
										else if (strcmp(patternType, "darkHorizontal") == 0 ) pattern = LXW_PATTERN_DARK_HORIZONTAL;
										else if (strcmp(patternType, "darkVertical") == 0 ) pattern = LXW_PATTERN_DARK_VERTICAL;
										else if (strcmp(patternType, "darkDown") == 0 ) pattern = LXW_PATTERN_DARK_DOWN;
										else if (strcmp(patternType, "darkUp") == 0 ) pattern = LXW_PATTERN_DARK_UP;
										else if (strcmp(patternType, "darkGrid") == 0 ) pattern = LXW_PATTERN_DARK_GRID;
										else if (strcmp(patternType, "darkTrellis") == 0 ) pattern = LXW_PATTERN_DARK_TRELLIS;
										else if (strcmp(patternType, "lightHorizontal") == 0 ) pattern = LXW_PATTERN_LIGHT_HORIZONTAL;
										else if (strcmp(patternType, "lightVertical") == 0 ) pattern = LXW_PATTERN_LIGHT_VERTICAL;
										else if (strcmp(patternType, "lightDown") == 0 ) pattern = LXW_PATTERN_LIGHT_DOWN;
										else if (strcmp(patternType, "lightUp") == 0 ) pattern = LXW_PATTERN_LIGHT_UP;
										else if (strcmp(patternType, "lightGrid") == 0 ) pattern = LXW_PATTERN_LIGHT_GRID;									
										else if (strcmp(patternType, "lightTrellis") == 0 ) pattern = LXW_PATTERN_LIGHT_TRELLIS;
										else if (strcmp(patternType, "gray125") == 0 ) pattern = LXW_PATTERN_GRAY_125;
										else if (strcmp(patternType, "gray0625") == 0 ) pattern = LXW_PATTERN_GRAY_0625;
									
										if (has_color) //cant get color from theme and indexed yet
											format_set_pattern(format, pattern);
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
											int l;
											for (l=0; l<sizeof(border_styles)/sizeof(struct border_style); l++)
												if (strcmp(style, border_styles[l].name) == 0){
													format_set_left(format, border_styles[l].style);
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
											int l;
											for (l=0; l<sizeof(border_styles)/sizeof(struct border_style); l++)
												if (strcmp(style, border_styles[l].name) == 0){
													format_set_right(format, border_styles[l].style);
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
											int l;
											for (l=0; l<sizeof(border_styles)/sizeof(struct border_style); l++)
												if (strcmp(style, border_styles[l].name) == 0){
													format_set_top(format, border_styles[l].style);
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
											int l;
											for (l=0; l<sizeof(border_styles)/sizeof(struct border_style); l++)
												if (strcmp(style, border_styles[l].name) == 0){
													format_set_bottom(format, border_styles[l].style);											
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
									uint8_t alignment = 0;
									if (strcmp(horizontal, "left") == 0) alignment = LXW_ALIGN_LEFT;
									else if (strcmp(horizontal, "center") == 0) alignment = LXW_ALIGN_CENTER;
									else if (strcmp(horizontal, "right") == 0) alignment = LXW_ALIGN_RIGHT;
									else if (strcmp(horizontal, "fill") == 0) alignment = LXW_ALIGN_FILL;
									else if (strcmp(horizontal, "justify") == 0) alignment = LXW_ALIGN_JUSTIFY;
									else if (strcmp(horizontal, "center_across") == 0) alignment = LXW_ALIGN_CENTER_ACROSS;
									else if (strcmp(horizontal, "distributed") == 0) alignment = LXW_ALIGN_DISTRIBUTED;
									format_set_align(format, alignment);
								}
								const char * vertical = ezxml_attr(alignment, "vertical");
								if (vertical){
									uint8_t alignment = 0;
									if (strcmp(vertical, "top") == 0) alignment = LXW_ALIGN_VERTICAL_TOP;
									else if (strcmp(vertical, "bottom") == 0) alignment = LXW_ALIGN_VERTICAL_BOTTOM;
									else if (strcmp(vertical, "center") == 0) alignment = LXW_ALIGN_VERTICAL_CENTER;
									else if (strcmp(vertical, "justify") == 0) alignment = LXW_ALIGN_VERTICAL_JUSTIFY;
									else if (strcmp(vertical, "distributed") == 0) alignment = LXW_ALIGN_VERTICAL_DISTRIBUTED;
									format_set_align(format, alignment);
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
								worksheet_write_string(ws, i, c, string, format);
						}
					} 
					else 
					if (strcmp(type, "n") == 0){
						//number type
						ezxml_t v = ezxml_child(cell, "v");
						if (v){
							const char * value = v->txt;
							if(value)
								worksheet_write_number(ws, i, c, atoi(value), format);
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
									worksheet_write_string(ws, i, c, string, format);
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
								worksheet_write_formula(ws, i, c, formula, format);
						}
					} 
				}
				else {
					ezxml_t v = ezxml_child(cell, "v");
					if (v){
						const char * value = v->txt;
						if(value)
							worksheet_write_number(ws, i, c, atoi(value), format);
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
