/**
 * File              : libxlsxwriter_open_template.h
 * Author            : Igor V. Sementsov <ig.kuzm@gmail.com>
 * Date              : 30.10.2022
 * Last Modified Date: 30.10.2022
 * Last Modified By  : Igor V. Sementsov <ig.kuzm@gmail.com>
 */

#ifndef _LXLSXWOT_H
#define _LXLSXWOT_H

#include <stdlib.h>
#include <stdio.h>
#include "zip.h"
#include "ezxml.h"
#include "Libxlsxwriter/include/xlsxwriter.h"

#ifdef __cplusplus
extern "C" {
#endif

	lxw_workbook *workbook_new_from_template(const char *filename, const char *template);


	/*
	 * IMP
	 */
	
	void parse_worksheet(ezxml_t xml, lxw_worksheet *ws, ezxml_t sst) {
		/*! TODO: add sheet properties
		 *  \todo add sheet properties
		 */

		//parse row
		ezxml_t row = ezxml_child(xml, "row");
		int i = 0;
		for(; row; row = row->next, i++){
			/*! TODO: add row prop
			 *  \todo add row prop
			 */
			ezxml_t cell = ezxml_child(xml, "c");
			int c = 0;
			for(; cell; cell = cell->next, c++){
				//cell type
				const char * type = ezxml_attr(cell, "t");
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
					const char * string = (ezxml_get(cell, "is", 0, "t", 0, "")->txt);
					if(string)
						worksheet_write_string(ws, i, c, string, NULL);
				} 
				else 
				if (strcmp(type, "n") == 0){
					//number type
					const char * value = (ezxml_get(cell, "v", 0, ""))->txt;
					if(value)
						worksheet_write_number(ws, i, c, atoi(value), NULL);
				}
				else 
				if (strcmp(type, "s") == 0){
					//string sst
					const char * value = (ezxml_get(cell, "v", 0, ""))->txt;
					const char * string = ezxml_get(sst, "si", atoi(value), "t", 0, "");
					if(string)
						worksheet_write_string(ws, i, c, string, NULL);
				}
				else 
				if (strcmp(type, "str") == 0){
					//formula string
					const char * formula = (ezxml_get(cell, "f", 0, ""))->txt;
					if(formula)
						worksheet_write_formula(ws, i, c, formula, NULL);
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
					parse_worksheet(sheet_xml, ws, sst);	
				}
			}
		}
			
		

		

		return wb;	
	}

#ifdef __cplusplus
}
#endif

#endif // _LXLSXWOT_H	
