/**
 * File              : test.c
 * Author            : Igor V. Sementsov <ig.kuzm@gmail.com>
 * Date              : 30.10.2022
 * Last Modified Date: 30.10.2022
 * Last Modified By  : Igor V. Sementsov <ig.kuzm@gmail.com>
 */

#include "libxlsxwriter_open_template.h"


int main(int argc, char* argv[])
{
	lxw_workbook *wb =
			workbook_new_from_template(
			"/Users/kuzmich/Desktop/out.xlsx", 
			"/Users/kuzmich/1.xlsx");
	workbook_close(wb);
	
	return 0;
}
