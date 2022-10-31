# Module for libxlsxwriter to open template xlsx file

	lxw_workbook *workbook_new_from_template(const char *filename, const char *template);

Header only file. To use - just add `#include libxlsxwriter_open_template.h`

```

#include "libxlsxwriter_open_template.h"


int main(int argc, char* argv[])
{
	lxw_workbook *wb =
			workbook_new_from_template("out.xlsx", "template.xlsx");
	
	//do your magic
	
	workbook_close(wb);
	
	return 0;
}


```
