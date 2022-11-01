# Module for libxlsxwriter to open template xlsx file

```c
lxw_workbook *workbook_new_from_template(const char *filename, const char *template);
```

![demo image](http://libxlsxwriter.github.io/demo.png)

## Header only file. 

To use - just add 
```c 
#include "libxlsxwriter_open_template.h"
```

### Example

```c
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

Uses [ezxml](https://github.com/lxfontes/ezxml) (MIT license),
[zip](https://github.com/kuba--/zip), miniz 
