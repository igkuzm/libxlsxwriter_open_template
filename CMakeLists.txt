cmake_minimum_required(VERSION 3.0)

set(TARGET cXLSX)

project(
	${TARGET} VERSION 1.0 
	DESCRIPTION "C library to handle with XLSX"
	HOMEPAGE_URL ""
	LANGUAGES C 
)

add_library(${TARGET} STATIC 
	cXLSX.c 
	ezxml.c 
	zip.c
)
target_link_libraries(${TARGET})

if(${CYD_BUILD_TEST})
	add_executable(${TARGET}_test test.c)
	target_link_libraries(${TARGET}_test ${TARGET})
endif()
