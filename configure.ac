AC_INIT([libxlsxwriter_open_template], [1.0])
AC_CONFIG_FILES([Makefile])


AC_CANONICAL_HOST
AC_CANONICAL_BUILD

dnl needed for CF_WITH_LIBTOOL
AC_CHECK_TOOL(AR, ar, ar)

AC_SUBST([CC])

# Detect the target system
case "${host_os}" in
    linux*)
        TARGETOS=Linux
        ;;
    cygwin*|mingw*)
        TARGETOS=Windows
        ;;
    darwin*)
        TARGETOS=MacOS
        ;;
    *)
        AC_MSG_ERROR(["OS $host_os is not supported"])
        ;;
esac

AC_SUBST([TARGETOS])

if test $host_os != $build_os
then
	CROSS_COMPILING=1
fi

AC_SUBST([CROSS_COMPILING])

#char ARCHS
AS_VAR_IF(ARCHS, [], AS_IF([test "$host_cpu" = "aarch64"], ARCHS=arm64, ARCHS=[$host_cpu]))
AC_SUBST([ARCHS])

AC_SUBST([CFLAGS])
AC_SUBST([LDFLAGS])

# Checks for programs.
AC_PROG_CC

AC_OUTPUT


########################################################################

echo ""
echo "        cXLSX	  $cXLSX_VERSION"
echo "        ======================"
echo ""
echo "        CC:               $CC"
echo "        HOST:             $host_os"
echo "        BUILD:            $build_os"
echo "        ARCHS:            $ARCHS"
echo "        TARGETOS:         $TARGETOS"
echo "        prefix:           $prefix"
echo ""
