/* Win95/VC6-CRT compatibility shims for the rust9x pancurses build.
 *
 * zig's bundled mingw-w64 headers are UCRT-flavored, but the vintage rust9x
 * link uses the VC6 (MSVC 6.0) static CRT (see .cargo/config.toml). The
 * PDCurses objects end up referencing a handful of UCRT-only symbols that
 * the VC6 CRT does not provide. This file supplies them; it deliberately
 * includes no CRT headers (the UCRT-flavored declarations would conflict
 * with these definitions).
 *
 * Compiled only for rust9x msvc targets (see build.rs).
 */

/* VC6's FILE array (`_iob` in C source; decorated `__iob`), 3 entries,
 * 32 bytes per FILE (4 pointers + 4 ints). Declared as a byte array since
 * we cannot include VC6's stdio.h here. */
extern unsigned char _iob[];

/* mingw's UCRT-mode <stdio.h> fetches stdout/stderr through
 * __acrt_iob_func(n) (declared dllimport, so the call goes through the
 * __imp_ import pointer). Same meaning as &VC6's _iob[n].
 * The .long data symbol must be decorated exactly:
 *   function  __acrt_iob_func  -> ___acrt_iob_func
 *   imp ptr                    -> __imp____acrt_iob_func
 */
void *__cdecl __acrt_iob_func(unsigned n) {
    return (void *)(_iob + n * 32);
}
__asm__(
    ".globl __imp____acrt_iob_func\n"
    "__imp____acrt_iob_func:\n"
    "  .long ___acrt_iob_func\n"
);

/* mingw's <time.h> maps time/localtime to their 64-bit-time variants
 * (_time64/_localtime64). VC6's CRT has neither; forward to its 32-bit
 * time/localtime (C names `time`/`localtime`; fine for the Win9x era).
 * Return types are opaque pointers so we need no struct definitions here;
 * PDCurses sees mingw's `struct tm *`, and the pointer itself is produced
 * by VC6's own localtime. */
extern long __cdecl time(long *t);
extern void *__cdecl localtime(const long *t);

long long __cdecl _time64(long long *t) {
    long v = time(0);
    if (t) {
        *t = v;
    }
    return v;
}

void *__cdecl _localtime64(const long long *t64) {
    long v = (long)*t64;
    return localtime(&v);
}

/* C99 vsnprintf: VC6's CRT predates it and only has _vsnprintf, which does
 * not NUL-terminate on truncation. Wrap it with C99 semantics (printw/
 * vwprintw in PDCurses use vsnprintf directly). */
extern int __cdecl _vsnprintf(char *buf, unsigned n, const char *fmt, char *ap);

int __cdecl vsnprintf(char *buf, unsigned n, const char *fmt, char *ap) {
    int r = _vsnprintf(buf, n, fmt, ap);
    if (n != 0 && (r < 0 || (unsigned)r >= n)) {
        buf[n - 1] = '\0';
    }
    return r;
}
