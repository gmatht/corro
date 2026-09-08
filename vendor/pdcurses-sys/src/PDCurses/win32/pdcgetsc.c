/* Public Domain Curses */

#include "pdcwin.h"

/* Windows 95's GetConsoleScreenBufferInfo is unreliable: srWindow is left
   uninitialized (values such as -4964), dwSize.Y comes back garbage (e.g.
   132) and dwMaximumWindowSize even worse (Y = 1). Only trust it on NT.
   On 9x hardcode the DOS box default 80x25, which matches the visible
   window and is what PDCurses creates its own screen buffer for anyway. */

int _pdc_is_win9x(void)
{
    return (GetVersion() & 0x80000000) != 0;
}

static void _pdc_safe_csbi(CONSOLE_SCREEN_BUFFER_INFO *scr)
{
    GetConsoleScreenBufferInfo(pdc_con_out, scr);

    if (_pdc_is_win9x())
    {
        scr->dwSize.X = 80;
        scr->dwSize.Y = 25;
        scr->srWindow.Left = 0;
        scr->srWindow.Top = 0;
        scr->srWindow.Right = 79;
        scr->srWindow.Bottom = 24;
        return;
    }

    if (scr->srWindow.Right < scr->srWindow.Left ||
        scr->srWindow.Bottom < scr->srWindow.Top ||
        scr->srWindow.Right - scr->srWindow.Left + 1 < 2 ||
        scr->srWindow.Right - scr->srWindow.Left + 1 > scr->dwSize.X ||
        scr->srWindow.Bottom - scr->srWindow.Top + 1 > scr->dwSize.Y)
    {
        scr->srWindow.Left = 0;
        scr->srWindow.Top = 0;
        if (scr->dwSize.X >= 2)
            scr->srWindow.Right = scr->dwSize.X - 1;
        if (scr->dwSize.Y >= 2)
            scr->srWindow.Bottom = scr->dwSize.Y - 1;
    }
}

/* get the cursor size/shape */

int PDC_get_cursor_mode(void)
{
    CONSOLE_CURSOR_INFO ci;
    
    PDC_LOG(("PDC_get_cursor_mode() - called\n"));

    GetConsoleCursorInfo(pdc_con_out, &ci);

    return ci.dwSize;
}

/* return number of screen rows */

int PDC_get_rows(void)
{
    CONSOLE_SCREEN_BUFFER_INFO scr;

    PDC_LOG(("PDC_get_rows() - called\n"));

    _pdc_safe_csbi(&scr);

    return scr.srWindow.Bottom - scr.srWindow.Top + 1;
}

/* return number of buffer rows */

int PDC_get_buffer_rows(void)
{
    CONSOLE_SCREEN_BUFFER_INFO scr;

    PDC_LOG(("PDC_get_buffer_rows() - called\n"));

    if (_pdc_is_win9x())
        return 25;

    GetConsoleScreenBufferInfo(pdc_con_out, &scr);

    return scr.dwSize.Y;
}

/* return width of screen/viewport */

int PDC_get_columns(void)
{
    CONSOLE_SCREEN_BUFFER_INFO scr;

    PDC_LOG(("PDC_get_columns() - called\n"));

    _pdc_safe_csbi(&scr);

    return scr.srWindow.Right - scr.srWindow.Left + 1;
}
