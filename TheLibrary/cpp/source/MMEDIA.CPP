#include "the.h"

// WIN ONLY
int TheVideo::setResolution(int width,int height, int colorBit) {
	//sprintf(tmpStr,"FullScreen proposed res::x=%d y=%d bit=%d",nWidth,nHeight,nColorBits);
	//Log::print(tmpStr);

	DEVMODE dm ;
	BOOL found = false ;

	// Check only width and height
	if (colorBit==-1) {
		dm.dmFields=DM_PELSWIDTH|DM_PELSHEIGHT;
		for (int i=0; EnumDisplaySettings(0,i,&dm); ++i) {
			if (
				dm.dmPelsWidth==width && dm.dmPelsHeight==height) {
				found = true ;
				break ;
			}
		}
	}
	// width and height and colorDepth
	else {
		dm.dmFields=DM_BITSPERPEL|DM_PELSWIDTH|DM_PELSHEIGHT;	
		for (int i=0; EnumDisplaySettings(0,i,&dm); ++i) {
			if (dm.dmBitsPerPel == colorBit &&
				dm.dmPelsWidth==width && dm.dmPelsHeight==height) {
				found = true ;
				break ;
			}
		}

	}
	// If we couldn't find the appropriate mode, give up
	if (!found)	{
		return false ;
	}
	// We found it... then change into it!
	// But, right before doing it, record current size and position of
	// top-level windows, so we can restore it later (the resolution change
	// causes windows to get resized).
	//RememberTopWindows(hMyWnd) ;
	// Now, finally change video mode

	LONG status = ChangeDisplaySettings( &dm, 0 );
	RegFlushKey( HKEY_LOCAL_MACHINE );
	return (DISP_CHANGE_SUCCESSFUL == status );

}
