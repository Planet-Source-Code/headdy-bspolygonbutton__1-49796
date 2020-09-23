Attribute VB_Name = "modSystemMetrics"
Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Enum SystemMetrics
   smScreenWidth = 0 'X Size of screen
   smScreenHeight = 1 'Y Size of Screen
   smVScrollArrowWidth = 2 'X Size of arrow in vertical scroll bar.
   smHScrollArrowHeight = 3 'Y Size of arrow in horizontal scroll bar
   smWinCaptionHeight = 4 'Height of windows caption
   smNonSizeBorderWidth = 5 'Width of non-sizable borders
   smNonSizeBorderHeight = 6 'Height of non-sizable borders
   smDialogBorderWidth = 7 'Width of dialog box borders
   smDialogBorderHeight = 8 'Height of dialog box borders
   smScrollBoxHeight = 9 'Height of scroll box on horizontal scroll bar
   smScrollBoxWidth = 10 ' Width of scroll box on horizontal scroll bar
   smIconWidth = 11 'Width of standard icon
   smIconHeight = 12 'Height of standard icon
   smCursorWidth = 13 'Width of standard cursor
   smCursorHeight = 14 'Height of standard cursor
   smMenuHeight = 15 'Height of menu
   smMaximisedWidth = 16 'Width of client area of maximized window
   smMaximisedHeight = 17 'Height of client area of maximized window
   smKanjiWindowHeight = 18 'Height of Kanji window
   smMousePresent = 19 'True is a mouse is present
   smScrollBarWidth = 20 'Height of arrow in vertical scroll bar
   smScrollBarHeight = 21 'Width of arrow in vertical scroll bar
   smDebugMode = 22 'True if deugging version of windows is running
   smMouseButtonsSwapped = 23 'True if left and right buttons are swapped.
   smMinimumWindowWidth = 28 'Minimum width of window
   smMinimumWindowHeight = 29 'Minimum height of window
   smTitlebarBitmapWidth = 30 'Width of title bar bitmaps
   smTitlebarBitmapHeight = 31 'height of title bar bitmaps
   smMinimumTrackingWidth = 34 'Minimum tracking width of window (when resizing)
   smMinimumTrackingHeight = 35 'Minimum tracking height of window (when resizing)
   smDoubleClickWidth = 36 'double click width
   smDoubleClickHeight = 37 'double click height
   smIconSpacingWidth = 38 'width between desktop icons
   smIconSpacingHeight = 39 'height between desktop icons
   smMenuPopupAlignToRight = 40 'Zero if popup menus are aligned to the left of the memu bar item. True if it is aligned to the right.
   smPenWindowsHandle = 41 'The handle of the pen windows DLL if loaded.
   smDoubleCharactersEnabled = 42 'True if double byte characteds are enabled
   smMouseButtonCount = 43 'Number of mouse buttons.
   smSystemMetricsCount = 44 'Number of system metrics
   smSecurityPresent = 44 'True if security is present on windows 95 system
   smSmallCaptionHeight = 51 'height of windows 95 small caption
   smMenuButtonWidth = 54 'width of button on menu bar
   smMenuButtonHeight = 55 'height of button on menu bar
   smMinimisedWidth = 57 'width of rectangle into which minimised windows must fit.
   smMinimisedHeight = 58 'height of rectangle into which minimised windows must fit.
   smMaximumTrackingHeight = 59 'maximum width when resizing win95 windows
   smMsximumTrackingWidth = 60 'maximum width when resizing win95 windows
   smDefaultMaximisedWidth = 61 'default width of win95 maximised window
   smDefaultMaximisedHeight = 62 'default height of win95 maximised window
   smNetworkPresent = 63 'bit 0 is set if a network is present.
   smBootMode = 67 'Windows 95 boot mode. 0 = normal, 1 = safe, 2 = safe with network
   smMenuCheckmarkWidth = 71 'width of menu checkmark bitmap
   smMenuCheckmarkHeight = 72 'height of menu checkmark bitmap
   smSlowMachine = 73 'true if machine is too slow to run win95.
   smMidEastEnabled = 74 'Hebrew and Arabic enabled for windows 95
End Enum


Public Function SystemMetric(smOption As SystemMetrics) As Long
   SystemMetric = GetSystemMetrics(smOption)
End Function

