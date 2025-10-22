// Program by Alexa Culley
// This program is built for The UPS Store to help manage various aspects of the business.
// This program will be used as my Capstone Internship Project for my Cybersecurity degree at Spokane Falls Community College. 
// 
// Fun Fact: Otiose means something/someone is "lazy" or "idle" and serves no practical purpose, which I thought was a funny/ironic name for a management app :)

#include "framework.h"
#include "Otiose.h"
#include <objidl.h>
#include <gdiplus.h>
#pragma comment(lib,"gdiplus.lib")
#include <string>
#include <fstream>
#include <ctime>
#include <sstream>
#include <iomanip>
#include <algorithm>

using namespace Gdiplus;
using std::min;

#define MAX_LOADSTRING 100

// Button and Control IDs
#define IDC_CLOCK_TOGGLE_BTN 1001
#define IDC_TAB_INVENTORY 1002
#define IDC_TAB_MAILBOX 1003
#define IDC_TAB_SCHEDULE 1004
#define IDC_TOGGLE_PANEL_BTN 1005

// Page enumeration
enum Page {
    PAGE_INVENTORY,
    PAGE_MAILBOX,
    PAGE_SCHEDULE
};

// Shift structure
struct Shift {
    std::string employeeName;
    std::string employeeID;
    time_t clockInTime;
    time_t clockOutTime;
    double hoursWorked;
    bool isActive;
};

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

// UI Elements
HWND hClockToggleBtn, hStatusText;
HWND hTabInventory, hTabMailbox, hTabSchedule;
HWND hTogglePanelBtn;
HWND hContentArea;

// Clock In/Out System Variables
Shift currentShift;
ULONG_PTR gdiplusToken; // For GDI+ initialization
HFONT hStatusFont; // Font for status text

// UI State
Page currentPage = PAGE_INVENTORY;
bool sidePanelOpen = true;
const int SIDE_PANEL_WIDTH = 250;
const int TAB_BAR_HEIGHT = 50;
const int TOGGLE_BTN_WIDTH = 30;

// Placeholder Employee Info
const std::string EMPLOYEE_NAME = "Alexa Culley"; // Placeholder, using me for now to track internship hours
const std::string EMPLOYEE_ID = "EMP001";         // Placeholder ID
const std::string TIMESHEET_FILE = "timesheet.csv";

// Forward declarations
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

// Clock In/Out functions 
void ClockIn();
void ClockOut();
void UpdateStatusDisplay(HWND hWnd);
void SaveShiftToCSV(const Shift& shift);
std::string TimeToString(time_t t);
void InitializeCSV();
void DrawColoredButton(LPDRAWITEMSTRUCT pDIS);

// UI functions
void SwitchPage(Page newPage);
void ToggleSidePanel(HWND hWnd);
void RepositionControls(HWND hWnd);
void DrawTabButton(LPDRAWITEMSTRUCT pDIS, bool isActive);
void DrawToggleButton(LPDRAWITEMSTRUCT pDIS);
void UpdateContentArea(HWND hWnd);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

	// Initialize GDI+
	GdiplusStartupInput gdiplusStartupInput;
	GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);

    // Initialize global strings
	LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadStringW(hInstance, IDC_OTIOSE, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

	// Initialize CSV file if it doesn't exist
	InitializeCSV();

    // Initialize current shift
    currentShift.employeeName = EMPLOYEE_NAME;
    currentShift.employeeID = EMPLOYEE_ID;
	currentShift.isActive = false;
	currentShift.clockInTime = 0;
	currentShift.clockOutTime = 0;
	currentShift.hoursWorked = 0.0;

	// Perform application initialization:
    if (!InitInstance(hInstance, nCmdShow))
    {
        GdiplusShutdown(gdiplusToken);
        return FALSE;
    }

	HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_OTIOSE));

    MSG msg;

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

	// Shutdown GDI+
	GdiplusShutdown(gdiplusToken);

    return (int) msg.wParam;
}

ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);
    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_OTIOSE));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
	wcex.hbrBackground = CreateSolidBrush(RGB(30, 30, 35)); // Dark grey/blue background
    wcex.lpszMenuName   = nullptr;
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // Store instance handle in our global variable

   HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, 1750, 1000, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   // Create Tab Buttons
   hTabInventory = CreateWindowW(
       L"BUTTON", L"Inventory",
       WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_OWNERDRAW,
       SIDE_PANEL_WIDTH + TOGGLE_BTN_WIDTH, 0, 200, TAB_BAR_HEIGHT,
	   hWnd, (HMENU)IDC_TAB_INVENTORY, hInstance, nullptr);

   hTabMailbox = CreateWindowW(
       L"BUTTON", L"Mailbox",
       WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_OWNERDRAW,
	   SIDE_PANEL_WIDTH + TOGGLE_BTN_WIDTH + 200, 0, 200, TAB_BAR_HEIGHT,
	   hWnd, (HMENU)IDC_TAB_MAILBOX, hInstance, nullptr);

   hTabSchedule = CreateWindowW(
       L"BUTTON", L"Schedule",
       WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_OWNERDRAW,
       SIDE_PANEL_WIDTH + TOGGLE_BTN_WIDTH + 400, 0, 200, TAB_BAR_HEIGHT,
	   hWnd, (HMENU)IDC_TAB_SCHEDULE, hInstance, nullptr);

   // Create Toggle Side Panel Button
   hTogglePanelBtn = CreateWindowW(
       L"BUTTON", L"\u2630",
       WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_OWNERDRAW,
       SIDE_PANEL_WIDTH, 0, TOGGLE_BTN_WIDTH, TAB_BAR_HEIGHT,
       hWnd, (HMENU)IDC_TOGGLE_PANEL_BTN, hInstance, nullptr);

   // Create Clock Toggle button in side panel
   hClockToggleBtn = CreateWindowW(
       L"BUTTON", L"Clock In",
       WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_OWNERDRAW,
       65, 80, 120, 120,
       hWnd, (HMENU)IDC_CLOCK_TOGGLE_BTN, hInstance, nullptr);

   // Create Status text display in side panel
   hStatusText = CreateWindowW(
       L"STATIC", L"Status: Currently Clocked Out", 
       WS_VISIBLE | WS_CHILD | SS_LEFT,
       10, 220, 230, 250,
	   hWnd, nullptr, hInstance, nullptr);

   // Create a font for the status text
   hStatusFont = CreateFontW(
       14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
       DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
       CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Segoe UI"
   );
   SendMessage(hStatusText, WM_SETFONT, (WPARAM)hStatusFont, TRUE);

   // Create main content area
   hContentArea = CreateWindowW(
       L"STATIC", L"",
       WS_VISIBLE | WS_CHILD | SS_LEFT,
       SIDE_PANEL_WIDTH + TOGGLE_BTN_WIDTH, TAB_BAR_HEIGHT,
       1750 - (SIDE_PANEL_WIDTH + TOGGLE_BTN_WIDTH), 1000 - TAB_BAR_HEIGHT,
	   hWnd, nullptr, hInstance, nullptr);

   UpdateContentArea(hWnd);

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_COMMAND:
        {
            int wmId = LOWORD(wParam);
            // Parse the menu selections:
            switch (wmId)
            {
            case IDC_CLOCK_TOGGLE_BTN:
                if (currentShift.isActive)
                {
                    // Currently clocked in, so clock out
                    ClockOut();
					UpdateStatusDisplay(hWnd);
                    SetWindowTextW(hClockToggleBtn, L"Clock In");
					InvalidateRect(hClockToggleBtn, NULL, TRUE); // Force button redraw
                }
                else
                {
                    // Currently clocked out, so clock in
                    ClockIn();
					UpdateStatusDisplay(hWnd);
                    SetWindowTextW(hClockToggleBtn, L"Clock Out");
                    InvalidateRect(hClockToggleBtn, NULL, TRUE); // Force button redraw
                }
				break;

            case IDC_TAB_INVENTORY:
                SwitchPage(PAGE_INVENTORY);
				break;

            case IDC_TAB_MAILBOX:
				SwitchPage(PAGE_MAILBOX);
				break;

			case IDC_TAB_SCHEDULE:
				SwitchPage(PAGE_SCHEDULE);
				break;

            case IDC_TOGGLE_PANEL_BTN:
				ToggleSidePanel(hWnd);
				break;

            case IDM_ABOUT:
				// Removed About - keeping in case of future use
                // DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;

            case IDM_EXIT:
				// Removed Menu - keeping in case of future use
                // DestroyWindow(hWnd);
                break;

            default:
                return DefWindowProc(hWnd, message, wParam, lParam);
            }
        }
        break;

    case WM_DRAWITEM:
    {
        LPDRAWITEMSTRUCT pDIS = (LPDRAWITEMSTRUCT)lParam;

        if (pDIS->CtlID == IDC_CLOCK_TOGGLE_BTN)
        {
            DrawColoredButton(pDIS);
            return TRUE;
        }
        else if (pDIS->CtlID == IDC_TAB_INVENTORY)
        {
            DrawTabButton(pDIS, currentPage == PAGE_INVENTORY);
            return TRUE;
        }
        else if (pDIS->CtlID == IDC_TAB_MAILBOX)
        {
            DrawTabButton(pDIS, currentPage == PAGE_MAILBOX);
            return TRUE;
        }
        else if (pDIS->CtlID == IDC_TAB_SCHEDULE)
        {
            DrawTabButton(pDIS, currentPage == PAGE_SCHEDULE);
            return TRUE;
        }
        else if (pDIS->CtlID == IDC_TOGGLE_PANEL_BTN)
        {
            DrawToggleButton(pDIS);
            return TRUE;
		}
    }
	break;

    case WM_CTLCOLORSTATIC:
    {
        HDC hdcStatic = (HDC)wParam;
        SetTextColor(hdcStatic, RGB(255, 255, 255)); // White text
        SetBkColor(hdcStatic, RGB(30, 30, 35)); // Match window background
        
		static HBRUSH hbrBkgnd = CreateSolidBrush(RGB(30, 30, 35));
		return (LRESULT)hbrBkgnd;
    }

    case WM_SIZE:
        RepositionControls(hWnd);
		break;

    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);

			// Draw side panel background
            if (sidePanelOpen)
            {
                RECT panelRect = { 0, 0, SIDE_PANEL_WIDTH, 10000 };
                HBRUSH panelBrush = CreateSolidBrush(RGB(40, 40, 45)); 
                FillRect(hdc, &panelRect, panelBrush);
                DeleteObject(panelBrush);
			}

            EndPaint(hWnd, &ps);
        }
        break;

    case WM_DESTROY:
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

// Clock In/Out Function Implementations

void ClockIn()
{
    currentShift.clockInTime = time(nullptr);
	currentShift.isActive = true;
    currentShift.clockOutTime = 0;
	currentShift.hoursWorked = 0.0;
}

void ClockOut()
{
    currentShift.clockOutTime = time(nullptr);
    currentShift.isActive = false;

    // Calculate hours worked
    double secondsWorked = difftime(currentShift.clockOutTime, currentShift.clockInTime);
    currentShift.hoursWorked = secondsWorked / 3600.0; // convert to hours

    // Save shift to CSV
    SaveShiftToCSV(currentShift);
}

void UpdateStatusDisplay(HWND hWnd)
{
    std::wstringstream status;

    if (currentShift.isActive)
    {
        status << L"Status: Clocked In\n\n";
        status << L"Employee:\n" << EMPLOYEE_NAME.c_str() << L"\n\n";
        status << L"Clocked In at: " << TimeToString(currentShift.clockInTime).c_str() << L"\n";
    }
    else if (currentShift.clockOutTime > 0)
    {
		status << L"Status: Clocked Out\n\n";
        status << L"Employee:\n" << EMPLOYEE_NAME.c_str() << L"\n\n";
        status << L"Last Shift:\n" << std::fixed << std::setprecision(2)
            << currentShift.hoursWorked << L" hours\n";
    }
    else
    {
		status << L"Status: Clocked Out\n\n";
		status << L"Employee:\n" << EMPLOYEE_NAME.c_str() << L"\n\n";
		status << L"Ready to clock in!";
    }

    SetWindowTextW(hStatusText, status.str().c_str());
}

void SaveShiftToCSV(const Shift& shift)
{
    std::ofstream file(TIMESHEET_FILE, std::ios::app);

    if (file.is_open())
    {
        file << shift.employeeName << ","
             << shift.employeeID << ","
             << TimeToString(shift.clockInTime) << ","
             << TimeToString(shift.clockOutTime) << ","
             << std::fixed << std::setprecision(2) << shift.hoursWorked << "\n";
        file.close();
    }
}

std::string TimeToString(time_t t)
{
    struct tm timeinfo;
    localtime_s(&timeinfo, &t);

    std::stringstream ss;
    ss << std::put_time(&timeinfo, "%Y-%m-%d %H:%M:%S");
    return ss.str();
}

void InitializeCSV()
{
    std::ifstream checkFile(TIMESHEET_FILE);

    if (!checkFile.good())
    {
		// File doesn't exist, create and write header
        std::ofstream file(TIMESHEET_FILE);
        if (file.is_open())
        {
            file << "Employee Name,Employee ID,Clock In Time,Clock Out Time,Hours Worked\n";
            file.close();
        }
    }
    checkFile.close();
}

void DrawColoredButton(LPDRAWITEMSTRUCT pDIS)
{
    HDC hdc = pDIS->hDC;
    RECT rect = pDIS->rcItem;

	// Clear the background first to prevent any artifacts
    HBRUSH hBackBrush = CreateSolidBrush(RGB(40, 40, 45));
	FillRect(hdc, &rect, hBackBrush);
    DeleteObject(hBackBrush);

	// Create GDI+ Graphics object for smoother rendering
	Graphics graphics(hdc);
    graphics.SetSmoothingMode(SmoothingModeAntiAlias);
    graphics.SetCompositingQuality(CompositingQualityHighQuality); // Better quality rendering
    graphics.SetPixelOffsetMode(PixelOffsetModeHighQuality); // smoother pixel rendering

	// Determine button color based on current state
    Color buttonColor;
    Color borderColor;

    if (currentShift.isActive)
    {
        // Clocked In - show red clock out button
        buttonColor = Color(255, 220, 53, 69); // Crimson Red (A, R, G, B)
		borderColor = Color(255, 180, 43, 59); // Darker Red for border
    }
    else
    {
        // Clocked Out - show blue clock in button
        buttonColor = Color(255, 13, 110, 253); // Nice Blue color
		borderColor = Color(255, 10, 88, 202); // Darker Blue for border
    }

    // Check if button is being pressed
    if (pDIS->itemState & ODS_SELECTED)
    {
        // Darken color when pressed
        BYTE r = (BYTE)(buttonColor.GetR() * 0.8);
        BYTE g = (BYTE)(buttonColor.GetG() * 0.8);
        BYTE b = (BYTE)(buttonColor.GetB() * 0.8);
        buttonColor = Color(255, r, g, b);
	}

    // Calculate center and radius for the circle
	int width = rect.right - rect.left;
	int height = rect.bottom - rect.top;
    int diameter = min(width, height);
	int centerX = rect.left + width / 2;
	int centerY = rect.top + height / 2;
	int radius = diameter / 2 - 5; // Padding of 5 pixels

    // Draw the filled circle
    SolidBrush brush(buttonColor);
    graphics.FillEllipse(&brush,
        centerX - radius,
		centerY - radius, 
		diameter - 10, // preserving the padding with 10 pixels
		diameter - 10);

    // Draw the border circle
    Pen borderPen(borderColor, 3.0f);
    graphics.DrawEllipse(&borderPen,
		centerX - radius,
		centerY - radius,
		diameter - 10,
		diameter - 10);

    // Draw the button text
    WCHAR buttonText[32];
    GetWindowTextW(pDIS->hwndItem, buttonText, 32);

    // Set up text rendering
	FontFamily fontFamily(L"Arial");
	Font font(&fontFamily, 16, FontStyleBold, UnitPixel);
	SolidBrush textBrush(Color(255, 255, 255, 255)); // White text

    // Center the text
	StringFormat stringFormat;
	stringFormat.SetAlignment(StringAlignmentCenter);
	stringFormat.SetLineAlignment(StringAlignmentCenter);

    RectF layoutRect(
        (REAL)(centerX - radius),
        (REAL)(centerY - radius),
        (REAL)(diameter - 10),
        (REAL)(diameter - 10)
    );

    // enable text anti-aliasing
	graphics.SetTextRenderingHint(TextRenderingHintAntiAlias);
    graphics.DrawString(buttonText, -1, &font, layoutRect, &stringFormat, &textBrush);
}

// UI Implementation

void SwitchPage(Page newPage)
{
    currentPage = newPage;

    // Force redraw of all tab buttons
    InvalidateRect(hTabInventory, NULL, TRUE);
	InvalidateRect(hTabMailbox, NULL, TRUE);
    InvalidateRect(hTabSchedule, NULL, TRUE);
    
    // Update content area
	UpdateContentArea(GetParent(hTabInventory));
}

void ToggleSidePanel(HWND hWnd)
{
    sidePanelOpen = !sidePanelOpen;
    RepositionControls(hWnd);
    InvalidateRect(hWnd, NULL, TRUE); // Redraw window to update side panel
}

void RepositionControls(HWND hWnd)
{
    RECT clientRect;
    GetClientRect(hWnd, &clientRect);

    int panelOffset = sidePanelOpen ? SIDE_PANEL_WIDTH : 0;

    // Reposition toggle button
    SetWindowPos(hTogglePanelBtn, NULL,
        panelOffset, 0, TOGGLE_BTN_WIDTH, TAB_BAR_HEIGHT, SWP_NOZORDER);

    // Reposition Tab Buttons
    int tabWidth = 200;
    SetWindowPos(hTabInventory, NULL,
        panelOffset + TOGGLE_BTN_WIDTH, 0, tabWidth, TAB_BAR_HEIGHT, SWP_NOZORDER);
    SetWindowPos(hTabMailbox, NULL,
        panelOffset + TOGGLE_BTN_WIDTH + tabWidth, 0, tabWidth, TAB_BAR_HEIGHT, SWP_NOZORDER);
    SetWindowPos(hTabSchedule, NULL,
        panelOffset + TOGGLE_BTN_WIDTH + tabWidth * 2, 0, tabWidth, TAB_BAR_HEIGHT, SWP_NOZORDER);
    
    // Reposition Content Area
	int contentX = panelOffset + TOGGLE_BTN_WIDTH;
	int contentWidth = clientRect.right - contentX;
	int contentHeight = clientRect.bottom - TAB_BAR_HEIGHT;
    SetWindowPos(hContentArea, NULL,
		contentX, TAB_BAR_HEIGHT, contentWidth, contentHeight, SWP_NOZORDER);

    // Show/hide side panel elements
    if (sidePanelOpen)
    {
        // Position elements within the side panel
        SetWindowPos(hClockToggleBtn, NULL, 65, 80, 120, 120, SWP_NOZORDER | SWP_SHOWWINDOW);
		SetWindowPos(hStatusText, NULL, 10, 220, 230, 250, SWP_NOZORDER | SWP_SHOWWINDOW);
    }
    else
    {
        // Move them off-screen when hidden
		ShowWindow(hClockToggleBtn, SW_HIDE);
		ShowWindow(hStatusText, SW_HIDE);
    }

    // Force a full window redraw to update teh side panel background
	InvalidateRect(hWnd, NULL, TRUE);
	UpdateWindow(hWnd);
}

void DrawTabButton(LPDRAWITEMSTRUCT pDIS, bool isActive)
{
    HDC hdc = pDIS->hDC;
    RECT rect = pDIS->rcItem;

	Graphics graphics(hdc);
	graphics.SetSmoothingMode(SmoothingModeAntiAlias);

	Color bgColor = isActive ? Color(255, 50, 50, 55) : Color(255, 35, 35, 40);
	Color textColor = isActive ? Color(255, 100, 180, 255) : Color(255, 200, 200, 200);

    Rect gdipRect(rect.left,
        rect.top,
        rect.right - rect.left,
        rect.bottom - rect.top);
	SolidBrush bgBrush(bgColor);
	graphics.FillRectangle(&bgBrush, gdipRect);

    if (isActive)
    {
        Pen accentPen(Color(255, 13, 110, 253), 3.0f);
        graphics.DrawLine(&accentPen, (INT)rect.left, (INT)rect.bottom - 2, (INT)rect.right, (INT)rect.bottom - 2);
    }

	WCHAR buttonText[32];
	GetWindowTextW(pDIS->hwndItem, buttonText, 32);

	FontFamily fontFamily(L"Segoe UI");
	Font font(&fontFamily, 14, isActive ? FontStyleBold : FontStyleRegular, UnitPixel);
    SolidBrush textBrush(textColor);

	StringFormat stringFormat;
	stringFormat.SetAlignment(StringAlignmentCenter);
    stringFormat.SetLineAlignment(StringAlignmentCenter);
    
    RectF layoutRect(
        (REAL)rect.left,
        (REAL)rect.top,
        (REAL)(rect.right - rect.left),
        (REAL)(rect.bottom - rect.top)
	);

	graphics.SetTextRenderingHint(TextRenderingHintAntiAlias);
	graphics.DrawString(buttonText, -1, &font, layoutRect, &stringFormat, &textBrush);
}

void DrawToggleButton(LPDRAWITEMSTRUCT pDIS)
{
    HDC hdc = pDIS->hDC;
    RECT rect = pDIS->rcItem;

    Graphics graphics(hdc);
    graphics.SetSmoothingMode(SmoothingModeAntiAlias);

    Color bgColor = Color(255, 35, 35, 40);

    Rect gdipRect(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top);
    SolidBrush bgBrush(bgColor);
    graphics.FillRectangle(&bgBrush, gdipRect);

    WCHAR buttonText[32];
	GetWindowTextW(pDIS->hwndItem, buttonText, 32);

    FontFamily fontFamily(L"Segoe UI");
    Font font(&fontFamily, 20, FontStyleBold, UnitPixel);
    SolidBrush textBrush(Color(255, 200, 200, 200));

    StringFormat stringFormat;
    stringFormat.SetAlignment(StringAlignmentCenter);
    stringFormat.SetLineAlignment(StringAlignmentCenter);

    RectF layoutRect(
        (REAL)rect.left,
        (REAL)rect.top,
        (REAL)(rect.right - rect.left),
        (REAL)(rect.bottom - rect.top)
    );

    graphics.SetTextRenderingHint(TextRenderingHintAntiAlias);
    graphics.DrawString(buttonText, -1, &font, layoutRect, &stringFormat, &textBrush);
}

void UpdateContentArea(HWND hWnd)
{
    std::wstringstream content;

    switch (currentPage)
    {
    case PAGE_INVENTORY:
        content << L"INVENTORY MANAGEMENT\n\n";
        content << L"This is where inventory tracking will go.\n";
        content << L"Track boxes, supplies, and retail items.";
        break;
    case PAGE_MAILBOX:
        content << L"MAILBOX MANAGEMENT\n\n";
		content << L"This is where mailbox management will go.\n";
        content << L"Track customer mailboxes and their details.";
        break;
    case PAGE_SCHEDULE:
        content << L"SCHEDULE MANAGEMENT\n\n";
		content << L"This is where schedule management will go.\n";
		content << L"Create and manage employee schedules.";
        break;
    }

    SetWindowTextW(hContentArea, content.str().c_str());
}