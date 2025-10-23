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
const std::string EMPLOYEE_NAME = "Toastiiieee"; // Placeholder, using me for now to track internship hours
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
void SwitchPage(Page newPage, HWND hWnd);
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
                SwitchPage(PAGE_INVENTORY, hWnd);
                break;

            case IDC_TAB_MAILBOX:
                SwitchPage(PAGE_MAILBOX, hWnd);
                break;

            case IDC_TAB_SCHEDULE:
                SwitchPage(PAGE_SCHEDULE, hWnd);
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
		// ctlID is a variable member of the DRAWITEMSTRUCT structure. pDIS is a pointer to a DRAWITEMSTRUCT structure.
		// thus, pDIS->CtlID gives us the control ID of the button being drawn. It's similar to using object.member in other languages; 
        // the difference here is that you're working with the pointer to the object, rather than the object itself.
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

    case WM_CTLCOLORBTN:
    {
        HDC hdcButton = (HDC)wParam;
        SetTextColor(hdcButton, RGB(30, 30, 35));
        SetBkMode(hdcButton, TRANSPARENT); // Make text background transparent
        return (LRESULT)GetStockObject(NULL_BRUSH); // Prevents solid background fill
    }

    case WM_CTLCOLORSTATIC:
    {
        HDC hdcStatic = (HDC)wParam;
        SetTextColor(hdcStatic, RGB(255, 255, 255)); // White text
        SetBkColor(hdcStatic, RGB(30, 30, 35)); // Match window background
        
		static HBRUSH hbrBkgnd = CreateSolidBrush(RGB(30, 30, 35));
		return (LRESULT)hbrBkgnd;
    }
    break;

    case WM_SIZE:
        RepositionControls(hWnd);
		break;

    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);

            RECT clientRect;
            GetClientRect(hWnd, &clientRect);

            // Draw lighter background behind the tab bar
            RECT tabBarBg = { 0, 0, clientRect.right, TAB_BAR_HEIGHT };
            HBRUSH tabBarBrush = CreateSolidBrush(RGB(45, 52, 60)); // Lighter than main window
            FillRect(hdc, &tabBarBg, tabBarBrush);
            DeleteObject(tabBarBrush);

            // Draw border between tab bar and content area, with gap for active tab
            int panelOffset = sidePanelOpen ? SIDE_PANEL_WIDTH : 0;
            int tabWidth = 200;
            int tabOverlap = 15;
            int totalTabWidth = tabWidth * 3;
            int availableWidth = clientRect.right;
            int startX = (availableWidth - totalTabWidth) / 2;

            if (startX < panelOffset + TOGGLE_BTN_WIDTH)
                startX = panelOffset + TOGGLE_BTN_WIDTH;

            // Calculate active tab position based on current page
            int activeTabLeft = 0;
            int activeTabRight = 0;

            if (currentPage == PAGE_INVENTORY)
            {
                activeTabLeft = startX;
                activeTabRight = startX + tabWidth;
            }
            else if (currentPage == PAGE_MAILBOX)
            {
                activeTabLeft = startX + tabWidth - tabOverlap;
                activeTabRight = startX + (tabWidth * 2) - tabOverlap;
            }
            else // PAGE_SCHEDULE
            {
                activeTabLeft = startX + (tabWidth * 2) - (tabOverlap * 2);
                activeTabRight = startX + (tabWidth * 3) - (tabOverlap * 2);
            }

            // Create GDI+ graphics for smooth line drawing
            Graphics graphics(hdc);
            graphics.SetSmoothingMode(SmoothingModeAntiAlias);

            Pen borderPen(Color(255, 60, 90, 95), 1.5f);

            // Draw border line from left edge to active tab (should go under inactive tabs)
            if (activeTabLeft > 0)
            {
                graphics.DrawLine(&borderPen, 0, TAB_BAR_HEIGHT - 1, activeTabLeft, TAB_BAR_HEIGHT - 1);
            }

            // Draw border line from active tab to right edge (should go under inactive tabs)
            if (activeTabRight < clientRect.right)
            {
                graphics.DrawLine(&borderPen, activeTabRight, TAB_BAR_HEIGHT - 1, clientRect.right, TAB_BAR_HEIGHT - 1);
            }

            // Draw side panel background
            if (sidePanelOpen)
            {
                RECT panelRect = { 0, 0, SIDE_PANEL_WIDTH, 10000 };
                HBRUSH panelBrush = CreateSolidBrush(RGB(40, 40, 45));
                FillRect(hdc, &panelRect, panelBrush);
                DeleteObject(panelBrush);
                // Draw line separating the side panel + toggle bar column from the main content window
                graphics.DrawLine(&borderPen, (REAL)(SIDE_PANEL_WIDTH + TOGGLE_BTN_WIDTH - 1), 0.0, (REAL)(SIDE_PANEL_WIDTH + TOGGLE_BTN_WIDTH - 1), (REAL)(clientRect.bottom));
            }
            else
            {
                // Draw line separating the side panel + toggle bar column from the main content window
                graphics.DrawLine(&borderPen, (REAL)(TOGGLE_BTN_WIDTH - 1), 0.0, (REAL)(TOGGLE_BTN_WIDTH - 1), (REAL)(clientRect.bottom));
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

void SwitchPage(Page newPage, HWND hWnd)
{
    currentPage = newPage;

    // Force redraw of all tab buttons
    InvalidateRect(hTabInventory, NULL, TRUE);
    InvalidateRect(hTabMailbox, NULL, TRUE);
    InvalidateRect(hTabSchedule, NULL, TRUE);

    // Update content area
    UpdateContentArea(GetParent(hContentArea));

    // Update layering of tab buttons
    RepositionControls(hWnd);

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
    SetWindowPos(hTogglePanelBtn, NULL, panelOffset, 0, TOGGLE_BTN_WIDTH, TAB_BAR_HEIGHT, SWP_NOZORDER);

    // Calculate centered position for tabs
	int tabWidth = 200;
    int tabOverlap = 15; // Pixels of overlap between tabs
	int totalTabWidth = tabWidth * 3;
    int availableWidth = clientRect.right;// - panelOffset - TOGGLE_BTN_WIDTH;
    int startX = (availableWidth - totalTabWidth) / 2; //panelOffset + TOGGLE_BTN_WIDTH + (availableWidth - totalTabWidth) / 2;

	// Ensure tabs don't go off-screen if window is too small
	if (startX < panelOffset + TOGGLE_BTN_WIDTH)
		startX = panelOffset + TOGGLE_BTN_WIDTH;

    // Reposition Tab Buttons
	// IMPORTANT: draw the active tab FIRST so it appears on top. This makes no sense and cost 3 hours of troubleshooting, but whatever. Fuck me I guess.
    if (currentPage == PAGE_INVENTORY)
    {
        // Position active tab first so it takes top priority
        SetWindowPos(hTabInventory, HWND_TOP, startX, 0, tabWidth, TAB_BAR_HEIGHT, SWP_SHOWWINDOW);
        // Position inactive tabs last
        SetWindowPos(hTabMailbox, NULL, startX + tabWidth - tabOverlap, 0, tabWidth, TAB_BAR_HEIGHT, SWP_NOACTIVATE);
        SetWindowPos(hTabSchedule, NULL, startX + (tabWidth * 2) - (tabOverlap * 2), 0, tabWidth, TAB_BAR_HEIGHT, SWP_NOACTIVATE);
    }
    else if (currentPage == PAGE_MAILBOX)
    {
        // Position active tab first so it takes top priority
        SetWindowPos(hTabMailbox, HWND_TOP, startX + tabWidth - tabOverlap, 0, tabWidth, TAB_BAR_HEIGHT, SWP_SHOWWINDOW);
        // Position inactive tabs last
        SetWindowPos(hTabInventory, NULL, startX, 0, tabWidth, TAB_BAR_HEIGHT, SWP_NOACTIVATE);
        SetWindowPos(hTabSchedule, NULL, startX + (tabWidth * 2) - (tabOverlap * 2), 0, tabWidth, TAB_BAR_HEIGHT, SWP_NOACTIVATE);
    }
    else // PAGE_SCHEDULE
    {
        // Position active tab first so it takes top priority
        SetWindowPos(hTabSchedule, HWND_TOP, startX + (tabWidth * 2) - (tabOverlap * 2), 0, tabWidth, TAB_BAR_HEIGHT, SWP_SHOWWINDOW);
        // Position inactive tabs last
        SetWindowPos(hTabMailbox, NULL, startX + tabWidth - tabOverlap, 0, tabWidth, TAB_BAR_HEIGHT, SWP_NOACTIVATE);
        SetWindowPos(hTabInventory, NULL, startX, 0, tabWidth, TAB_BAR_HEIGHT, SWP_NOACTIVATE);
    }
    
    // Reposition Content Area
	int contentX = panelOffset + TOGGLE_BTN_WIDTH;
	int contentWidth = clientRect.right - contentX;
	int contentHeight = clientRect.bottom - TAB_BAR_HEIGHT;
    SetWindowPos(hContentArea, NULL, contentX, TAB_BAR_HEIGHT, contentWidth, contentHeight, SWP_NOZORDER);

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

    // Force a full window redraw to update the side panel background
	InvalidateRect(hWnd, NULL, TRUE);
	UpdateWindow(hWnd);
}

void DrawTabButton(LPDRAWITEMSTRUCT pDIS, bool isActive)
{
    HDC hdc = pDIS->hDC;
    RECT rect = pDIS->rcItem;

	Graphics graphics(hdc);
	graphics.SetSmoothingMode(SmoothingModeAntiAlias);
	graphics.SetPixelOffsetMode(PixelOffsetModeHighQuality);

    // DON'T clear the background - let it be transparent for overlapping!
    // Only clear if this is the active tab (so it draws on top cleanly)
    if (isActive)
    {
        Color clearColor(0, 30, 30, 35);
        SolidBrush clearBrush(clearColor);
        graphics.FillRectangle(&clearBrush, (INT)rect.left, (INT)rect.top, (INT)(rect.right - rect.left), (INT)(rect.bottom - rect.top));
    }
	Color bgColor = isActive ? Color(255, 30, 30, 35) : Color(255, 35, 35, 40);
	Color textColor = isActive ? Color(255, 100, 180, 255) : Color(255, 200, 200, 200);

	// Create trapezoid shape (wider at bottom, narrower at top)
	int indent = 8; // How much to indent the top corners
	int cornerRadius = 6; // Radius for rounded corners

    REAL left = (REAL)rect.left;
    REAL right = (REAL)rect.right - 1;
    REAL top = (REAL)rect.top + 5;
    REAL bottom = (REAL)rect.bottom;

    // Create a graphics path for the trapezoid with rounded corners
    GraphicsPath path;

    // Start from top-left corner (indented, with rounded corner)
    RectF topLeftArc(left + indent, top, (REAL)(cornerRadius * 2), (REAL)(cornerRadius * 2));
    path.AddArc(topLeftArc, 180.0f, 90.0f);

    // Top edge (narrower at top)
    path.AddLine(left + indent + cornerRadius, top, right - indent - cornerRadius, top);

    // Top-right rounded corner
    RectF topRightArc(right - indent - cornerRadius * 2, top, (REAL)(cornerRadius * 2), (REAL)(cornerRadius * 2));
    path.AddArc(topRightArc, 270.0f, 90.0f);

    // Right edge (angled outward toward bottom)
    path.AddLine(right - indent, top + cornerRadius, right, bottom);

    // Bottom edge (full width at bottom)
    path.AddLine(right, bottom, left, bottom);

    // Left edge (angled outward toward bottom)
    path.AddLine(left, bottom, left + indent, top + cornerRadius);

    path.CloseFigure();

    // Fill the trapezoid
    SolidBrush bgBrush(bgColor);
    graphics.FillPath(&bgBrush, &path);

    // Draw border - but skip the bottom line if active!
    Color borderColor = isActive ? Color(255, 60, 90, 95) : Color(255, 50, 50, 55);
    Pen borderPen(borderColor, 1.5f);
    
    if (isActive)
    {
        // Draw only the top and sides for active tab (no bottom border)
        GraphicsPath borderPath;

        // Top-left arc
        RectF topLeftBorderArc(left + indent, top, (REAL)(cornerRadius * 2), (REAL)(cornerRadius * 2));
        borderPath.AddArc(topLeftBorderArc, 180.0f, 90.0f);

        // Top edge
        borderPath.AddLine(left + indent + cornerRadius, top, right - indent - cornerRadius, top);

        // Top-right arc
        RectF topRightBorderArc(right - indent - cornerRadius * 2, top, (REAL)(cornerRadius * 2), (REAL)(cornerRadius * 2));
        borderPath.AddArc(topRightBorderArc, 270.0f, 90.0f);

        // Right edge (down to bottom)
        borderPath.AddLine(right - indent, top + cornerRadius, right, bottom);

        // Draw the border path (no bottom line!)
        graphics.DrawPath(&borderPen, &borderPath);

        // Left edge separately
        graphics.DrawLine(&borderPen, left, bottom, left + indent, top + cornerRadius);
    }
    else
    {
        // Inactive tabs get full border
        graphics.DrawPath(&borderPen, &path);

        // Also draw the horizontal border line underneath inactive tabs
        Pen horizontalBorderPen(Color(255, 60, 90, 95), 2.0f);
        graphics.DrawLine(&horizontalBorderPen, left, bottom, right, bottom);
    }

    // Draw accent line at bottom for active tab
    //if (isActive)
    //{
    //    Pen accentPen(Color(255, 13, 110, 253), 3.0f);
    //    graphics.DrawLine(&accentPen, (INT)left, (INT)bottom - 1, (INT)right, (INT)bottom - 1);
    //}

    // Draw button text
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

	// Using GDI+ Rect instead of just casting to avoid potential issues
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
