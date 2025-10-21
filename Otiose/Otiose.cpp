// Program by Alexa Culley
// This program is built for The UPS Store to help manage various aspects of the business.
// This program will be used as my Capstone Internship Project for my Cybersecurity degree at Spokane Falls Community College. 
// 
// Fun Fact: Otiose means "lazy" or "idle", which I thought was funny/ironic for a management app :)

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

// Button IDs
#define IDC_CLOCK_TOGGLE_BTN 1001

// Shift structure
struct Shift {
    std::string employeeName;
    std::string employeeID;
    time_t clockInTime;
    time_t clockOutTime;
    double hoursWorked;
    bool isActive = false; // true if clocked in, false if clocked out
};

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

// Clock In/Out System Variables
Shift currentShift;
HWND hClockToggleBtn, hStatusText;
ULONG_PTR gdiplusToken; // For GDI+ initialization
const std::string EMPLOYEE_NAME = "Alexa Culley"; // Placeholder, using me for now to track internship hours
const std::string EMPLOYEE_ID = "EMP001";         // Placeholder ID
const std::string TIMESHEET_FILE = "timesheet.csv";

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

// function prototypes 
void ClockIn();
void ClockOut();
void UpdateStatusDisplay(HWND hWnd);
void SaveShiftToCSV(const Shift& shift);
std::string TimeToString(time_t t);
void InitializeCSV();
void DrawColoredButton(LPDRAWITEMSTRUCT pDIS);

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
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_OTIOSE);
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

   // Create Clock Toggle button
   hClockToggleBtn = CreateWindowW(
       L"BUTTON", L"Clock In",
       WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_OWNERDRAW,
       125, 50, 150, 120, // make it square so the circle fits properly
       hWnd, (HMENU)IDC_CLOCK_TOGGLE_BTN, hInstance, nullptr);

   // Create Status text display
   hStatusText = CreateWindowW(
       L"STATIC", L"Status: Currently Clocked Out", 
       WS_VISIBLE | WS_CHILD | SS_LEFT,
       50, 190, 300, 200, 
	   hWnd, nullptr, hInstance, nullptr);

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

            case IDM_ABOUT:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;

            case IDM_EXIT:
                DestroyWindow(hWnd);
                break;

            default:
                return DefWindowProc(hWnd, message, wParam, lParam);
            }
        }
        break;

    case WM_DRAWITEM:
    {
        // Handle drawing the colored button
        LPDRAWITEMSTRUCT pDIS = (LPDRAWITEMSTRUCT)lParam;
        if (pDIS->CtlID == IDC_CLOCK_TOGGLE_BTN)
        {
            DrawColoredButton(pDIS);
            return TRUE;
        }
    }
	break;

    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            // TODO: Add any drawing code that uses hdc here...
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
        status << L"Status: Currently Clocked In\n\n";
        status << L"Employee: " << EMPLOYEE_NAME.c_str() << L"\n";
        status << L"Clocked In at: " << TimeToString(currentShift.clockInTime).c_str() << L"\n";
    }
    else if (currentShift.clockOutTime > 0)
    {
		status << L"Status: Currently Clocked Out\n\n";
        status << L"Employee: " << EMPLOYEE_NAME.c_str() << L"\n";
		status << L"Clock In Time: " << TimeToString(currentShift.clockInTime).c_str() << L"\n";
        status << L"Clock Out Time: " << TimeToString(currentShift.clockOutTime).c_str() << L"\n";
		status << L"Hours Worked: " << std::fixed << std::setprecision(2) 
            << currentShift.hoursWorked << L" hours\n";
    }
    else
    {
		status << L"Status: Currently Clocked Out\n\n";
		status << L"Employee: " << EMPLOYEE_NAME.c_str() << L"\n";
		status << L"Ready to Clock In!\n";
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
    HBRUSH hBackBrush = (HBRUSH)GetStockObject(WHITE_BRUSH);
	FillRect(hdc, &rect, hBackBrush);

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