// Program by Alexa Culley
// This program is built for The UPS Store to help manage various aspects of the business.
// This program will be used as my Capstone Internship Project for my Cybersecurity degree at Spokane Falls Community College. 
// 
// Fun Fact: Otiose means "lazy" or "idle", which I thought was funny/ironic for a management app :)

#include "framework.h"
#include "Otiose.h"
#include <string>
#include <fstream>
#include <ctime>
#include <sstream>
#include <iomanip>

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
    bool isActive; // true if clocked in, flase if clocked out
};

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

// Clock In/Out System Variables
Shift currentShift;
HWND hClockToggleBtn, hStatusText;
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
      CW_USEDEFAULT, 0, 600, 400, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   // Create Clock Toggle button
   hClockToggleBtn = CreateWindowW(
       L"BUTTON", L"Clock In",
       WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_OWNERDRAW,
       50, 50, 150, 120, // make it square so the circle fits properly
       hWnd, (HMENU)IDC_CLOCK_TOGGLE_BTN, hInstance, nullptr);

   // Create Status text display
   hStatusText = CreateWindowW(
       L"STATIC", L"Status: Currently Clocked Out", 
       WS_VISIBLE | WS_CHILD | SS_LEFT,
       50, 190, 500, 200, 
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
                }
                else
                {
                    // Currently clocked out, so clock in
                    ClockIn();
					UpdateStatusDisplay(hWnd);
                    SetWindowTextW(hClockToggleBtn, L"Clock Out");
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

	// Determine button color based on current state
    COLORREF buttonColor;
    COLORREF borderColor;
	COLORREF textColor = RGB(255, 255, 255); // White text

    if (currentShift.isActive)
    {
        // Clocked In - show red clock out button
        buttonColor = RGB(220, 53, 69); // Crimson Red
		borderColor = RGB(180, 43, 59); // Darker Red for border
    }
    else
    {
        // Clocked Out - show blue clock in button
        buttonColor = RGB(13, 110, 253); // Nice Blue color
		borderColor = RGB(10, 88, 202); // Darker Blue for border
    }

    // Check if button is being pressed
    if (pDIS->itemState & ODS_SELECTED)
    {
        // Darken color when pressed
		int r = (int)(GetRValue(buttonColor) * 0.8);
		int g = (int)(GetGValue(buttonColor) * 0.8);
		int b = (int)(GetBValue(buttonColor) * 0.8);
		buttonColor = RGB(r, g, b);
	}

    // Calculate center and radius for the circle
	int width = rect.right - rect.left;
	int height = rect.bottom - rect.top;
    int diameter = min(width, height);
	int centerX = rect.left + width / 2;
	int centerY = rect.top + height / 2;
	int radius = diameter / 2 - 5; // Padding of 5 pixels

    // Create circular region for the button
    HRGN hRegion = CreateEllipticRgn(
        centerX - radius, centerY - radius,
		centerX + radius, centerY + radius
    );

    // Clip to the circular region
	SelectClipRgn(hdc, hRegion);

    // Fill the circle with the button color
	HBRUSH hBrush = CreateSolidBrush(buttonColor);
	HBRUSH hOldBrush = (HBRUSH)SelectObject(hdc, hBrush);
	HPEN hPen = CreatePen(PS_SOLID, 1, borderColor);
	HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);

	Ellipse(hdc, centerX - radius, centerY - radius, 
        centerX + radius, centerY + radius);

    SelectObject(hdc, hOldPen);
	SelectObject(hdc, hOldBrush);
	DeleteObject(hPen);
    DeleteObject(hBrush);

    // Draw border circle
	HPEN hBorderPen = CreatePen(PS_SOLID, 3, borderColor);
	hOldPen = (HPEN)SelectObject(hdc, hPen);
    SelectObject(hdc, hBorderPen);
	SelectObject(hdc, GetStockObject(NULL_BRUSH)); // No fill for border

    Ellipse(hdc, centerX - radius, centerY - radius, 
		centerX + radius, centerY + radius);
    
	SelectObject(hdc, hOldPen);
	DeleteObject(hBorderPen);

	// Remove clipping region
	SelectClipRgn(hdc, NULL);
	DeleteObject(hRegion);

    // Draw the button text
    WCHAR buttonText[32];
	GetWindowTextW(pDIS->hwndItem, buttonText, 32);

    SetTextColor(hdc, textColor);
    SetBkMode(hdc, TRANSPARENT);

    // Create a font for the button
    HFONT hFont = CreateFontW(
        16, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        ANTIALIASED_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial"
    );
	HFONT hOldFont = (HFONT)SelectObject(hdc, hFont);

    // Create a rect for the text centered in the circle
    RECT textRect;
	textRect.left = centerX - radius;
	textRect.top = centerY - radius;
	textRect.right = centerX + radius;
	textRect.bottom = centerY + radius;

	DrawTextW(hdc, buttonText, -1, &textRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

    SelectObject(hdc, hOldFont);
    DeleteObject(hFont);
}