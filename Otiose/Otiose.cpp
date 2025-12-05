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
#include <commctrl.h>
#pragma comment(lib, "comctl32.lib")
#include <shellapi.h>
#include <uxtheme.h>
#pragma comment(lib, "uxtheme.lib")
#include <string>
#include <fstream>
#include <ctime>
#include <sstream>
#include <iomanip>
#include <algorithm>
#include <vector>

using namespace Gdiplus;
using std::min;

#define MAX_LOADSTRING 100

// Button and Control IDs
#define IDC_CLOCK_TOGGLE_BTN 1001
#define IDC_TAB_INVENTORY 1002
#define IDC_TAB_MAILBOX 1003
#define IDC_TAB_SCHEDULE 1004
#define IDC_TOGGLE_PANEL_BTN 1005
#define IDC_CATEGORY_BTN_BASE 2000  // Base ID for category buttons (2000, 2001, 2002, etc.)

// Inventory management button IDs
#define IDC_ADD_ITEM_BTN 3001
#define IDC_EDIT_ITEM_BTN 3002
#define IDC_DELETE_ITEM_BTN 3003
#define IDC_INVENTORY_LIST 3004
#define IDC_SEARCH_BOX 3005
#define IDC_SEARCH_LABEL 3006

// Dialog control IDs
#define IDC_EDIT_NAME 4001
#define IDC_EDIT_BULK 4002
#define IDC_EDIT_QTY 4003
#define IDC_EDIT_PRICE 4004
#define IDC_EDIT_LOCATION 4005
#define IDC_EDIT_LINK 4006
#define IDC_EDIT_NOTES 4007
#define IDC_DIALOG_OK 4010
#define IDC_DIALOG_CANCEL 4011

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

// Inventory item structure
struct InventoryItem {
    std::wstring name;
    int bulkCount;            // Bulk storage count
    int quantity;             // Current quantity on floor/shelf
    double price;
    std::wstring storageLocation;  // Where item is stored
    std::wstring orderLink;   // Optional URL for ordering more
    std::wstring notes;       // Additional notes
    time_t lastModified;      // Auto-set timestamp when item is created/edited
};

// Category enumeration for clarity
enum Category {
    CAT_RETAIL = 0,
    CAT_SUPPLIES = 1,
    CAT_BOXES = 2
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
HWND hCategoryPanel;
std::vector<HWND> hCategoryButtons;  // Dynamic list of category buttons
std::vector<std::wstring> categoryNames;  // Names of categories
int activeCategory = 0;  // Index of currently active category

// Inventory management UI elements
HWND hInventoryList;        // ListView for inventory items
HWND hAddItemBtn;           // Add new item button
HWND hEditItemBtn;          // Edit selected item button
HWND hDeleteItemBtn;        // Delete selected item button
HWND hSearchBox;            // Search/filter text box
HWND hSearchLabel;          // Label for search box
int selectedItemIndex = -1; // Currently selected item in list (-1 = none)

// Sorting state
int sortColumn = -1;        // Currently sorted column (-1 = none)
bool sortAscending = true;  // Sort direction

// Search/filter state
std::wstring searchFilter;  // Current search filter text

// Inventory storage - one vector per category
std::vector<InventoryItem> retailInventory;
std::vector<InventoryItem> suppliesInventory;
std::vector<InventoryItem> boxesInventory;

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
const int CATEGORY_PANEL_WIDTH = 200;

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
void DrawCategoryButton(LPDRAWITEMSTRUCT pDIS, bool isActive);
void UpdateContentArea(HWND hWnd);
void UpdateCategoryPanel(HWND hWnd);
void CreateCategoryButtons(HWND hWnd);
void SwitchCategory(int categoryIndex);

// Inventory functions
void InitializeInventory();
std::vector<InventoryItem>& GetCurrentInventory();
void SaveInventoryToCSV();
void LoadInventoryFromCSV();
std::wstring FormatInventoryDisplay();
void RefreshInventoryList();
void CreateInventoryManagementControls(HWND hWnd);
void ShowInventoryControls(bool show);
void DrawInventoryButton(LPDRAWITEMSTRUCT pDIS, const wchar_t* text, COLORREF bgColor);
INT_PTR CALLBACK ItemDialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);
void ShowAddItemDialog(HWND hWnd);
void ShowEditItemDialog(HWND hWnd);
void DeleteSelectedItem(HWND hWnd);
void SortInventoryList(int column);
void FilterInventoryList(const std::wstring& filter);
void SetupListViewColumns(HWND hListView);
void ResizeLastColumn();
int GetSelectedInventoryIndex();

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPWSTR    lpCmdLine,
    _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // Initialize Common Controls (required for ListView)
    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_LISTVIEW_CLASSES;
    InitCommonControlsEx(&icex);

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

    // Initialize default categories
    categoryNames.push_back(L"Retail");
    categoryNames.push_back(L"Supplies");
    categoryNames.push_back(L"Boxes");

    // Initialize inventory (load from CSV or create sample data)
    InitializeInventory();

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

    return (int)msg.wParam;
}

ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.cbClsExtra = 0;
    wcex.cbWndExtra = 0;
    wcex.hInstance = hInstance;
    wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_OTIOSE));
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = CreateSolidBrush(RGB(30, 30, 35)); // Dark grey/blue background
    wcex.lpszMenuName = nullptr;
    wcex.lpszClassName = szWindowClass;
    wcex.hIconSm = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

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
        1750 - (SIDE_PANEL_WIDTH + TOGGLE_BTN_WIDTH + CATEGORY_PANEL_WIDTH), 1000 - TAB_BAR_HEIGHT,
        hWnd, nullptr, hInstance, nullptr);

    // Create category panel (right side, starts below tabs)
    hCategoryPanel = CreateWindowW(
        L"STATIC", L"",
        WS_VISIBLE | WS_CHILD | SS_OWNERDRAW,
        1750 - CATEGORY_PANEL_WIDTH, TAB_BAR_HEIGHT,
        CATEGORY_PANEL_WIDTH, 1000 - TAB_BAR_HEIGHT,
        hWnd, nullptr, hInstance, nullptr);

    // Create category buttons
    CreateCategoryButtons(hWnd);

    // Create inventory management controls
    CreateInventoryManagementControls(hWnd);

    UpdateContentArea(hWnd);
    UpdateCategoryPanel(hWnd);

    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    int wmId = LOWORD(wParam);

    switch (message)
    {
    case WM_COMMAND:
    {

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

        case IDC_ADD_ITEM_BTN:
            ShowAddItemDialog(hWnd);
            break;

        case IDC_EDIT_ITEM_BTN:
            ShowEditItemDialog(hWnd);
            break;

        case IDC_DELETE_ITEM_BTN:
            DeleteSelectedItem(hWnd);
            break;

        case IDC_SEARCH_BOX:
            // Handle search box text changes
            if (HIWORD(wParam) == EN_CHANGE)
            {
                wchar_t searchText[256];
                GetWindowTextW(hSearchBox, searchText, 256);
                searchFilter = searchText;
                RefreshInventoryList();
            }
            break;

        default:
            // Check if it's a category button
            if (wmId >= IDC_CATEGORY_BTN_BASE && wmId < IDC_CATEGORY_BTN_BASE + 100)
            {
                int categoryIndex = wmId - IDC_CATEGORY_BTN_BASE;
                SwitchCategory(categoryIndex);
                // Redraw all category buttons
                for (HWND btn : hCategoryButtons)
                {
                    InvalidateRect(btn, NULL, TRUE);
                }
                break;
            }
            return DefWindowProc(hWnd, message, wParam, lParam);
        }
        // Removed About - keeping in case of future use
        // DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
        break;

    case IDM_EXIT:
        // Removed Menu - keeping in case of future use
        // DestroyWindow(hWnd);
        break;

    case IDM_ABOUT:
        // Removed About - keeping in case of future use
        // DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
        break;
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
        else if (pDIS->CtlID >= IDC_CATEGORY_BTN_BASE && pDIS->CtlID < IDC_CATEGORY_BTN_BASE + 100)
        {
            int categoryIndex = pDIS->CtlID - IDC_CATEGORY_BTN_BASE;
            DrawCategoryButton(pDIS, categoryIndex == activeCategory);
            return TRUE;
        }
        else if (pDIS->CtlID == IDC_ADD_ITEM_BTN)
        {
            DrawInventoryButton(pDIS, L"+ Add Item", RGB(40, 160, 80));
            return TRUE;
        }
        else if (pDIS->CtlID == IDC_EDIT_ITEM_BTN)
        {
            DrawInventoryButton(pDIS, L"Edit", RGB(60, 130, 200));
            return TRUE;
        }
        else if (pDIS->CtlID == IDC_DELETE_ITEM_BTN)
        {
            DrawInventoryButton(pDIS, L"Delete", RGB(200, 60, 60));
            return TRUE;
        }
    }
    break;

    case WM_NOTIFY:
    {
        LPNMHDR pnmh = (LPNMHDR)lParam;
        if (pnmh->idFrom == IDC_INVENTORY_LIST)
        {
            switch (pnmh->code)
            {
            case LVN_COLUMNCLICK:
            {
                // User clicked a column header - sort by that column
                LPNMLISTVIEW pnmlv = (LPNMLISTVIEW)lParam;
                SortInventoryList(pnmlv->iSubItem);
            }
            break;

            case LVN_ITEMCHANGED:
            {
                // Selection changed
                LPNMLISTVIEW pnmlv = (LPNMLISTVIEW)lParam;
                if (pnmlv->uNewState & LVIS_SELECTED)
                {
                    selectedItemIndex = pnmlv->iItem;
                    EnableWindow(hEditItemBtn, TRUE);
                    EnableWindow(hDeleteItemBtn, TRUE);
                }
                else if (ListView_GetSelectedCount(hInventoryList) == 0)
                {
                    selectedItemIndex = -1;
                    EnableWindow(hEditItemBtn, FALSE);
                    EnableWindow(hDeleteItemBtn, FALSE);
                }
            }
            break;

            case NM_DBLCLK:
            {
                // Double-click to edit
                if (selectedItemIndex >= 0)
                {
                    ShowEditItemDialog(hWnd);
                }
            }
            break;

            case NM_CLICK:
            {
                // Check if user clicked on the Order Link column
                LPNMITEMACTIVATE pnmia = (LPNMITEMACTIVATE)lParam;
                if (pnmia->iSubItem == 5 && pnmia->iItem >= 0)  // Column 5 is Order Link
                {
                    // Get the actual URL from the inventory item
                    LVITEM lvi;
                    lvi.mask = LVIF_PARAM;
                    lvi.iItem = pnmia->iItem;
                    lvi.iSubItem = 0;

                    if (ListView_GetItem(hInventoryList, &lvi))
                    {
                        int inventoryIndex = (int)lvi.lParam;
                        std::vector<InventoryItem>& inventory = GetCurrentInventory();

                        if (inventoryIndex >= 0 && inventoryIndex < (int)inventory.size())
                        {
                            const std::wstring& link = inventory[inventoryIndex].orderLink;
                            if (!link.empty())
                            {
                                ShellExecuteW(NULL, L"open", link.c_str(), NULL, NULL, SW_SHOWNORMAL);
                            }
                        }
                    }
                }
            }
            break;

            case NM_CUSTOMDRAW:
            {
                LPNMLVCUSTOMDRAW lplvcd = (LPNMLVCUSTOMDRAW)lParam;

                switch (lplvcd->nmcd.dwDrawStage)
                {
                case CDDS_PREPAINT:
                    return CDRF_NOTIFYITEMDRAW;

                case CDDS_ITEMPREPAINT:
                    // Just request subitem notifications
                    return CDRF_NOTIFYSUBITEMDRAW;

                case CDDS_SUBITEM | CDDS_ITEMPREPAINT:
                {
                    int visualRow = (int)lplvcd->nmcd.dwItemSpec;

                    // Query the actual selection state from the ListView
                    UINT state = ListView_GetItemState(hInventoryList, visualRow, LVIS_SELECTED);
                    bool isSelected = (state & LVIS_SELECTED) != 0;

                    if (isSelected)
                    {
                        lplvcd->clrText = RGB(255, 255, 255);
                        lplvcd->clrTextBk = RGB(60, 100, 160);
                    }
                    else
                    {
                        // Alternating row colors
                        lplvcd->clrText = RGB(220, 220, 220);
                        lplvcd->clrTextBk = (visualRow % 2 == 0) ? RGB(45, 45, 50) : RGB(50, 50, 55);
                    }

                    // Make the Order Link column (5) look like a hyperlink (unless selected)
                    if (lplvcd->iSubItem == 5 && !isSelected)
                    {
                        wchar_t linkText[512];
                        ListView_GetItemText(hInventoryList, visualRow, 5, linkText, 512);
                        if (wcslen(linkText) > 0)
                        {
                            lplvcd->clrText = RGB(100, 180, 255);
                        }
                    }

                    return CDRF_NEWFONT;
                }

                default:
                    return CDRF_DODEFAULT;
                }
            }
            }
        }

        // Handle header custom draw for dark theme
        HWND hHeader = ListView_GetHeader(hInventoryList);
        if (pnmh->hwndFrom == hHeader && pnmh->code == NM_CUSTOMDRAW)
        {
            LPNMCUSTOMDRAW pnmcd = (LPNMCUSTOMDRAW)lParam;

            switch (pnmcd->dwDrawStage)
            {
            case CDDS_PREPAINT:
                return CDRF_NOTIFYITEMDRAW;

            case CDDS_ITEMPREPAINT:
            {
                // Draw dark header background
                HBRUSH hBrush = CreateSolidBrush(RGB(50, 50, 55));
                FillRect(pnmcd->hdc, &pnmcd->rc, hBrush);
                DeleteObject(hBrush);

                // Draw header text
                SetTextColor(pnmcd->hdc, RGB(200, 200, 200));
                SetBkMode(pnmcd->hdc, TRANSPARENT);

                // Get header item text
                HDITEM hdi = { 0 };
                wchar_t headerText[64] = { 0 };
                hdi.mask = HDI_TEXT;
                hdi.pszText = headerText;
                hdi.cchTextMax = 64;
                Header_GetItem(hHeader, pnmcd->dwItemSpec, &hdi);

                // Draw the text centered vertically, left-aligned with padding
                RECT textRect = pnmcd->rc;
                textRect.left += 6;
                DrawTextW(pnmcd->hdc, headerText, -1, &textRect, DT_SINGLELINE | DT_VCENTER | DT_LEFT);

                // Draw a subtle border
                HPEN hPen = CreatePen(PS_SOLID, 1, RGB(70, 70, 75));
                HPEN hOldPen = (HPEN)SelectObject(pnmcd->hdc, hPen);
                MoveToEx(pnmcd->hdc, pnmcd->rc.right - 1, pnmcd->rc.top, NULL);
                LineTo(pnmcd->hdc, pnmcd->rc.right - 1, pnmcd->rc.bottom);
                SelectObject(pnmcd->hdc, hOldPen);
                DeleteObject(hPen);

                return CDRF_SKIPDEFAULT;
            }
            }
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

    case WM_CTLCOLOREDIT:
    {
        // Style the search box
        HDC hdcEdit = (HDC)wParam;
        SetTextColor(hdcEdit, RGB(220, 220, 220)); // Light text
        SetBkColor(hdcEdit, RGB(50, 50, 55)); // Dark background

        static HBRUSH hbrEditBkgnd = CreateSolidBrush(RGB(50, 50, 55));
        return (LRESULT)hbrEditBkgnd;
    }

    case WM_CTLCOLORSTATIC:
    {
        HDC hdcStatic = (HDC)wParam;

        // Default styling for static controls
        SetTextColor(hdcStatic, RGB(255, 255, 255)); // White text
        SetBkColor(hdcStatic, RGB(30, 30, 35)); // Match window background

        static HBRUSH hbrBkgnd = CreateSolidBrush(RGB(30, 30, 35));
        return (LRESULT)hbrBkgnd;
    }
    break;

    case WM_CTLCOLORLISTBOX:
    {
        HDC hdcListBox = (HDC)wParam;
        SetTextColor(hdcListBox, RGB(220, 220, 220)); // Light gray text
        SetBkColor(hdcListBox, RGB(45, 45, 50)); // Dark background

        static HBRUSH hbrListBkgnd = CreateSolidBrush(RGB(45, 45, 50));
        return (LRESULT)hbrListBkgnd;
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
        HBRUSH tabBarBrush = CreateSolidBrush(RGB(45, 45, 50)); // Lighter than main window
        FillRect(hdc, &tabBarBg, tabBarBrush);
        DeleteObject(tabBarBrush);

        // Draw side panel background FIRST
        if (sidePanelOpen)
        {
            RECT panelRect = { 0, 0, SIDE_PANEL_WIDTH, 10000 };
            HBRUSH panelBrush = CreateSolidBrush(RGB(40, 40, 45));
            FillRect(hdc, &panelRect, panelBrush);
            DeleteObject(panelBrush);
        }

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

        Pen borderPen(Color(255, 60, 60, 65), 2.0f);

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

// Wide string version for inventory display (shorter format)
std::wstring TimeToWString(time_t t)
{
    if (t == 0) return L"";

    struct tm timeinfo;
    localtime_s(&timeinfo, &t);

    std::wstringstream ss;
    ss << std::put_time(&timeinfo, L"%m/%d/%y %I:%M %p");
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

    // Update category panel visibility
    UpdateCategoryPanel(GetParent(hCategoryPanel));

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
    int contentWidth = clientRect.right - contentX - CATEGORY_PANEL_WIDTH;
    int contentHeight = clientRect.bottom - TAB_BAR_HEIGHT;
    SetWindowPos(hContentArea, NULL, contentX, TAB_BAR_HEIGHT, contentWidth, contentHeight, SWP_NOZORDER);

    // Reposition Category Panel (right side) and category buttons
    int categoryX = clientRect.right - CATEGORY_PANEL_WIDTH;
    SetWindowPos(hCategoryPanel, NULL, categoryX, TAB_BAR_HEIGHT,
        CATEGORY_PANEL_WIDTH, contentHeight, SWP_NOZORDER);

    // Reposition category buttons
    int buttonY = TAB_BAR_HEIGHT + 10;
    int buttonHeight = 50;
    int buttonSpacing = 10;

    for (size_t i = 0; i < hCategoryButtons.size(); i++)
    {
        SetWindowPos(hCategoryButtons[i], NULL,
            categoryX + 25, buttonY,
            150, buttonHeight, SWP_NOZORDER);
        buttonY += buttonHeight + buttonSpacing;
    }

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

    // Reposition inventory management controls
    if (currentPage == PAGE_INVENTORY)
    {
        int listX = contentX + 10;
        int searchY = TAB_BAR_HEIGHT + 10;
        int searchHeight = 25;
        int searchLabelWidth = 60;
        int searchBoxWidth = 250;

        // Position search label and box at top
        SetWindowPos(hSearchLabel, NULL, listX, searchY + 3, searchLabelWidth, searchHeight, SWP_NOZORDER);
        SetWindowPos(hSearchBox, NULL, listX + searchLabelWidth + 5, searchY, searchBoxWidth, searchHeight, SWP_NOZORDER);

        int listY = searchY + searchHeight + 10;
        int listWidth = contentWidth - 20;
        int listHeight = contentHeight - searchHeight - 80;  // Leave room for search and buttons

        SetWindowPos(hInventoryList, NULL, listX, listY, listWidth, listHeight, SWP_NOZORDER);

        // Resize last column to fill remaining width
        ResizeLastColumn();

        // Position buttons at bottom of content area
        int btnY = listY + listHeight + 10;
        int btnWidth = 100;
        int btnHeight = 35;
        int btnSpacing = 10;

        SetWindowPos(hAddItemBtn, NULL, listX, btnY, btnWidth, btnHeight, SWP_NOZORDER);
        SetWindowPos(hEditItemBtn, NULL, listX + btnWidth + btnSpacing, btnY, btnWidth, btnHeight, SWP_NOZORDER);
        SetWindowPos(hDeleteItemBtn, NULL, listX + (btnWidth + btnSpacing) * 2, btnY, btnWidth, btnHeight, SWP_NOZORDER);

        ShowInventoryControls(true);
    }
    else
    {
        ShowInventoryControls(false);
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

    // Draw line connecting edge of screen to corner of tab
    graphics.DrawLine(&borderPen, left - indent * 10, bottom, left, bottom);

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
        Pen horizontalBorderPen(Color(255, 60, 60, 65), 2.0f);
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
        content << L"INVENTORY MANAGEMENT - " << categoryNames[activeCategory] << L"\n\n";
        content << FormatInventoryDisplay();
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

void UpdateCategoryPanel(HWND hWnd)
{
    // Show/hide category panel based on current page
    if (currentPage == PAGE_INVENTORY)
    {
        ShowWindow(hCategoryPanel, SW_SHOW);
        for (HWND btn : hCategoryButtons)
        {
            ShowWindow(btn, SW_SHOW);
        }
    }
    else
    {
        ShowWindow(hCategoryPanel, SW_HIDE);
        for (HWND btn : hCategoryButtons)
        {
            ShowWindow(btn, SW_HIDE);
        }
    }
}

void CreateCategoryButtons(HWND hWnd)
{
    // Clear existing buttons
    for (HWND btn : hCategoryButtons)
    {
        DestroyWindow(btn);
    }
    hCategoryButtons.clear();

    // Create a button for each category
    int buttonY = TAB_BAR_HEIGHT + 10;
    int buttonHeight = 50;
    int buttonSpacing = 10;
    int categoryX = 1750 - CATEGORY_PANEL_WIDTH;

    for (size_t i = 0; i < categoryNames.size(); i++)
    {
        HWND hBtn = CreateWindowW(
            L"BUTTON", categoryNames[i].c_str(),
            WS_VISIBLE | WS_CHILD | BS_OWNERDRAW,  // Removed WS_TABSTOP to prevent focus
            categoryX + 25, buttonY,
            150, buttonHeight,
            hWnd, (HMENU)(IDC_CATEGORY_BTN_BASE + i), hInst, nullptr);

        hCategoryButtons.push_back(hBtn);
        buttonY += buttonHeight + buttonSpacing;
    }
}

void SwitchCategory(int categoryIndex)
{
    if (categoryIndex >= 0 && categoryIndex < (int)categoryNames.size())
    {
        activeCategory = categoryIndex;
        selectedItemIndex = -1;  // Clear selection when switching categories

        // Clear search filter and reset sort
        searchFilter.clear();
        SetWindowTextW(hSearchBox, L"");
        sortColumn = -1;
        sortAscending = true;

        // Update the inventory display to show items from this category
        RefreshInventoryList();
        // Disable edit/delete buttons since nothing is selected
        EnableWindow(hEditItemBtn, FALSE);
        EnableWindow(hDeleteItemBtn, FALSE);
    }
}

void DrawCategoryButton(LPDRAWITEMSTRUCT pDIS, bool isActive)
{
    HDC hdc = pDIS->hDC;
    RECT rect = pDIS->rcItem;

    Graphics graphics(hdc);
    graphics.SetSmoothingMode(SmoothingModeAntiAlias);

    REAL left = (REAL)rect.left;
    REAL top = (REAL)rect.top;
    REAL right = (REAL)rect.right - 1;
    REAL bottom = (REAL)rect.bottom - 1;

    // Clear the entire button area first with the background color
    Color clearColor(255, 35, 35, 40); // Match category panel background
    SolidBrush clearBrush(clearColor);
    graphics.FillRectangle(&clearBrush, left, top, right - left, bottom - top);

    // Button colors
    Color bgColor = isActive ? Color(255, 70, 130, 200) : Color(255, 45, 45, 50);
    Color textColor = isActive ? Color(255, 255, 255, 255) : Color(255, 180, 180, 180);

    if (pDIS->itemState & ODS_SELECTED)
    {
        // Darken when pressed
        BYTE r = (BYTE)(bgColor.GetR() * 0.8);
        BYTE g = (BYTE)(bgColor.GetG() * 0.8);
        BYTE b = (BYTE)(bgColor.GetB() * 0.8);
        bgColor = Color(255, r, g, b);
    }

    // Draw rounded rectangle
    int cornerRadius = 6;
    GraphicsPath path;

    RectF topLeft(left, top, cornerRadius * 2, cornerRadius * 2);
    RectF topRight(right - cornerRadius * 2, top, cornerRadius * 2, cornerRadius * 2);
    RectF bottomRight(right - cornerRadius * 2, bottom - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2);
    RectF bottomLeft(left, bottom - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2);

    path.AddArc(topLeft, 180, 90);
    path.AddArc(topRight, 270, 90);
    path.AddArc(bottomRight, 0, 90);
    path.AddArc(bottomLeft, 90, 90);
    path.CloseFigure();

    // Fill button
    SolidBrush brush(bgColor);
    graphics.FillPath(&brush, &path);

    // Draw border for active button
    if (isActive)
    {
        Pen borderPen(Color(255, 50, 100, 180), 2.0f);
        graphics.DrawPath(&borderPen, &path);
    }

    // Draw text
    WCHAR buttonText[64];
    GetWindowTextW(pDIS->hwndItem, buttonText, 64);

    FontFamily fontFamily(L"Segoe UI");
    Font font(&fontFamily, 14, isActive ? FontStyleBold : FontStyleRegular, UnitPixel);
    SolidBrush textBrush(textColor);

    StringFormat stringFormat;
    stringFormat.SetAlignment(StringAlignmentCenter);
    stringFormat.SetLineAlignment(StringAlignmentCenter);

    RectF layoutRect((REAL)rect.left, (REAL)rect.top, (REAL)(rect.right - rect.left), (REAL)(rect.bottom - rect.top));

    graphics.SetTextRenderingHint(TextRenderingHintAntiAlias);
    graphics.DrawString(buttonText, -1, &font, layoutRect, &stringFormat, &textBrush);

    // Suppress the focus rectangle - this prevents the blue outline!
    if (pDIS->itemState & ODS_FOCUS)
    {
        // Normally Windows would draw a focus rect here, but we're owner-drawing so we skip it
        // Our active state styling already shows which button is selected
    }
}

void DrawAddCategoryButton(LPDRAWITEMSTRUCT pDIS)
{
    HDC hdc = pDIS->hDC;
    RECT rect = pDIS->rcItem;

    Graphics graphics(hdc);
    graphics.SetSmoothingMode(SmoothingModeAntiAlias);
    graphics.SetCompositingQuality(CompositingQualityHighQuality);
    graphics.SetPixelOffsetMode(PixelOffsetModeHighQuality);

    // Clear background
    Color bgColor = Color(255, 50, 120, 200); // Nice blue color for add button
    if (pDIS->itemState & ODS_SELECTED)
    {
        bgColor = Color(255, 40, 96, 160); // Darker when pressed
    }

    // Draw rounded rectangle button
    int cornerRadius = 8;
    GraphicsPath path;

    RectF rectF((REAL)rect.left, (REAL)rect.top, (REAL)(rect.right - rect.left), (REAL)(rect.bottom - rect.top));

    // Create rounded rectangle path
    REAL left = rectF.X;
    REAL top = rectF.Y;
    REAL right = rectF.X + rectF.Width;
    REAL bottom = rectF.Y + rectF.Height;

    RectF topLeft(left, top, cornerRadius * 2, cornerRadius * 2);
    RectF topRight(right - cornerRadius * 2, top, cornerRadius * 2, cornerRadius * 2);
    RectF bottomRight(right - cornerRadius * 2, bottom - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2);
    RectF bottomLeft(left, bottom - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2);

    path.AddArc(topLeft, 180, 90);
    path.AddArc(topRight, 270, 90);
    path.AddArc(bottomRight, 0, 90);
    path.AddArc(bottomLeft, 90, 90);
    path.CloseFigure();

    // Fill button
    SolidBrush brush(bgColor);
    graphics.FillPath(&brush, &path);

    // Draw border
    Pen borderPen(Color(255, 40, 90, 150), 2.0f);
    graphics.DrawPath(&borderPen, &path);

    // Draw the "+" text
    WCHAR buttonText[32];
    GetWindowTextW(pDIS->hwndItem, buttonText, 32);

    FontFamily fontFamily(L"Segoe UI");
    Font font(&fontFamily, 24, FontStyleBold, UnitPixel);
    SolidBrush textBrush(Color(255, 255, 255, 255));

    StringFormat stringFormat;
    stringFormat.SetAlignment(StringAlignmentCenter);
    stringFormat.SetLineAlignment(StringAlignmentCenter);

    RectF layoutRect((REAL)rect.left, (REAL)rect.top, (REAL)(rect.right - rect.left), (REAL)(rect.bottom - rect.top));

    graphics.SetTextRenderingHint(TextRenderingHintAntiAlias);
    graphics.DrawString(buttonText, -1, &font, layoutRect, &stringFormat, &textBrush);
}

// ============================================================================
// INVENTORY MANAGEMENT FUNCTIONS
// ============================================================================

// Returns a reference to the inventory vector for the currently active category
std::vector<InventoryItem>& GetCurrentInventory()
{
    switch (activeCategory)
    {
    case CAT_RETAIL:
        return retailInventory;
    case CAT_SUPPLIES:
        return suppliesInventory;
    case CAT_BOXES:
        return boxesInventory;
    default:
        return retailInventory;
    }
}

// Initialize inventory with sample data (or load from CSV)
void InitializeInventory()
{
    // Try to load from CSV first
    LoadInventoryFromCSV();

    // If no data was loaded, add some sample items
    // Structure: { name, bulkCount, quantity, price, storageLocation, orderLink, notes, lastModified }
    time_t now = time(nullptr);

    if (retailInventory.empty())
    {
        InventoryItem item1 = { L"Bubble Mailer - Small", 100, 50, 2.99, L"Shelf A1", L"https://www.uline.com/BL/S-9854", L"", now };
        InventoryItem item2 = { L"Bubble Mailer - Medium", 75, 35, 3.99, L"Shelf A2", L"", L"", now };
        InventoryItem item3 = { L"Bubble Mailer - Large", 50, 25, 4.99, L"Shelf A3", L"", L"", now };
        InventoryItem item4 = { L"Packing Tape", 200, 100, 5.99, L"Back Room", L"https://www.uline.com/BL/S-423", L"Reorder when below 50", now };
        InventoryItem item5 = { L"Fragile Stickers", 500, 200, 0.25, L"Counter", L"", L"", now };
        retailInventory.push_back(item1);
        retailInventory.push_back(item2);
        retailInventory.push_back(item3);
        retailInventory.push_back(item4);
        retailInventory.push_back(item5);
    }

    if (suppliesInventory.empty())
    {
        InventoryItem item1 = { L"Receipt Paper Roll", 48, 24, 8.99, L"Office Cabinet", L"", L"", now };
        InventoryItem item2 = { L"Printer Ink - Black", 12, 6, 34.99, L"Office Cabinet", L"", L"For main printer", now };
        InventoryItem item3 = { L"Printer Ink - Color", 8, 4, 44.99, L"Office Cabinet", L"", L"For main printer", now };
        InventoryItem item4 = { L"Shipping Labels", 1000, 500, 0.15, L"Label Station", L"", L"", now };
        InventoryItem item5 = { L"Box Cutter Blades", 100, 50, 0.50, L"Tool Drawer", L"", L"", now };
        suppliesInventory.push_back(item1);
        suppliesInventory.push_back(item2);
        suppliesInventory.push_back(item3);
        suppliesInventory.push_back(item4);
        suppliesInventory.push_back(item5);
    }

    if (boxesInventory.empty())
    {
        InventoryItem item1 = { L"Small Box (8x6x4)", 150, 75, 1.99, L"Box Wall - Bottom", L"https://www.uline.com/BL/S-4511", L"", now };
        InventoryItem item2 = { L"Medium Box (12x10x8)", 100, 50, 3.49, L"Box Wall - Middle", L"", L"Best seller", now };
        InventoryItem item3 = { L"Large Box (18x14x12)", 60, 30, 4.99, L"Box Wall - Top", L"", L"", now };
        InventoryItem item4 = { L"Extra Large Box (24x18x18)", 30, 15, 7.99, L"Back Room", L"", L"", now };
        InventoryItem item5 = { L"Picture Box (30x24x3)", 20, 10, 9.99, L"Specialty Section", L"", L"For framed items", now };
        InventoryItem item6 = { L"Mirror Box (48x6x30)", 16, 8, 14.99, L"Specialty Section", L"", L"For mirrors/glass", now };
        boxesInventory.push_back(item1);
        boxesInventory.push_back(item2);
        boxesInventory.push_back(item3);
        boxesInventory.push_back(item4);
        boxesInventory.push_back(item5);
        boxesInventory.push_back(item6);
    }
}

// Format the current category's inventory for display
std::wstring FormatInventoryDisplay()
{
    std::wstringstream display;
    std::vector<InventoryItem>& inventory = GetCurrentInventory();

    if (inventory.empty())
    {
        display << L"No items in this category.\n";
        return display.str();
    }

    // Header
    display << std::left << std::setw(30) << L"Item Name"
        << std::setw(12) << L"Bulk"
        << std::setw(10) << L"Qty"
        << std::setw(10) << L"Price"
        << L"Location\n";
    display << L"--------------------------------------------------------------------------------\n";

    // Items
    for (const auto& item : inventory)
    {
        display << std::left << std::setw(30) << item.name
            << std::setw(12) << item.bulkCount
            << std::setw(10) << item.quantity
            << L"$" << std::fixed << std::setprecision(2) << std::setw(8) << item.price
            << item.storageLocation << L"\n";
    }

    display << L"\n--------------------------------------------------------------------------------\n";
    display << L"Total Items: " << inventory.size() << L"\n";

    return display.str();
}

// Save all inventory to CSV files
void SaveInventoryToCSV()
{
    // Helper lambda to convert wstring to string
    auto wstringToString = [](const std::wstring& wstr) -> std::string {
        if (wstr.empty()) return "";
        int size = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, nullptr, 0, nullptr, nullptr);
        std::string str(size - 1, 0);
        WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, &str[0], size, nullptr, nullptr);
        return str;
        };

    // Helper lambda to save a single category
    auto saveCategory = [&wstringToString](const std::string& filename, const std::vector<InventoryItem>& inventory) {
        std::ofstream file(filename);
        if (file.is_open())
        {
            // Write header
            file << "Name,BulkCount,Quantity,Price,StorageLocation,OrderLink,Notes,LastModified\n";

            // Write items
            for (const auto& item : inventory)
            {
                file << wstringToString(item.name) << ","
                    << item.bulkCount << ","
                    << item.quantity << ","
                    << std::fixed << std::setprecision(2) << item.price << ","
                    << wstringToString(item.storageLocation) << ","
                    << wstringToString(item.orderLink) << ","
                    << wstringToString(item.notes) << ","
                    << item.lastModified << "\n";
            }
            file.close();
        }
        };

    saveCategory("inventory_retail.csv", retailInventory);
    saveCategory("inventory_supplies.csv", suppliesInventory);
    saveCategory("inventory_boxes.csv", boxesInventory);
}

// Load inventory from CSV files
void LoadInventoryFromCSV()
{
    // Helper lambda to convert string to wstring
    auto stringToWstring = [](const std::string& str) -> std::wstring {
        if (str.empty()) return L"";
        int size = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, nullptr, 0);
        std::wstring wstr(size - 1, 0);
        MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, &wstr[0], size);
        return wstr;
        };

    // Helper lambda to load a single category
    auto loadCategory = [&stringToWstring](const std::string& filename, std::vector<InventoryItem>& inventory) {
        std::ifstream file(filename);
        if (!file.is_open()) return;

        std::string line;
        bool isHeader = true;

        while (std::getline(file, line))
        {
            if (isHeader) { isHeader = false; continue; }  // Skip header row
            if (line.empty()) continue;

            InventoryItem item;
            std::stringstream ss(line);
            std::string token;

            // Parse CSV fields
            if (std::getline(ss, token, ',')) item.name = stringToWstring(token);
            if (std::getline(ss, token, ',')) item.bulkCount = std::stoi(token);
            if (std::getline(ss, token, ',')) item.quantity = std::stoi(token);
            if (std::getline(ss, token, ',')) item.price = std::stod(token);
            if (std::getline(ss, token, ',')) item.storageLocation = stringToWstring(token);
            if (std::getline(ss, token, ',')) item.orderLink = stringToWstring(token);
            if (std::getline(ss, token, ',')) item.notes = stringToWstring(token);
            if (std::getline(ss, token, ',')) item.lastModified = std::stoll(token);

            inventory.push_back(item);
        }
        file.close();
        };

    loadCategory("inventory_retail.csv", retailInventory);
    loadCategory("inventory_supplies.csv", suppliesInventory);
    loadCategory("inventory_boxes.csv", boxesInventory);
}

// ============================================================================
// INVENTORY MANAGEMENT UI FUNCTIONS
// ============================================================================

// Create the inventory management controls (listbox and buttons)
void CreateInventoryManagementControls(HWND hWnd)
{
    // Create search label
    hSearchLabel = CreateWindowW(
        L"STATIC", L"Search:",
        WS_CHILD | SS_LEFT,
        0, 0, 60, 25,
        hWnd, (HMENU)IDC_SEARCH_LABEL, hInst, nullptr);

    // Create search box
    hSearchBox = CreateWindowW(
        L"EDIT", L"",
        WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
        0, 0, 250, 25,
        hWnd, (HMENU)IDC_SEARCH_BOX, hInst, nullptr);

    // Create inventory ListView (instead of listbox)
    hInventoryList = CreateWindowExW(
        0,
        WC_LISTVIEW, L"",
        WS_CHILD | WS_BORDER | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS,
        0, 0, 100, 100,  // Position will be set in RepositionControls
        hWnd, (HMENU)IDC_INVENTORY_LIST, hInst, nullptr);

    // Disable Windows visual theme on ListView so our custom colors work
    SetWindowTheme(hInventoryList, L"", L"");

    // Enable full row select (no grid lines - they don't theme well)
    ListView_SetExtendedListViewStyle(hInventoryList,
        LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER);

    // Set dark theme colors for ListView
    ListView_SetBkColor(hInventoryList, RGB(45, 45, 50));
    ListView_SetTextBkColor(hInventoryList, RGB(45, 45, 50));
    ListView_SetTextColor(hInventoryList, RGB(220, 220, 220));

    // Set up the columns
    SetupListViewColumns(hInventoryList);

    // Set fonts
    HFONT hListFont = CreateFontW(
        14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Segoe UI"
    );
    SendMessage(hInventoryList, WM_SETFONT, (WPARAM)hListFont, TRUE);
    SendMessage(hSearchBox, WM_SETFONT, (WPARAM)hListFont, TRUE);
    SendMessage(hSearchLabel, WM_SETFONT, (WPARAM)hListFont, TRUE);

    // Create Add button
    hAddItemBtn = CreateWindowW(
        L"BUTTON", L"+ Add Item",
        WS_CHILD | BS_OWNERDRAW,
        0, 0, 100, 35,
        hWnd, (HMENU)IDC_ADD_ITEM_BTN, hInst, nullptr);

    // Create Edit button
    hEditItemBtn = CreateWindowW(
        L"BUTTON", L"Edit",
        WS_CHILD | BS_OWNERDRAW,
        0, 0, 100, 35,
        hWnd, (HMENU)IDC_EDIT_ITEM_BTN, hInst, nullptr);

    // Create Delete button
    hDeleteItemBtn = CreateWindowW(
        L"BUTTON", L"Delete",
        WS_CHILD | BS_OWNERDRAW,
        0, 0, 100, 35,
        hWnd, (HMENU)IDC_DELETE_ITEM_BTN, hInst, nullptr);

    // Initially disable edit and delete buttons (nothing selected)
    EnableWindow(hEditItemBtn, FALSE);
    EnableWindow(hDeleteItemBtn, FALSE);

    // Populate the list with initial data
    RefreshInventoryList();
}

// Set up ListView columns
void SetupListViewColumns(HWND hListView)
{
    LVCOLUMN lvc;
    lvc.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;

    // Column 0: Name
    lvc.iSubItem = 0;
    lvc.pszText = (LPWSTR)L"Name";
    lvc.cx = 170;
    lvc.fmt = LVCFMT_LEFT;
    ListView_InsertColumn(hListView, 0, &lvc);

    // Column 1: Bulk Count
    lvc.iSubItem = 1;
    lvc.pszText = (LPWSTR)L"Bulk Count";
    lvc.cx = 80;
    lvc.fmt = LVCFMT_RIGHT;
    ListView_InsertColumn(hListView, 1, &lvc);

    // Column 2: Quantity
    lvc.iSubItem = 2;
    lvc.pszText = (LPWSTR)L"Qty";
    lvc.cx = 50;
    lvc.fmt = LVCFMT_RIGHT;
    ListView_InsertColumn(hListView, 2, &lvc);

    // Column 3: Price
    lvc.iSubItem = 3;
    lvc.pszText = (LPWSTR)L"Price";
    lvc.cx = 65;
    lvc.fmt = LVCFMT_RIGHT;
    ListView_InsertColumn(hListView, 3, &lvc);

    // Column 4: Storage Location
    lvc.iSubItem = 4;
    lvc.pszText = (LPWSTR)L"Storage Location";
    lvc.cx = 130;
    lvc.fmt = LVCFMT_LEFT;
    ListView_InsertColumn(hListView, 4, &lvc);

    // Column 5: Order Link (clickable)
    lvc.iSubItem = 5;
    lvc.pszText = (LPWSTR)L"Order Link";
    lvc.cx = 100;
    lvc.fmt = LVCFMT_LEFT;
    ListView_InsertColumn(hListView, 5, &lvc);

    // Column 6: Notes
    lvc.iSubItem = 6;
    lvc.pszText = (LPWSTR)L"Notes";
    lvc.cx = 150;
    lvc.fmt = LVCFMT_LEFT;
    ListView_InsertColumn(hListView, 6, &lvc);

    // Column 7: Last Edited
    lvc.iSubItem = 7;
    lvc.pszText = (LPWSTR)L"Last Edited";
    lvc.cx = 130;
    lvc.fmt = LVCFMT_LEFT;
    ListView_InsertColumn(hListView, 7, &lvc);
}

// Resize the last column to fill remaining ListView width
void ResizeLastColumn()
{
    if (!hInventoryList) return;

    RECT listRect;
    GetClientRect(hInventoryList, &listRect);
    int listWidth = listRect.right - listRect.left;

    // Calculate total width of columns 0-6
    int usedWidth = 0;
    for (int i = 0; i < 7; i++)
    {
        usedWidth += ListView_GetColumnWidth(hInventoryList, i);
    }

    // Set last column (7) to fill remaining space
    int lastColWidth = listWidth - usedWidth - 4;  // -4 for border/padding
    if (lastColWidth < 100) lastColWidth = 100;    // Minimum width

    ListView_SetColumnWidth(hInventoryList, 7, lastColWidth);
}

// Show or hide inventory management controls
void ShowInventoryControls(bool show)
{
    int showCmd = show ? SW_SHOW : SW_HIDE;
    ShowWindow(hInventoryList, showCmd);
    ShowWindow(hAddItemBtn, showCmd);
    ShowWindow(hEditItemBtn, showCmd);
    ShowWindow(hDeleteItemBtn, showCmd);

    // Hide the static content area when showing inventory controls
    ShowWindow(hContentArea, show ? SW_HIDE : SW_SHOW);

    // Show/hide search controls
    ShowWindow(hSearchLabel, showCmd);
    ShowWindow(hSearchBox, showCmd);
}

// Helper function to check if item matches search filter
bool ItemMatchesFilter(const InventoryItem& item, const std::wstring& filter)
{
    if (filter.empty()) return true;

    // Convert filter to lowercase for case-insensitive search
    std::wstring lowerFilter = filter;
    std::transform(lowerFilter.begin(), lowerFilter.end(), lowerFilter.begin(), ::towlower);

    // Check each field
    std::wstring lowerName = item.name;
    std::transform(lowerName.begin(), lowerName.end(), lowerName.begin(), ::towlower);
    if (lowerName.find(lowerFilter) != std::wstring::npos) return true;

    std::wstring lowerLocation = item.storageLocation;
    std::transform(lowerLocation.begin(), lowerLocation.end(), lowerLocation.begin(), ::towlower);
    if (lowerLocation.find(lowerFilter) != std::wstring::npos) return true;

    std::wstring lowerNotes = item.notes;
    std::transform(lowerNotes.begin(), lowerNotes.end(), lowerNotes.begin(), ::towlower);
    if (lowerNotes.find(lowerFilter) != std::wstring::npos) return true;

    // Check quantity and bulk count as strings
    std::wstring qtyStr = std::to_wstring(item.quantity);
    if (qtyStr.find(filter) != std::wstring::npos) return true;

    std::wstring bulkStr = std::to_wstring(item.bulkCount);
    if (bulkStr.find(filter) != std::wstring::npos) return true;

    return false;
}

// Refresh the inventory ListView with current category items
void RefreshInventoryList()
{
    // Clear the ListView
    ListView_DeleteAllItems(hInventoryList);

    std::vector<InventoryItem>& inventory = GetCurrentInventory();

    // Create a filtered and sorted copy for display
    std::vector<std::pair<int, InventoryItem*>> filteredItems;
    for (size_t i = 0; i < inventory.size(); i++)
    {
        if (ItemMatchesFilter(inventory[i], searchFilter))
        {
            filteredItems.push_back(std::make_pair((int)i, &inventory[i]));
        }
    }

    // Sort if a column is selected
    if (sortColumn >= 0)
    {
        std::sort(filteredItems.begin(), filteredItems.end(),
            [](const std::pair<int, InventoryItem*>& a, const std::pair<int, InventoryItem*>& b) {
                int cmp = 0;
                switch (sortColumn)
                {
                case 0: // Name
                    cmp = _wcsicmp(a.second->name.c_str(), b.second->name.c_str());
                    break;
                case 1: // Bulk Count
                    cmp = a.second->bulkCount - b.second->bulkCount;
                    break;
                case 2: // Quantity
                    cmp = a.second->quantity - b.second->quantity;
                    break;
                case 3: // Price
                    if (a.second->price < b.second->price) cmp = -1;
                    else if (a.second->price > b.second->price) cmp = 1;
                    else cmp = 0;
                    break;
                case 4: // Storage Location
                    cmp = _wcsicmp(a.second->storageLocation.c_str(), b.second->storageLocation.c_str());
                    break;
                case 6: // Notes
                    cmp = _wcsicmp(a.second->notes.c_str(), b.second->notes.c_str());
                    break;
                case 7: // Last Edited
                    if (a.second->lastModified < b.second->lastModified) cmp = -1;
                    else if (a.second->lastModified > b.second->lastModified) cmp = 1;
                    else cmp = 0;
                    break;
                }
                return sortAscending ? (cmp < 0) : (cmp > 0);
            });
    }

    // Add items to ListView
    int rowIndex = 0;
    for (const auto& pair : filteredItems)
    {
        const InventoryItem& item = *pair.second;

        // Insert the item (first column)
        LVITEM lvi;
        lvi.mask = LVIF_TEXT | LVIF_PARAM;
        lvi.iItem = rowIndex;
        lvi.iSubItem = 0;
        lvi.pszText = (LPWSTR)item.name.c_str();
        lvi.lParam = pair.first;  // Store original index for editing/deleting
        ListView_InsertItem(hInventoryList, &lvi);

        // Set subitems (other columns)
        wchar_t bulkStr[32];
        swprintf_s(bulkStr, L"%d", item.bulkCount);
        ListView_SetItemText(hInventoryList, rowIndex, 1, bulkStr);

        wchar_t qtyStr[32];
        swprintf_s(qtyStr, L"%d", item.quantity);
        ListView_SetItemText(hInventoryList, rowIndex, 2, qtyStr);

        wchar_t priceStr[32];
        swprintf_s(priceStr, L"$%.2f", item.price);
        ListView_SetItemText(hInventoryList, rowIndex, 3, priceStr);

        ListView_SetItemText(hInventoryList, rowIndex, 4, (LPWSTR)item.storageLocation.c_str());

        // Order Link - show shortened version or placeholder
        if (!item.orderLink.empty())
        {
            ListView_SetItemText(hInventoryList, rowIndex, 5, (LPWSTR)L"[Click to Order]");
        }
        else
        {
            ListView_SetItemText(hInventoryList, rowIndex, 5, (LPWSTR)L"");
        }

        // Notes column
        ListView_SetItemText(hInventoryList, rowIndex, 6, (LPWSTR)item.notes.c_str());

        // Last Edited column
        std::wstring lastEditStr = TimeToWString(item.lastModified);
        ListView_SetItemText(hInventoryList, rowIndex, 7, (LPWSTR)lastEditStr.c_str());

        rowIndex++;
    }

    // Clear selection
    selectedItemIndex = -1;
    EnableWindow(hEditItemBtn, FALSE);
    EnableWindow(hDeleteItemBtn, FALSE);
}

// Sort the inventory list by the specified column
void SortInventoryList(int column)
{
    if (sortColumn == column)
    {
        // Same column - toggle direction
        sortAscending = !sortAscending;
    }
    else
    {
        // New column - default to ascending
        sortColumn = column;
        sortAscending = true;
    }

    RefreshInventoryList();
}

// Get the original inventory index from the selected ListView item
int GetSelectedInventoryIndex()
{
    if (selectedItemIndex < 0) return -1;

    LVITEM lvi;
    lvi.mask = LVIF_PARAM;
    lvi.iItem = selectedItemIndex;
    lvi.iSubItem = 0;

    if (ListView_GetItem(hInventoryList, &lvi))
    {
        return (int)lvi.lParam;
    }
    return -1;
}

// Draw inventory management buttons with custom colors
void DrawInventoryButton(LPDRAWITEMSTRUCT pDIS, const wchar_t* text, COLORREF bgColor)
{
    HDC hdc = pDIS->hDC;
    RECT rect = pDIS->rcItem;

    Graphics graphics(hdc);
    graphics.SetSmoothingMode(SmoothingModeAntiAlias);

    // Check if button is disabled
    bool isDisabled = (pDIS->itemState & ODS_DISABLED) != 0;

    // Adjust color for disabled state
    BYTE r = GetRValue(bgColor);
    BYTE g = GetGValue(bgColor);
    BYTE b = GetBValue(bgColor);

    if (isDisabled)
    {
        r = (BYTE)(r * 0.4);
        g = (BYTE)(g * 0.4);
        b = (BYTE)(b * 0.4);
    }
    else if (pDIS->itemState & ODS_SELECTED)
    {
        r = (BYTE)(r * 0.7);
        g = (BYTE)(g * 0.7);
        b = (BYTE)(b * 0.7);
    }

    Color buttonColor(255, r, g, b);

    // Draw rounded rectangle
    int cornerRadius = 6;
    GraphicsPath path;

    REAL left = (REAL)rect.left;
    REAL top = (REAL)rect.top;
    REAL right = (REAL)rect.right - 1;
    REAL bottom = (REAL)rect.bottom - 1;

    RectF topLeft(left, top, cornerRadius * 2.0f, cornerRadius * 2.0f);
    RectF topRight(right - cornerRadius * 2, top, cornerRadius * 2.0f, cornerRadius * 2.0f);
    RectF bottomRight(right - cornerRadius * 2, bottom - cornerRadius * 2, cornerRadius * 2.0f, cornerRadius * 2.0f);
    RectF bottomLeft(left, bottom - cornerRadius * 2, cornerRadius * 2.0f, cornerRadius * 2.0f);

    path.AddArc(topLeft, 180, 90);
    path.AddArc(topRight, 270, 90);
    path.AddArc(bottomRight, 0, 90);
    path.AddArc(bottomLeft, 90, 90);
    path.CloseFigure();

    // Fill button
    SolidBrush brush(buttonColor);
    graphics.FillPath(&brush, &path);

    // Draw text
    FontFamily fontFamily(L"Segoe UI");
    Font font(&fontFamily, 12, FontStyleBold, UnitPixel);
    Color textColor = isDisabled ? Color(255, 120, 120, 120) : Color(255, 255, 255, 255);
    SolidBrush textBrush(textColor);

    StringFormat stringFormat;
    stringFormat.SetAlignment(StringAlignmentCenter);
    stringFormat.SetLineAlignment(StringAlignmentCenter);

    RectF layoutRect((REAL)rect.left, (REAL)rect.top, (REAL)(rect.right - rect.left), (REAL)(rect.bottom - rect.top));

    graphics.SetTextRenderingHint(TextRenderingHintAntiAlias);
    graphics.DrawString(text, -1, &font, layoutRect, &stringFormat, &textBrush);
}

// Global variables for dialog data passing
static InventoryItem dialogItem;
static bool dialogIsEdit = false;

// Dialog procedure for Add/Edit item dialog
INT_PTR CALLBACK ItemDialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_INITDIALOG:
    {
        // Center the dialog on the parent window
        HWND hParent = GetParent(hDlg);
        RECT rcParent, rcDlg;
        GetWindowRect(hParent, &rcParent);
        GetWindowRect(hDlg, &rcDlg);
        int x = rcParent.left + (rcParent.right - rcParent.left - (rcDlg.right - rcDlg.left)) / 2;
        int y = rcParent.top + (rcParent.bottom - rcParent.top - (rcDlg.bottom - rcDlg.top)) / 2;
        SetWindowPos(hDlg, NULL, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER);

        // Populate fields if editing
        if (dialogIsEdit)
        {
            SetWindowTextW(hDlg, L"Edit Item");
            SetDlgItemTextW(hDlg, IDC_EDIT_NAME, dialogItem.name.c_str());
            SetDlgItemInt(hDlg, IDC_EDIT_BULK, dialogItem.bulkCount, FALSE);
            SetDlgItemInt(hDlg, IDC_EDIT_QTY, dialogItem.quantity, FALSE);

            wchar_t priceStr[32];
            swprintf_s(priceStr, L"%.2f", dialogItem.price);
            SetDlgItemTextW(hDlg, IDC_EDIT_PRICE, priceStr);

            SetDlgItemTextW(hDlg, IDC_EDIT_LOCATION, dialogItem.storageLocation.c_str());
        }
        else
        {
            SetWindowTextW(hDlg, L"Add New Item");
            SetDlgItemInt(hDlg, IDC_EDIT_QTY, 0, FALSE);
            SetDlgItemTextW(hDlg, IDC_EDIT_PRICE, L"0.00");
        }
        return TRUE;
    }

    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case IDC_DIALOG_OK:
        {
            // Get values from dialog
            wchar_t buffer[256];

            GetDlgItemTextW(hDlg, IDC_EDIT_NAME, buffer, 256);
            dialogItem.name = buffer;

            dialogItem.bulkCount = GetDlgItemInt(hDlg, IDC_EDIT_BULK, NULL, FALSE);

            dialogItem.quantity = GetDlgItemInt(hDlg, IDC_EDIT_QTY, NULL, FALSE);

            GetDlgItemTextW(hDlg, IDC_EDIT_PRICE, buffer, 256);
            dialogItem.price = _wtof(buffer);

            GetDlgItemTextW(hDlg, IDC_EDIT_LOCATION, buffer, 256);
            dialogItem.storageLocation = buffer;

            // Validate
            if (dialogItem.name.empty())
            {
                MessageBoxW(hDlg, L"Please enter an item name.", L"Validation Error", MB_OK | MB_ICONWARNING);
                return TRUE;
            }

            EndDialog(hDlg, IDOK);
            return TRUE;
        }

        case IDC_DIALOG_CANCEL:
            EndDialog(hDlg, IDCANCEL);
            return TRUE;
        }
        break;

    case WM_CLOSE:
        EndDialog(hDlg, IDCANCEL);
        return TRUE;
    }
    return FALSE;
}

// Global to track dialog result
static INT_PTR g_dialogResult = IDCANCEL;
static HWND g_hCurrentDialog = NULL;

// Custom dialog window procedure
LRESULT CALLBACK ItemDialogWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case IDC_DIALOG_OK:
        {
            // Get values from dialog
            wchar_t buffer[256];
            wchar_t linkBuffer[512];
            wchar_t notesBuffer[512];

            GetDlgItemTextW(hWnd, IDC_EDIT_NAME, buffer, 256);
            dialogItem.name = buffer;

            dialogItem.bulkCount = GetDlgItemInt(hWnd, IDC_EDIT_BULK, NULL, FALSE);

            dialogItem.quantity = GetDlgItemInt(hWnd, IDC_EDIT_QTY, NULL, FALSE);

            GetDlgItemTextW(hWnd, IDC_EDIT_PRICE, buffer, 256);
            dialogItem.price = _wtof(buffer);

            GetDlgItemTextW(hWnd, IDC_EDIT_LOCATION, buffer, 256);
            dialogItem.storageLocation = buffer;

            GetDlgItemTextW(hWnd, IDC_EDIT_LINK, linkBuffer, 512);
            dialogItem.orderLink = linkBuffer;

            GetDlgItemTextW(hWnd, IDC_EDIT_NOTES, notesBuffer, 512);
            dialogItem.notes = notesBuffer;

            // Validate
            if (dialogItem.name.empty())
            {
                MessageBoxW(hWnd, L"Please enter an item name.", L"Validation Error", MB_OK | MB_ICONWARNING);
                return 0;
            }

            g_dialogResult = IDOK;
            PostMessage(hWnd, WM_CLOSE, 0, 0);
            return 0;
        }

        case IDC_DIALOG_CANCEL:
            g_dialogResult = IDCANCEL;
            PostMessage(hWnd, WM_CLOSE, 0, 0);
            return 0;
        }
        break;

    case WM_CLOSE:
        // Re-enable parent and destroy
    {
        HWND hParent = GetParent(hWnd);
        if (hParent)
        {
            EnableWindow(hParent, TRUE);
            SetForegroundWindow(hParent);
        }
        g_hCurrentDialog = NULL;  // Signal the message loop to exit
        DestroyWindow(hWnd);
    }
    return 0;

    case WM_DESTROY:
        // Don't call PostQuitMessage here - that's only for the main window!
        // The dialog loop checks g_hCurrentDialog to know when to exit
        return 0;
    }
    return DefWindowProc(hWnd, message, wParam, lParam);
}

// Register dialog window class (call once at startup)
void RegisterDialogClass()
{
    static bool registered = false;
    if (registered) return;

    WNDCLASSEXW wcex = { 0 };
    wcex.cbSize = sizeof(WNDCLASSEX);
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = ItemDialogWndProc;
    wcex.hInstance = hInst;
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = (HBRUSH)(COLOR_3DFACE + 1);
    wcex.lpszClassName = L"OtioseItemDialog";

    RegisterClassExW(&wcex);
    registered = true;
}

// Create and show the dialog
HWND CreateItemDialog(HWND hParent)
{
    RegisterDialogClass();

    // Dialog dimensions
    int dlgWidth = 450;
    int dlgHeight = 400;
    int margin = 15;
    int labelWidth = 100;
    int editWidth = dlgWidth - labelWidth - margin * 3;
    int rowHeight = 25;
    int rowSpacing = 32;

    HWND hDlg = CreateWindowExW(
        WS_EX_DLGMODALFRAME,
        L"OtioseItemDialog",
        dialogIsEdit ? L"Edit Item" : L"Add New Item",
        WS_POPUP | WS_CAPTION | WS_SYSMENU,
        0, 0, dlgWidth, dlgHeight,
        hParent, NULL, hInst, NULL);

    if (!hDlg) return NULL;

    // Create a font for controls
    HFONT hFont = CreateFontW(14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Segoe UI");

    int y = margin;

    // Name field
    HWND hLabel = CreateWindowW(L"STATIC", L"Name:", WS_CHILD | WS_VISIBLE,
        margin, y + 3, labelWidth, rowHeight, hDlg, NULL, hInst, NULL);
    SendMessage(hLabel, WM_SETFONT, (WPARAM)hFont, TRUE);

    HWND hEdit = CreateWindowW(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
        margin + labelWidth, y, editWidth, rowHeight, hDlg, (HMENU)IDC_EDIT_NAME, hInst, NULL);
    SendMessage(hEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
    y += rowSpacing;

    // Bulk Count field
    hLabel = CreateWindowW(L"STATIC", L"Bulk Count:", WS_CHILD | WS_VISIBLE,
        margin, y + 3, labelWidth, rowHeight, hDlg, NULL, hInst, NULL);
    SendMessage(hLabel, WM_SETFONT, (WPARAM)hFont, TRUE);

    hEdit = CreateWindowW(L"EDIT", L"0", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER,
        margin + labelWidth, y, 80, rowHeight, hDlg, (HMENU)IDC_EDIT_BULK, hInst, NULL);
    SendMessage(hEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
    y += rowSpacing;

    // Quantity field
    hLabel = CreateWindowW(L"STATIC", L"Quantity:", WS_CHILD | WS_VISIBLE,
        margin, y + 3, labelWidth, rowHeight, hDlg, NULL, hInst, NULL);
    SendMessage(hLabel, WM_SETFONT, (WPARAM)hFont, TRUE);

    hEdit = CreateWindowW(L"EDIT", L"0", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER,
        margin + labelWidth, y, 80, rowHeight, hDlg, (HMENU)IDC_EDIT_QTY, hInst, NULL);
    SendMessage(hEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
    y += rowSpacing;

    // Price field
    hLabel = CreateWindowW(L"STATIC", L"Price ($):", WS_CHILD | WS_VISIBLE,
        margin, y + 3, labelWidth, rowHeight, hDlg, NULL, hInst, NULL);
    SendMessage(hLabel, WM_SETFONT, (WPARAM)hFont, TRUE);

    hEdit = CreateWindowW(L"EDIT", L"0.00", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
        margin + labelWidth, y, 80, rowHeight, hDlg, (HMENU)IDC_EDIT_PRICE, hInst, NULL);
    SendMessage(hEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
    y += rowSpacing;

    // Storage Location field
    hLabel = CreateWindowW(L"STATIC", L"Storage Location:", WS_CHILD | WS_VISIBLE,
        margin, y + 3, labelWidth, rowHeight, hDlg, NULL, hInst, NULL);
    SendMessage(hLabel, WM_SETFONT, (WPARAM)hFont, TRUE);

    hEdit = CreateWindowW(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
        margin + labelWidth, y, editWidth, rowHeight, hDlg, (HMENU)IDC_EDIT_LOCATION, hInst, NULL);
    SendMessage(hEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
    y += rowSpacing;

    // Order Link field (optional)
    hLabel = CreateWindowW(L"STATIC", L"Order Link:", WS_CHILD | WS_VISIBLE,
        margin, y + 3, labelWidth, rowHeight, hDlg, NULL, hInst, NULL);
    SendMessage(hLabel, WM_SETFONT, (WPARAM)hFont, TRUE);

    hEdit = CreateWindowW(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
        margin + labelWidth, y, editWidth, rowHeight, hDlg, (HMENU)IDC_EDIT_LINK, hInst, NULL);
    SendMessage(hEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
    y += rowSpacing;

    // Notes field
    hLabel = CreateWindowW(L"STATIC", L"Notes:", WS_CHILD | WS_VISIBLE,
        margin, y + 3, labelWidth, rowHeight, hDlg, NULL, hInst, NULL);
    SendMessage(hLabel, WM_SETFONT, (WPARAM)hFont, TRUE);

    hEdit = CreateWindowW(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
        margin + labelWidth, y, editWidth, rowHeight, hDlg, (HMENU)IDC_EDIT_NOTES, hInst, NULL);
    SendMessage(hEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
    y += rowSpacing + 10;

    // Buttons
    int btnWidth = 80;
    int btnHeight = 30;
    int btnX = dlgWidth - margin * 2 - btnWidth * 2 - 10;

    HWND hBtn = CreateWindowW(L"BUTTON", L"OK", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
        btnX, y, btnWidth, btnHeight, hDlg, (HMENU)IDC_DIALOG_OK, hInst, NULL);
    SendMessage(hBtn, WM_SETFONT, (WPARAM)hFont, TRUE);

    hBtn = CreateWindowW(L"BUTTON", L"Cancel", WS_CHILD | WS_VISIBLE,
        btnX + btnWidth + 10, y, btnWidth, btnHeight, hDlg, (HMENU)IDC_DIALOG_CANCEL, hInst, NULL);
    SendMessage(hBtn, WM_SETFONT, (WPARAM)hFont, TRUE);

    // Center dialog
    RECT rcParent, rcDlg;
    GetWindowRect(hParent, &rcParent);
    GetWindowRect(hDlg, &rcDlg);
    int x = rcParent.left + (rcParent.right - rcParent.left - (rcDlg.right - rcDlg.left)) / 2;
    int yPos = rcParent.top + (rcParent.bottom - rcParent.top - (rcDlg.bottom - rcDlg.top)) / 2;
    SetWindowPos(hDlg, NULL, x, yPos, 0, 0, SWP_NOSIZE | SWP_NOZORDER);

    return hDlg;
}

// Run a modal dialog loop
INT_PTR RunItemDialog(HWND hParent)
{
    g_dialogResult = IDCANCEL;

    HWND hDlg = CreateItemDialog(hParent);
    if (!hDlg) return IDCANCEL;

    g_hCurrentDialog = hDlg;

    // Populate fields if editing
    if (dialogIsEdit)
    {
        SetDlgItemTextW(hDlg, IDC_EDIT_NAME, dialogItem.name.c_str());
        SetDlgItemInt(hDlg, IDC_EDIT_BULK, dialogItem.bulkCount, FALSE);
        SetDlgItemInt(hDlg, IDC_EDIT_QTY, dialogItem.quantity, FALSE);

        wchar_t priceStr[32];
        swprintf_s(priceStr, L"%.2f", dialogItem.price);
        SetDlgItemTextW(hDlg, IDC_EDIT_PRICE, priceStr);

        SetDlgItemTextW(hDlg, IDC_EDIT_LOCATION, dialogItem.storageLocation.c_str());
        SetDlgItemTextW(hDlg, IDC_EDIT_LINK, dialogItem.orderLink.c_str());
        SetDlgItemTextW(hDlg, IDC_EDIT_NOTES, dialogItem.notes.c_str());
    }

    // Disable parent and show dialog
    EnableWindow(hParent, FALSE);
    ShowWindow(hDlg, SW_SHOW);
    UpdateWindow(hDlg);

    // Run modal message loop for this dialog
    MSG msg;
    while (g_hCurrentDialog && GetMessage(&msg, NULL, 0, 0))
    {
        // Handle Tab key navigation
        if (!IsDialogMessage(hDlg, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }

        // Check if dialog was closed
        if (!IsWindow(hDlg) || !g_hCurrentDialog)
        {
            break;
        }
    }

    return g_dialogResult;
}

// Show the Add Item dialog
void ShowAddItemDialog(HWND hWnd)
{
    dialogIsEdit = false;
    dialogItem = InventoryItem();  // Clear the item

    if (RunItemDialog(hWnd) == IDOK)
    {
        // Set creation timestamp
        dialogItem.lastModified = time(nullptr);

        // Add the new item to current category
        std::vector<InventoryItem>& inventory = GetCurrentInventory();
        inventory.push_back(dialogItem);

        // Save and refresh
        SaveInventoryToCSV();
        RefreshInventoryList();
    }
}

// Show the Edit Item dialog
void ShowEditItemDialog(HWND hWnd)
{
    if (selectedItemIndex < 0) return;

    // Get the actual index in the inventory (accounts for sorting/filtering)
    int actualIndex = GetSelectedInventoryIndex();

    std::vector<InventoryItem>& inventory = GetCurrentInventory();

    if (actualIndex < 0 || actualIndex >= (int)inventory.size())
    {
        MessageBoxW(hWnd, L"Please select a valid item to edit.", L"Edit Item", MB_OK | MB_ICONINFORMATION);
        return;
    }

    dialogIsEdit = true;
    dialogItem = inventory[actualIndex];

    if (RunItemDialog(hWnd) == IDOK)
    {
        // Update the timestamp
        dialogItem.lastModified = time(nullptr);

        // Update the item
        inventory[actualIndex] = dialogItem;

        // Save and refresh
        SaveInventoryToCSV();
        RefreshInventoryList();
    }
}

// Delete the selected item
void DeleteSelectedItem(HWND hWnd)
{
    if (selectedItemIndex < 0) return;

    // Get the actual index in the inventory (accounts for sorting/filtering)
    int actualIndex = GetSelectedInventoryIndex();

    std::vector<InventoryItem>& inventory = GetCurrentInventory();

    if (actualIndex < 0 || actualIndex >= (int)inventory.size())
    {
        MessageBoxW(hWnd, L"Please select a valid item to delete.", L"Delete Item", MB_OK | MB_ICONINFORMATION);
        return;
    }

    // Confirm deletion
    std::wstring msg = L"Are you sure you want to delete \"" + inventory[actualIndex].name + L"\"?";
    if (MessageBoxW(hWnd, msg.c_str(), L"Confirm Delete", MB_YESNO | MB_ICONQUESTION) == IDYES)
    {
        inventory.erase(inventory.begin() + actualIndex);

        // Save and refresh
        SaveInventoryToCSV();
        RefreshInventoryList();
    }
}