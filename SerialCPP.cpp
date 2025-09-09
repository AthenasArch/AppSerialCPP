// main.cpp - Terminal Serial Win32 (corrigido e robusto)
#include <windows.h>
#include <commctrl.h>
#include <tchar.h>
#include <strsafe.h>
#include <string>
#include <vector>
#include <sstream>
#include <setupapi.h>
#include <initguid.h>
#include <devguid.h>
#include <regstr.h>
#include <thread>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "setupapi.lib")

#define ID_COMBOBOX_COMPORT 101
#define ID_COMBOBOX_BAUDRATE 102
#define ID_BTN_CONNECT 103
#define ID_BTN_SEND 104
#define ID_TERMINAL 105
#define ID_EDIT_SEND1 106
#define ID_EDIT_SEND2 107
#define ID_RADIO_SEND1 108
#define ID_RADIO_SEND2 109

HWND hComboComPort, hComboBaudRate, hBtnConnect, hBtnSend;
HWND hTerminal, hEditSend1, hEditSend2;
HWND hRadioSend1, hRadioSend2;
HANDLE hSerial = INVALID_HANDLE_VALUE;
bool isConnected = false;
bool isReceiving = false;
std::thread serialThread;

// ==== Utils UTF-8 <-> UTF-16 ====
static std::string WideToUtf8(const std::wstring& w) {
    if (w.empty()) return {};
    int len = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), nullptr, 0, nullptr, nullptr);
    std::string out(len, 0);
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), &out[0], len, nullptr, nullptr);
    return out;
}
static std::wstring Utf8ToWide(const char* data, int bytes) {
    if (bytes <= 0) return {};
    int wlen = MultiByteToWideChar(CP_UTF8, 0, data, bytes, nullptr, 0);
    std::wstring w(wlen, 0);
    MultiByteToWideChar(CP_UTF8, 0, data, bytes, &w[0], wlen);
    return w;
}

// ==== UI ====
void AppendToTerminal(const std::wstring& text) {
    int len = GetWindowTextLength(hTerminal);
    SendMessage(hTerminal, EM_SETSEL, len, len);
    SendMessage(hTerminal, EM_REPLACESEL, FALSE, (LPARAM)text.c_str());
}

// ==== Serial config helpers ====
static void SetupSerialTimeouts_BlockingOnChars(HANDLE h) {
    // Leitura "quase bloqueante" até chegar algo, sem travar por muito tempo.
    COMMTIMEOUTS t = {};
    t.ReadIntervalTimeout = 1;          // intervalo entre bytes
    t.ReadTotalTimeoutMultiplier = 0;   // por byte
    t.ReadTotalTimeoutConstant = 50;    // timeout total (ms) por chamada
    t.WriteTotalTimeoutMultiplier = 0;
    t.WriteTotalTimeoutConstant = 100;
    SetCommTimeouts(h, &t);
}

static bool ConfigurePort(HANDLE h, DWORD baud) {
    // Buffer interno de driver
    SetupComm(h, 4096, 4096);

    // Limpa qualquer lixo pendente
    PurgeComm(h, PURGE_RXCLEAR | PURGE_TXCLEAR | PURGE_RXABORT | PURGE_TXABORT);

    DCB dcb = {};
    dcb.DCBlength = sizeof(DCB);
    if (!GetCommState(h, &dcb)) return false;

    dcb.BaudRate = baud;
    dcb.fBinary = TRUE;
    dcb.fParity = FALSE;
    dcb.ByteSize = 8;
    dcb.Parity = NOPARITY;
    dcb.StopBits = ONESTOPBIT;

    // Sem handshake por padrão; habilite se precisar
    dcb.fOutxCtsFlow = FALSE;
    dcb.fOutxDsrFlow = FALSE;
    dcb.fOutX = FALSE;
    dcb.fInX = FALSE;
    dcb.fDtrControl = DTR_CONTROL_ENABLE;  // muitos dispositivos precisam
    dcb.fRtsControl = RTS_CONTROL_ENABLE;  // idem

    if (!SetCommState(h, &dcb)) return false;

    SetupSerialTimeouts_BlockingOnChars(h);
    // Sobe DTR/RTS explicitamente (opcional, alguns devices precisam)
    EscapeCommFunction(h, SETDTR);
    EscapeCommFunction(h, SETRTS);

    return true;
}

// ==== Serial I/O loop ====
void SerialReadLoop() {
    const DWORD BUF = 1024;
    char buffer[BUF];
    DWORD bytesRead = 0;

    while (isReceiving && hSerial != INVALID_HANDLE_VALUE) {
        // Tenta ler um bloco; retorna após timeout curto se não houver dados
        if (ReadFile(hSerial, buffer, BUF, &bytesRead, nullptr)) {
            if (bytesRead > 0) {
                // Converte exatamente "bytesRead" bytes de UTF-8 -> wide
                std::wstring w = Utf8ToWide(buffer, (int)bytesRead);
                if (!w.empty()) {
                    AppendToTerminal(L"[RX] " + w);
                }
                else {
                    // Dados binários/bytes não UTF-8? Mostra em HEX simples
                    std::wstringstream ss;
                    ss << L"[RX HEX] ";
                    for (DWORD i = 0; i < bytesRead; ++i) {
                        wchar_t tmp[8];
                        StringCchPrintf(tmp, 8, L"%02X ", (unsigned char)buffer[i]);
                        ss << tmp;
                    }
                    ss << L"\r\n";
                    AppendToTerminal(ss.str());
                }
            }
        }
        else {
            // Erro transitório/porta desconectada?
            DWORD err = GetLastError();
            if (err == ERROR_OPERATION_ABORTED) break; // ao fechar
            // pequena pausa para não ficar em busy loop em caso de erro
            Sleep(5);
        }
    }
}

// ==== UI: listar portas ====
void ListComPorts(HWND hComboBox) {
    SendMessage(hComboBox, CB_RESETCONTENT, 0, 0);
    HDEVINFO hDevInfo = SetupDiGetClassDevs(&GUID_DEVCLASS_PORTS, nullptr, nullptr, DIGCF_PRESENT);
    if (hDevInfo == INVALID_HANDLE_VALUE) return;

    SP_DEVINFO_DATA devInfoData = { sizeof(SP_DEVINFO_DATA) };
    for (DWORD i = 0; SetupDiEnumDeviceInfo(hDevInfo, i, &devInfoData); i++) {
        TCHAR buffer[256];
        if (SetupDiGetDeviceRegistryProperty(hDevInfo, &devInfoData, SPDRP_FRIENDLYNAME, nullptr, (PBYTE)buffer, sizeof(buffer), nullptr)) {
            std::wstring name(buffer);
            if (name.find(L"(COM") != std::wstring::npos) {
                SendMessage(hComboBox, CB_ADDSTRING, 0, (LPARAM)name.c_str());
            }
        }
    }
    SetupDiDestroyDeviceInfoList(hDevInfo);
    SendMessage(hComboBox, CB_SETCURSEL, 0, 0);
}

void PopulateBaudRates(HWND hComboBox) {
    const std::vector<std::wstring> baudRates = { L"9600", L"19200", L"38400", L"57600", L"115200", L"230400", L"460800", L"921600" };
    for (const auto& rate : baudRates)
        SendMessage(hComboBox, CB_ADDSTRING, 0, (LPARAM)rate.c_str());
    SendMessage(hComboBox, CB_SETCURSEL, 4, 0); // 115200 por padrão
}

bool OpenSerialPort(const std::wstring& portName, DWORD baudRate) {
    std::wstring full = L"\\\\.\\" + portName; // COM10+ requer \\.\ prefixo
    hSerial = CreateFileW(full.c_str(), GENERIC_READ | GENERIC_WRITE, 0, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hSerial == INVALID_HANDLE_VALUE) return false;
    if (!ConfigurePort(hSerial, baudRate)) {
        CloseHandle(hSerial);
        hSerial = INVALID_HANDLE_VALUE;
        return false;
    }
    isConnected = true;
    isReceiving = true;
    serialThread = std::thread(SerialReadLoop);
    return true;
}

void CloseSerialPort() {
    isReceiving = false;
    if (hSerial != INVALID_HANDLE_VALUE) {
        // aborta I/O bloqueante
        CancelIo(hSerial);
    }
    if (serialThread.joinable()) serialThread.join();
    if (isConnected && hSerial != INVALID_HANDLE_VALUE) {
        CloseHandle(hSerial);
        hSerial = INVALID_HANDLE_VALUE;
    }
    isConnected = false;
}

void SendSelectedMessage() {
    if (!isConnected) {
        MessageBox(nullptr, L"Conecte-se a uma porta COM primeiro.", L"Erro", MB_OK | MB_ICONERROR);
        return;
    }
    // Pega texto wide
    wchar_t buf1[1024] = {}, buf2[1024] = {};
    GetWindowTextW(hEditSend1, buf1, 1024);
    GetWindowTextW(hEditSend2, buf2, 1024);
    std::wstring wmsg = (SendMessage(hRadioSend1, BM_GETCHECK, 0, 0) == BST_CHECKED) ? buf1 : buf2;

    // Opcional: adicionar CRLF
    // wmsg += L"\r\n";

    // Converte para UTF-8 e envia bytes
    std::string bytes = WideToUtf8(wmsg);
    DWORD written = 0;
    if (WriteFile(hSerial, bytes.data(), (DWORD)bytes.size(), &written, nullptr)) {
        AppendToTerminal(L"[TX] " + wmsg + L"\r\n");
    }
    else {
        AppendToTerminal(L"[ERRO TX]\r\n");
    }
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE:
        hComboComPort = CreateWindowW(L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST, 10, 10, 220, 200, hwnd, (HMENU)ID_COMBOBOX_COMPORT, nullptr, nullptr);
        hComboBaudRate = CreateWindowW(L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST, 240, 10, 140, 200, hwnd, (HMENU)ID_COMBOBOX_BAUDRATE, nullptr, nullptr);
        hBtnConnect = CreateWindowW(L"BUTTON", L"Conectar", WS_CHILD | WS_VISIBLE, 10, 40, 100, 26, hwnd, (HMENU)ID_BTN_CONNECT, nullptr, nullptr);

        hEditSend1 = CreateWindowW(L"EDIT", nullptr, WS_CHILD | WS_VISIBLE | WS_BORDER, 10, 80, 370, 22, hwnd, (HMENU)ID_EDIT_SEND1, nullptr, nullptr);
        hEditSend2 = CreateWindowW(L"EDIT", nullptr, WS_CHILD | WS_VISIBLE | WS_BORDER, 10, 130, 370, 22, hwnd, (HMENU)ID_EDIT_SEND2, nullptr, nullptr);
        hRadioSend1 = CreateWindowW(L"BUTTON", L"Usar Caixa 1", WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON, 10, 105, 120, 20, hwnd, (HMENU)ID_RADIO_SEND1, nullptr, nullptr);
        hRadioSend2 = CreateWindowW(L"BUTTON", L"Usar Caixa 2", WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON, 140, 105, 120, 20, hwnd, (HMENU)ID_RADIO_SEND2, nullptr, nullptr);
        SendMessage(hRadioSend1, BM_SETCHECK, BST_CHECKED, 0);

        hBtnSend = CreateWindowW(L"BUTTON", L"Enviar", WS_CHILD | WS_VISIBLE, 280, 105, 100, 25, hwnd, (HMENU)ID_BTN_SEND, nullptr, nullptr);

        hTerminal = CreateWindowW(L"EDIT", nullptr, WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY,
            10, 160, 370, 260, hwnd, (HMENU)ID_TERMINAL, nullptr, nullptr);

        ListComPorts(hComboComPort);
        PopulateBaudRates(hComboBaudRate);
        break;

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case ID_BTN_CONNECT: {
            if (isConnected) {
                CloseSerialPort();
                SetWindowTextW(hBtnConnect, L"Conectar");
                AppendToTerminal(L"[INFO] Porta desconectada\r\n");
            }
            else {
                wchar_t portName[100] = {};
                wchar_t baudStr[20] = {};
                GetWindowTextW(hComboComPort, portName, 100);
                GetWindowTextW(hComboBaudRate, baudStr, 20);
                DWORD baud = _wtoi(baudStr);

                std::wstring portStr(portName);
                size_t b = portStr.find(L"(COM");
                size_t e = portStr.find(L")", b == std::wstring::npos ? 0 : b);
                std::wstring portOnly = (b != std::wstring::npos && e != std::wstring::npos && e > b) ?
                    portStr.substr(b + 1, e - (b + 1)) : portStr; // "COMx"

                if (OpenSerialPort(portOnly, baud)) {
                    SetWindowTextW(hBtnConnect, L"Desconectar");
                    AppendToTerminal(L"[OK] Conectado em " + portOnly + L" @ " + std::to_wstring(baud) + L"\r\n");
                }
                else {
                    AppendToTerminal(L"[ERRO] Falha ao conectar\r\n");
                }
            }
        } break;

        case ID_BTN_SEND:
            SendSelectedMessage();
            break;
        }
        break;

    case WM_DESTROY:
        CloseSerialPort();
        PostQuitMessage(0);
        break;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE, LPSTR, int nCmdShow) {
    WNDCLASSW wc = {};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"SerialApp";
    RegisterClassW(&wc);

    DWORD style = (WS_OVERLAPPEDWINDOW & ~WS_THICKFRAME); // remove redimensionar
    HWND hwnd = CreateWindowW(L"SerialApp", L"Terminal Serial (Win32)",
        style, CW_USEDEFAULT, CW_USEDEFAULT, 400, 460, nullptr, nullptr, hInstance, nullptr);

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return 0;
}
