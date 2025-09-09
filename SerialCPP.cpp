// main.cpp - Terminal Serial Win32 (corrigido e robusto)
// Objetivo: terminal serial simples em Win32 puro, com envio (UTF-8) e recepção,
//           conversão correta UTF-8<->UTF-16, configuração sólida de porta
//           (DTR/RTS, timeouts, purge) e UI responsiva.

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

#define USE_TERMINAL_DEBUG

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "setupapi.lib")

// ---- IDs dos controles da janela ----
#define ID_COMBOBOX_COMPORT 101
#define ID_COMBOBOX_BAUDRATE 102
#define ID_BTN_CONNECT       103
#define ID_BTN_SEND          104
#define ID_TERMINAL          105
#define ID_EDIT_SEND1        106
#define ID_EDIT_SEND2        107
#define ID_RADIO_SEND1       108
#define ID_RADIO_SEND2       109

// ---- Handles globais dos controles ----
HWND hComboComPort, hComboBaudRate, hBtnConnect, hBtnSend;
HWND hTerminal, hEditSend1, hEditSend2;
HWND hRadioSend1, hRadioSend2;

// ---- Estado da serial ----
HANDLE hSerial = INVALID_HANDLE_VALUE;
bool isConnected = false;   // conectado à COM?
bool isReceiving = false;   // thread de RX rodando?
std::thread serialThread;   // thread de leitura assíncrona


static std::string WideToUtf8(const std::wstring& w);
static std::wstring Utf8ToWide(const char* data, int bytes);
void AppendToTerminal(const std::wstring& text);
static void SetupSerialTimeouts_BlockingOnChars(HANDLE h);
static bool ConfigurePort(HANDLE h, DWORD baud);
void SerialReadLoop();
void ListComPorts(HWND hComboBox);
void PopulateBaudRates(HWND hComboBox);
bool OpenSerialPort(const std::wstring& portName, DWORD baudRate);
void CloseSerialPort();
void SendSelectedMessage();
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
static void InitDebugConsole(void);
void RefreshComPortsAndKeepSelection(HWND hComboBox);


// ============================================================================
//                                 WinMain
// ============================================================================
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE, LPSTR, int nCmdShow) {
    // Registra classe da janela principal:
    
#ifdef USE_TERMINAL_DEBUG
    InitDebugConsole(); 
    printf("\r\r\n\n**** Inicializando Interface Serial CPP ****\r\r\n\n");
#endif /* USE_TERMINAL_DEBUG */

    WNDCLASSW wc = {};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"SerialApp";
    RegisterClassW(&wc);

    // Janela fixa (sem redimensionar) para simplificar layout:
    DWORD style = (WS_OVERLAPPEDWINDOW & ~WS_THICKFRAME);

    HWND hwnd = CreateWindowW(
        L"SerialApp", L"Terminal Serial (Win32)",
        style,
        CW_USEDEFAULT, CW_USEDEFAULT, 400, 460,
        nullptr, nullptr, hInstance, nullptr);

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    // Loop de mensagens padrão:
    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return 0;
}

/*
    O modo Win32 GUI (subsystem WINDOWS), não existe console anexado por padrão, entao, apenas 
    Para debugar informações, estou forçandoa criação de um terminal.
*/
static void InitDebugConsole(void)
{
    // cria console
    AllocConsole();

    // redireciona stdout/stderr para o console
    FILE* fp;
    freopen_s(&fp, "CONOUT$", "w", stdout);
    freopen_s(&fp, "CONOUT$", "w", stderr);

    // opcional: também ler do console
    freopen_s(&fp, "CONIN$", "r", stdin);

    // em Unicode, você pode ajustar o modo wide:
    // _setmode(_fileno(stdout), _O_U16TEXT);
}


// ============================================================================
//                        Utils de encoding UTF-8 <-> UTF-16
// ============================================================================
// Observação: a UI Win32 usa UTF-16 (wide). Dispositivos seriais, geralmente,
//             trocam bytes "crus" (texto ASCII/UTF-8). Aqui padronizamos:
//
//  - No ENVIO:   wide (UI) -> UTF-8 (bytes) -> WriteFile
//  - Na LEITURA: bytes -> tentamos decodificar como UTF-8 -> wide (UI)
//                Se não for texto válido, mostramos em HEX.

static std::string WideToUtf8(const std::wstring& w) {
    if (w.empty()) return {};
    // Tamanho necessário em bytes:
    int len = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), nullptr, 0, nullptr, nullptr);
    std::string out(len, 0);
    // Converte exatamente 'size' (sem depender de NUL no meio):
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), &out[0], len, nullptr, nullptr);
    return out;
}

static std::wstring Utf8ToWide(const char* data, int bytes) {
    if (bytes <= 0) return {};
    // Decodifica exatamente 'bytes' recebidos (não usamos -1):
    int wlen = MultiByteToWideChar(CP_UTF8, 0, data, bytes, nullptr, 0);
    std::wstring w(wlen, 0);
    MultiByteToWideChar(CP_UTF8, 0, data, bytes, &w[0], wlen);
    return w;
}

// ============================================================================
//                                UI helpers
// ============================================================================

void AppendToTerminal(const std::wstring& text) {
    // Acrescenta no fim do EDIT multi-line (read-only) sem apagar conteúdo.
    int len = GetWindowTextLength(hTerminal);
    SendMessage(hTerminal, EM_SETSEL, len, len);
    // EM_REPLACESEL aceita texto wide diretamente.
    SendMessage(hTerminal, EM_REPLACESEL, FALSE, (LPARAM)text.c_str());
}

// ============================================================================
//                        Configuração/Timeouts da Serial
// ============================================================================
static void SetupSerialTimeouts_BlockingOnChars(HANDLE h) {
    // Modelo de leitura: "quase bloqueante" com timeout curto.
    // - Se há bytes chegando, ReadFile retorna rapidamente.
    // - Se não há, devolve após ~50ms, permitindo a thread checar flags/fechamento.
    COMMTIMEOUTS t = {};
    t.ReadIntervalTimeout = 1;    // tolera intervalos pequenos entre bytes
    t.ReadTotalTimeoutMultiplier = 0;    // sem multiplicador por byte
    t.ReadTotalTimeoutConstant = 50;   // timeout total por chamada (ms)
    t.WriteTotalTimeoutMultiplier = 0;
    t.WriteTotalTimeoutConstant = 100;  // timeout de escrita conservador
    SetCommTimeouts(h, &t);
}

static bool ConfigurePort(HANDLE h, DWORD baud) {
    // Sugere buffers internos do driver (entrada/saída):
    SetupComm(h, 4096, 4096);

    // Limpa buffers e aborta I/O pendente para começar "do zero":
    PurgeComm(h, PURGE_RXCLEAR | PURGE_TXCLEAR | PURGE_RXABORT | PURGE_TXABORT);

    // Carrega configuração atual, para então ajustar:
    DCB dcb = {};
    dcb.DCBlength = sizeof(DCB);
    if (!GetCommState(h, &dcb)) return false;

    // Parâmetros clássicos 8N1 a 'baud':
    dcb.BaudRate = baud;
    dcb.fBinary = TRUE;       // obrigatório
    dcb.fParity = FALSE;      // sem bit de paridade
    dcb.ByteSize = 8;
    dcb.Parity = NOPARITY;
    dcb.StopBits = ONESTOPBIT;

    // Handshake desativado por padrão (CTS/DSR/XON/XOFF):
    // Se precisar (ex.: modem/rádio), habilite conforme o hardware.
    dcb.fOutxCtsFlow = FALSE;
    dcb.fOutxDsrFlow = FALSE;
    dcb.fOutX = FALSE;
    dcb.fInX = FALSE;

    // MUITOS adaptadores/boards só transmitem com DTR/RTS ativos:
    dcb.fDtrControl = DTR_CONTROL_ENABLE;
    dcb.fRtsControl = RTS_CONTROL_ENABLE;

    if (!SetCommState(h, &dcb)) return false;

    // Define timeouts de I/O:
    SetupSerialTimeouts_BlockingOnChars(h);

    // Garante linhas DTR/RTS em nível alto (redundante ao SetCommState, mas seguro):
    EscapeCommFunction(h, SETDTR);
    EscapeCommFunction(h, SETRTS);

    return true;
}

// ============================================================================
//                              Thread de Leitura
// ============================================================================
// Lê blocos do driver serial, tenta decodificar como UTF-8 e joga no terminal.
// Se os bytes não formarem texto UTF-8 válido, exibe em HEX (debug útil p/ binário).
void SerialReadLoop() {
    const DWORD BUF = 1024;
    char  buffer[BUF];
    DWORD bytesRead = 0;

    while (isReceiving && hSerial != INVALID_HANDLE_VALUE) {
        // ReadFile retorna imediatamente se houver dados, ou após ~50ms se não houver.
        if (ReadFile(hSerial, buffer, BUF, &bytesRead, nullptr)) {
            if (bytesRead > 0) {
                // Tenta decodificar exatamente 'bytesRead' como UTF-8:
                std::wstring w = Utf8ToWide(buffer, (int)bytesRead);
                if (!w.empty()) {
                    // Mostra como texto (sem forçar CRLF: respeita o que veio do device)
                    AppendToTerminal(L"[RX] " + w);
                }
                else {
                    // UTF-8 inválido: mostra payload em HEX (ex.: dados binários/frames)
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
            // Falha de leitura: pode ser fechamento (CancelIo) ou erro transitório.
            DWORD err = GetLastError();
            if (err == ERROR_OPERATION_ABORTED) break; // saímos ao fechar a porta
            // Evita busy loop caso o driver esteja sinalizando erro repetidamente:
            Sleep(5);
        }
    }
}

// ============================================================================
//                       UI: Descoberta/Listagem de portas
// ============================================================================
// Percorre a classe de dispositivos "Ports (COM & LPT)" e joga os friendly names
// no ComboBox (ex.: "USB-Serial (COM6)"). Isso facilita extrair "COMx" depois.
void ListComPorts(HWND hComboBox) {
    SendMessage(hComboBox, CB_RESETCONTENT, 0, 0);
    printf("Vai listar as portas COM e LPT conectadas.\r\n");
    HDEVINFO hDevInfo = SetupDiGetClassDevs(&GUID_DEVCLASS_PORTS, nullptr, nullptr, DIGCF_PRESENT);
    if (hDevInfo == INVALID_HANDLE_VALUE) return;

    SP_DEVINFO_DATA devInfoData = { sizeof(SP_DEVINFO_DATA) };
    for (DWORD i = 0; SetupDiEnumDeviceInfo(hDevInfo, i, &devInfoData); i++) {
        TCHAR buffer[256];
        if (SetupDiGetDeviceRegistryProperty(hDevInfo, &devInfoData, SPDRP_FRIENDLYNAME,
            nullptr, (PBYTE)buffer, sizeof(buffer), nullptr)) {
            std::wstring name(buffer);
            if (name.find(L"(COM") != std::wstring::npos) {
                SendMessage(hComboBox, CB_ADDSTRING, 0, (LPARAM)name.c_str());
            }
        }
    }
    SetupDiDestroyDeviceInfoList(hDevInfo);

    // Seleciona o primeiro item por padrão (se houver):
    SendMessage(hComboBox, CB_SETCURSEL, 0, 0);
}

void RefreshComPortsAndKeepSelection(HWND hComboBox) {
    // guarda texto selecionado (se houver)
    int sel = (int)SendMessage(hComboBox, CB_GETCURSEL, 0, 0);
    std::wstring prev;
    if (sel != CB_ERR) {
        int len = (int)SendMessage(hComboBox, CB_GETLBTEXTLEN, sel, 0);
        if (len > 0) {
            prev.resize(len);
            SendMessage(hComboBox, CB_GETLBTEXT, sel, (LPARAM)prev.data());
        }
    }

    // evita flicker enquanto repovoa
    SendMessage(hComboBox, WM_SETREDRAW, FALSE, 0);
    ListComPorts(hComboBox);

    // tenta restaurar seleção anterior
    if (!prev.empty()) {
        int idx = (int)SendMessage(hComboBox, CB_FINDSTRINGEXACT, (WPARAM)-1, (LPARAM)prev.c_str());
        if (idx != CB_ERR) {
            SendMessage(hComboBox, CB_SETCURSEL, idx, 0);
        }
        else {
            // se não achou, selecione o primeiro se existir
            if (SendMessage(hComboBox, CB_GETCOUNT, 0, 0) > 0) {
                SendMessage(hComboBox, CB_SETCURSEL, 0, 0);
            }
        }
    }
    else {
        // sem seleção prévia: selecione o primeiro se existir
        if (SendMessage(hComboBox, CB_GETCOUNT, 0, 0) > 0) {
            SendMessage(hComboBox, CB_SETCURSEL, 0, 0);
        }
    }

    SendMessage(hComboBox, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(hComboBox, nullptr, TRUE);
}


// Preenche baud rates comuns (inclui alguns altos para adaptadores modernos).
void PopulateBaudRates(HWND hComboBox) {
    const std::vector<std::wstring> baudRates = {
        L"9600", L"19200", L"38400", L"57600",
        L"115200", L"230400", L"460800", L"921600"
    };
    for (const auto& rate : baudRates)
        SendMessage(hComboBox, CB_ADDSTRING, 0, (LPARAM)rate.c_str());

    // 115200 é um padrão razoável:
    SendMessage(hComboBox, CB_SETCURSEL, 4, 0);
}

// Abre a porta (ex.: "COM6") com CreateFile e aplica configuração via DCB.
// Observação: portas COM >= 10 exigem o prefixo "\\.\".
bool OpenSerialPort(const std::wstring& portName, DWORD baudRate) {
    std::wstring full = L"\\\\.\\" + portName; // COM10+ requer \\.\ prefixo

    hSerial = CreateFileW(full.c_str(),
        GENERIC_READ | GENERIC_WRITE,   // leitura e escrita
        0,                              // sem compartilhamento
        nullptr, OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hSerial == INVALID_HANDLE_VALUE) return false;

    if (!ConfigurePort(hSerial, baudRate)) {
        CloseHandle(hSerial);
        hSerial = INVALID_HANDLE_VALUE;
        return false;
    }

    // Marca estado e inicia thread de recepção:
    isConnected = true;
    isReceiving = true;
    serialThread = std::thread(SerialReadLoop);
    return true;
}




// Fecha a porta com segurança:
// - sinaliza a thread para parar (isReceiving=false)
// - aborta ReadFile pendente (CancelIo)
// - junta a thread (join)
// - fecha handle
void CloseSerialPort() {
    isReceiving = false;

    if (hSerial != INVALID_HANDLE_VALUE) {
        CancelIo(hSerial); // faz ReadFile retornar com ERROR_OPERATION_ABORTED
    }
    if (serialThread.joinable())
        serialThread.join();

    if (isConnected && hSerial != INVALID_HANDLE_VALUE) {
        CloseHandle(hSerial);
        hSerial = INVALID_HANDLE_VALUE;
    }
    isConnected = false;
}

// Envia o conteúdo de uma das caixas de texto.
// Fluxo: wide (UI) -> UTF-8 (bytes) -> WriteFile.
// Opcionalmente, pode-se anexar "\r\n" se o dispositivo exigir ENTER.
void SendSelectedMessage() {
    if (!isConnected) {
        MessageBox(nullptr, L"Conecte-se a uma porta COM primeiro.", L"Erro", MB_OK | MB_ICONERROR);
        return;
    }

    // Captura o texto das caixas:
    wchar_t buf1[1024] = {}, buf2[1024] = {};
    GetWindowTextW(hEditSend1, buf1, 1024);
    GetWindowTextW(hEditSend2, buf2, 1024);

    // Qual caixa usar?
    std::wstring wmsg = (SendMessage(hRadioSend1, BM_GETCHECK, 0, 0) == BST_CHECKED) ? buf1 : buf2;

    // (Opcional) Se quiser garantir quebra de linha no device:
    // wmsg += L"\r\n";

    // Converte para UTF-8 e envia:
    std::string bytes = WideToUtf8(wmsg);
    DWORD written = 0;
    if (WriteFile(hSerial, bytes.data(), (DWORD)bytes.size(), &written, nullptr)) {
        AppendToTerminal(L"[TX] " + wmsg + L"\r\n");
    }
    else {
        AppendToTerminal(L"[ERRO TX]\r\n");
    }
}




// ============================================================================
//                              Janela / Mensageria
// ============================================================================
// Esta função recebe TODAS as mensagens destinadas à janela principal.
// Ela é chamada pelo loop de mensagens (GetMessage/DispatchMessage) no WinMain.
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {

        // ------------------------------------------------------------------------
        // Mensagem enviada quando a janela é criada (após CreateWindow/ShowWindow).
        // Momento ideal para criar CONTROLES FILHOS (botões, edits, combos, etc.).
        // ------------------------------------------------------------------------
    case WM_CREATE:

        // ---- Criação dos controles (COMBOBOX de portas COM) ----
        // - WS_CHILD | WS_VISIBLE  -> controle é filho da janela e visível.
        // - CBS_DROPDOWNLIST       -> estilo "somente seleção" (sem edição).
        // Parâmetros (x, y, w, h) posicionam o controle na janela.
        hComboComPort = CreateWindowW(
            L"COMBOBOX", nullptr,
            WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
            10, 10, 220, 200,                // posição e tamanho
            hwnd, (HMENU)ID_COMBOBOX_COMPORT,// ID único do controle
            nullptr, nullptr);                // sem menu/extra data

        // ---- COMBOBOX de baud rates ----
        hComboBaudRate = CreateWindowW(
            L"COMBOBOX", nullptr,
            WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
            240, 10, 140, 200,
            hwnd, (HMENU)ID_COMBOBOX_BAUDRATE,
            nullptr, nullptr);

        // ---- Botão "Conectar" ----
        hBtnConnect = CreateWindowW(
            L"BUTTON", L"Conectar",
            WS_CHILD | WS_VISIBLE,
            10, 40, 100, 26,
            hwnd, (HMENU)ID_BTN_CONNECT,
            nullptr, nullptr);

        // ---- Caixa de texto para envio 1 ----
        // WS_BORDER dá borda fina; é um EDIT de linha única (sem ES_MULTILINE).
        hEditSend1 = CreateWindowW(
            L"EDIT", nullptr,
            WS_CHILD | WS_VISIBLE | WS_BORDER,
            10, 80, 370, 22,
            hwnd, (HMENU)ID_EDIT_SEND1,
            nullptr, nullptr);

        // ---- Caixa de texto para envio 2 ----
        hEditSend2 = CreateWindowW(
            L"EDIT", nullptr,
            WS_CHILD | WS_VISIBLE | WS_BORDER,
            10, 130, 370, 22,
            hwnd, (HMENU)ID_EDIT_SEND2,
            nullptr, nullptr);

        // ---- Rádio para escolher a caixa 1 como fonte do envio ----
        // BS_AUTORADIOBUTTON faz o Windows desmarcar o outro automaticamente
        // quando marcamos este (desde que estejam no mesmo "grupo" de rádios).
        hRadioSend1 = CreateWindowW(
            L"BUTTON", L"Usar Caixa 1",
            WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON,
            10, 105, 120, 20,
            hwnd, (HMENU)ID_RADIO_SEND1,
            nullptr, nullptr);

        // ---- Rádio para escolher a caixa 2 ----
        hRadioSend2 = CreateWindowW(
            L"BUTTON", L"Usar Caixa 2",
            WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON,
            140, 105, 120, 20,
            hwnd, (HMENU)ID_RADIO_SEND2,
            nullptr, nullptr);

        // Define a Caixa 1 como seleção inicial.
        SendMessage(hRadioSend1, BM_SETCHECK, BST_CHECKED, 0);

        // ---- Botão "Enviar" ----
        hBtnSend = CreateWindowW(
            L"BUTTON", L"Enviar",
            WS_CHILD | WS_VISIBLE,
            280, 105, 100, 25,
            hwnd, (HMENU)ID_BTN_SEND,
            nullptr, nullptr);

        // ---- Janela "terminal" (área de log) ----
        // ES_MULTILINE   -> múltiplas linhas
        // ES_AUTOVSCROLL -> rolagem automática conforme texto cresce
        // ES_READONLY    -> usuário não edita (só leitura)
        // WS_VSCROLL     -> mostra barra de rolagem vertical
        hTerminal = CreateWindowW(
            L"EDIT", nullptr,
            WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL |
            ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY,
            10, 160, 370, 260,
            hwnd, (HMENU)ID_TERMINAL,
            nullptr, nullptr);

        // ---- Preenche os combos com dados iniciais ----
        // - Lista portas COM detectadas (ex.: "USB-SERIAL (COM6)").
        // - Lista baud rates comuns e seleciona 115200 por padrão.
        ListComPorts(hComboComPort);
        PopulateBaudRates(hComboBaudRate);

        // Fim do tratamento da criação: retornamos pois já lidamos com WM_CREATE.
        break;

        // ------------------------------------------------------------------------
        // Mensagens de COMANDO de controles (botões, menus, combos, edits, etc).
        // wParam (LOWORD) = ID do controle que gerou o comando.
        // ------------------------------------------------------------------------
    case WM_COMMAND:
        switch (LOWORD(wParam)) {

        case ID_COMBOBOX_COMPORT:

            // printf("\r\nAlgo aconteceu na combobox de Portas COM");
            // ...e foi o evento "abrindo o dropdown"
            if (HIWORD(wParam) == CBN_DROPDOWN) {
                printf("User Pediu para listar portas\r\n");
                RefreshComPortsAndKeepSelection(hComboComPort);
            }
            break;

            // ---- Clique no botão "Conectar"/"Desconectar" ----
        case ID_BTN_CONNECT: {
            if (isConnected) {
                // Já está conectado → então vamos desconectar.
                // Fecha a thread de RX, cancela I/O e libera handle.
                CloseSerialPort();

                // Atualiza rótulo do botão e loga no terminal.
                SetWindowTextW(hBtnConnect, L"Conectar");
                AppendToTerminal(L"[INFO] Porta desconectada\r\n");
            }
            else {
                // Ainda não está conectado → vamos tentar abrir a porta escolhida.

                // 1) Lê o "friendly name" do combo de portas (ex.: "USB-Serial (COM6)").
                wchar_t portName[100] = {};
                GetWindowTextW(hComboComPort, portName, 100);

                // 2) Lê o baud selecionado (string -> número).
                wchar_t baudStr[20] = {};
                GetWindowTextW(hComboBaudRate, baudStr, 20);
                DWORD baud = _wtoi(baudStr);

                // 3) Extrai apenas "COMx" do friendly name.
                //    Se o nome tiver "(COM6)", capturamos o conteúdo entre parênteses.
                std::wstring portStr(portName);
                size_t b = portStr.find(L"(COM");                 // posição do "(COM"
                size_t e = portStr.find(L")", (b == std::wstring::npos ? 0 : b)); // fecha ")"

                // Se encontrou "(COMx)", extrai "COMx"; caso contrário, usa a string inteira.
                std::wstring portOnly =
                    (b != std::wstring::npos && e != std::wstring::npos && e > b)
                    ? portStr.substr(b + 1, e - (b + 1))  // remove parênteses
                    : portStr;

                // 4) Tenta abrir a porta e configurar (baud, 8N1, DTR/RTS, timeouts).
                if (OpenSerialPort(portOnly, baud)) {
                    // Sucesso: atualiza botão para "Desconectar" e loga.
                    SetWindowTextW(hBtnConnect, L"Desconectar");
                    AppendToTerminal(L"[OK] Conectado em " + portOnly + L" @ " + std::to_wstring(baud) + L"\r\n");
                }
                else {
                    // Falha ao abrir/configurar.
                    AppendToTerminal(L"[ERRO] Falha ao conectar\r\n");
                }
            }
        } break;

            // ---- Clique no botão "Enviar" ----
        case ID_BTN_SEND:
            // Encaminha para a rotina de envio: pega texto da caixa escolhida,
            // converte para UTF-8 e manda via WriteFile.
            SendSelectedMessage();
            break;
        }
        break;

        // ------------------------------------------------------------------------
        // A janela está sendo destruída (usuário fechou, Alt+F4, etc.)
        // Limpeza geral: encerre a serial, pare threads, avise ao sistema para sair.
        // ------------------------------------------------------------------------
    case WM_DESTROY:
        CloseSerialPort();       // garante que nada fica pendurado
        PostQuitMessage(0);      // pede para o loop principal encerrar
        break;
    }

    // Para mensagens que não tratamos, delegamos ao DefWindowProc
    // (o Windows fornece comportamento padrão, como mover/redimensionar, etc.).
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

