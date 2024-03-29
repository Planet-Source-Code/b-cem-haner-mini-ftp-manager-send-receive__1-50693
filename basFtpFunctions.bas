Attribute VB_Name = "IDBAS_WININET"
Option Explicit
'WININET API declarations (provided by WININET.DLL)
'For more information on doing FTP using WININET see:
'http://www.microsoft.com/workshop/prog/sdk/docs/wininet/inetr004.htm

' ------------------------------------------------------------------------
'
'    WININET.TXT -- WININET API Declarations for Visual Basic
'
'              Copyright (C) 1998 Microsoft Corporation
'
'  This file is required for the Visual Basic 6.0 version of the APILoader.
'  This file is backwards compatible with previous releases
'  of the APILoader with the exception that Constants are no longer declared
'  as Global or Public in this file.
'
'  This file contains only the Const, Type,
'  and Declare statements for the WININET APIs.
'
'  You have a royalty-free right to use, modify, reproduce and distribute
'  this file (and/or any modified version) in any way you find useful,
'  provided that you agree that Microsoft has no warranty, obligation or
'  liability for its contents.  Refer to the Microsoft Windows Programmer's
'  Reference for further information.
'
' ------------------------------------------------------------------------

Const MAX_PATH = 260
Const NO_ERROR = 0
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_TEMPORARY = &H100
Const FILE_ATTRIBUTE_COMPRESSED = &H800
Const FILE_ATTRIBUTE_OFFLINE = &H1000


Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Const ERROR_NO_MORE_FILES = 18

Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String, ByRef lpdwCurrentDirectory As Long) As Boolean

' Initializes an application's use of the Win32 Internet functions
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

' User agent constant.
Const scUserAgent = "vb wininet"

' Use registry access settings.
Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Const INTERNET_OPEN_TYPE_DIRECT = 1
Const INTERNET_OPEN_TYPE_PROXY = 3
Const INTERNET_INVALID_PORT_NUMBER = 0

' Opens a HTTP session for a given site.
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean

' Number of the TCP/IP port on the server to connect to.
Const INTERNET_DEFAULT_FTP_PORT = 21
Const INTERNET_DEFAULT_GOPHER_PORT = 70
Const INTERNET_DEFAULT_HTTP_PORT = 80
Const INTERNET_DEFAULT_HTTPS_PORT = 443
Const INTERNET_DEFAULT_SOCKS_PORT = 1080

Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6
Const INTERNET_OPTION_SEND_TIMEOUT = 5

Const INTERNET_OPTION_USERNAME = 28
Const INTERNET_OPTION_PASSWORD = 29
Const INTERNET_OPTION_PROXY_USERNAME = 43
Const INTERNET_OPTION_PROXY_PASSWORD = 44

' Type of service to access.
Const INTERNET_SERVICE_FTP = 1
Const INTERNET_SERVICE_GOPHER = 2
Const INTERNET_SERVICE_HTTP = 3

' Opens an HTTP request handle.
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000

' Sends the specified request to the HTTP server.
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As String, ByVal lOptionalLength As Long) As Integer

' Queries for information about an HTTP request.
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer

' InternetErrorDlg
Private Declare Function InternetErrorDlg Lib "wininet.dll" (ByVal hWnd As Long, ByVal hInternet As Long, ByVal dwError As Long, ByVal dwFlags As Long, ByVal lppvData As Long) As Long

' InternetErrorDlg constants
Const FLAGS_ERROR_UI_FILTER_FOR_ERRORS = &H1
Const FLAGS_ERROR_UI_FLAGS_CHANGE_OPTIONS = &H2
Const FLAGS_ERROR_UI_FLAGS_GENERATE_DATA = &H4
Const FLAGS_ERROR_UI_FLAGS_NO_UI = &H8
Const FLAGS_ERROR_UI_SERIALIZE_DIALOGS = &H10

Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long

' The possible values for the lInfoLevel parameter include:
Const HTTP_QUERY_CONTENT_TYPE = 1
Const HTTP_QUERY_CONTENT_LENGTH = 5
Const HTTP_QUERY_EXPIRES = 10
Const HTTP_QUERY_LAST_MODIFIED = 11
Const HTTP_QUERY_PRAGMA = 17
Const HTTP_QUERY_VERSION = 18
Const HTTP_QUERY_STATUS_CODE = 19
Const HTTP_QUERY_STATUS_TEXT = 20
Const HTTP_QUERY_RAW_HEADERS = 21
Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Const HTTP_QUERY_FORWARDED = 30
Const HTTP_QUERY_SERVER = 37
Const HTTP_QUERY_USER_AGENT = 39
Const HTTP_QUERY_SET_COOKIE = 43
Const HTTP_QUERY_REQUEST_METHOD = 45
Const HTTP_STATUS_DENIED = 401
Const HTTP_STATUS_PROXY_AUTH_REQ = 407

' Add this flag to the about flags to get request header.
Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000
Const HTTP_QUERY_FLAG_NUMBER = &H20000000

' Reads data from a handle opened by the HttpOpenRequest function.
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer

Type INTERNET_BUFFERS
    dwStructSize As Long        ' used for API versioning. Set to sizeof(INTERNET_BUFFERS)
    Next As Long                ' INTERNET_BUFFERS chain of buffers
    lpcszHeader As Long       ' pointer to headers (may be NULL)
    dwHeadersLength As Long     ' length of headers if not NULL
    dwHeadersTotal As Long      ' size of headers if not enough buffer
    lpvBuffer As Long           ' pointer to data buffer (may be NULL)
    dwBufferLength As Long      ' length of data buffer if not NULL
    dwBufferTotal As Long       ' total size of chunk, or content-length if not chunked
    dwOffsetLow As Long         ' used for read-ranges (only used in HttpSendRequest2)
    dwOffsetHigh As Long
End Type

Private Declare Function HttpSendRequestEx Lib "wininet.dll" Alias "HttpSendRequestExA" (ByVal hHttpRequest As Long, lpBuffersIn As INTERNET_BUFFERS, ByVal lpBuffersOut As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function HttpEndRequest Lib "wininet.dll" Alias "HttpEndRequestA" (ByVal hHttpRequest As Long, ByVal lpBuffersOut As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetWriteFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumberOfBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Long, ByVal sFileName As String, ByVal lAccess As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
Private Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByVal lBufferLength As Long) As Integer
Private Declare Function InternetSetOptionStr Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hInternet As Long, ByVal lOption As Long, ByVal sBuffer As String, ByVal lBufferLength As Long) As Integer

' Closes a single Internet handle or a subtree of Internet handles.
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

' Queries an Internet option on the specified handle
Private Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" (ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long) As Integer

' Returns the version number of Wininet.dll.
Const INTERNET_OPTION_VERSION = 40

' Contains the version number of the DLL that contains the Windows Internet
' functions (Wininet.dll). This structure is used when passing the
' INTERNET_OPTION_VERSION flag to the InternetQueryOption function.
Type tWinInetDLLVersion
    lMajorVersion As Long
    lMinorVersion As Long
End Type

' Adds one or more HTTP request headers to the HTTP request handle.
Private Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lModifiers As Long) As Integer

' Flags to modify the semantics of this function. Can be a combination of these values:

' Adds the header only if it does not already exist; otherwise, an error is returned.
Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000

' Adds the header if it does not exist. Used with REPLACE.
Const HTTP_ADDREQ_FLAG_ADD = &H20000000

' Replaces or removes a header. If the header value is empty and the header is found,
' it is removed. If not empty, the header value is replaced
Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000

' Internet Errors
Const INTERNET_ERROR_BASE = 12000

Const ERROR_INTERNET_OUT_OF_HANDLES = (INTERNET_ERROR_BASE + 1)
Const ERROR_INTERNET_TIMEOUT = (INTERNET_ERROR_BASE + 2)
Const ERROR_INTERNET_EXTENDED_ERROR = (INTERNET_ERROR_BASE + 3)
Const ERROR_INTERNET_INTERNAL_ERROR = (INTERNET_ERROR_BASE + 4)
Const ERROR_INTERNET_INVALID_URL = (INTERNET_ERROR_BASE + 5)
Const ERROR_INTERNET_UNRECOGNIZED_SCHEME = (INTERNET_ERROR_BASE + 6)
Const ERROR_INTERNET_NAME_NOT_RESOLVED = (INTERNET_ERROR_BASE + 7)
Const ERROR_INTERNET_PROTOCOL_NOT_FOUND = (INTERNET_ERROR_BASE + 8)
Const ERROR_INTERNET_INVALID_OPTION = (INTERNET_ERROR_BASE + 9)
Const ERROR_INTERNET_BAD_OPTION_LENGTH = (INTERNET_ERROR_BASE + 10)
Const ERROR_INTERNET_OPTION_NOT_SETTABLE = (INTERNET_ERROR_BASE + 11)
Const ERROR_INTERNET_SHUTDOWN = (INTERNET_ERROR_BASE + 12)
Const ERROR_INTERNET_INCORRECT_USER_NAME = (INTERNET_ERROR_BASE + 13)
Const ERROR_INTERNET_INCORRECT_PASSWORD = (INTERNET_ERROR_BASE + 14)
Const ERROR_INTERNET_LOGIN_FAILURE = (INTERNET_ERROR_BASE + 15)
Const ERROR_INTERNET_INVALID_OPERATION = (INTERNET_ERROR_BASE + 16)
Const ERROR_INTERNET_OPERATION_CANCELLED = (INTERNET_ERROR_BASE + 17)
Const ERROR_INTERNET_INCORRECT_HANDLE_TYPE = (INTERNET_ERROR_BASE + 18)
Const ERROR_INTERNET_INCORRECT_HANDLE_STATE = (INTERNET_ERROR_BASE + 19)
Const ERROR_INTERNET_NOT_PROXY_REQUEST = (INTERNET_ERROR_BASE + 20)
Const ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND = (INTERNET_ERROR_BASE + 21)
Const ERROR_INTERNET_BAD_REGISTRY_PARAMETER = (INTERNET_ERROR_BASE + 22)
Const ERROR_INTERNET_NO_DIRECT_ACCESS = (INTERNET_ERROR_BASE + 23)
Const ERROR_INTERNET_NO_CONTEXT = (INTERNET_ERROR_BASE + 24)
Const ERROR_INTERNET_NO_CALLBACK = (INTERNET_ERROR_BASE + 25)
Const ERROR_INTERNET_REQUEST_PENDING = (INTERNET_ERROR_BASE + 26)
Const ERROR_INTERNET_INCORRECT_FORMAT = (INTERNET_ERROR_BASE + 27)
Const ERROR_INTERNET_ITEM_NOT_FOUND = (INTERNET_ERROR_BASE + 28)
Const ERROR_INTERNET_CANNOT_CONNECT = (INTERNET_ERROR_BASE + 29)
Const ERROR_INTERNET_CONNECTION_ABORTED = (INTERNET_ERROR_BASE + 30)
Const ERROR_INTERNET_CONNECTION_RESET = (INTERNET_ERROR_BASE + 31)
Const ERROR_INTERNET_FORCE_RETRY = (INTERNET_ERROR_BASE + 32)
Const ERROR_INTERNET_INVALID_PROXY_REQUEST = (INTERNET_ERROR_BASE + 33)
Const ERROR_INTERNET_NEED_UI = (INTERNET_ERROR_BASE + 34)

Const ERROR_INTERNET_HANDLE_EXISTS = (INTERNET_ERROR_BASE + 36)
Const ERROR_INTERNET_SEC_CERT_DATE_INVALID = (INTERNET_ERROR_BASE + 37)
Const ERROR_INTERNET_SEC_CERT_CN_INVALID = (INTERNET_ERROR_BASE + 38)
Const ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR = (INTERNET_ERROR_BASE + 39)
Const ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR = (INTERNET_ERROR_BASE + 40)
Const ERROR_INTERNET_MIXED_SECURITY = (INTERNET_ERROR_BASE + 41)
Const ERROR_INTERNET_CHG_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 42)
Const ERROR_INTERNET_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 43)
Const ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED = (INTERNET_ERROR_BASE + 44)
Const ERROR_INTERNET_INVALID_CA = (INTERNET_ERROR_BASE + 45)
Const ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP = (INTERNET_ERROR_BASE + 46)
Const ERROR_INTERNET_ASYNC_THREAD_FAILED = (INTERNET_ERROR_BASE + 47)
Const ERROR_INTERNET_REDIRECT_SCHEME_CHANGE = (INTERNET_ERROR_BASE + 48)
Const ERROR_INTERNET_DIALOG_PENDING = (INTERNET_ERROR_BASE + 49)
Const ERROR_INTERNET_RETRY_DIALOG = (INTERNET_ERROR_BASE + 50)
Const ERROR_INTERNET_HTTPS_HTTP_SUBMIT_REDIR = (INTERNET_ERROR_BASE + 52)
Const ERROR_INTERNET_INSERT_CDROM = (INTERNET_ERROR_BASE + 53)

' FTP API errors

Const ERROR_FTP_TRANSFER_IN_PROGRESS = (INTERNET_ERROR_BASE + 110)
Const ERROR_FTP_DROPPED = (INTERNET_ERROR_BASE + 111)
Const ERROR_FTP_NO_PASSIVE_MODE = (INTERNET_ERROR_BASE + 112)

' gopher API errors

Const ERROR_GOPHER_PROTOCOL_ERROR = (INTERNET_ERROR_BASE + 130)
Const ERROR_GOPHER_NOT_FILE = (INTERNET_ERROR_BASE + 131)
Const ERROR_GOPHER_DATA_ERROR = (INTERNET_ERROR_BASE + 132)
Const ERROR_GOPHER_END_OF_DATA = (INTERNET_ERROR_BASE + 133)
Const ERROR_GOPHER_INVALID_LOCATOR = (INTERNET_ERROR_BASE + 134)
Const ERROR_GOPHER_INCORRECT_LOCATOR_TYPE = (INTERNET_ERROR_BASE + 135)
Const ERROR_GOPHER_NOT_GOPHER_PLUS = (INTERNET_ERROR_BASE + 136)
Const ERROR_GOPHER_ATTRIBUTE_NOT_FOUND = (INTERNET_ERROR_BASE + 137)
Const ERROR_GOPHER_UNKNOWN_LOCATOR = (INTERNET_ERROR_BASE + 138)

' HTTP API errors

Const ERROR_HTTP_HEADER_NOT_FOUND = (INTERNET_ERROR_BASE + 150)
Const ERROR_HTTP_DOWNLEVEL_SERVER = (INTERNET_ERROR_BASE + 151)
Const ERROR_HTTP_INVALID_SERVER_RESPONSE = (INTERNET_ERROR_BASE + 152)
Const ERROR_HTTP_INVALID_HEADER = (INTERNET_ERROR_BASE + 153)
Const ERROR_HTTP_INVALID_QUERY_REQUEST = (INTERNET_ERROR_BASE + 154)
Const ERROR_HTTP_HEADER_ALREADY_EXISTS = (INTERNET_ERROR_BASE + 155)
Const ERROR_HTTP_REDIRECT_FAILED = (INTERNET_ERROR_BASE + 156)
Const ERROR_HTTP_NOT_REDIRECTED = (INTERNET_ERROR_BASE + 160)
Const ERROR_HTTP_COOKIE_NEEDS_CONFIRMATION = (INTERNET_ERROR_BASE + 161)
Const ERROR_HTTP_COOKIE_DECLINED = (INTERNET_ERROR_BASE + 162)
Const ERROR_HTTP_REDIRECT_NEEDS_CONFIRMATION = (INTERNET_ERROR_BASE + 168)

' additional Internet API error codes

Const ERROR_INTERNET_SECURITY_CHANNEL_ERROR = (INTERNET_ERROR_BASE + 157)
Const ERROR_INTERNET_UNABLE_TO_CACHE_FILE = (INTERNET_ERROR_BASE + 158)
Const ERROR_INTERNET_TCPIP_NOT_INSTALLED = (INTERNET_ERROR_BASE + 159)
Const ERROR_INTERNET_DISCONNECTED = (INTERNET_ERROR_BASE + 163)
Const ERROR_INTERNET_SERVER_UNREACHABLE = (INTERNET_ERROR_BASE + 164)
Const ERROR_INTERNET_PROXY_SERVER_UNREACHABLE = (INTERNET_ERROR_BASE + 165)

Const ERROR_INTERNET_BAD_AUTO_PROXY_SCRIPT = (INTERNET_ERROR_BASE + 166)
Const ERROR_INTERNET_UNABLE_TO_DOWNLOAD_SCRIPT = (INTERNET_ERROR_BASE + 167)
Const ERROR_INTERNET_SEC_INVALID_CERT = (INTERNET_ERROR_BASE + 169)
Const ERROR_INTERNET_SEC_CERT_REVOKED = (INTERNET_ERROR_BASE + 170)

' InternetAutodial specific errors

Const ERROR_INTERNET_FAILED_DUETOSECURITYCHECK = (INTERNET_ERROR_BASE + 171)

Const INTERNET_ERROR_LAST = ERROR_INTERNET_FAILED_DUETOSECURITYCHECK

'
' flags common to open functions (not InternetOpen()):
'

Const INTERNET_FLAG_RELOAD = &H80000000             ' retrieve the original item

'
' flags for InternetOpenUrl():
'

Const INTERNET_FLAG_RAW_DATA = &H40000000           ' FTP/gopher find: receive the item as raw (structured) data
Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000   ' FTP: use existing InternetConnect handle for server if possible

'
' flags for InternetOpen():
'

Const INTERNET_FLAG_ASYNC = &H10000000              ' this request is asynchronous (where supported)

'
' protocol-specific flags:
'

Const INTERNET_FLAG_PASSIVE = &H8000000             ' used for FTP connections

'
' additional cache flags
'

Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000      ' don't write this item to the cache
Const INTERNET_FLAG_DONT_CACHE = INTERNET_FLAG_NO_CACHE_WRITE
Const INTERNET_FLAG_MAKE_PERSISTENT = &H2000000     ' make this item persistent in cache
Const INTERNET_FLAG_FROM_CACHE = &H1000000          ' use offline semantics
Const INTERNET_FLAG_OFFLINE = INTERNET_FLAG_FROM_CACHE

'
' additional flags
'

Const INTERNET_FLAG_SECURE = &H800000               ' use PCT/SSL if applicable (HTTP)
Const INTERNET_FLAG_KEEP_CONNECTION = &H400000      ' use keep-alive semantics
Const INTERNET_FLAG_NO_AUTO_REDIRECT = &H200000     ' don't handle redirections automatically
Const INTERNET_FLAG_READ_PREFETCH = &H100000        ' do background read prefetch
Const INTERNET_FLAG_NO_COOKIES = &H80000            ' no automatic cookie handling
Const INTERNET_FLAG_NO_AUTH = &H40000               ' no automatic authentication handling
Const INTERNET_FLAG_CACHE_IF_NET_FAIL = &H10000     ' return cache file if net request fails

'
' Security Ignore Flags, Allow HttpOpenRequest to overide
'  Secure Channel (SSL/PCT) failures of the following types.
'

Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP = &H8000       ' ex: https:// to http://
Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS = &H4000      ' ex: http:// to https://
Const INTERNET_FLAG_IGNORE_CERT_DATE_INVALID = &H2000      ' expired X509 Cert.
Const INTERNET_FLAG_IGNORE_CERT_CN_INVALID = &H1000        ' bad common name in X509 Cert.

'
' more caching flags
'

Const INTERNET_FLAG_RESYNCHRONIZE = &H800           ' asking wininet to update an item if it is newer
Const INTERNET_FLAG_HYPERLINK = &H400               ' asking wininet to do hyperlinking semantic which works right for scripts
Const INTERNET_FLAG_NO_UI = &H200                   ' no cookie popup
Const INTERNET_FLAG_PRAGMA_NOCACHE = &H100          ' asking wininet to add "pragma: no-cache"
Const INTERNET_FLAG_CACHE_ASYNC = &H80              ' ok to perform lazy cache-write
Const INTERNET_FLAG_FORMS_SUBMIT = &H40             ' this is a forms submit
Const INTERNET_FLAG_NEED_FILE = &H10                ' need a file for this request
Const INTERNET_FLAG_MUST_CACHE_REQUEST = INTERNET_FLAG_NEED_FILE

'
' flags for FTP
'

Const INTERNET_FLAG_TRANSFER_ASCII = &H1
Const INTERNET_FLAG_TRANSFER_BINARY = &H2

'
' flags field masks
'

Const SECURITY_INTERNET_MASK = INTERNET_FLAG_IGNORE_CERT_CN_INVALID Or INTERNET_FLAG_IGNORE_CERT_DATE_INVALID Or INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS Or INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP

Const INTERNET_FLAGS_MASK = INTERNET_FLAG_RELOAD Or INTERNET_FLAG_RAW_DATA Or INTERNET_FLAG_EXISTING_CONNECT Or INTERNET_FLAG_ASYNC Or INTERNET_FLAG_PASSIVE Or INTERNET_FLAG_NO_CACHE_WRITE Or INTERNET_FLAG_MAKE_PERSISTENT Or INTERNET_FLAG_FROM_CACHE Or INTERNET_FLAG_SECURE Or INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_NO_AUTO_REDIRECT Or INTERNET_FLAG_READ_PREFETCH Or INTERNET_FLAG_NO_COOKIES Or INTERNET_FLAG_NO_AUTH Or INTERNET_FLAG_CACHE_IF_NET_FAIL Or SECURITY_INTERNET_MASK Or INTERNET_FLAG_RESYNCHRONIZE Or INTERNET_FLAG_HYPERLINK Or INTERNET_FLAG_NO_UI Or INTERNET_FLAG_PRAGMA_NOCACHE Or INTERNET_FLAG_CACHE_ASYNC Or INTERNET_FLAG_FORMS_SUBMIT Or INTERNET_FLAG_NEED_FILE Or INTERNET_FLAG_TRANSFER_BINARY Or INTERNET_FLAG_TRANSFER_ASCII

Const INTERNET_ERROR_MASK_INSERT_CDROM = &H1

Const INTERNET_OPTIONS_MASK = (Not INTERNET_FLAGS_MASK)

'
' common per-API flags (new APIs)
'

Const WININET_API_FLAG_ASYNC = &H1                  ' force async operation
Const WININET_API_FLAG_SYNC = &H4                   ' force sync operation
Const WININET_API_FLAG_USE_CONTEXT = &H8            ' use value supplied in dwContext (even if 0)

'
' INTERNET_NO_CALLBACK - if this value is presented as the dwContext parameter
' then no call-backs will be made for that API
'

Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SystemTime) As Long
Public Type SystemTime
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Const INTERNET_NO_CALLBACK = 0

Public Type tFtpFile
    fileName As String
    LastWriteTime As Date
    FileSize As Long
    isDirectory As Boolean
End Type

Public Enum eFtpTransferType
    FTP_TRANSFER_TYPE_ASCII = &H1
    FTP_TRANSFER_TYPE_BINARY = &H0
End Enum

Public fileinfo As tFtpFile



Public Function GetFtpFile(HostName As String, UserName As String, UserPassword As String, HostFilename As String, localFilename As String, Optional DeleteHost As Boolean = False, Optional TransferMode As eFtpTransferType = FTP_TRANSFER_TYPE_ASCII) As Long
Dim hInternet As Long
Dim hFTP As Long
Dim DeleteSuccess As Boolean
    hInternet = InternetOpen(App.Title, 0, "", "", 0)
    hFTP = InternetConnect(hInternet, HostName, INTERNET_DEFAULT_FTP_PORT, UserName, UserPassword, INTERNET_SERVICE_FTP, 0, 0)
    GetFtpFile = FtpGetFile(hFTP, HostFilename, localFilename, False, 0, INTERNET_FLAG_DONT_CACHE + TransferMode, 0)
    If DeleteHost = True Then
        DeleteSuccess = FtpDeleteFile(hFTP, HostFilename)
    End If
    Call InternetCloseHandle(hInternet)
End Function
Public Function PutFtpFile(HostName As String, UserName As String, UserPassword As String, HostFilename As String, localFilename As String, Optional TransferMode As eFtpTransferType = FTP_TRANSFER_TYPE_ASCII) As Long
Dim hInternet As Long
Dim hFTP As Long
    hInternet = InternetOpen(App.Title, 0, "", "", 0)
    hFTP = InternetConnect(hInternet, HostName, INTERNET_DEFAULT_FTP_PORT, UserName, UserPassword, INTERNET_SERVICE_FTP, 0, 0)
    PutFtpFile = FtpPutFile(hFTP, localFilename, HostFilename, INTERNET_FLAG_DONT_CACHE + TransferMode, 0)
    Call InternetCloseHandle(hInternet)
End Function
Public Function DeleteFtpFile(HostName As String, UserName As String, UserPassword As String, HostFilename As String) As Boolean
Dim hInternet As Long
Dim hFTP As Long
    hInternet = InternetOpen(App.Title, 0, "", "", 0)
    hFTP = InternetConnect(hInternet, HostName, INTERNET_DEFAULT_FTP_PORT, UserName, UserPassword, INTERNET_SERVICE_FTP, 0, 0)
    DeleteFtpFile = FtpDeleteFile(hFTP, HostFilename)
    Call InternetCloseHandle(hInternet)
End Function
Public Function TouchFtpFile(HostName As String, UserName As String, UserPassword As String, HostFilename As String) As Boolean
Dim localFilename As String
    localFilename = "C:\Touchfile.txt"
    Open localFilename For Output As #1
        Print #1, ""
    Close #1
    If PutFtpFile(HostName, UserName, UserPassword, HostFilename, localFilename, FTP_TRANSFER_TYPE_ASCII) = 1 Then
        TouchFtpFile = True
    Else
        TouchFtpFile = False
    End If
    Kill localFilename
End Function
Public Function GetFtpDirectory(ByVal HostName As String, ByVal UserName As String, ByVal password As String, ByVal directory As String, ByRef fileinfo() As tFtpFile, ByRef fileCount As Integer, Optional searchPattern As String = "*") As Boolean
Dim hInternet As Long
Dim hFTP As Long
Dim hFind As Long
Dim findfile As WIN32_FIND_DATA
Dim flags As Long
Dim content As Long
Dim Count As Integer
Dim ret As Long
Dim startingPoint As Integer
Dim startingPoint2 As Integer
Dim LocalFileTime As SystemTime

    'ReDim fileinfo(1 To 1024)
    
    hInternet = InternetOpen(App.Title, 0, "", "", 0)
    hFTP = InternetConnect(hInternet, HostName, INTERNET_DEFAULT_FTP_PORT, UserName, password, INTERNET_SERVICE_FTP, 0, 0)
    Call FtpSetCurrentDirectory(hFTP, directory)
    hFind = FtpFindFirstFile(hFTP, searchPattern, findfile, flags, content)
    If hFind = 0 Then
        'ReDim fileinfo(0 To 0)
        GetFtpDirectory = False
        fileCount = 0
        Exit Function
    End If
    
    Count = 1
    fileinfo(Count).fileName = Trim(Mid(findfile.cFileName, 1, InStr(1, findfile.cFileName, Chr(0), vbTextCompare) - 1))
    If findfile.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY Then
        fileinfo(Count).isDirectory = True
    Else
        fileinfo(Count).isDirectory = False
    End If
    fileinfo(Count).FileSize = findfile.nFileSizeLow
        Call FileTimeToLocalFileTime(findfile.ftLastWriteTime, findfile.ftLastWriteTime)
        If Not FileTimeToSystemTime(findfile.ftLastWriteTime, LocalFileTime) Then
            fileinfo(Count).LastWriteTime = Format(LocalFileTime.wMonth, "00") & "/" & Format(LocalFileTime.wDay, "00") & "/" & Format(LocalFileTime.wYear, "0000") & " " & Format(LocalFileTime.wHour, "00") & ":" & Format(LocalFileTime.wMinute, "00")
        End If
    ret = 1
    Do While ret <> 0
        ret = InternetFindNextFile(hFind, findfile)
        If ret <> 0 Then
            Count = Count + 1
            fileinfo(Count).fileName = Trim(Mid(findfile.cFileName, 1, InStr(1, findfile.cFileName, Chr(0), vbTextCompare) - 1))
            If findfile.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY Then
                fileinfo(Count).isDirectory = True
            Else
                fileinfo(Count).isDirectory = False
            End If
            fileinfo(Count).FileSize = findfile.nFileSizeLow
            Call FileTimeToLocalFileTime(findfile.ftLastWriteTime, findfile.ftLastWriteTime)
            If Not FileTimeToSystemTime(findfile.ftLastWriteTime, LocalFileTime) Then
                fileinfo(Count).LastWriteTime = Format(LocalFileTime.wMonth, "00") & "/" & Format(LocalFileTime.wDay, "00") & "/" & Format(LocalFileTime.wYear, "0000") & " " & Format(LocalFileTime.wHour, "00") & ":" & Format(LocalFileTime.wMinute, "00")
            End If
        End If
    Loop
    
    fileCount = Count
    Call InternetCloseHandle(hInternet)
    GetFtpDirectory = True
End Function
