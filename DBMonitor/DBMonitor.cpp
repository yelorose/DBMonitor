// DBMonitor.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "DBMonitor.h"
#include "VisitData.h"
 #include <WinInet.h>
 #pragma comment(lib,"WinINet.lib")


#ifdef _DEBUG
#define new DEBUG_NEW
#endif



// The one and only application object

CWinApp theApp;

using namespace std;

BOOL GetURL(LPCSTR pszHostAddress,UINT nHostPort, LPCSTR pSendBuf,char *pRevBuf,int nRevBufLen)
{
	INT		nRet = -1;
	HINTERNET	hInetOpen = NULL;
	HINTERNET	hInetConn = NULL;
	HINTERNET	hInetReq = NULL;
	BOOL	bRet = FALSE;
	//CHAR	szBuf[1024] = {0};
	DWORD	dwReadLen = 0L;
	int nType=0;
	
	//初始化INTERNET函数，目前暂不支持异步IO
	hInetOpen = InternetOpen(NULL, INTERNET_OPEN_TYPE_DIRECT, NULL, NULL, 0);
	if (!hInetOpen)
	{
		printf("%s : %d, CreateURL:%s (Type%d, %s : %d) error",__FILE__,__LINE__,pSendBuf,nType,pszHostAddress,nHostPort);
		goto HRUERR;
	}
	
	//建立连接缓冲
	hInetConn = InternetConnect(hInetOpen, pszHostAddress, nHostPort, NULL, NULL, INTERNET_SERVICE_HTTP,
		INTERNET_FLAG_PASSIVE | INTERNET_FLAG_DONT_CACHE | INTERNET_FLAG_NO_CACHE_WRITE | INTERNET_FLAG_PRAGMA_NOCACHE, 0);
	if (!hInetConn)
	{
		printf("%s : %d, ConnecteURL:%s (Type%d, %s : %d) error",__FILE__,__LINE__,pSendBuf,nType,pszHostAddress,nHostPort);
		goto HRUERR;
	}
	
	//请示服务对象
	hInetReq = HttpOpenRequest(hInetConn, "GET", pSendBuf, NULL, NULL, NULL,
		INTERNET_FLAG_RELOAD | INTERNET_FLAG_KEEP_CONNECTION | INTERNET_FLAG_DONT_CACHE, 0);
	if (!hInetReq)
	{
		printf("%s : %d, OpenURL:%s (Type%d, %s : %d) error",__FILE__,__LINE__,pSendBuf,nType,pszHostAddress,nHostPort);
		goto HRUERR;
	}
	
	bRet = HttpSendRequest(hInetReq, NULL, 0, NULL, 0);	
	if(bRet)
	{
		CHAR	szStatusCode[10] = {0};
		DWORD	dwLen = 0L, dwIndex = 0L;
		
		dwLen = sizeof(szStatusCode);
		bRet = HttpQueryInfo(hInetReq, HTTP_QUERY_STATUS_CODE, szStatusCode, &dwLen, &dwIndex);
		
		{
			dwLen = sizeof(szStatusCode);
			dwIndex = 0L;
			bRet = HttpQueryInfo(hInetReq, HTTP_QUERY_CONTENT_LENGTH, szStatusCode, &dwLen, &dwIndex);
			InternetReadFile(hInetReq, pRevBuf, nRevBufLen, &dwLen);
			Sleep(0);
		}
		nRet = 0;
	}
	else
	{
		printf("%s : %d, GetURL:%s (Type%d, %s : %d) error",__FILE__,__LINE__,pSendBuf,nType,pszHostAddress,nHostPort);
		nRet = -2;
	}
	
HRUERR:
	if (hInetReq) InternetCloseHandle(hInetReq);
	if (hInetConn) InternetCloseHandle(hInetConn);
	if (hInetOpen) InternetCloseHandle(hInetOpen);
	
	if (nRet==0) return TRUE;
	return FALSE;
}

// 检查网站是否正常工作
BOOL CheckWeb()
{
	char szBuf[512]={0};
	// http://192.168.0.20/text1.aspx
	GetURL("192.168.0.20",80,"/text1.aspx",szBuf,512);
	printf(szBuf);
	printf("\r\n");
	if (strstr(szBuf, "JYwebWorking!")!=NULL)
	{
		printf("网站工作正常！\r\n");
		return TRUE;
	}
	else
	{
		printf("!!网站不干活了\r\n");
		return FALSE;
	}
}

BOOL CheckDB()
{
	CVisitData m_db;

	BOOL bRet;
	bRet = m_db.Initialize("192.168.0.201", "monitor", "monitor", "QPPlatformDB");
	if (!bRet)
	{
		printf("create com fail!");
		return 1;
	}
	if (!m_db.SearchData("select username from monitor where id=1"))
	{
		printf("exec select fail!");
		return 1;
	}
	while (! m_db.IsRecordSetEOF())
	{
		CString str = m_db.GetFieldValue_String("username");
		if (str.Compare("sysmonitor")==0)
		{
			printf("OK!");
			return 0;
		}
		else
			printf("Fail!");
			return 1;
	}
	
	printf("Fail!");
	return 1;
}

int _tmain(int argc, TCHAR* argv[], TCHAR* envp[])
{
	int nRetCode = 0;

	// initialize MFC and print and error on failure
	if (!AfxWinInit(::GetModuleHandle(NULL), NULL, ::GetCommandLine(), 0))
	{
		// TODO: change error code to suit your needs
		_tprintf(_T("Fatal Error: MFC initialization failed\n"));
		nRetCode = 1;
	}
	else
	{
		// TODO: code your application's behavior here.
	}

	AfxEnableControlContainer();
	CoInitialize(NULL);

	// 检查网站是否正常工作
	CheckWeb();

	getchar();
	return 0;

	//检查数据库是否工作正常
	CheckDB();
	getchar();

	return nRetCode;
}
