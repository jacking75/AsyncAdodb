# AsyncAdodb
1 header file로 만든 C++ ADO 라이브러리이다(include 디렉토리에 있는 "AdoManager.h").  
이 라이브러리는 [네이버의 온라인 서버 제작자 모임](https://cafe.naver.com/ongameserver/3412) 의 멤버인 김영찬님이 공개한 라이브러리를 수정한 버전이다.  

       
  
## [펌] ADO 연결 문자열
출처는 알 수 없음.  
  
ADO를 사용하려면 ConnectionString이 필요한다. 이걸 실제로 잘 만들어 내는 방법에 대한 이야기이다.  
실제로 접속 문자열만 완성 되면 모두 1가지 DB인 것처럼 사용할 수 있다.  
  
로컬의 FILE DSN은 물론 SQL Server, Oracle, Text 파일, Excel 파일, Access나 dBASE 파일까지 연결하는 방법에 대해서 언급하겠다.  
일단 연결이 되고나면 심지어 Excel 파일에 대해서도 SELECT와 같은 SQL문이 먹혀 들어가는 아주 강력한 것이 바로 ADO 이다.  

      
### SQL SERVER에 대한 접속

```
"Provider=sqloledb;Data Source=155.230.29.10;User Id=sa;Password=pass"
```  
ConnectionString을 위와 같이 하면 된다. 실제로 Id랑 소스 , 패스워드는 맞게 고쳐 줘야한다.  

### Txt 파일 , dBase 파일 , Access 파일 , Excel 파일 
이건 OLE DB를 통해 다 함께 지워이 된다. 먼저 다음과 같은 함수를 알아두자  
```
#include "odbcinst.h"
CString GetDriverStringX(CString drivername) 
{
    char szBuf[2001];
    WORD cbBufMax = 2000;
    WORD cbBufOut;
    char *pszBuf = szBuf;
    CString sDriver; // return value
    CString tDriver = drivername ;

if(!SQLGetInstalledDrivers(szBuf,cbBufMax,& cbBufOut))
        return "";
    
    do
    {   
        if( strstr( pszBuf, tDriver ) != 0 )
        {
            // Found !
            sDriver = CString( pszBuf );
            break;// NOT BREAK
        }
        pszBuf = strchr( pszBuf, '\0' ) + 1;
        
    }
    while( pszBuf[1] != '\0' );
  
    return sDriver;
}
``` 
  
driver 이름을 적으면 실제로 시스템에 설치된 드라이버의 스트링을 얻어온다.  
파라메터에 "Access", "dBase" , "Text", "Excel" 이라고 넣어주면 각각의 드라이버 이름을 얻어올 수 있다.  
  
그 다음은 이렇게 한다.  
```
CString ConnectionString;
CString driverStr = GetDriverStringX("Excel");
CString pathStr="C:\Example.xls";
ConnectionString.Format( "DRIVER={%s};DBQ=%s;",driverStr,pathStr);
```  
  
지금까지 해서 5가지를 알아봤다 . 다음은 FileDSN 이다.  
  
### System File DSN 
제어판의 관리 도구를 통해서 설정한 DSN의 값을 컴퓨터로 부터 얻는 방법이다.  
```
void OnDsn(void)
{
	/* HKEY_CURRENT_USER\SOFTWARE\ODBC\ODBC.INI */
	LONG lResult;
	HKEY hKey;
	CHAR achClass[MAX_PATH] = "";
	DWORD cchClassName=MAX_PATH,cSubKeys,cbMaxSubKey,cchMaxClass;
	DWORD cValues,cchMaxValue,cbMaxValueData,cbSecurityDescriptor;
	FILETIME ftLastWriteTime,FileTime;

	RegOpenKeyEx(HKEY_CURRENT_USER,"SOFTWARE\\ODBC\\ODBC.INI",0,KEY_READ,&hKey);
	RegQueryInfoKey(hKey,achClass,&cchClassName,NULL,&cSubKeys,&cbMaxSubKey,&cchMaxClass,&cValues, &cchMaxValue,&cbMaxValueData,&cbSecurityDescriptor,&ftLastWriteTime);

	DWORD dwBuffer;
	unsigned char szSubKeyName[MAX_PATH];
	for (DWORD nCount=0 ; nCount<cSubKeys ; nCount++)
	{
		dwBuffer = MAX_PATH;
		lResult = RegEnumKeyEx(hKey,nCount,(LPTSTR)
		szSubKeyName,&dwBuffer,NULL,NULL,0,&FileTime);
		
		if( strcmp((char*)szSubKeyName,"ODBC Data Sources") && strcmp((char*)szSubKeyName,"ODBC File DSN") )
		{
			AfxMessageBox( (LPCTSTR)szSubKeyName );
		}
	}

	RegCloseKey(hKey);
}
```  
  
중간에 AfxMessageBox로 출력하는 값이 그게 바로 DSN 이다.  
실제로 스트링을 만들려면 다음과 같이 한다.  
```
CString ConnectionString ="DSN=dsnString;";
```  
  
FileDSN을 얻는 방법이 1가지 더 있다.  
바로 CDatabase의 함수를 호출하는 방법이다.  
```
#include "afxdb.h"
void DSNDialog(void);
{
    CDatabase db;
    db.OpenEx(NULL);
        // 제어판에서 볼수 있는 DSN 설정 다이얼로그가 뜹니다

    CString driverStr = db.GetConnect ();
         // 다이얼로그에서 고른 ODBC Driver에 대해서 접속문자열을 가져 옵니다.

    // "ODBC;"를 삭제한다.
         // 실제로 ADO로 연결하기 위해서 앞의 "ODBC;"를 삭제해줍니다
    driverStr.Replace ("ODBC;","");

	ConnectionString= driverStr;
}
```  
  
지금까지 ADO의 접속 문자열을 만드는 방법을 실제의 코드로 알아봤다.  
텍스트 파일 , Access 파일, dBASE 파일은 물론, SQL Server, 등등.  
OLE DB와 ODBC가 지원하는 모든 데이터베이스에 실제로 접근할 수 있다.  
  
마지막으로 1가지만 더 이야기 하겠다
```
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\db1.mdb;Persist Security Info=False;"
```  
ADO의 접속 문자열에 Provider란게 있는데, 여기 뒤에 뭐가 들어갈지 처음 하면 정말 찾기 어렵다. 아무리 찾아도 안 나온다.  
맨땅에 헤딩하는 것을 방지하기 위해서 Provider이름을 제공하는 웹페이지를 알려 주겠다.  
http://msdn.microsoft.com/library/default.asp?url=/library/en-us/ado270/htm/mdrefodbcprovspec.asp  
다른데가 아니고 바로 MSDN ADO 홈페이지이다.  
접속 문자열에 대한 더 정확한 설명이 나올 것이다.    
  
  
## 참고 자료
- [MySQL OLEDB, ODBC(DSN 등록 안함) 연결하기](https://m.blog.naver.com/kilsu1024/110162885226 )
- [ado ms-sql, mysql](http://sakbals.tistory.com/entry/%EC%B4%88%EA%B8%89-%EA%B0%84%EB%8B%A8%ED%95%98%EA%B2%8C-ADO-oledb-%EC%82%AC%EC%9A%A9%ED%95%98%EA%B8%B0 )
- [C++로 Mysql을 ADO로 연결하려고 하는데 출력시 한글이 깨집니다](https://kldp.org/node/142606 )
- [msado15.dll을 이용해서 MFC에서 ADO로 데이테베이스 억세스하기](http://wwwi.tistory.com/80 )
- [(일어) Connector/ODBC アプリケーション別情報](http://download.nust.na/pub6/mysql/doc/refman/5.1/ja/myodbc-usagenotes-apptips.html )
- [(일어) 접속 문자열 리스트](http://kojama.la.coocan.jp/works/rdbms/conn/connstr.html )
- [Tutorial: Moving from MySQL to ADODB](http://web.unife.it/lib/adodb/docs/tute.htm )
- [ADO Connection Strings](https://www.codeproject.com/Articles/2304/ADO-Connection-Strings )




